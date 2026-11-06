"""Exercise the affected control flow without Windows or provider dependencies."""
import ast
import logging
import os
from pathlib import Path
import re
import shutil
import tempfile
import time
import unittest
from types import SimpleNamespace
from unittest.mock import Mock

SOURCE = Path(__file__).resolve().parents[1] / 'organize_sharepoint_files.py'
TREE = ast.parse(SOURCE.read_text())
SERVICE = next(node for node in TREE.body if isinstance(node, ast.ClassDef))
METHODS = {node.name: node for node in SERVICE.body if isinstance(node, ast.FunctionDef)}


def load_nodes(nodes, namespace):
    exec(compile(ast.fix_missing_locations(ast.Module(body=nodes, type_ignores=[])),
                 str(SOURCE), 'exec'), namespace)
    return namespace


class RegressionTests(unittest.TestCase):
    def test_download_isolation_newest_selection_and_cleanup(self):
        with tempfile.TemporaryDirectory() as root:
            root = Path(root)
            unrelated = root / 'unrelated.pdf'
            unrelated.write_text('keep')
            service = SimpleNamespace(download_dir=root)
            driver = Mock()
            destination = None

            def configure(command, params):
                nonlocal destination
                self.assertEqual(command, 'Browser.setDownloadBehavior')
                destination = Path(params['downloadPath'])
                self.assertEqual(destination.parent, root)
                for name, mtime in [('old.pdf', 10), ('new.pdf', 20)]:
                    path = destination / name
                    path.write_text(name)
                    os.utime(path, (mtime, mtime))
                # A newer unrelated file must never be selected or deleted.
                (root / 'concurrent.pdf').write_text('keep too')
                (destination / 'pending.crdownload').write_text('partial')

            driver.execute_cdp_cmd.side_effect = configure
            sleeps = []

            def finish_download(_):
                sleeps.append(True)
                (destination / 'pending.crdownload').unlink()

            namespace = dict(Path=Path, tempfile=tempfile, shutil=shutil,
                             time=SimpleNamespace(monotonic=time.monotonic, sleep=finish_download),
                             WebDriverWait=lambda *_: SimpleNamespace(until=lambda _: Mock()),
                             EC=Mock(), By=Mock(), ActionChains=Mock())
            load_nodes([METHODS['download_sharepoint_file']], namespace)
            downloaded = namespace['download_sharepoint_file'](
                service, driver, Mock(), {'title': 'matter', '_elements': {'title': Mock()}})
            self.assertEqual(downloaded, destination / 'new.pdf')
            self.assertTrue(sleeps)
            # Execute the caller's actual finally block (also used on scan errors).
            cleanup = next(node.finalbody for node in ast.walk(TREE)
                           if isinstance(node, ast.Try) and node.finalbody
                           and 'shutil.rmtree(downloaded_path.parent)' in ast.unparse(node.finalbody))
            load_nodes(cleanup, dict(downloaded_path=downloaded, shutil=shutil,
                                     self=SimpleNamespace(logging_service=Mock()), logging=logging))
            self.assertFalse(destination.exists())
            self.assertEqual(unrelated.read_text(), 'keep')
            self.assertTrue((root / 'concurrent.pdf').exists())

    def test_download_setup_failure_cleans_only_owned_directory(self):
        with tempfile.TemporaryDirectory() as root:
            root = Path(root)
            (root / 'keep.pdf').write_text('keep')
            driver = Mock()
            driver.execute_cdp_cmd.side_effect = RuntimeError('CDP failed')
            namespace = dict(Path=Path, tempfile=tempfile, shutil=shutil)
            load_nodes([METHODS['download_sharepoint_file']], namespace)
            with self.assertRaisesRegex(RuntimeError, 'CDP failed'):
                namespace['download_sharepoint_file'](SimpleNamespace(download_dir=root), driver, None, {})
            self.assertEqual([p.name for p in root.iterdir()], ['keep.pdf'])

    def verify(self, title, rename_result):
        class NoAlert(Exception):
            pass
        folder = {'title': title, 'format': 'PDF', 'author': 'Writer'}
        service = SimpleNamespace(get_folder_by_title=Mock(return_value=folder),
                                  logging_service=Mock(), rename_sharepoint_folder=Mock(return_value=rename_result))
        driver = Mock()
        type(driver.switch_to).alert = property(lambda _: (_ for _ in ()).throw(NoAlert()))
        driver.execute_script.side_effect = lambda script: 'reload' if 'performance' in script else 'complete'
        namespace = dict(re=re, logging=logging, NoAlertPresentException=NoAlert,
                         UnexpectedAlertPresentException=type('Alert', (Exception,), {}),
                         StaleElementReferenceException=type('Stale', (Exception,), {}),
                         TimeoutException=TimeoutError, By=Mock(), EC=Mock(), ActionChains=Mock())
        load_nodes([METHODS['verify_folder_catalog_update']], namespace)
        result = namespace['verify_folder_catalog_update'](
            service, driver, SimpleNamespace(until=lambda predicate: predicate(driver)),
            'Original', '9781234567890', ['pdf'], {'author': 'Writer'})
        return result, service, folder

    def test_rejected_rename_is_failure(self):
        result, service, _ = self.verify('Original', False)
        self.assertIsNone(result)
        service.rename_sharepoint_folder.assert_called_once()

    def test_verified_folder_is_returned(self):
        result, service, folder = self.verify('9781234567890', True)
        self.assertIs(result, folder)
        service.rename_sharepoint_folder.assert_not_called()

    def test_failed_verification_clears_context_before_folder_update(self):
        block = next(node for node in ast.walk(METHODS['update_folder_from_catalog'])
                     if isinstance(node, ast.Try) and 'verified_folder = ' in ast.unparse(node.body))
        wrapper = ast.FunctionDef(name='check', args=ast.arguments(posonlyargs=[], args=[],
                                  kwonlyargs=[], kw_defaults=[], defaults=[]), body=[block], decorator_list=[])
        for result in (None, {'title': '9781234567890'}):
            with self.subTest(result=result):
                service = SimpleNamespace(verify_folder_catalog_update=Mock(return_value=result),
                                          _catalog_log_context={'isbn': 'test'})
                folder = Mock()
                namespace = dict(self=service, driver=None, wait=None, folder=folder,
                                 original_folder_title='Original', isbn='9781234567890',
                                 formats_to_apply=[], expected_fields={}, TimeoutException=TimeoutError)
                load_nodes([wrapper], namespace)
                returned = namespace['check']()
                if result is None:
                    self.assertIs(returned, False)
                    self.assertIsNone(service._catalog_log_context)
                    folder.update.assert_not_called()
                else:
                    folder.update.assert_called_once_with(result)

    def test_failed_update_excludes_current_casefolded_title(self):
        block = next(node for node in ast.walk(TREE) if isinstance(node, ast.If)
                     and ast.unparse(node.test).startswith('not self.update_folder_from_catalog('))
        for result in (False, True):
            processed = {'original'}
            service = SimpleNamespace(update_folder_from_catalog=Mock(return_value=result))
            load_nodes([block], dict(self=service, driver=None, wait=None, result={},
                                    parent_folder={'title': 'Straße'}, processed_titles=processed))
            self.assertEqual(processed, {'original', 'strasse'} if not result else {'original'})


if __name__ == '__main__':
    unittest.main()
