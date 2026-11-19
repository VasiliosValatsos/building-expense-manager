import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from openpyxl import Workbook
import main

class SaveTests(unittest.TestCase):
    def test_second_replace_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            paths=[Path(tmp)/f'{i}.xlsx' for i in range(2)]
            books=[Workbook(),Workbook()]
            for book in books:
                self.addCleanup(book.close)
            for book,path in zip(books,paths): book.save(path)
            before=[p.read_bytes() for p in paths]
            for book in books: book.active['A1']='unsaved change'
            original=os.replace
            calls=0
            def replace(src,dst):
                nonlocal calls
                calls+=1
                if calls==2: raise OSError('simulated failure')
                return original(src,dst)
            with patch.object(main.os,'replace',side_effect=replace):
                with self.assertRaises(OSError):
                    main.save_workbooks_safely(books[0],paths[0],books[1],paths[1])
            self.assertEqual(before,[p.read_bytes() for p in paths])

    def fresh_pair(self, directory):
        paths = [Path(directory) / name for name in ("common.xlsx", "ledger.xlsx")]
        books = [Workbook(), Workbook()]
        for book, path in zip(books, paths):
            self.addCleanup(book.close)
            book.active.title = "December 2025"
            book.save(path)
            book.active["A1"] = "New"
        return paths, books, tuple(main.workbook_revision(path) for path in paths)

    def test_error_staging_second_workbook_keeps_originals(self):
        with tempfile.TemporaryDirectory() as directory:
            paths, (common, ledger), revisions = self.fresh_pair(directory)
            with patch.object(ledger, "save", side_effect=PermissionError("cannot stage ledger")):
                with self.assertRaises(PermissionError):
                    main.save_workbooks_safely(common, paths[0], ledger, paths[1], revisions=revisions)
            self.assertEqual(tuple(main.workbook_revision(path) for path in paths), revisions)
            self.assertFalse(list(Path(directory).glob(".bem-*")))

    def test_failed_rollback_keeps_recovery_copy(self):
        with tempfile.TemporaryDirectory() as directory:
            paths, (common, ledger), revisions = self.fresh_pair(directory)
            replace = os.replace
            failures = []

            def fail_commit_and_restore(source, target):
                if Path(target).resolve() == paths[1].resolve():
                    failures.append("second workbook")
                    raise PermissionError("locked")
                if Path(source).name.startswith(".bem-rollback-"):
                    self.assertEqual(Path(target).resolve(), paths[0].resolve())
                    failures.append("rollback")
                    raise PermissionError("locked")
                return replace(source, target)

            with patch.object(main.os, "replace", side_effect=fail_commit_and_restore):
                with self.assertRaisesRegex(RuntimeError, "restore these backups") as error:
                    main.save_workbooks_safely(common, paths[0], ledger, paths[1], revisions=revisions)
            self.assertEqual(failures, ["second workbook", "rollback"])
            backups = list(Path(directory).glob(".bem-rollback-*.xlsx"))
            self.assertEqual(len(backups), 1)
            self.assertIn(str(backups[0].resolve()), str(error.exception))
            self.assertEqual(main.workbook_revision(backups[0]), revisions[0])
            self.assertEqual(main.workbook_revision(paths[1]), revisions[1])

    def test_external_changes_and_mismatched_months(self):
        with tempfile.TemporaryDirectory() as directory:
            paths, (common, ledger), revisions = self.fresh_pair(directory)
            paths[1].write_bytes(b"External update")
            with self.assertRaisesRegex(ValueError, "changed on disk"):
                main.save_workbooks_safely(common, paths[0], ledger, paths[1], revisions=revisions)
            self.assertEqual(paths[1].read_bytes(), b"External update")
            self.assertEqual(main.workbook_revision(paths[0]), revisions[0])
            ledger.active.title = "January 2026"
            with self.assertRaisesRegex(ValueError, "different month"):
                main.save_workbooks_safely(common, paths[0], ledger, paths[1])

    def test_another_instance_holding_lock_cannot_be_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            paths, (common, ledger), revisions = self.fresh_pair(directory)
            with main.workbook_locks(paths):
                with self.assertRaisesRegex(ValueError, "is being updated"):
                    main.save_workbooks_safely(common, paths[0], ledger, paths[1], revisions=revisions)
            self.assertEqual(tuple(main.workbook_revision(path) for path in paths), revisions)
