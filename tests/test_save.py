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
            for book in books: book.close()
