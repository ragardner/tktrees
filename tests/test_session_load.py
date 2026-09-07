# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import ast
import json
import os
import re
import tempfile
import unittest
from pathlib import Path

from src.session import ListOps, Session, fail, ok
from tests.fixtures import ANIMALS_ID_COL, ANIMALS_PARENT_COLS, ANIMALS_TABLE

REPO_ROOT = Path(__file__).resolve().parents[1]
_DATE = re.compile(r"^\d{4}/\d{2}/\d{2}$")
_REBIND = re.compile(r"sheet\.MT\.data\s*=")
_UNPACK_REBIND = re.compile(r"sheet\.MT\.data\s*,[^=\n]*=")
_EDITOR_FILES = (
    REPO_ROOT / "src" / "tree_editor.py",
    REPO_ROOT / "src" / "app.py",
    REPO_ROOT / "src" / "widgets.py",
)


class TestOkFail(unittest.TestCase):
    def test_ok_empty_result(self):
        out = ok()
        self.assertEqual(
            out,
            {"ok": True, "error": None, "warnings": [], "result": {}},
        )

    def test_ok_with_result_and_warnings(self):
        out = ok(result={"ids": 3}, warnings=[" - note"])
        self.assertTrue(out["ok"])
        self.assertIsNone(out["error"])
        self.assertEqual(out["result"], {"ids": 3})
        self.assertEqual(out["warnings"], [" - note"])

    def test_fail_code_and_extra(self):
        out = fail("need_columns", "need id and parents", id="Lion")
        self.assertFalse(out["ok"])
        self.assertEqual(out["result"], {})
        self.assertEqual(out["error"]["code"], "need_columns")
        self.assertEqual(out["error"]["message"], "need id and parents")
        self.assertEqual(out["error"]["id"], "Lion")
        self.assertEqual(out["warnings"], [])


class TestSessionIdentity(unittest.TestCase):
    def test_initial_state(self):
        s = Session()
        self.assertEqual(s.data, [])
        self.assertIsInstance(s.ops, ListOps)
        self.assertIs(s.ops.s, s)
        self.assertEqual(s.nodes, {})
        self.assertEqual(s.rns, {})
        self.assertEqual(s.headers, [])
        self.assertEqual(s.ic, 0)
        self.assertEqual(s.pc, 0)
        self.assertEqual(s.hiers, [])
        self.assertEqual(s.row_len, 0)
        self.assertEqual(s.changelog, [])
        self.assertEqual(s.warnings, [])
        self.assertEqual(s.tagged_ids, set())
        self.assertTrue(s.auto_sort_nodes_bool)
        self.assertFalse(s.allow_spaces_ids_var)
        self.assertFalse(s.unsaved)
        self.assertIsNone(s.filepath)
        self.assertFalse(s.has_document)
        self.assertEqual(s.changelog_at_open, 0)
        self.assertEqual(
            s.status_dict(),
            {
                "file": None,
                "unsaved": False,
                "hierarchy": None,
                "ids": 0,
                "undo": 0,
            },
        )

    def test_listops_mutates_same_list(self):
        s = Session()
        table = id(s.data)
        s.ops.insert_rows([row[:] for row in ANIMALS_TABLE[1:]])
        self.assertEqual(id(s.data), table)
        self.assertEqual(len(s.data), 3)
        self.assertEqual(s.data[2][0], "Lion")

        s.ops.insert_rows([["Tiger", "Cats", "Tiger"]], idx=2)
        self.assertEqual(id(s.data), table)
        self.assertEqual(s.data[2][0], "Tiger")
        self.assertEqual(s.data[3][0], "Lion")

        s.ops.insert_cols(3, n=1)
        self.assertEqual(id(s.data), table)
        self.assertEqual(s.data[0][3], "")
        self.assertEqual(len(s.data[0]), 4)

        s.ops.delete_cols([3])
        self.assertEqual(id(s.data), table)
        self.assertEqual(len(s.data[0]), 3)

        s.ops.delete_rows([2])
        self.assertEqual(id(s.data), table)
        self.assertEqual([row[0] for row in s.data], ["Animals", "Cats", "Lion"])

    def test_rebind_is_a_new_list_until_set_records(self):
        s = Session()
        old = s.data
        s.data = [row[:] for row in ANIMALS_TABLE]
        self.assertIsNot(s.data, old)
        self.assertEqual(s.data[0][0], "ID")

    def test_clear_empties_document_not_new_sheet(self):
        s = Session()
        s.data = [row[:] for row in ANIMALS_TABLE]
        s.nodes = {"lion": object()}
        s.has_document = True
        s.filepath = "/tmp/x.xlsx"
        s.unsaved = True
        s.allow_spaces_ids_var = True
        s.clear()
        self.assertEqual(s.data, [])
        self.assertEqual(s.nodes, {})
        self.assertFalse(s.has_document)
        self.assertIsNone(s.filepath)
        self.assertFalse(s.unsaved)
        self.assertTrue(s.allow_spaces_ids_var)


class TestChangelogHooks(unittest.TestCase):
    def test_append_calls_hooks_and_sets_unsaved(self):
        s = Session()
        unsaved_hits = []
        rows = []
        s.on_increment_unsaved = lambda: unsaved_hits.append(1)
        s.on_changelog_row = lambda: rows.append(1)
        s.changelog_append("Add ID", "Name: Lion Parent: Cats column #3 named: PARENT_1", "", "")
        self.assertEqual(len(s.changelog), 1)
        ch = s.changelog[0]
        self.assertRegex(ch.at, _DATE)
        self.assertEqual(ch.label, "Add ID")
        self.assertEqual(ch.rows[0].what, "Name: Lion Parent: Cats column #3 named: PARENT_1")
        self.assertEqual(ch.rows[0].old, "")
        self.assertEqual(ch.rows[0].new, "")
        self.assertTrue(s.unsaved)
        self.assertEqual(unsaved_hits, [1])
        self.assertEqual(rows, [1])

    def test_append_no_unsaved_still_counts_a_row(self):
        s = Session()
        unsaved_hits = []
        rows = []
        s.on_increment_unsaved = lambda: unsaved_hits.append(1)
        s.on_changelog_row = lambda: rows.append(1)
        s.changelog_append_no_unsaved("Cut and paste ID |", "Lion", "Cats", "Tigers")
        self.assertFalse(s.unsaved)
        self.assertEqual(unsaved_hits, [])
        self.assertEqual(s.changelog, [])
        self.assertIsNotNone(s._pending)

    def test_singular_rewrites_type_without_row_hook(self):
        s = Session()
        unsaved_hits = []
        rows = []
        s.on_increment_unsaved = lambda: unsaved_hits.append(1)
        s.on_changelog_row = lambda: rows.append(1)
        s.changelog_append("Delete ID |", "Lion", "", "")
        unsaved_hits.clear()
        rows.clear()
        s.changelog_singular("Delete ID")
        self.assertEqual(s.changelog[-1].label, "Delete ID")
        self.assertEqual(s.changelog[-1].rows[0].what, "Lion")
        self.assertFalse(s.changelog[-1].has_summary)
        self.assertEqual(unsaved_hits, [1])
        self.assertEqual(rows, [1])


class TestSessionModuleIsolation(unittest.TestCase):
    def test_session_source_does_not_import_gui(self):
        src = (REPO_ROOT / "src" / "session.py").read_text(encoding="utf-8")
        tree = ast.parse(src)
        names = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                names.add(node.module.split(".")[0])
        forbidden = {
            "app",
            "tree_editor",
            "toplevels",
            "widgets",
            "tree_compare",
        }
        self.assertEqual(names & forbidden, set())

    def test_cli_source_does_not_import_gui(self):
        src = (REPO_ROOT / "src" / "cli.py").read_text(encoding="utf-8")
        tree = ast.parse(src)
        names = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                names.add(node.module.split(".")[0])
        forbidden = {
            "app",
            "tree_editor",
            "toplevels",
            "widgets",
            "tree_compare",
        }
        self.assertEqual(names & forbidden, set())


class TestNoEditorDataRebind(unittest.TestCase):
    def test_editor_paths_do_not_rebind_mt_data(self):
        leftovers = []
        for path in _EDITOR_FILES:
            text = path.read_text(encoding="utf-8")
            for i, line in enumerate(text.splitlines(), 1):
                stripped = line.lstrip()
                if stripped.startswith("#"):
                    continue
                if _REBIND.search(line) or _UNPACK_REBIND.search(line):
                    leftovers.append(f"{path.relative_to(REPO_ROOT)}:{i}:{line.rstrip()}")
        self.assertEqual(leftovers, [], msg="rebind of sheet.MT.data; use set_records\n" + "\n".join(leftovers))

    def test_sheetops_insert_rows_does_not_pass_end_string(self):
        text = (REPO_ROOT / "src" / "tree_editor.py").read_text(encoding="utf-8")
        self.assertNotIn(
            'idx="end" if idx is None',
            text,
            'tksheet treats string idx as an Excel letter; alpha2idx("END") is 3747',
        )


class TestAnimalsFixture(unittest.TestCase):
    def test_documentation_animals_shape(self):
        self.assertEqual(ANIMALS_TABLE[0], ["ID", "Parent", "Name"])
        self.assertEqual(ANIMALS_TABLE[1][0], "Animals")
        self.assertEqual(ANIMALS_TABLE[2][0], "Cats")
        self.assertEqual(ANIMALS_TABLE[3][0], "Lion")
        self.assertEqual(ANIMALS_TABLE[3][1], "Cats")


class TestLoadTable(unittest.TestCase):
    def test_animals_ids_and_parents(self):
        s = Session()
        out = s.load_table(
            ANIMALS_TABLE,
            id_col=ANIMALS_ID_COL,
            parent_cols=ANIMALS_PARENT_COLS,
            fmt=0,
        )
        self.assertTrue(out["ok"], out)
        self.assertTrue(s.has_document)
        self.assertEqual(sorted(s.nodes), ["animals", "cats", "lion"])
        self.assertEqual(s.nodes["animals"].name, "Animals")
        self.assertEqual(s.nodes["animals"].ps[s.pc], "")
        self.assertEqual(s.nodes["cats"].ps[s.pc], "animals")
        self.assertEqual(s.nodes["lion"].ps[s.pc], "cats")
        self.assertEqual(s.nodes["cats"].cn[s.pc], ["lion"])
        self.assertEqual(s.data[s.rns["lion"]][1], "Cats")
        self.assertEqual(s.data[s.rns["lion"]][2], "Lion")
        self.assertEqual(out["result"]["ids"], 3)
        self.assertEqual(out["result"]["format"], 0)
        self.assertEqual(out["result"]["id_column"]["name"], "ID")
        self.assertEqual(out["result"]["id_column"]["letter"], "A")
        self.assertEqual(out["result"]["hierarchies"][0]["name"], "Parent")
        self.assertEqual(ANIMALS_TABLE[0], ["ID", "Parent", "Name"])

    def test_need_columns_does_not_clear_an_open_document(self):
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        out = s.load_table(ANIMALS_TABLE, fmt=0)
        self.assertFalse(out["ok"])
        self.assertEqual(out["error"]["code"], "need_columns")
        self.assertIn("lion", s.nodes)

    def test_duplicate_id_warning_text(self):
        s = Session()
        table = [
            ["ID", "Parent"],
            ["A", ""],
            ["A", ""],
        ]
        out = s.load_table(table, id_col=0, parent_cols=[1], fmt=0)
        self.assertTrue(out["ok"], out)
        self.assertIn("a", s.nodes)
        self.assertIn("a_duplicated_1", s.nodes)
        self.assertTrue(
            any("renamed due to repeat occurrence at row #3" in w for w in s.warnings),
            s.warnings,
        )

    def test_missing_parent_row_is_added(self):
        s = Session()
        table = [
            ["ID", "Parent", "Name"],
            ["Lion", "Cats", "Lion"],
        ]
        out = s.load_table(table, id_col=0, parent_cols=[1], fmt=0)
        self.assertTrue(out["ok"], out)
        self.assertIn("cats", s.nodes)
        self.assertTrue(
            any("missing from ID column, new row added" in w for w in s.warnings),
            s.warnings,
        )

    def test_new_has_default_columns_and_no_ids(self):
        s = Session()
        out = s.new()
        self.assertTrue(out["ok"], out)
        self.assertTrue(s.has_document)
        self.assertIsNone(s.filepath)
        self.assertEqual([h.name for h in s.headers], ["ID", "DETAIL_1", "PARENT_1"])
        self.assertEqual(out["result"]["ids"], 0)
        self.assertEqual(s.data, [])
        out = s.close()
        self.assertTrue(out["ok"])
        self.assertFalse(s.has_document)

    def test_new_blocked_when_unsaved(self):
        s = Session()
        s.new()
        s.changelog_append("Add ID", "Lion", "", "")
        out = s.new()
        self.assertFalse(out["ok"])
        self.assertEqual(out["error"]["code"], "unsaved")
        out = s.new(discard=True)
        self.assertTrue(out["ok"])
        self.assertEqual(s.changelog, [])


class TestSaveAndLoadPath(unittest.TestCase):
    def test_csv_roundtrip_animals(self):
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.csv")
            out = s.save_path(path, overwrite=True)
            self.assertTrue(out["ok"], out)
            self.assertEqual(out["result"]["kind"], "csv")
            self.assertTrue(os.path.isfile(path))
            exists = s.save_path(path)
            self.assertFalse(exists["ok"])
            self.assertEqual(exists["error"]["code"], "file_exists")
            s2 = Session()
            loaded = s2.load_path(path, id_col=0, parent_cols=[1])
            self.assertTrue(loaded["ok"], loaded)
            self.assertEqual(sorted(s2.nodes), ["animals", "cats", "lion"])
            self.assertEqual(s2.nodes["lion"].ps[s2.pc], "cats")
            self.assertEqual(os.path.abspath(path), s2.filepath)

    def test_json_roundtrip_with_program_data(self):
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        s.add("Tiger", "Cats")
        s.delete(["Lion", "Tiger"])
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.json")
            out = s.save_path(path, overwrite=True)
            self.assertTrue(out["ok"], out)
            with open(path, encoding="utf-8") as fh:
                raw = json.loads(fh.read())
            self.assertGreater(len(raw["program_data"]), 0)
            blob = raw["changelog"]
            self.assertEqual(blob[-1][1], "Delete 2 IDs")
            self.assertEqual(len(raw["changelog"][0]), 5)
            s2 = Session()
            loaded = s2.load_path(path)
            self.assertTrue(loaded["ok"], loaded)
            self.assertEqual(sorted(s2.nodes), ["animals", "cats"])
            self.assertEqual(s2.changelog[-1].label, "Delete 2 IDs")
            self.assertEqual(s2.changelog[-1].n, 2)
            self.assertNotIn("lion", s2.nodes)
            self.assertNotIn("tiger", s2.nodes)

    def test_xlsx_roundtrip_with_program_data(self):
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.xlsx")
            out = s.save_path(path, overwrite=True)
            self.assertTrue(out["ok"], out)
            s2 = Session()
            loaded = s2.load_path(path)
            self.assertTrue(loaded["ok"], loaded)
            self.assertEqual(sorted(s2.nodes), ["animals", "cats", "lion"])

    def test_load_path_missing_file(self):
        s = Session()
        out = s.load_path("/this/path/does/not/exist.csv", id_col=0, parent_cols=[1])
        self.assertFalse(out["ok"])
        self.assertEqual(out["error"]["code"], "file_not_found")

    def test_csv_without_columns_is_need_columns(self):
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.csv")
            s.save_path(path, overwrite=True)
            s2 = Session()
            out = s2.load_path(path)
            self.assertFalse(out["ok"])
            self.assertEqual(out["error"]["code"], "need_columns")

    def test_load_program_data_rebuild_uses_file_allow_spaces(self):
        src = Session()
        src.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        d = src.program_data_dict()
        d["allow_spaces_ids"] = True
        d["nodes"] = "corrupt"
        d["records"] = [
            ["Hello World", "", "x"],
            ["Child", "Hello World", "y"],
        ]
        s = Session()
        self.assertFalse(s.allow_spaces_ids_var)
        out = s.load_program_data(d)
        self.assertTrue(out["ok"], out)
        self.assertIn("hello world", s.nodes)
        self.assertNotIn("helloworld", s.nodes)


class TestGitignoreAllowsTests(unittest.TestCase):
    def test_test_modules_are_not_gitignored(self):
        gitignore = (REPO_ROOT / ".gitignore").read_text(encoding="utf-8")
        for line in gitignore.splitlines():
            stripped = line.strip()
            if stripped.startswith("#") or not stripped:
                continue
            self.assertFalse(
                stripped == "test*.py*" or stripped == "test*.py",
                "gitignore must not ignore tests/test_*.py; use /test*.py* for root scratch files",
            )


if __name__ == "__main__":
    os.chdir(REPO_ROOT)
    unittest.main()
