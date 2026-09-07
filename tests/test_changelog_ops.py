# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Changelog grouping, import, persist, and GUI shim writers."""

from __future__ import annotations

import os
import tempfile
import unittest

from src.changelog import display_rows, flatten_changelog, load_changelog
from tests.helpers import animals, assert_changelog_sound, assert_one_action, assert_sound, forest


class TestPasteShimsGroup(unittest.TestCase):
    def test_member_rows_then_summary_are_one_change(self):
        s = animals()
        n_log, n_vs = len(s.changelog), len(s.vs)
        s.changelog_append_no_unsaved("Copy and paste ID |", "Lion", "from", "to")
        s.changelog_append_no_unsaved("Copy and paste ID |", "Tiger", "from", "to")
        s.changelog_append("Copy and paste 2 IDs", "", "", "")
        self.assertEqual(len(s.changelog), n_log + 1)
        ch = s.changelog[-1]
        self.assertEqual(ch.label, "Copy and paste 2 IDs")
        self.assertEqual(ch.n, 2)
        self.assertTrue(ch.has_summary)
        self.assertEqual(ch.rows[0].display_type, "Copy and paste ID |")
        self.assertEqual(ch.rows[-1].display_type, "Copy and paste 2 IDs")
        assert_changelog_sound(self, s)
        s.vs.append({"type": "paste id", "rows": [], "required_data": s.snapshot_required_data()})
        self.assertEqual(len(s.vs), n_vs + 1)

    def test_singular_one_member_has_no_summary_pipe(self):
        s = animals()
        s.changelog_append_no_unsaved("Cut and paste ID |", "Lion", "Cats", "Dogs")
        s.changelog_singular("Cut and paste ID")
        ch = s.changelog[-1]
        self.assertEqual(ch.label, "Cut and paste ID")
        self.assertFalse(ch.has_summary)
        self.assertEqual(ch.rows[0].display_type, "Cut and paste ID")
        self.assertEqual(len(ch.rows), 1)
        assert_changelog_sound(self, s)

    def test_edit_cell_pipe_then_singular(self):
        s = animals()
        s.changelog_append_no_unsaved(
            "Edit cell |",
            "ID: Lion column #3 named: Name with type: Text",
            "Lion",
            "Panthera",
        )
        s.changelog_singular("Edit cell")
        ch = s.changelog[-1]
        self.assertEqual(ch.label, "Edit cell")
        self.assertEqual(ch.n, 1)
        self.assertFalse(ch.has_summary)
        assert_changelog_sound(self, s)

    def test_display_rows_are_existing_objects(self):
        s = animals()
        s.add("Tiger", "Cats")
        s.delete(["Lion", "Tiger"])
        flat = display_rows(s.changelog)
        self.assertGreater(len(flat), len(s.changelog))
        self.assertIs(flat[-1], s.changelog[-1].rows[-1])
        self.assertIs(flat[-1][2], s.changelog[-1].rows[-1].what)


class TestImportMoreTypes(unittest.TestCase):
    def test_import_move_and_copy(self):
        s = animals()
        s.move("Lion", parent="Animals")
        move_row = list(s.changelog[-1].rows[0])
        s.undo()
        out = s.import_changes([move_row])
        self.assertEqual(out["result"]["applied"], 1, out)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        assert_sound(self, s, "import-move")

        s2 = animals()
        col = len(s2.headers)
        s2.add_hier_col(col, "H2")
        s2.pc = col
        s2.add("Animals", "")
        s2.pc = s2.hiers[0]
        s2.copy("Lion", parent="Animals", hier=col)
        copy_row = list(s2.changelog[-1].rows[0])
        s2.undo()
        self.assertIsNone(s2.nodes["lion"].ps[col])
        out = s2.import_changes([copy_row])
        self.assertEqual(out["result"]["applied"], 1, out)
        self.assertEqual(s2.nodes["lion"].ps[col], "animals")
        assert_sound(self, s2, "import-copy")

    def test_import_delete_children_and_all_hier(self):
        s = animals()
        s.add("Cub", "Lion")
        s.delete(["Lion"], children=True)
        row = list(s.changelog[-1].rows[0])
        s2 = animals()
        s2.add("Cub", "Lion")
        out = s2.import_changes([row])
        self.assertEqual(out["result"]["applied"], 1, out)
        self.assertNotIn("lion", s2.nodes)
        self.assertNotIn("cub", s2.nodes)
        assert_sound(self, s2, "import-del-children")

    def test_import_grouped_delete_replays_each_member(self):
        s = animals()
        s.add("Tiger", "Cats")
        s.delete(["Lion", "Tiger"])
        rows, _ = flatten_changelog([s.changelog[-1]])
        s2 = animals()
        s2.add("Tiger", "Cats")
        out = s2.import_changes(rows)
        self.assertGreaterEqual(out["result"]["applied"], 2, out)
        self.assertNotIn("lion", s2.nodes)
        self.assertNotIn("tiger", s2.nodes)
        assert_sound(self, s2, "import-grouped")

    def test_import_is_one_changelog_action(self):
        s = animals()
        s.add("Tiger", "Cats")
        row = list(s.changelog[-1].rows[0])
        s2 = animals()
        n_log = len(s2.changelog)
        s2.import_changes([row])
        self.assertEqual(len(s2.changelog), n_log + 1)
        self.assertEqual(s2.changelog[-1].origin, "import")
        self.assertTrue(s2.changelog[-1].has_summary)
        assert_changelog_sound(self, s2)

    def test_prune_changelog_is_one_undoable_action(self):
        s = animals()
        s.add("Tiger", "Cats")
        s.add("Wolf", "Cats")
        n = len(s.changelog)
        out = s.prune_changelog(0)
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["n"], 1)
        self.assertEqual(len(s.changelog), n)
        self.assertEqual(s.changelog[-1].label, "Pruned changelog")

    def test_failed_import_does_not_leave_pending_snapshot(self):
        s = animals()
        n_vs = len(s.vs)
        out = s.import_changes([["2026/01/01", "Rename ID", "Nope", "Nope", "Ghost"]])
        self.assertEqual(out["result"]["applied"], 0)
        self.assertEqual(len(s.vs), n_vs)
        self.assertIn("lion", s.nodes)


class TestPersistAndList(unittest.TestCase):
    def test_json_program_data_keeps_grouped_changelog(self):
        s = animals()
        s.add("Tiger", "Cats")
        s.delete(["Lion", "Tiger"])
        ch = s.changelog[-1]
        self.assertEqual(ch.n, 2)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "a.json")
            self.assertTrue(s.save_path(path, overwrite=True)["ok"])
            from src.session import Session

            s2 = Session()
            out = s2.load_path(path)
            self.assertTrue(out["ok"], out)
            self.assertEqual(s2.changelog[-1].n, 2)
            self.assertEqual(s2.changelog[-1].label, "Delete 2 IDs")
            assert_changelog_sound(self, s2)
            assert_sound(self, s2, "reload")

    def test_changelog_list_counts_actions_not_rows(self):
        s = animals()
        s.add("Tiger", "Cats")
        s.delete(["Lion", "Tiger"])
        listed = s.changelog_list(limit=0)
        self.assertEqual(listed["result"]["total"], len(s.changelog))
        last = listed["result"]["entries"][-1]
        self.assertEqual(last["n"], 2)
        self.assertEqual(last["type"], "Delete 2 IDs")
        self.assertEqual(len(last["rows"]), 3)

    def test_export_flatten_then_load_regroups(self):
        s = animals()
        s.delete(["Lion"])
        s.add("Tiger", "Cats")
        rows, idx = flatten_changelog(s.changelog)
        back = load_changelog(rows)
        self.assertEqual(len(back), len(s.changelog))
        self.assertEqual(flatten_changelog(back)[0], rows)


class TestReplaceAndMergeChangelog(unittest.TestCase):
    def test_replace_many_cells_is_one_action(self):
        s = forest()
        n_log, n_vs = len(s.changelog), len(s.vs)
        out = s.replace_mapping({"lion": "Simba"})
        self.assertTrue(out["ok"], out)
        self.assertGreater(out["result"]["cells_changed"], 1)
        assert_one_action(self, s, n_log, n_vs, "replace")
        self.assertTrue(s.changelog[-1].has_summary)
        assert_sound(self, s, "replace")

    def test_merge_is_one_import_like_action(self):
        s = animals()
        n_log = len(s.changelog)
        incoming = [
            ["ID", "Parent", "Name"],
            ["Lion", "Cats", "Panthera leo"],
            ["Plants", "", "All plants"],
        ]
        out = s.merge_from_rows(incoming, fmt=0, id_col=0, parent_cols=[1])
        self.assertTrue(out["ok"], out)
        self.assertEqual(len(s.changelog), n_log + 1)
        self.assertEqual(s.changelog[-1].origin, "merge")
        self.assertTrue(s.changelog[-1].has_summary)
        assert_sound(self, s, "merge-log")
