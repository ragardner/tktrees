# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import unittest

from src.session import Session
from tests.helpers import assert_sound, forest, two_hier


def _forest():
    return forest()


def _two():
    return two_hier()


class TestSoundnessAfterLoad(unittest.TestCase):
    def test_forest_and_two_hier_are_sound(self):
        s = _forest()
        assert_sound(self, s, "forest")
        t = _two()
        assert_sound(self, t, "two-hier")
        self.assertEqual(t.nodes["b"].ps[t.hiers[0]], "a")
        self.assertEqual(t.nodes["b"].ps[t.hiers[1]], "root")
        self.assertEqual(t.nodes["c"].ps[t.hiers[0]], "")
        self.assertEqual(t.nodes["c"].ps[t.hiers[1]], "root")
        self.assertEqual(t.nodes["d"].ps[t.hiers[0]], "c")
        # Empty parent cell and no children in that hierarchy → not in it (associate).
        self.assertIsNone(t.nodes["d"].ps[t.hiers[1]])

    def test_self_parent_is_cleared_with_warning(self):
        s = Session()
        out = s.load_table(
            [["ID", "Parent"], ["Loop", "Loop"], ["Kid", "Loop"]],
            id_col=0,
            parent_cols=[1],
            fmt=0,
        )
        self.assertTrue(out["ok"], out)
        self.assertTrue(any("same as parent" in w for w in s.warnings), s.warnings)
        self.assertEqual(s.nodes["loop"].ps[s.pc], "")
        assert_sound(self, s, "self-parent")

    def test_spaces_warning_and_stripped_id(self):
        s = Session()
        out = s.load_table(
            [["ID", "Parent"], ["Bad ID", ""], ["Child", "Bad ID"]],
            id_col=0,
            parent_cols=[1],
            fmt=0,
        )
        self.assertTrue(out["ok"], out)
        self.assertTrue(any("Spaces" in w for w in s.warnings), s.warnings)
        self.assertIn("badid", s.nodes)
        self.assertNotIn("bad id", s.nodes)
        assert_sound(self, s, "spaces")


class TestDeletesAndOrder(unittest.TestCase):
    def test_orphan_nested_appends_children_as_tops(self):
        s = _forest()
        s.set_option("auto-sort", "off")
        before = list(s.topnodes_order[s.pc])
        s.delete(["Cats"], orphan=True)
        self.assertNotIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertEqual(s.nodes["tiger"].ps[s.pc], "")
        self.assertEqual(
            list(s.topnodes_order[s.pc]),
            before + ["lion", "tiger"],
        )
        assert_sound(self, s, "orphan")

    def test_orphan_top_same_as_promote_for_children_but_keeps_no_parent_row(self):
        s = _forest()
        s.set_option("auto-sort", "off")
        s.delete(["Moss"], orphan=True)
        self.assertNotIn("moss", s.nodes)
        self.assertNotIn("moss", s.topnodes_order[s.pc])
        assert_sound(self, s, "orphan-leaf-top")

    def test_delete_children_removes_descendants_keeps_siblings(self):
        s = _forest()
        s.delete(["Cats"], children=True)
        self.assertNotIn("cats", s.nodes)
        self.assertNotIn("lion", s.nodes)
        self.assertNotIn("tiger", s.nodes)
        self.assertIn("dogs", s.nodes)
        self.assertIn("wolf", s.nodes)
        assert_sound(self, s, "delete-children")

    def test_delete_from_one_hier_keeps_row_in_other(self):
        s = _two()
        s.pc = s.hiers[0]
        out = s.delete(["B"])
        self.assertTrue(out["ok"], out)
        self.assertIn("b", s.nodes)
        self.assertIsNone(s.nodes["b"].ps[s.hiers[0]])
        self.assertEqual(s.nodes["b"].ps[s.hiers[1]], "root")
        self.assertIn("B", out["result"]["removed_from_hierarchy"])
        assert_sound(self, s, "delete-one-hier")

    def test_delete_all_hierarchies_removes_row(self):
        s = _two()
        s.pc = s.hiers[0]
        out = s.delete(["B"], all_hierarchies=True)
        self.assertNotIn("b", s.nodes)
        self.assertIn("B", out["result"]["removed_entirely"])
        self.assertNotIn("b", s.rns)
        assert_sound(self, s, "delete-all-hier")

    def test_delete_all_hier_orphan_makes_children_tops_everywhere(self):
        s = _two()
        s.pc = s.hiers[0]
        s.delete(["A"], orphan=True, all_hierarchies=True)
        self.assertNotIn("a", s.nodes)
        self.assertEqual(s.nodes["b"].ps[s.hiers[0]], "")
        assert_sound(self, s, "orphan-all")

    def test_delete_not_in_current_hierarchy(self):
        s = _two()
        col = len(s.headers)
        s.add_hier_col(col, "H3")
        s.pc = col
        out = s.delete(["D"])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["not_in_hierarchy"], ["D"])
        self.assertIn("d", s.nodes)
        assert_sound(self, s, "not-in-hier")

    def test_tagged_id_cleared_on_full_delete(self):
        s = _forest()
        s.tag(["Lion"])
        s.delete(["Lion"])
        self.assertNotIn("lion", s.tagged_ids)
        assert_sound(self, s)


class TestMoveCopyCycle(unittest.TestCase):
    def test_move_self_parent_is_cycle(self):
        s = _forest()
        out = s.move("Lion", parent="Lion")
        self.assertEqual(out["error"]["code"], "cycle")
        assert_sound(self, s, "cycle-self")

    def test_move_into_current_parent_fails(self):
        s = _forest()
        out = s.move("Lion", parent="Cats")
        self.assertEqual(out["error"]["code"], "already_in_hierarchy")
        assert_sound(self, s)

    def test_move_between_hierarchies(self):
        s = _two()
        s.pc = s.hiers[0]
        # C is top in H1; move B (under A in H1) to C in H1
        out = s.move("B", parent="C")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["b"].ps[s.hiers[0]], "c")
        self.assertNotIn("b", s.nodes["a"].cn[s.hiers[0]])
        self.assertIn("b", s.nodes["c"].cn[s.hiers[0]])
        self.assertEqual(s.nodes["b"].ps[s.hiers[1]], "root")
        assert_sound(self, s, "move-h1")

    def test_move_from_other_hier_to_current(self):
        s = _two()
        h1 = s.hiers[0]
        col = len(s.headers)
        s.add_hier_col(col, "H3")
        s.pc = h1
        out = s.move("D", parent="Root", from_hier=col)
        self.assertEqual(out["error"]["code"], "not_in_hierarchy")
        s.pc = col
        s.add("Root", "")
        out = s.move("D", parent="Root", from_hier=h1, hier=col)
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["d"].ps[col], "root")
        assert_sound(self, s, "move-across")

    def test_copy_same_hierarchy_fails(self):
        s = _two()
        h1, _h2 = s.hiers
        s.pc = h1
        out = s.copy("B", parent="Root", hier=h1)
        self.assertEqual(out["error"]["code"], "already_in_hierarchy")

    def test_copy_into_other_hier_as_child(self):
        s = _two()
        h1 = s.hiers[0]
        col = len(s.headers)
        s.add_hier_col(col, "H3")
        s.pc = col
        s.add("Root", "")
        s.pc = h1
        out = s.copy("D", parent="Root", hier=col)
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["d"].ps[col], "root")
        self.assertEqual(s.nodes["d"].ps[h1], "c")
        assert_sound(self, s, "copy-h3")

    def test_copy_top_appends_topnodes_when_auto_sort_off(self):
        s = _two()
        h1 = s.hiers[0]
        col = len(s.headers)
        s.add_hier_col(col, "H3")
        s.set_option("auto-sort", "off")
        s.pc = h1
        s.copy("A", top=True, hier=col)
        self.assertEqual(s.nodes["a"].ps[col], "")
        self.assertEqual(s.topnodes_order[col][-1], "a")
        assert_sound(self, s, "copy-top")

    def test_cut_paste_children_and_sound(self):
        s = _forest()
        s.cut_paste_children("Dogs", "Cats", s.pc, snapshot=False)
        self.assertEqual(s.nodes["wolf"].ps[s.pc], "cats")
        self.assertEqual(s.nodes["dogs"].cn[s.pc], [])
        self.assertIn("wolf", s.nodes["cats"].cn[s.pc])
        assert_sound(self, s, "cut-children")

    def test_add_existing_id_to_second_hierarchy(self):
        s = _two()
        h1 = s.hiers[0]
        col = len(s.headers)
        s.add_hier_col(col, "H3")
        s.pc = col
        s.add("Root", "")
        out = s.add("D", "Root")
        self.assertTrue(out["ok"], out)
        self.assertFalse(out["result"]["added"][0]["created_row"])
        self.assertEqual(s.nodes["d"].ps[col], "root")
        self.assertEqual(s.nodes["d"].ps[h1], "c")
        assert_sound(self, s, "add-second-hier")


class TestUndoRestores(unittest.TestCase):
    def test_undo_set_detail(self):
        s = _forest()
        old = s.data[s.rns["lion"]][2]
        s.set_detail("Lion", 2, "Panthera")
        self.assertEqual(s.data[s.rns["lion"]][2], "Panthera")
        out = s.undo()
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["undone"], "ctrl x, v, del key")
        self.assertEqual(s.data[s.rns["lion"]][2], old)
        assert_sound(self, s, "undo-set")

    def test_undo_rename(self):
        s = _forest()
        s.rename("Cats", "Felines")
        s.undo()
        self.assertIn("cats", s.nodes)
        self.assertNotIn("felines", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "cats")
        self.assertEqual(s.data[s.rns["lion"]][s.pc], "Cats")
        assert_sound(self, s, "undo-rename")

    def test_undo_delete_promotes_back(self):
        s = _forest()
        s.delete(["Cats"])
        self.assertNotIn("cats", s.nodes)
        s.undo()
        self.assertIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "cats")
        assert_sound(self, s, "undo-delete")

    def test_undo_sort_restores_row_order(self):
        s = _forest()
        before = [row[0] for row in s.data]
        s.sort_sheet("Name", "DESCENDING")
        self.assertNotEqual([row[0] for row in s.data], before)
        s.undo()
        self.assertEqual([row[0] for row in s.data], before)
        assert_sound(self, s, "undo-sort")

    def test_undo_move(self):
        s = _forest()
        s.move("Lion", parent="Animals")
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        s.undo()
        self.assertEqual(s.nodes["lion"].ps[s.pc], "cats")
        self.assertIn("lion", s.nodes["cats"].cn[s.pc])
        assert_sound(self, s, "undo-move")

    def test_undo_add_col_and_del_col(self):
        s = _forest()
        n = len(s.headers)
        s.add_col(n, "Habitat", "Text")
        self.assertEqual(len(s.headers), n + 1)
        s.undo()
        self.assertEqual(len(s.headers), n)
        self.assertEqual(s.row_len, n)
        assert_sound(self, s, "undo-add-col")
        s.add_col(n, "Habitat", "Text")
        s.del_cols([n])
        self.assertEqual(len(s.headers), n)
        s.undo()
        self.assertEqual(s.headers[n].name, "Habitat")
        assert_sound(self, s, "undo-del-col")

    def test_nothing_to_undo(self):
        s = _forest()
        s.vs.clear()
        self.assertEqual(s.undo()["error"]["code"], "nothing_to_undo")


class TestColumnsAndOptions(unittest.TestCase):
    def test_add_hier_then_operations_stay_sound(self):
        s = _forest()
        col = len(s.headers)
        s.add_hier_col(col, "H2")
        self.assertIn(col, s.hiers)
        for node in s.nodes.values():
            self.assertIsNone(node.ps[col])
            self.assertEqual(node.cn[col], [])
        s.pc = col
        s.add("Life", "")
        self.assertEqual(s.nodes["life"].ps[col], "")
        assert_sound(self, s, "new-hier")

    def test_spaces_not_allowed_until_option(self):
        s = _forest()
        self.assertEqual(s.add("Bad ID", "Cats")["error"]["code"], "spaces_not_allowed")
        s.set_option("allow-spaces-ids", "on")
        out = s.add("Bad ID", "Cats")
        self.assertTrue(out["ok"], out)
        self.assertIn("bad id", s.nodes)
        assert_sound(self, s, "spaces-id")

    def test_replace_id_rebuilds_identity(self):
        s = _forest()
        out = s.replace_mapping({"lion": "Simba"})
        self.assertTrue(out["ok"], out)
        self.assertGreaterEqual(out["result"]["cells_changed"], 1)
        self.assertIn("simba", s.nodes)
        self.assertNotIn("lion", s.nodes)
        assert_sound(self, s, "replace-id")

    def test_merge_overwrite_detail_and_add_id(self):
        s = _forest()
        incoming = [
            ["ID", "Parent", "Name"],
            ["Lion", "Cats", "Panthera leo"],
            ["NewTop", "", "Brand new"],
        ]
        out = s.merge_from_rows(incoming, fmt=0, id_col=0, parent_cols=[1])
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.data[s.rns["lion"]][2], "Panthera leo")
        self.assertIn("newtop", s.nodes)
        self.assertGreaterEqual(out["result"]["ids_added"], 1)
        self.assertGreaterEqual(out["result"]["details_written"], 1)
        assert_sound(self, s, "merge")

    def test_sort_children_oak_numeric(self):
        s = _forest()
        s.set_option("auto-sort", "off")
        s.nodes["oak"].cn[s.pc] = ["ch10", "ch2", "ch3"]
        s.sort_children("Oak")
        self.assertEqual(
            [s.nodes[i].name for i in s.nodes["oak"].cn[s.pc]],
            ["Ch2", "Ch3", "Ch10"],
        )
        assert_sound(self, s, "sort-children")
