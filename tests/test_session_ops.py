# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Order, tops, undo pairing, and multi-hierarchy mutations."""

from __future__ import annotations

import unittest

from tests.helpers import (
    animals,
    assert_one_action,
    assert_sound,
    children_of,
    forest,
    names,
    tops_of,
    two_hier,
)


def _h2_setup():
    """Forest plus a second hierarchy with mixed branched/leaf children."""
    s = forest()
    col = len(s.headers)
    s.add_hier_col(col, "H2")
    s.pc = col
    s.add("Life", "")
    s.add("Animals", "Life")
    s.add("Zebra", "Animals")
    s.add("Aardvark", "Animals")
    s.add("Cats", "Animals")
    s.add("Lion", "Cats")
    return s, col


class TestDeletePromotesAndSorts(unittest.TestCase):
    def test_promote_resorts_parent_children(self):
        s = forest()
        self.assertEqual(children_of(s, "Animals"), ["Cats", "Dogs", "Aardvark", "Zebra"])
        n_log, n_vs = len(s.changelog), len(s.vs)
        s.delete(["Cats"])
        assert_one_action(self, s, n_log, n_vs, "promote", "Delete ID")
        # Lion/Tiger become leaves under Animals; Dogs stays branched.
        self.assertEqual(children_of(s, "Animals"), ["Dogs", "Aardvark", "Lion", "Tiger", "Zebra"])
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        self.assertEqual(s.nodes["tiger"].ps[s.pc], "animals")
        assert_sound(self, s, "promote")

    def test_delete_last_branched_child_resorts_grandparent(self):
        s = forest()
        s.add("MidLeaf", "Life")
        # Animals and Plants are branched; MidLeaf is a leaf.
        self.assertEqual(children_of(s, "Life"), ["Animals", "Plants", "MidLeaf"])
        s.delete(["Oak"], children=True)
        self.assertEqual(s.nodes["plants"].cn[s.pc], [])
        # Plants is now a leaf, so it sorts after MidLeaf.
        self.assertEqual(children_of(s, "Life"), ["Animals", "MidLeaf", "Plants"])
        assert_sound(self, s, "gp-sort")

    def test_delete_leaf_keeps_parent_order(self):
        s = forest()
        before = children_of(s, "Animals")
        s.delete(["Aardvark"])
        self.assertEqual(children_of(s, "Animals"), [n for n in before if n != "Aardvark"])
        assert_sound(self, s, "leaf")

    def test_delete_two_is_one_undo_and_restores_order(self):
        s = forest()
        before = children_of(s, "Animals")
        n_log = len(s.changelog)
        s.delete(["Lion", "Tiger"])
        self.assertEqual(len(s.changelog), n_log + 1)
        self.assertEqual(s.changelog[-1].n, 2)
        self.assertEqual(children_of(s, "Cats"), [])
        s.undo()
        self.assertEqual(children_of(s, "Animals"), before)
        self.assertEqual(children_of(s, "Cats"), ["Lion", "Tiger"])
        assert_sound(self, s, "undo-two")

    def test_orphan_nested_auto_sort_makes_leaf_tops(self):
        s = forest()
        s.delete(["Cats"], orphan=True)
        self.assertNotIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertEqual(s.nodes["tiger"].ps[s.pc], "")
        # Lion/Tiger are leaf tops; Life still branched.
        tops = tops_of(s)
        self.assertIn("Lion", tops)
        self.assertIn("Tiger", tops)
        self.assertLess(tops.index("Life"), tops.index("Lion"))
        assert_sound(self, s, "orphan-auto")

    def test_delete_children_does_not_touch_other_hier_order(self):
        s, col = _h2_setup()
        h1 = s.hiers[0]
        before_h2 = list(s.nodes["animals"].cn[col])
        s.pc = h1
        s.delete(["Tiger"], children=True)
        self.assertEqual(list(s.nodes["animals"].cn[col]), before_h2)
        self.assertNotIn("tiger", s.nodes)
        assert_sound(self, s, "children-other-hier")


class TestDeleteAllHierarchiesOrder(unittest.TestCase):
    def test_other_hierarchy_parent_resorts_when_child_becomes_leaf(self):
        s, col = _h2_setup()
        # H2 Animals children: Cats (branched, has Lion), then Aardvark, Zebra.
        self.assertEqual(children_of(s, "Animals", col), ["Cats", "Aardvark", "Zebra"])
        h1_cats = list(s.nodes["cats"].cn[s.hiers[0]])
        s.pc = s.hiers[0]
        s.delete(["Lion"], all_hierarchies=True)
        self.assertNotIn("lion", s.nodes)
        # H1: Cats still has Tiger, so still branched.
        self.assertEqual(names(s, [i for i in h1_cats if i != "lion"]), children_of(s, "Cats", s.hiers[0]))
        # H2: Cats is now a leaf, so Animals children are all leaves: Aardvark, Cats, Zebra.
        self.assertEqual(children_of(s, "Animals", col), ["Aardvark", "Cats", "Zebra"])
        assert_sound(self, s, "all-hier-resort")

    def test_one_hier_delete_does_not_resort_other(self):
        s, col = _h2_setup()
        before_h2 = children_of(s, "Animals", col)
        s.pc = s.hiers[0]
        s.delete(["Lion"])
        self.assertIsNone(s.nodes["lion"].ps[s.hiers[0]])
        self.assertEqual(s.nodes["lion"].ps[col], "cats")
        self.assertEqual(children_of(s, "Animals", col), before_h2)
        assert_sound(self, s, "one-hier-leave-other")

    def test_delete_all_hier_top_promotes_in_every_hier(self):
        s = two_hier()
        h1, h2 = s.hiers
        s.pc = h1
        s.delete(["A"], all_hierarchies=True)
        self.assertNotIn("a", s.nodes)
        self.assertEqual(s.nodes["b"].ps[h1], "root")
        self.assertEqual(s.nodes["b"].ps[h2], "root")
        assert_sound(self, s, "del-a-all")

    def test_undo_delete_all_hier_restores_both(self):
        s, col = _h2_setup()
        before_h1 = children_of(s, "Cats", s.hiers[0])
        before_h2 = children_of(s, "Animals", col)
        s.pc = s.hiers[0]
        s.delete(["Lion"], all_hierarchies=True)
        s.undo()
        self.assertEqual(children_of(s, "Cats", s.hiers[0]), before_h1)
        self.assertEqual(children_of(s, "Animals", col), before_h2)
        self.assertIn("lion", s.nodes)
        assert_sound(self, s, "undo-all-hier")

    def test_orphan_all_hier_tops_and_order(self):
        s, col = _h2_setup()
        s.pc = col
        s.delete(["Cats"], orphan=True, all_hierarchies=True)
        self.assertNotIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[col], "")
        self.assertEqual(s.nodes["lion"].ps[s.hiers[0]], "")
        assert_sound(self, s, "orphan-all-h2")


class TestTopnodesOrderOps(unittest.TestCase):
    def test_delete_top_all_hier_updates_each_list(self):
        s = two_hier()
        s.set_option("auto-sort", "off")
        h1, h2 = s.hiers
        s.pc = h1
        before_h2 = list(s.topnodes_order[h2])
        s.delete(["C"], all_hierarchies=True)
        self.assertNotIn("c", s.nodes)
        self.assertNotIn("c", s.topnodes_order[h1])
        # C was not a top in H2 (under Root); H2 tops unchanged.
        self.assertEqual(list(s.topnodes_order[h2]), before_h2)
        # D was under C in H1 and is promoted to top.
        self.assertEqual(s.nodes["d"].ps[h1], "")
        self.assertIn("d", s.topnodes_order[h1])
        assert_sound(self, s, "top-all-hier")

    def test_undo_delete_top_restores_topnodes_order(self):
        s = forest()
        s.set_option("auto-sort", "off")
        before = list(s.topnodes_order[s.pc])
        s.delete(["Moss"])
        self.assertNotIn("moss", s.topnodes_order[s.pc])
        s.undo()
        self.assertEqual(list(s.topnodes_order[s.pc]), before)
        assert_sound(self, s, "undo-top")

    def test_add_to_second_hier_as_top_appends_only_that_list(self):
        s = two_hier()
        s.set_option("auto-sort", "off")
        h1, h2 = s.hiers
        before_h1 = list(s.topnodes_order[h1])
        s.pc = h2
        s.add("D", "")
        self.assertEqual(s.nodes["d"].ps[h2], "")
        self.assertEqual(s.topnodes_order[h2][-1], "d")
        self.assertEqual(list(s.topnodes_order[h1]), before_h1)
        assert_sound(self, s, "add-top-h2")

    def test_del_hierarchy_column_drops_topnodes_key(self):
        s = forest()
        col = len(s.headers)
        s.add_hier_col(col, "H2")
        s.set_option("auto-sort", "off")
        self.assertIn(col, s.topnodes_order)
        s.del_cols([col])
        self.assertNotIn(col, s.hiers)
        self.assertNotIn(col, s.topnodes_order)
        assert_sound(self, s, "del-hier-col")


class TestMoveCopyOrder(unittest.TestCase):
    def test_move_resorts_new_parent_and_grandparent(self):
        s = forest()
        s.move("Lion", parent="Animals")
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        self.assertEqual(children_of(s, "Animals"), ["Cats", "Dogs", "Aardvark", "Lion", "Zebra"])
        self.assertEqual(children_of(s, "Cats"), ["Tiger"])
        assert_sound(self, s, "move-sort")

    def test_move_to_top_auto_sort(self):
        s = forest()
        s.move("Lion", top=True)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertIn("Lion", tops_of(s))
        self.assertNotIn("lion", s.nodes["cats"].cn[s.pc])
        assert_sound(self, s, "move-top")

    def test_copy_into_h2_resorts_dest_not_source(self):
        s = forest()
        col = len(s.headers)
        s.add_hier_col(col, "H2")
        s.pc = col
        s.add("Life", "")
        s.add("Animals", "Life")
        s.add("Zebra", "Animals")
        src = list(s.nodes["cats"].cn[s.hiers[0]])
        s.pc = s.hiers[0]
        s.copy("Lion", parent="Animals", hier=col)
        self.assertEqual(list(s.nodes["cats"].cn[s.hiers[0]]), src)
        self.assertEqual(children_of(s, "Animals", col), ["Lion", "Zebra"])
        assert_sound(self, s, "copy-h2")

    def test_cut_paste_children_resorts_new_parent(self):
        s = forest()
        s.cut_paste_children("Dogs", "Cats", s.pc)
        self.assertEqual(children_of(s, "Cats"), ["Lion", "Tiger", "Wolf"])
        self.assertEqual(children_of(s, "Dogs"), [])
        assert_sound(self, s, "cut-children-sort")

    def test_sort_later_then_flush_matches_immediate_sort(self):
        s = forest()
        col = len(s.headers)
        s.add_hier_col(col, "H2")
        s.pc = col
        s.add("Life", "")
        s.add("Animals", "Life")
        s.add("Zebra", "Animals")
        s.sort_later_dct = None
        s.copy_paste("Lion", s.hiers[0], "Animals", snapshot=True, sort_later=True)
        s.copy_paste("Aardvark", s.hiers[0], "Animals", snapshot=True, sort_later=True)
        # Not flushed yet: append order.
        self.assertEqual(s.nodes["animals"].cn[col][-2:], ["lion", "aardvark"])
        s.apply_sort_later()
        self.assertEqual(children_of(s, "Animals", col), ["Aardvark", "Lion", "Zebra"])
        assert_sound(self, s, "sort-later")


class TestUndoPairsChangelog(unittest.TestCase):
    def test_each_mutate_is_one_changelog_and_one_snapshot(self):
        s = animals()
        cases = [
            (lambda: s.add("Tiger", "Cats"), "Add ID"),
            (lambda: s.rename("Tiger", "Tigers"), "Rename ID"),
            (lambda: s.set_detail("Tigers", 2, "Panthera"), "Edit cell"),
            (lambda: s.move("Tigers", parent="Animals"), "Cut and paste ID"),
            (lambda: s.delete(["Tigers"]), "Delete ID"),
        ]
        for fn, label in cases:
            n_log, n_vs = len(s.changelog), len(s.vs)
            out = fn()
            self.assertTrue(out["ok"], (label, out))
            assert_one_action(self, s, n_log, n_vs, label, label)
            assert_sound(self, s, label)

    def test_undo_restores_after_chain(self):
        s = animals()
        s.add("Tiger", "Cats")
        s.move("Tiger", parent="Animals")
        s.delete(["Tiger"])
        self.assertNotIn("tiger", s.nodes)
        s.undo()
        self.assertIn("tiger", s.nodes)
        self.assertEqual(s.nodes["tiger"].ps[s.pc], "animals")
        s.undo()
        self.assertEqual(s.nodes["tiger"].ps[s.pc], "cats")
        s.undo()
        self.assertNotIn("tiger", s.nodes)
        assert_sound(self, s, "chain-undo")


class TestColumnsStaySound(unittest.TestCase):
    def test_add_and_delete_detail_column(self):
        s = forest()
        n = len(s.headers)
        s.add_col(n, "Habitat", "Text")
        s.set_detail("Lion", n, "Savanna")
        assert_sound(self, s, "add-detail")
        s.del_cols([n])
        self.assertEqual(len(s.headers), n)
        assert_sound(self, s, "del-detail")

    def test_rename_id_rewrites_parent_cells_everywhere(self):
        s = two_hier()
        s.rename("Root", "Base")
        self.assertNotIn("root", s.nodes)
        for h in s.hiers:
            for iid, node in s.nodes.items():
                if node.ps[h] == "base":
                    self.assertEqual(s.data[s.rns[iid]][h], "Base")
        assert_sound(self, s, "rename-root")
