# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import unittest

from src.session import Session
from tests.fixtures import ANIMALS_TABLE, FOREST_TABLE


def _load(rows):
    s = Session()
    out = s.load_table(rows, id_col=0, parent_cols=[1], fmt=0)
    assert out["ok"], out
    return s


def _names(s, iids):
    return [s.nodes[i].name for i in iids]


def _tops(s):
    return _names(s, s.top_iids())


def _children(s, parent):
    return _names(s, s.nodes[parent.lower()].cn[s.pc])


def _walk_session(s):
    out = []
    for iid in s.top_iids():
        stack = [iid]
        while stack:
            cur = stack.pop()
            out.append(s.nodes[cur].name)
            stack.extend(reversed(s.nodes[cur].cn[s.pc]))
    return out


class TestAutoSortOrder(unittest.TestCase):
    def test_tops_are_branched_then_leaves_with_numeric_sort_key(self):
        s = _load(FOREST_TABLE)
        self.assertTrue(s.auto_sort_nodes_bool)
        # Branched tops first (Life), then leaf tops by sort_key: Item2, Item10, Moss.
        self.assertEqual(_tops(s), ["Life", "Item2", "Item10", "Moss"])
        self.assertEqual(
            s.topnodes_order[s.pc],
            ["life", "item2", "item10", "moss"],
        )

    def test_children_branched_then_leaves(self):
        s = _load(FOREST_TABLE)
        self.assertEqual(_children(s, "Animals"), ["Cats", "Dogs", "Aardvark", "Zebra"])
        self.assertEqual(_children(s, "Cats"), ["Lion", "Tiger"])
        self.assertEqual(_children(s, "Oak"), ["Ch2", "Ch3", "Ch10"])

    def test_tree_walk_parent_before_descendants(self):
        s = _load(FOREST_TABLE)
        walk = _walk_session(s)
        self.assertEqual(walk[0], "Life")
        self.assertLess(walk.index("Animals"), walk.index("Cats"))
        self.assertLess(walk.index("Cats"), walk.index("Lion"))
        self.assertLess(walk.index("Dogs"), walk.index("Wolf"))
        self.assertLess(walk.index("Plants"), walk.index("Oak"))
        self.assertLess(walk.index("Oak"), walk.index("Ch2"))
        self.assertLess(walk.index("Ch2"), walk.index("Ch10"))
        self.assertEqual(walk[-1], "Moss")

    def test_tree_dump_roots_match_top_iids(self):
        s = _load(FOREST_TABLE)
        dump = s.tree_dump(depth=None, force=True, details=False)
        self.assertTrue(dump["ok"], dump)
        roots = [n["id"] for n in dump["result"]["nodes"]]
        self.assertEqual(roots, ["Life", "Item2", "Item10", "Moss"])
        animals = dump["result"]["nodes"][0]["children"][0]
        self.assertEqual(animals["id"], "Animals")
        self.assertEqual([c["id"] for c in animals["children"]], ["Cats", "Dogs", "Aardvark", "Zebra"])

    def test_sort_sheet_walk_keeps_depth_order(self):
        s = _load(FOREST_TABLE)
        s.sort_sheet("Name", "DESCENDING")
        s.sort_sheet_walk()
        names = [row[0] for row in s.data]
        self.assertLess(names.index("Life"), names.index("Animals"))
        self.assertLess(names.index("Animals"), names.index("Zebra"))
        self.assertLess(names.index("Cats"), names.index("Tiger"))
        self.assertLess(names.index("Oak"), names.index("Ch10"))

    def test_add_with_auto_sort_on_reinserts_in_sorted_cn(self):
        s = _load(FOREST_TABLE)
        s.add("Puma", "Cats")
        self.assertEqual(_children(s, "Cats"), ["Lion", "Puma", "Tiger"])
        s.add("Bear", "Animals")
        # Bear is a leaf; branched Cats/Dogs still come first.
        self.assertEqual(_children(s, "Animals"), ["Cats", "Dogs", "Aardvark", "Bear", "Zebra"])


class TestTopnodesOrderManual(unittest.TestCase):
    def test_turning_auto_sort_off_snapshots_sorted_tops(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        self.assertFalse(s.auto_sort_nodes_bool)
        self.assertEqual(
            list(s.topnodes_order[s.pc]),
            ["life", "item2", "item10", "moss"],
        )
        self.assertEqual(_tops(s), ["Life", "Item2", "Item10", "Moss"])

    def test_add_top_appends_when_auto_sort_off(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        s.add("Zoo", "")
        self.assertEqual(
            list(s.topnodes_order[s.pc]),
            ["life", "item2", "item10", "moss", "zoo"],
        )
        self.assertEqual(_tops(s), ["Life", "Item2", "Item10", "Moss", "Zoo"])
        s.add("Ant", "Animals")
        self.assertEqual(_children(s, "Animals")[-1], "Ant")
        self.assertNotIn("ant", s.topnodes_order[s.pc])

    def test_before_after_splice_tops_and_children(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        s.add("Mid", "", before="Item10")
        self.assertEqual(
            list(s.topnodes_order[s.pc]),
            ["life", "item2", "mid", "item10", "moss"],
        )
        s.add("Lynx", "Cats", after="Lion")
        self.assertEqual(_children(s, "Cats"), ["Lion", "Lynx", "Tiger"])
        mixed = s.add("Nope", "Cats", before="Moss")
        self.assertEqual(mixed["error"]["code"], "usage")
        missing = s.add("Ghost", "", before="Missing")
        self.assertEqual(missing["error"]["code"], "id_not_found")

    def test_before_rejected_when_auto_sort_on(self):
        s = _load(FOREST_TABLE)
        out = s.add("Mid", "", before="Moss")
        self.assertEqual(out["error"]["code"], "auto_sort_on")

    def test_delete_top_promotes_children_onto_topnodes_order(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        before = list(s.topnodes_order[s.pc])
        s.delete(["Life"])
        self.assertNotIn("life", s.topnodes_order[s.pc])
        self.assertEqual(s.nodes["animals"].ps[s.pc], "")
        self.assertEqual(s.nodes["plants"].ps[s.pc], "")
        # Promoted children appended in cn order (Animals, Plants).
        self.assertEqual(
            list(s.topnodes_order[s.pc]),
            [k for k in before if k != "life"] + ["animals", "plants"],
        )
        self.assertEqual(_tops(s), ["Item2", "Item10", "Moss", "Animals", "Plants"])

    def test_delete_top_with_auto_sort_on_resorts_new_tops(self):
        s = _load(FOREST_TABLE)
        s.delete(["Life"])
        # Animals/Plants now branched tops; Item2/Item10/Moss still leaf tops.
        self.assertEqual(_tops(s), ["Animals", "Plants", "Item2", "Item10", "Moss"])

    def test_rename_rewrites_topnodes_order_key(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        s.rename("Item2", "Item20")
        self.assertNotIn("item2", s.topnodes_order[s.pc])
        self.assertIn("item20", s.topnodes_order[s.pc])
        self.assertEqual(s.topnodes_order[s.pc].index("item20"), 1)

    def test_undo_add_top_restores_topnodes_order(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        before = list(s.topnodes_order[s.pc])
        s.add("Zoo", "")
        self.assertIn("zoo", s.topnodes_order[s.pc])
        undone = s.undo()
        self.assertTrue(undone["ok"], undone)
        self.assertEqual(list(s.topnodes_order[s.pc]), before)
        self.assertNotIn("zoo", s.nodes)

    def test_turning_auto_sort_on_sorts_insertion_order_children(self):
        s = _load(ANIMALS_TABLE)
        s.set_option("auto-sort", "off")
        s.add("Zebra", "Animals")
        s.add("Aardvark", "Animals")
        self.assertEqual(_children(s, "Animals"), ["Cats", "Zebra", "Aardvark"])
        s.set_option("auto-sort", "on")
        self.assertEqual(_children(s, "Animals"), ["Cats", "Aardvark", "Zebra"])

    def test_move_to_top_appends_when_auto_sort_off(self):
        s = _load(FOREST_TABLE)
        s.set_option("auto-sort", "off")
        s.move("Lion", top=True)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertEqual(s.topnodes_order[s.pc][-1], "lion")
        self.assertEqual(_tops(s)[-1], "Lion")
        self.assertNotIn("lion", s.nodes["cats"].cn[s.pc])

    def test_copy_into_new_hierarchy_leaves_source_cn_alone(self):
        s = _load(FOREST_TABLE)
        before = list(s.nodes["cats"].cn[s.pc])
        col = len(s.headers)
        s.add_hier_col(col, "PARENT_2")
        s.copy("Lion", top=True, hier=col)
        self.assertEqual(s.nodes["cats"].cn[s.pc], before)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "cats")
        self.assertEqual(s.nodes["lion"].ps[col], "")
        s.pc = col
        self.assertIn("lion", list(s.top_iids()))
