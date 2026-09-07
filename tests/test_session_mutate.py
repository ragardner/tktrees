# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import unittest

from src.session import Session
from tests.fixtures import ANIMALS_TABLE


def _animals():
    s = Session()
    out = s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
    assert out["ok"], out
    return s


class TestAddRenameSet(unittest.TestCase):
    def test_add_child_and_changelog(self):
        s = _animals()
        out = s.add("Tiger", "Cats")
        self.assertTrue(out["ok"], out)
        self.assertTrue(out["result"]["added"][0]["created_row"])
        self.assertEqual(s.nodes["tiger"].ps[s.pc], "cats")
        self.assertIn("tiger", s.nodes["cats"].cn[s.pc])
        self.assertEqual(s.changelog[-1].label, "Add ID")
        self.assertIn("Name: Tiger Parent: Cats column #2 named: Parent", s.changelog[-1].rows[0].what)

    def test_add_label_row_undoes_with_add(self):
        s = _animals()
        n_log = len(s.changelog)
        n_vs = len(s.vs)
        self.assertTrue(s.add("Tiger", "Cats")["ok"])
        rn = s.rns["tiger"]
        old = f"{s.data[rn][2]}"
        s.data[rn][2] = "Panthera tigris"
        s.changelog[-1].add_row(
            "Edit cell",
            "ID: Tiger column #3 named: Name with type: Text",
            old,
            "Panthera tigris",
        )
        self.assertEqual(len(s.changelog), n_log + 1)
        self.assertEqual(len(s.vs), n_vs + 1)
        self.assertEqual([r.display_type for r in s.changelog[-1].rows], ["Add ID", "Edit cell"])
        out = s.undo()
        self.assertTrue(out["ok"], out)
        self.assertNotIn("tiger", s.nodes)
        self.assertEqual(len(s.changelog), n_log)
        self.assertEqual(len(s.vs), n_vs)

    def test_add_top_uses_na_parent_text(self):
        s = _animals()
        out = s.add("Plants", "")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["plants"].ps[s.pc], "")
        self.assertIn("n/a - Top ID", s.changelog[-1].rows[0].what)

    def test_add_already_in_hierarchy(self):
        s = _animals()
        out = s.add("Lion", "Cats")
        self.assertFalse(out["ok"])
        self.assertEqual(out["error"]["code"], "already_in_hierarchy")
        self.assertEqual(out["error"]["message"], "ID already in hierarchy   ")

    def test_add_empty_id(self):
        s = _animals()
        out = s.add("", "Cats")
        self.assertEqual(out["error"]["code"], "empty_id")

    def test_rename_and_parent_cells(self):
        s = _animals()
        out = s.rename("Cats", "Felines")
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["old"], "Cats")
        self.assertEqual(out["result"]["new"], "Felines")
        self.assertIn("felines", s.nodes)
        self.assertNotIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "felines")
        self.assertEqual(s.data[s.rns["lion"]][s.pc], "Felines")
        self.assertEqual(s.changelog[-1].label, "Rename ID")
        self.assertEqual(s.changelog[-1].rows[0].what, "Cats")
        self.assertEqual(s.changelog[-1].rows[0].new, "Felines")

    def test_rename_missing_and_clash(self):
        s = _animals()
        self.assertEqual(s.rename("Nope", "X")["error"]["code"], "id_not_found")
        self.assertEqual(s.rename("Lion", "Cats")["error"]["code"], "id_exists")
        self.assertEqual(s.rename("Lion", "")["error"]["code"], "empty_id")

    def test_set_detail(self):
        s = _animals()
        out = s.set_detail("Lion", 2, "Panthera leo")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.data[s.rns["lion"]][2], "Panthera leo")
        self.assertEqual(s.changelog[-1].label, "Edit cell")
        self.assertEqual(out["result"]["old"], "Lion")
        self.assertEqual(out["result"]["new"], "Panthera leo")
        bad = s.set_detail("Lion", 0, "X")
        self.assertEqual(bad["error"]["code"], "usage")


class TestDelete(unittest.TestCase):
    def test_delete_promotes_children(self):
        s = _animals()
        out = s.delete(["Cats"])
        self.assertTrue(out["ok"], out)
        self.assertNotIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        self.assertIn("Lion", out["result"]["promoted"])
        self.assertIn("Cats", out["result"]["removed_entirely"])
        self.assertEqual(s.changelog[-1].label, "Delete ID")
        self.assertIn("ID: Cats", s.changelog[-1].rows[0].what)

    def test_delete_two_ids_is_one_change_and_one_undo(self):
        s = _animals()
        s.add("Tiger", "Cats")
        n_before = len(s.changelog)
        out = s.delete(["Lion", "Tiger"])
        self.assertTrue(out["ok"], out)
        self.assertEqual(len(s.changelog), n_before + 1)
        ch = s.changelog[-1]
        self.assertEqual(ch.label, "Delete 2 IDs")
        self.assertEqual(ch.n, 2)
        self.assertTrue(ch.has_summary)
        self.assertEqual(ch.rows[0].display_type, "Delete ID |")
        self.assertEqual(ch.rows[-1].display_type, "Delete 2 IDs")
        self.assertNotIn("lion", s.nodes)
        self.assertNotIn("tiger", s.nodes)
        undone = s.undo()
        self.assertTrue(undone["ok"], undone)
        self.assertIn("lion", s.nodes)
        self.assertIn("tiger", s.nodes)
        self.assertEqual(len(s.changelog), n_before)

    def test_delete_missing_is_ok(self):
        s = _animals()
        out = s.delete(["Nope"])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["not_found"], ["Nope"])
        self.assertIn("lion", s.nodes)

    def test_delete_children(self):
        s = _animals()
        s.add("Cub", "Lion")
        out = s.delete(["Lion"], children=True)
        self.assertTrue(out["ok"], out)
        self.assertNotIn("lion", s.nodes)
        self.assertNotIn("cub", s.nodes)
        self.assertIn("Cub", out["result"]["deleted_descendants"])

    def test_delete_orphan(self):
        s = _animals()
        out = s.delete(["Cats"], orphan=True)
        self.assertTrue(out["ok"], out)
        self.assertNotIn("cats", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertIn("Lion", out["result"]["orphaned"])
        self.assertEqual(s.changelog[-1].label, "Delete ID, orphan children")

    def test_delete_children_and_orphan_is_usage(self):
        s = _animals()
        out = s.delete(["Cats"], children=True, orphan=True)
        self.assertEqual(out["error"]["code"], "usage")
        self.assertIn("cats", s.nodes)

    def test_delete_orphan_multiple_ids_keeps_rns(self):
        s = _animals()
        out = s.delete(["Animals", "Cats"], orphan=True)
        self.assertTrue(out["ok"], out)
        self.assertNotIn("animals", s.nodes)
        self.assertNotIn("cats", s.nodes)
        self.assertIn("lion", s.nodes)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertEqual([row[0] for row in s.data], ["Lion"])
        self.assertEqual(s.rns["lion"], 0)


class TestMoveCopyColumns(unittest.TestCase):
    def test_move_reparents_and_changelog(self):
        s = _animals()
        out = s.move("Lion", parent="Animals")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        self.assertNotIn("lion", s.nodes["cats"].cn[s.pc])
        self.assertIn("lion", s.nodes["animals"].cn[s.pc])
        self.assertEqual(s.data[s.rns["lion"]][s.pc], "Animals")
        self.assertEqual(s.changelog[-1].label, "Cut and paste ID")
        self.assertIn("Old parent: Cats", s.changelog[-1].rows[0].old)
        self.assertIn("New parent: Animals", s.changelog[-1].rows[0].new)

    def test_move_to_top(self):
        s = _animals()
        out = s.move("Lion", top=True)
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "")
        self.assertEqual(s.data[s.rns["lion"]][s.pc], "")

    def test_cut_paste_children(self):
        s = _animals()
        out = s.cut_paste_children("Cats", "Animals", s.pc)
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["lion"].ps[s.pc], "animals")
        self.assertIn("lion", s.nodes["animals"].cn[s.pc])
        self.assertEqual(s.nodes["cats"].cn[s.pc], [])
        empty = s.cut_paste_children("Cats", "Animals", s.pc)
        self.assertFalse(empty["ok"])

    def test_cut_paste_children_snapshot_is_undoable(self):
        s = _animals()
        self.assertTrue(s.cut_paste_children("Cats", "Animals", s.pc)["ok"])
        self.assertTrue(s.vs)
        self.assertEqual(s.vs[-1]["type"], "paste id")
        self.assertTrue(s.undo()["ok"])
        self.assertEqual(s.nodes["lion"].ps[s.pc], "cats")
        self.assertEqual(s.data[s.rns["lion"]][s.pc], "Cats")

    def test_move_already_has_parent(self):
        s = _animals()
        out = s.move("Lion", parent="Cats")
        self.assertFalse(out["ok"])
        self.assertEqual(out["error"]["code"], "already_in_hierarchy")
        self.assertEqual(out["error"]["message"], "ID: Lion already has this parent   ")

    def test_copy_into_new_hierarchy(self):
        s = _animals()
        col = len(s.headers)
        s.add_hier_col(col, "PARENT_2")
        dest = s.hiers[-1]
        out = s.copy("Lion", top=True, hier=dest)
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.nodes["lion"].ps[dest], "")
        self.assertEqual(s.nodes["lion"].ps[1], "cats")
        self.assertEqual(s.changelog[-1].label, "Copy and paste ID")
        self.assertIn("From column #", s.changelog[-1].rows[0].old)
        self.assertIn(s.headers[s.hiers[0]].name, s.changelog[-1].rows[0].old)

    def test_add_detail_column_changelog_is_zero_based(self):
        s = _animals()
        col = len(s.headers)
        out = s.add_col(col, "Habitat", "Text")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.headers[col].name, "Habitat")
        self.assertEqual(len(s.data[0]), 4)
        self.assertEqual(s.data[s.rns["lion"]][col], "")
        self.assertEqual(s.changelog[-1].label, "Add new detail column")
        self.assertEqual(
            s.changelog[-1].rows[0].what,
            f"Column #{col} with name: Habitat and type: Text",
        )

    def test_add_hierarchy_column_changelog_is_one_based(self):
        s = _animals()
        col = len(s.headers)
        out = s.add_hier_col(col, "PARENT_2")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.headers[col].type_, "Parent")
        self.assertIn(col, s.hiers)
        self.assertIsNone(s.nodes["lion"].ps[col])
        self.assertEqual(s.changelog[-1].label, "Add new hierarchy column")
        self.assertEqual(s.changelog[-1].rows[0].what, f"Column #{col + 1} with name: PARENT_2")

    def test_rename_and_delete_detail_column(self):
        s = _animals()
        s.add_col(3, "Habitat", "Text")
        out = s.rename_col(3, "Biome")
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.headers[3].name, "Biome")
        self.assertEqual(s.changelog[-1].label, "Column rename")
        gone = s.del_cols([3])
        self.assertTrue(gone["ok"], gone)
        self.assertEqual([h.name for h in s.headers], ["ID", "Parent", "Name"])
        self.assertEqual(len(s.data[0]), 3)
        self.assertEqual(s.changelog[-1].label, "Delete columns")


class TestSortTagOptionValidation(unittest.TestCase):
    def test_sort_sheet_by_name(self):
        s = _animals()
        out = s.sort_sheet("Name", "DESCENDING")
        self.assertTrue(out["ok"], out)
        names = [row[0] for row in s.data]
        self.assertEqual(names[0], "Lion")
        self.assertEqual(s.changelog[-1].label, "Sort sheet")
        self.assertIn("DESCENDING", s.changelog[-1].rows[0].what)

    def test_sort_tree_walk_keeps_parent_before_child(self):
        s = _animals()
        s.sort_sheet("Name", "DESCENDING")
        out = s.sort_sheet_walk()
        self.assertTrue(out["ok"], out)
        names = [row[0] for row in s.data]
        self.assertEqual(names.index("Animals"), 0)
        self.assertLess(names.index("Cats"), names.index("Lion"))
        self.assertEqual(s.changelog[-1].rows[0].what, "Sorted sheet in tree walk order")

    def test_sort_children(self):
        s = _animals()
        s.set_option("auto-sort", "off")
        s.add("Zebra", "Animals")
        s.add("Aardvark", "Animals")
        out = s.sort_children("Animals")
        self.assertTrue(out["ok"], out)
        # Nodes with children sort before leaves, matching sort_node_cn.
        self.assertEqual(
            [s.nodes[ik].name for ik in s.nodes["animals"].cn[s.pc]],
            ["Cats", "Aardvark", "Zebra"],
        )
        s.add("Cub", "Lion")
        s.add("Kitten", "Lion")
        s.sort_children("Lion")
        self.assertEqual(
            [s.nodes[ik].name for ik in s.nodes["lion"].cn[s.pc]],
            ["Cub", "Kitten"],
        )

    def test_tag_untag_clear(self):
        s = _animals()
        out = s.tag(["Lion", "Nope", "Lion"])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["tagged"], ["Lion"])
        self.assertEqual(out["result"]["already"], ["Lion"])
        self.assertEqual(out["result"]["not_found"], ["Nope"])
        self.assertEqual(s.tags_list()["result"]["ids"], ["Lion"])
        un = s.untag(["Lion", "Cats"])
        self.assertEqual(un["result"]["untagged"], ["Lion"])
        self.assertEqual(un["result"]["not_tagged"], ["Cats"])
        s.tag(["Lion", "Cats"])
        cleared = s.clear_tags()
        self.assertEqual(cleared["result"]["cleared"], 2)
        self.assertEqual(s.tagged_ids, set())

    def test_tag_from_terms_substring(self):
        s = _animals()
        out = s.tag_from_terms(["lion", "family"], in_="any", exact=False)
        self.assertTrue(out["ok"], out)
        self.assertIn("Lion", out["result"]["ids_tagged"])
        self.assertIn("Cats", out["result"]["ids_tagged"])

    def test_option_auto_sort_and_unknown(self):
        s = _animals()
        out = s.set_option("auto-sort", "off")
        self.assertTrue(out["ok"], out)
        self.assertFalse(out["result"]["auto-sort"])
        self.assertIn(s.pc, s.topnodes_order)
        bad = s.set_option("nope", "on")
        self.assertEqual(bad["error"]["code"], "unknown_option")
        s.set_option("allow-spaces-ids", True)
        self.assertTrue(s.allow_spaces_ids_var)
        s.set_option("save-program-data", "off")
        self.assertFalse(s.save_xlsx_with_program_data)
        self.assertFalse(s.save_json_with_program_data)

    def test_validation_prepends_empty_and_clears_bad_cells(self):
        s = _animals()
        s.set_detail("Lion", 2, "Lion")
        out = s.set_validation(2, ["a", "b"])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["validation"], ["", "a", "b"])
        self.assertEqual(s.data[s.rns["lion"]][2], "")
        self.assertEqual(s.changelog[-1].label, "Edit validation")
        parent_col = s.set_validation(s.pc, ["x"])
        self.assertEqual(parent_col["error"]["code"], "validation")
        cleared = s.set_validation(2, [])
        self.assertEqual(cleared["result"]["validation"], [])

