# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import csv
import os
import tempfile
import unittest

from src.session import Session, compare_files
from tests.fixtures import ANIMALS_TABLE, FOREST_TABLE, TWO_HIER_TABLE


def _animals():
    s = Session()
    out = s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
    assert out["ok"], out
    return s


def _forest():
    s = Session()
    out = s.load_table(FOREST_TABLE, id_col=0, parent_cols=[1], fmt=0)
    assert out["ok"], out
    return s


def _two_hier():
    s = Session()
    out = s.load_table(TWO_HIER_TABLE, id_col=0, parent_cols=[1, 2], fmt=0)
    assert out["ok"], out
    return s


def _names(rows):
    return [row["id"] for row in rows]


def _write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(rows)


class TestInspect(unittest.TestCase):
    def test_columns_list_and_resolve(self):
        s = _animals()
        cols = s.columns_list()["result"]["columns"]
        self.assertEqual([c["role"] for c in cols], ["id", "parent", "detail"])
        self.assertEqual(s.resolve_column("Name")["result"]["index"], 2)
        self.assertEqual(s.resolve_column("C")["result"]["index"], 2)
        self.assertEqual(s.resolve_column("2")["result"]["index"], 2)
        self.assertEqual(s.resolve_hier("Parent")["result"]["index"], 1)
        self.assertEqual(s.resolve_hier("Name")["error"]["code"], "unknown_hierarchy")
        self.assertEqual(s.resolve_id("Lion")["result"]["id"], "Lion")
        self.assertEqual(s.resolve_id("Nope")["error"]["code"], "id_not_found")

    def test_get_and_tree_and_find(self):
        s = _animals()
        one = s.get_ids(["Lion"])
        self.assertTrue(one["ok"], one)
        self.assertEqual(one["result"]["id"], "Lion")
        self.assertEqual(one["result"]["parents"]["Parent"], "Cats")
        self.assertEqual(one["result"]["details"]["Name"], "Lion")
        missing = s.get_ids(["Nope", "Also"])
        self.assertEqual(missing["error"]["code"], "id_not_found")
        multi = s.get_ids(["Lion", "Nope"])
        self.assertTrue(multi["ok"], multi)
        self.assertEqual(multi["result"]["not_found"], ["Nope"])
        tree = s.tree_dump(under="Animals", depth=1)
        self.assertTrue(tree["ok"], tree)
        self.assertEqual(tree["result"]["nodes"][0]["id"], "Animals")
        self.assertEqual([c["id"] for c in tree["result"]["nodes"][0]["children"]], ["Cats"])
        self.assertEqual(tree["result"]["nodes"][0]["children"][0]["children"], [])
        hits = s.find("lion")
        self.assertTrue(hits["ok"], hits)
        kinds = {(h["id"], h["type"]) for h in hits["result"]["hits"]}
        self.assertIn(("Lion", "id"), kinds)
        self.assertIn(("Lion", "detail"), kinds)
        lion_hits = [h for h in hits["result"]["hits"] if h["id"] == "Lion"]
        self.assertTrue(lion_hits)
        self.assertEqual(lion_hits[0]["level"], 3)

    def test_changelog_and_hier(self):
        s = _animals()
        s.add("Tiger", "Cats")
        log = s.changelog_list(limit=1)
        self.assertEqual(log["result"]["entries"][0]["type"], "Add ID")
        self.assertTrue(log["result"]["total"] >= 1)
        self.assertEqual(s.hier_get()["result"]["hierarchy"], "Parent")
        self.assertEqual(s.set_hier("Parent")["result"]["index"], 1)
        self.assertEqual(s.warnings_list()["result"]["warnings"], [])


class TestGetQuery(unittest.TestCase):
    def test_level_index_matches_node_level_and_tree_order(self):
        s = _forest()
        by_level, levels = s._level_index()
        for iid, node in s.nodes.items():
            if node.ps[s.pc] is None:
                self.assertNotIn(iid, levels)
                continue
            self.assertEqual(levels[iid], s._node_level(iid))
        self.assertEqual([s.nodes[i].name for i in by_level[1]], ["Life", "Item2", "Item10", "Moss"])
        self.assertEqual([s.nodes[i].name for i in by_level[3]], ["Cats", "Dogs", "Aardvark", "Zebra", "Oak"])
        stopped, stopped_map = s._level_index(stop_level=2)
        self.assertNotIn(3, stopped)
        self.assertNotIn("lion", stopped_map)
        self.assertNotIn("cats", stopped_map)
        self.assertIn("animals", stopped_map)
        under_animals, _ = s._level_index(under="animals", stop_level=3)
        self.assertEqual([s.nodes[i].name for i in under_animals[3]], ["Cats", "Dogs", "Aardvark", "Zebra"])

    def test_get_query_level_and_contains(self):
        s = _animals()
        out = s.get_query(level=3)
        self.assertTrue(out["ok"], out)
        self.assertEqual(_names(out["result"]["ids"]), ["Lion"])
        self.assertEqual(out["result"]["ids"][0]["level"], 3)
        self.assertEqual(out["result"]["ids"][0]["details"]["Name"], "Lion")
        hit = s.get_query(level=3, contains="lion")
        self.assertEqual(_names(hit["result"]["ids"]), ["Lion"])
        miss = s.get_query(level=2, contains="lion")
        self.assertEqual(miss["result"]["ids"], [])
        miss_exact = s.get_query(level=3, contains="lio", exact=True)
        self.assertEqual(miss_exact["result"]["ids"], [])
        family = s.get_query(contains="family")
        self.assertEqual(_names(family["result"]["ids"]), ["Cats"])
        name_col = s.resolve_column("Name")["result"]["index"]
        col_hit = s.get_query(contains="lion", column=name_col)
        self.assertEqual(_names(col_hit["result"]["ids"]), ["Lion"])
        ids_only = s.get_query(level=3, ids_only=True)
        self.assertEqual(ids_only["result"]["ids"], [{"id": "Lion", "level": 3}])

    def test_get_query_under_and_named_filter(self):
        s = _forest()
        under = s.get_query(under="animals", level=3)
        self.assertEqual(_names(under["result"]["ids"]), ["Cats", "Dogs", "Aardvark", "Zebra"])
        listed = s.get_query(["Lion", "Oak", "Nope"], level=4)
        self.assertEqual(_names(listed["result"]["ids"]), ["Lion"])
        self.assertEqual(listed["result"]["not_found"], ["Nope"])
        limited = s.get_query(level=3, limit=2)
        self.assertEqual(limited["result"]["returned"], 2)
        self.assertEqual(limited["result"]["total"], 5)
        self.assertTrue(limited["result"]["truncated"])
        self.assertEqual(_names(limited["result"]["ids"]), ["Cats", "Dogs"])

    def test_get_query_hierarchy(self):
        s = _two_hier()
        h1 = s.resolve_hier("H1")["result"]["index"]
        h2 = s.resolve_hier("H2")["result"]["index"]
        at_h1 = s.get_query(level=3, hier=h1)
        self.assertEqual(_names(at_h1["result"]["ids"]), ["B"])
        at_h2 = s.get_query(level=2, hier=h2)
        self.assertEqual(_names(at_h2["result"]["ids"]), ["B", "C"])

    def test_find_level_under_column(self):
        s = _animals()
        hits = s.find("lion", in_="detail", level=3)
        self.assertTrue(hits["ok"], hits)
        self.assertEqual({h["id"] for h in hits["result"]["hits"]}, {"Lion"})
        self.assertEqual(hits["result"]["hits"][0]["level"], 3)
        none = s.find("lion", in_="detail", level=2)
        self.assertEqual(none["result"]["hits"], [])
        under = s.find("lion", in_="detail", under="animals")
        self.assertEqual({h["id"] for h in under["result"]["hits"]}, {"Lion"})
        name_col = s.resolve_column("Name")["result"]["index"]
        col = s.find("lion", column=name_col)
        self.assertTrue(all(h["column"] == "Name" for h in col["result"]["hits"]))
        blocked = s.find("lion", level=3, all_hier=True)
        self.assertEqual(blocked["error"]["code"], "usage")


class TestReplaceExportUndo(unittest.TestCase):
    def test_replace_mapping_detail_and_undo(self):
        s = _animals()
        out = s.replace_mapping({"cat family": "Felis"})
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["cells_changed"], 1)
        self.assertEqual(s.data[s.rns["cats"]][2], "Felis")
        self.assertEqual(s.changelog[-1].label, "Edit cell")
        undone = s.undo()
        self.assertTrue(undone["ok"], undone)
        self.assertEqual(undone["result"]["undone"], "ctrl x, v, del key")
        self.assertEqual(s.data[s.rns["cats"]][2], "Cat family")
        empty = s.replace_mapping({"zzzz": "nope"})
        self.assertEqual(empty["result"]["cells_changed"], 0)
        self.assertEqual(s.undo()["error"]["code"], "nothing_to_undo")

    def test_undo_add(self):
        s = _animals()
        s.add("Tiger", "Cats")
        self.assertIn("tiger", s.nodes)
        out = s.undo()
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["undone"], "add id")
        self.assertNotIn("tiger", s.nodes)

    def test_export_changelog_and_flat(self):
        s = _animals()
        s.add("Tiger", "Cats")
        with tempfile.TemporaryDirectory() as tmp:
            log_path = os.path.join(tmp, "changes.csv")
            out = s.export_changelog(log_path)
            self.assertTrue(out["ok"], out)
            self.assertEqual(out["result"]["rows"], len(s.changelog))
            with open(log_path, newline="", encoding="utf-8") as fh:
                rows = list(csv.reader(fh))
            self.assertEqual(len(rows[0]), 5)
            self.assertEqual(rows[-1][1], "Add ID")
            exists = s.export_changelog(log_path)
            self.assertEqual(exists["error"]["code"], "file_exists")
            flat_path = os.path.join(tmp, "flat.csv")
            flat = s.export_flat(flat_path, hier="Parent")
            self.assertTrue(flat["ok"], flat)
            self.assertGreaterEqual(flat["result"]["rows"], 2)
            self.assertGreaterEqual(flat["result"]["cols"], 1)


class TestMergeImportCompare(unittest.TestCase):
    def test_merge_adds_id(self):
        s = _animals()
        incoming = [
            ["ID", "Parent", "Name"],
            ["Plants", "", "All plants"],
        ]
        out = s.merge_from_rows(incoming, fmt=0, id_col=0, parent_cols=[1])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["ids_added"], 1)
        self.assertIn("plants", s.nodes)
        none = s.merge_from_rows(incoming, fmt=0, id_col=0, parent_cols=[1], add_ids=False)
        self.assertEqual(none["error"]["code"], "no_changes")

    def test_import_delete_applies(self):
        s = _animals()
        s.delete(["Lion"])
        row = list(s.changelog[-1].rows[0])
        self.assertEqual(row[1], "Delete ID")
        s2 = _animals()
        out = s2.import_changes([row])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["applied"], 1, out)
        self.assertEqual(out["result"]["failed"], 0, out)
        self.assertNotIn("lion", s2.nodes)
        self.assertEqual(s2.changelog[-1].origin, "import")
        self.assertEqual(s2.changelog[-1].kind, "Delete ID")

    def test_import_add_id(self):
        s = _animals()
        s.add("Tiger", "Cats")
        row = list(s.changelog[-1].rows[0])
        s.undo()
        self.assertNotIn("tiger", s.nodes)
        out = s.import_changes([row])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["applied"], 1)
        self.assertIsNone(out["result"]["rows"][0]["reason"])
        self.assertIn("tiger", s.nodes)
        again = s.import_changes([row])
        self.assertEqual(again["result"]["applied"], 0)
        self.assertEqual(again["result"]["failed"], 1)

    def test_compare_identical_and_diff(self):
        with tempfile.TemporaryDirectory() as tmp:
            a = os.path.join(tmp, "a.csv")
            b = os.path.join(tmp, "b.csv")
            _write_csv(a, ANIMALS_TABLE)
            _write_csv(b, ANIMALS_TABLE)
            same = compare_files(a, b, id_a=0, parents_a=[1], id_b=0, parents_b=[1])
            self.assertTrue(same["ok"], same)
            self.assertTrue(same["result"]["identical"], same)
            changed = [list(r) for r in ANIMALS_TABLE]
            changed[3][2] = "Panthera leo"
            _write_csv(b, changed)
            diff = compare_files(a, b, id_a=0, parents_a=[1], id_b=0, parents_b=[1])
            self.assertTrue(diff["ok"], diff)
            self.assertFalse(diff["result"]["identical"])
            self.assertEqual(diff["result"]["detail_diffs"][0]["id"], "Lion")
            missing = compare_files(a, os.path.join(tmp, "nope.csv"), id_a=0, parents_a=[1], id_b=0, parents_b=[1])
            self.assertEqual(missing["error"]["code"], "file_not_found")
            need = compare_files(a, b)
            self.assertEqual(need["error"]["code"], "need_columns")

    def test_import_rename_set_delete_and_add_col(self):
        s = _animals()
        s.rename("Lion", "Leo")
        rename_row = list(s.changelog[-1].rows[0])
        s.set_detail("Leo", 2, "renamed")
        set_row = list(s.changelog[-1].rows[0])
        s.add_col(len(s.headers), "NOTE", "Text")
        col_row = list(s.changelog[-1].rows[0])
        s.delete(["Leo"])
        del_row = list(s.changelog[-1].rows[0])
        s2 = _animals()
        out = s2.import_changes([rename_row, set_row, col_row, del_row])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["applied"], 4, out)
        self.assertEqual(out["result"]["failed"], 0, out)
        self.assertNotIn("lion", s2.nodes)
        self.assertNotIn("leo", s2.nodes)
        self.assertIn("NOTE", [h.name for h in s2.headers])

    def test_merge_no_overwrite_details_and_add_column(self):
        s = _animals()
        incoming = [
            ["ID", "Parent", "Name", "Extra"],
            ["Lion", "Cats", "CHANGED", "x"],
            ["Plants", "", "All plants", ""],
        ]
        blocked = s.merge_from_rows(
            incoming,
            fmt=0,
            id_col=0,
            parent_cols=[1],
            add_ids=False,
            add_dcols=False,
            overwrite_details=False,
            overwrite_parents=False,
        )
        self.assertEqual(blocked["error"]["code"], "no_changes")
        out = s.merge_from_rows(incoming, fmt=0, id_col=0, parent_cols=[1], overwrite_details=False)
        self.assertTrue(out["ok"], out)
        self.assertEqual(s.data[s.rns["lion"]][2], "Lion")
        self.assertIn("plants", s.nodes)
        self.assertIn("Extra", [h.name for h in s.headers])

    def test_export_flat_reverse_and_index(self):
        s = _animals()
        with tempfile.TemporaryDirectory() as tmp:
            plain = os.path.join(tmp, "plain.csv")
            rev = os.path.join(tmp, "rev.csv")
            indexed = os.path.join(tmp, "idx.csv")
            self.assertTrue(s.export_flat(plain, hier="Parent", overwrite=True)["ok"])
            self.assertTrue(s.export_flat(rev, hier="Parent", reverse=True, overwrite=True)["ok"])
            self.assertTrue(s.export_flat(indexed, hier="Parent", index=True, overwrite=True)["ok"])
            with open(plain, newline="", encoding="utf-8") as fh:
                p = list(csv.reader(fh))
            with open(rev, newline="", encoding="utf-8") as fh:
                r = list(csv.reader(fh))
            with open(indexed, newline="", encoding="utf-8") as fh:
                i = list(csv.reader(fh))
            self.assertNotEqual(p[0], r[0])
            self.assertEqual(i[0][0], "Index")
            self.assertGreaterEqual(len(p), 2)
