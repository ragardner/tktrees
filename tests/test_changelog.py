# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import json
import unittest

from src.changelog import (
    EMPTY,
    ORIGIN_IMPORT,
    ORIGIN_MERGE,
    ORIGIN_USER,
    Change,
    ChangeBuilder,
    ChangeRow,
    display_rows,
    flatten_change,
    flatten_changelog,
    load_changelog,
    regroup_tuples,
)
from src.constants import changelog_header
from src.functions import full_sheet_to_dict


def _finish(builder, at="2026/09/06"):
    return builder.finish(at)


class TestChangeRowSequence(unittest.TestCase):
    def test_getitem_iter_share_strings(self):
        what = "ID: Cats parent: Animals column #2 named: Parent"
        ch = _finish(ChangeBuilder().add("Delete ID", what, "", ""))
        row = ch.rows[0]
        self.assertIs(row.what, what)
        self.assertIs(row[2], what)
        self.assertIs(row.kind, row.display_type)
        self.assertIs(row[0], ch.at)
        self.assertEqual(len(row), 5)
        self.assertEqual(tuple(row), (ch.at, "Delete ID", what, "", ""))

    def test_popup_list_holds_row_objects(self):
        ch = _finish(ChangeBuilder().add("Add ID", "Lion", "", ""))
        data = display_rows([ch])
        self.assertIs(data[0], ch.rows[0])
        self.assertIs(data[0][2], ch.rows[0].what)

    def test_add_row_stays_on_same_change(self):
        what = "ID: Tiger column #3 named: Name with type: Text"
        ch = _finish(ChangeBuilder().add("Add ID", "Name: Tiger", "", ""))
        row = ch.add_row("Edit cell", what, "", "Panthera")
        self.assertEqual(len(ch.rows), 2)
        self.assertFalse(ch.has_summary)
        self.assertEqual(ch.n, 2)
        self.assertEqual(ch.label, "Add ID")
        self.assertEqual(row.display_type, "Edit cell")
        self.assertIs(row.change, ch)
        self.assertEqual(
            flatten_change(ch),
            [
                ("2026/09/06", "Add ID", "Name: Tiger", "", ""),
                ("2026/09/06", "Edit cell", what, "", "Panthera"),
            ],
        )
        data = display_rows([ch])
        self.assertEqual(len(data), 2)
        self.assertIs(data[1], row)


class TestBakeDisplayType(unittest.TestCase):
    def test_user_singular_no_pipe(self):
        ch = _finish(ChangeBuilder().add("Add ID", "Name: Lion", "", ""))
        self.assertEqual(ch.label, "Add ID")
        self.assertFalse(ch.has_summary)
        self.assertEqual(ch.n, 1)
        self.assertEqual(ch.rows[0].display_type, "Add ID")
        self.assertEqual(flatten_change(ch), [("2026/09/06", "Add ID", "Name: Lion", "", "")])

    def test_user_multi_delete_pipes_and_summary(self):
        b = ChangeBuilder()
        b.add("Delete ID", "ID: Cats", "", "")
        b.add("Delete ID", "ID: Dogs", "", "")
        b.summary("Delete 2 IDs")
        ch = _finish(b)
        self.assertTrue(ch.has_summary)
        self.assertEqual(ch.n, 2)
        self.assertEqual(ch.kind, "Delete ID")
        self.assertEqual(ch.rows[0].display_type, "Delete ID |")
        self.assertEqual(ch.rows[1].display_type, "Delete ID |")
        self.assertIs(ch.rows[2].display_type, ch.label)
        flat = flatten_change(ch)
        self.assertEqual(
            [r[1] for r in flat],
            ["Delete ID |", "Delete ID |", "Delete 2 IDs"],
        )

    def test_import_n1_still_has_summary(self):
        b = ChangeBuilder(origin=ORIGIN_IMPORT)
        b.add("Edit cell", "ID: Lion column #3", "old", "new")
        b.summary("Imported 1 changes from: ", "Unsuccessful: 0 Unnecessary: 0")
        ch = _finish(b)
        self.assertTrue(ch.has_summary)
        self.assertEqual(ch.n, 1)
        self.assertEqual(ch.rows[0].display_type, "Imported change | Edit cell")
        self.assertEqual(ch.rows[1].display_type, "Imported 1 changes from: ")
        self.assertEqual(ch.origin, ORIGIN_IMPORT)

    def test_merge_mixed_kinds(self):
        b = ChangeBuilder(origin=ORIGIN_MERGE)
        b.add("Add new detail column", "Column #4", "", "")
        b.add("Edit cell", "ID: Lion", "", "x")
        b.add("Add ID", "Name: Tiger", "", "")
        b.summary("Merged sheets making 3 changes", "With file: other.csv")
        ch = _finish(b)
        self.assertEqual(ch.kind, "Add new detail column")
        types = [r.display_type for r in ch.rows]
        self.assertEqual(
            types,
            [
                "Merge | Add new detail column",
                "Merge | Edit cell",
                "Merge | Add ID",
                "Merged sheets making 3 changes",
            ],
        )

    def test_force_singular_strips_pipe_intent(self):
        b = ChangeBuilder()
        b.add("Delete ID", "Lion", "", "")
        b.force_singular = True
        b.label = "Delete ID"
        ch = _finish(b)
        self.assertFalse(ch.has_summary)
        self.assertEqual(ch.rows[0].display_type, "Delete ID")
        self.assertEqual(ch.n, 1)

    def test_cli_batch_add_label(self):
        b = ChangeBuilder()
        b.add("Add ID", "Tiger", "", "")
        b.add("Add ID", "Lion", "", "")
        b.summary("Add 2 IDs")
        ch = _finish(b)
        self.assertEqual([r[1] for r in flatten_change(ch)], ["Add ID |", "Add ID |", "Add 2 IDs"])


class TestRegroupRoundTrip(unittest.TestCase):
    def test_round_trip_cases(self):
        cases = []
        cases.append(_finish(ChangeBuilder().add("Add ID", "Lion", "", "")))
        cases.append(_finish(ChangeBuilder().add("Edit cell", "ID: Lion", "a", "b")))
        b = ChangeBuilder()
        b.add("Delete ID", "Cats", "", "")
        b.add("Delete ID", "Dogs", "", "")
        b.summary("Delete 2 IDs")
        cases.append(_finish(b))
        b = ChangeBuilder()
        b.add("Edit cell", "c1", "o", "n")
        b.add("Edit cell", "c2", "o", "n")
        b.summary("Edit 2 cells")
        cases.append(_finish(b))
        b = ChangeBuilder(origin=ORIGIN_IMPORT)
        b.add("Edit cell", "x", "o", "n")
        b.summary("Imported 1 changes from: ")
        cases.append(_finish(b))
        b = ChangeBuilder(origin=ORIGIN_MERGE)
        b.add("Add new detail column", "c", "", "")
        b.add("Edit cell", "e", "", "v")
        b.add("Add ID", "id", "", "")
        b.summary("Merged sheets making 3 changes")
        cases.append(_finish(b))
        b = ChangeBuilder()
        b.add("Delete ID from all hierarchies", "Cats", "", "")
        b.add("Delete ID from all hierarchies", "Dogs", "", "")
        b.summary("Deleted 2 IDs from all hierarchies")
        cases.append(_finish(b))
        b = ChangeBuilder()
        b.add("Add ID", "A", "", "")
        b.add("Add ID", "B", "", "")
        b.summary("Add 2 IDs")
        cases.append(_finish(b))
        cases.append(
            _finish(ChangeBuilder().add("Pruned changelog", "From: a To: b", "", ""))
        )
        for ch in cases:
            flat, idx = flatten_changelog([ch])
            back = regroup_tuples(flat)
            self.assertEqual(back, [ch], msg=repr(ch))
            self.assertEqual(flatten_changelog(back)[0], flat)
            self.assertEqual(len(idx), len(flat))
            self.assertTrue(all(i == 0 for i in idx))

    def test_leftover_members_without_summary(self):
        rows = [
            ("2026/09/06", "Delete ID |", "Cats", "", ""),
            ("2026/09/06", "Delete ID |", "Dogs", "", ""),
        ]
        groups = regroup_tuples(rows)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].n, 2)
        self.assertFalse(groups[0].has_summary)

    def test_origin_change_splits_buf(self):
        rows = [
            ("2026/09/06", "Delete ID |", "Cats", "", ""),
            ("2026/09/06", "Imported change | Edit cell", "x", "o", "n"),
        ]
        groups = regroup_tuples(rows)
        self.assertEqual(len(groups), 2)
        self.assertEqual(groups[0].origin, ORIGIN_USER)
        self.assertEqual(groups[1].origin, ORIGIN_IMPORT)


class TestLoadChangelog(unittest.TestCase):
    def test_empty(self):
        self.assertEqual(load_changelog([]), [])
        self.assertEqual(load_changelog(None), [])

    def test_format2_and_corrupt(self):
        ch = _finish(ChangeBuilder().add("Add ID", "Lion", "", ""))
        loaded = load_changelog([ch.to_dict()])
        self.assertEqual(loaded, [ch])
        warnings = []
        self.assertEqual(load_changelog([{"rows": []}], warnings=warnings), [])
        self.assertTrue(warnings)
        warnings = []
        self.assertEqual(load_changelog([{"rows": [{"kind": "x"}]}], warnings=warnings), [])
        self.assertTrue(warnings)

    def test_old_five_tuples(self):
        raw = [["2026/09/06", "Add ID", "Lion", "", ""]]
        loaded = load_changelog(raw)
        self.assertEqual(loaded[0].label, "Add ID")
        self.assertEqual(list(loaded[0].rows[0])[1:], ["Add ID", "Lion", "", ""])

    def test_old_loader_len_gt_5_on_dict(self):
        ch = _finish(ChangeBuilder().add("Add ID", "Lion", "", ""))
        blob = ch.to_dict()
        self.assertGreater(len(blob), 5)

    def test_eq_ignores_backpointer(self):
        a = _finish(ChangeBuilder().add("Add ID", "Lion", "", ""))
        b = _finish(ChangeBuilder().add("Add ID", "Lion", "", ""))
        self.assertEqual(a, b)
        self.assertIsNot(a.rows[0].change, b.rows[0].change)


class TestChangelogJsonExport(unittest.TestCase):
    def test_all_json_formats_dump_display_rows(self):
        ch = _finish(ChangeBuilder().add("Add ID", "Name: Lion Parent: Cats", "", ""))
        rows = display_rows([ch])
        self.assertIsInstance(rows[0], ChangeRow)
        expected = list(rows[0])
        for fmt in (1, 2, 3, 4):
            payload = full_sheet_to_dict(
                changelog_header,
                rows,
                include_headers=True,
                format_=fmt,
            )
            text = json.dumps(payload, indent=4)
            self.assertIsInstance(text, str)
            back = json.loads(text)
            self.assertIn("records", back)
        fmt3 = full_sheet_to_dict(changelog_header, rows, include_headers=True, format_=3)
        self.assertEqual(fmt3["records"][0], list(changelog_header))
        self.assertEqual(fmt3["records"][1], expected)
        self.assertTrue(all(isinstance(cell, str) for cell in fmt3["records"][1]))


class TestEmptyShared(unittest.TestCase):
    def test_summary_empty_fields_are_empty_constant(self):
        b = ChangeBuilder()
        b.add("Delete ID", "Cats", "", "")
        b.add("Delete ID", "Dogs", "", "")
        b.summary("Delete 2 IDs")
        ch = _finish(b)
        self.assertIs(ch.rows[2].what, EMPTY)
        self.assertIs(ch.rows[2].old, EMPTY)
        self.assertIs(ch.rows[2].new, EMPTY)


if __name__ == "__main__":
    unittest.main()
