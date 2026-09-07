# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import csv
import io
import json
import os
import tempfile
import unittest

from src.cli import COMMAND_HELP, HELP_LINES, run_cli
from tests.fixtures import ANIMALS_TABLE

ENVELOPE_KEYS = {"ok", "command", "error", "warnings", "result", "status"}
STATUS_KEYS = {"file", "unsaved", "hierarchy", "ids", "undo"}

RESULT_KEYS = {
    "open": {"ids", "format", "id_column", "columns", "hierarchies", "warnings"},
    "new": {"ids", "format", "id_column", "columns", "hierarchies", "warnings"},
    "save": {"file", "kind"},
    "save-as": {"file", "kind"},
    "close": set(),
    "quit": set(),
    "status": {
        "file",
        "unsaved",
        "hierarchy",
        "ids",
        "undo",
        "warnings_count",
        "columns",
        "hierarchies",
        "options",
    },
    "help": {"help"},
    "format": {"format"},
    "tree": {"hierarchy", "depth", "under", "truncated", "returned", "total", "nodes"},
    "get": {"id", "parents", "children", "details", "tagged"},
    "find": {"hits", "truncated", "returned", "total"},
    "columns": {"columns"},
    "changelog": {"entries", "truncated", "returned", "total"},
    "warnings": {"warnings"},
    "hier": {"hierarchy", "index", "letter"},
    "add": {"added"},
    "rename": {"old", "new"},
    "set": {"id", "column", "old", "new"},
    "move": {"moved"},
    "copy": {"copied"},
    "delete": {
        "requested",
        "removed_entirely",
        "removed_from_hierarchy",
        "promoted",
        "deleted_descendants",
        "orphaned",
        "not_in_hierarchy",
        "not_found",
    },
    "column add": {"name", "index", "letter", "role", "type"},
    "column rename": {"old", "new", "index"},
    "column delete": {"deleted", "not_found"},
    "column validate": {"column", "index", "validation"},
    "option": {"allow-spaces-ids", "allow-spaces-columns", "auto-sort", "save-program-data"},
    "option set": {"allow-spaces-ids", "allow-spaces-columns", "auto-sort", "save-program-data"},
    "tag": {"tagged", "already", "not_found"},
    "untag": {"untagged", "not_tagged", "not_found"},
    "tags": {"ids"},
    "sort": {"kind", "column", "descending"},
    "undo": {"undone", "undo"},
    "merge": {
        "ids_added",
        "details_written",
        "parents_written",
        "detail_columns_added",
        "parent_columns_added",
    },
    "replace": {"file", "cells_changed"},
    "import-changes": {"applied", "unnecessary", "failed", "rows"},
    "export-changes": {"file", "rows"},
    "export-flat": {"file", "rows", "cols"},
    "compare": {
        "identical",
        "warnings_a",
        "warnings_b",
        "headers",
        "ids_only_a",
        "ids_only_b",
        "parent_diffs",
        "detail_diffs",
    },
}


def _run(args, check=True):
    out = io.StringIO()
    argv = ["TKTREES.pyw", "cli", "--json", *args]
    try:
        run_cli(argv, outfile=out)
        code = 0
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
    text = out.getvalue()
    envs = [json.loads(line) for line in text.splitlines() if line]
    if check:
        assert code == 0, (code, envs)
    return code, envs


def _write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(rows)


def _open_animals(tmp):
    path = os.path.join(tmp, "animals.csv")
    _write_csv(path, ANIMALS_TABLE)
    quoted = json.dumps(path)
    return quoted, path


def _assert_envelope(test, env, command, ok=True):
    test.assertEqual(set(env), ENVELOPE_KEYS)
    test.assertEqual(set(env["status"]), STATUS_KEYS)
    test.assertEqual(env["command"], command)
    test.assertIsInstance(env["warnings"], list)
    test.assertIsInstance(env["result"], dict)
    if ok:
        test.assertTrue(env["ok"])
        test.assertIsNone(env["error"])
        test.assertTrue(
            RESULT_KEYS[command] <= set(env["result"]),
            msg=f"{command}: missing {RESULT_KEYS[command] - set(env['result'])}",
        )
    else:
        test.assertFalse(env["ok"])
        test.assertIsInstance(env["error"], dict)
        test.assertIn("code", env["error"])
        test.assertEqual(env["result"], {})


class TestHelp(unittest.TestCase):
    def test_catalog_and_command_help(self):
        code, envs = _run(["-c", "help"])
        self.assertEqual(code, 0)
        _assert_envelope(self, envs[0], "help")
        for line in HELP_LINES:
            self.assertIn(line.split(" —")[0], envs[0]["result"]["help"])
        code, envs = _run(["-c", "help add"])
        text = envs[0]["result"]["help"]
        self.assertIn("syntax:", text)
        self.assertIn("flags:", text)
        self.assertIn("example:", text)
        self.assertIn("result:", text)
        self.assertIn("GUI:", text)
        self.assertIn("errors:", text)
        self.assertEqual(set(COMMAND_HELP), set(RESULT_KEYS) | {"column", "exit"})


class TestEnvelopeCatalog(unittest.TestCase):
    def test_new_status_add_get_tree_columns_hier(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "status",
                "-c",
                "add Animals",
                "-c",
                "add Cats --parent Animals",
                "-c",
                "get Cats",
                "-c",
                "tree --under Animals --depth 1",
                "-c",
                "columns",
                "-c",
                "hier",
                "-c",
                "hier --list",
                "-c",
                "warnings",
                "-c",
                "changelog --limit 0",
                "-c",
                "option",
                "-c",
                "format json",
            ]
        )
        self.assertEqual(code, 0)
        _assert_envelope(self, envs[0], "new")
        _assert_envelope(self, envs[1], "status")
        _assert_envelope(self, envs[2], "add")
        _assert_envelope(self, envs[4], "get")
        _assert_envelope(self, envs[5], "tree")
        _assert_envelope(self, envs[6], "columns")
        _assert_envelope(self, envs[7], "hier")
        self.assertEqual(set(envs[8]["result"]), {"hierarchies", "current"})
        _assert_envelope(self, envs[9], "warnings")
        _assert_envelope(self, envs[10], "changelog")
        _assert_envelope(self, envs[11], "option")
        _assert_envelope(self, envs[12], "format")
        self.assertEqual(envs[0]["result"]["ids"], 0)
        self.assertEqual(envs[4]["result"]["parents"]["PARENT_1"], "Animals")
        self.assertEqual(envs[5]["result"]["under"], "Animals")

    def test_mutate_sort_tag_undo_set(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "add Animals",
                "-c",
                "add Cats --parent Animals",
                "-c",
                "add Lion --parent Cats",
                "-c",
                "set Lion DETAIL_1 big",
                "-c",
                "rename Lion Cub",
                "-c",
                "tag Cub",
                "-c",
                "tags",
                "-c",
                "untag Cub",
                "-c",
                "sort --column DETAIL_1 --desc",
                "-c",
                "undo",
                "-c",
                "find big",
                "-c",
                "option set auto-sort off",
                "-c",
                "column add Habitat",
                "-c",
                "column rename Habitat Biome",
                "-c",
                "column validate Biome --set a,b",
                "-c",
                "column delete Biome",
                "-c",
                "delete Cub --orphan",
            ]
        )
        self.assertEqual(code, 0)
        _assert_envelope(self, envs[4], "set")
        _assert_envelope(self, envs[5], "rename")
        _assert_envelope(self, envs[6], "tag")
        _assert_envelope(self, envs[7], "tags")
        _assert_envelope(self, envs[8], "untag")
        _assert_envelope(self, envs[9], "sort")
        _assert_envelope(self, envs[10], "undo")
        _assert_envelope(self, envs[11], "find")
        _assert_envelope(self, envs[12], "option set")
        _assert_envelope(self, envs[13], "column add")
        _assert_envelope(self, envs[14], "column rename")
        _assert_envelope(self, envs[15], "column validate")
        _assert_envelope(self, envs[16], "column delete")
        _assert_envelope(self, envs[17], "delete")
        self.assertEqual(envs[5]["result"]["new"], "Cub")
        self.assertEqual(envs[13]["result"]["role"], "detail")
        self.assertEqual(envs[15]["result"]["validation"][0], "")
        self.assertEqual(envs[16]["result"]["deleted"][0]["name"], "Biome")

    def test_move_copy_and_failure_result_empty(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "add Animals",
                "-c",
                "add Cats --parent Animals",
                "-c",
                "column add PARENT_2 --hier",
                "-c",
                "copy Cats --top --hier PARENT_2",
                "-c",
                "move Cats --parent Animals",
            ],
            check=False,
        )
        self.assertEqual(code, 1)
        _assert_envelope(self, envs[4], "copy")
        self.assertEqual(envs[4]["result"]["copied"][0]["to_parent"], "")
        _assert_envelope(self, envs[5], "move", ok=False)
        self.assertEqual(envs[5]["error"]["code"], "already_in_hierarchy")

    def test_open_save_merge_replace_import_export_compare(self):
        with tempfile.TemporaryDirectory() as tmp:
            quoted, _path = _open_animals(tmp)
            extra = os.path.join(tmp, "extra.csv")
            _write_csv(extra, [["ID", "Parent", "Name"], ["Plants", "", "All plants"]])
            mapping = os.path.join(tmp, "map.csv")
            _write_csv(mapping, [["cat family", "Felis"]])
            out_csv = os.path.join(tmp, "out.csv")
            changes = os.path.join(tmp, "changes.csv")
            flat = os.path.join(tmp, "flat.csv")
            other = os.path.join(tmp, "b.csv")
            changed = [list(r) for r in ANIMALS_TABLE]
            changed[3][2] = "Panthera leo"
            _write_csv(other, changed)
            code, envs = _run(
                [
                    "-c",
                    f"open {quoted} --id ID --parents Parent",
                    "-c",
                    f"save-as {json.dumps(out_csv)} --overwrite",
                    "-c",
                    f"merge {json.dumps(extra)} --id ID --parents Parent",
                    "-c",
                    f"replace --from-file {json.dumps(mapping)}",
                    "-c",
                    "add Tiger --parent Cats",
                    "-c",
                    f"export-changes {json.dumps(changes)} --overwrite",
                    "-c",
                    f"export-flat {json.dumps(flat)} --hier Parent --overwrite",
                    "-c",
                    f"compare {quoted} {json.dumps(other)} --id-a ID --parents-a Parent --id-b ID --parents-b Parent",
                ]
            )
            self.assertEqual(code, 0)
            _assert_envelope(self, envs[0], "open")
            _assert_envelope(self, envs[1], "save-as")
            _assert_envelope(self, envs[2], "merge")
            _assert_envelope(self, envs[3], "replace")
            _assert_envelope(self, envs[5], "export-changes")
            _assert_envelope(self, envs[6], "export-flat")
            _assert_envelope(self, envs[7], "compare")
            self.assertEqual(envs[0]["result"]["ids"], 3)
            self.assertEqual(envs[2]["result"]["ids_added"], 1)
            self.assertEqual(envs[3]["result"]["cells_changed"], 1)
            self.assertFalse(envs[7]["result"]["identical"])
            with open(changes, newline="", encoding="utf-8") as fh:
                rows = list(csv.reader(fh))
            code, envs = _run(
                [
                    "-c",
                    f"open {quoted} --id ID --parents Parent --discard",
                    "-c",
                    f"import-changes {json.dumps(changes)}",
                ]
            )
            _assert_envelope(self, envs[1], "import-changes")
            self.assertGreaterEqual(envs[1]["result"]["applied"] + envs[1]["result"]["failed"] + envs[1]["result"]["unnecessary"], len(rows))

    def test_delete_from_file_strips_whitespace(self):
        with tempfile.TemporaryDirectory() as tmp:
            quoted, _path = _open_animals(tmp)
            listed = os.path.join(tmp, "ids.csv")
            _write_csv(listed, [[" Lion "], ["Nope"]])
            code, envs = _run(
                [
                    "-c",
                    f"open {quoted} --id ID --parents Parent",
                    "-c",
                    f"delete --from-file {json.dumps(listed)}",
                ]
            )
            self.assertEqual(code, 0)
            self.assertEqual(envs[1]["result"]["listed"], 2)
            self.assertEqual(envs[1]["result"]["listed_deleted"], 1)
            self.assertEqual(envs[1]["result"]["not_found"], ["Nope"])
            self.assertIn("Lion", envs[1]["result"]["removed_entirely"] + envs[1]["result"]["removed_from_hierarchy"])

    def test_script_and_close_and_no_session(self):
        with tempfile.TemporaryDirectory() as tmp:
            script = os.path.join(tmp, "cmds.txt")
            with open(script, "w", encoding="utf-8") as fh:
                fh.write("# comment\n\nnew\nadd Animals\n")
            code, envs = _run(["--script", script, "-c", "get Animals", "-c", "close --discard"])
            self.assertEqual(code, 0)
            _assert_envelope(self, envs[0], "new")
            _assert_envelope(self, envs[2], "get")
            _assert_envelope(self, envs[3], "close")
            code, envs = _run(["-c", "tree"], check=False)
            self.assertEqual(code, 1)
            _assert_envelope(self, envs[0], "tree", ok=False)
            self.assertEqual(envs[0]["error"]["code"], "no_session")

    def test_query_get_and_find_level(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "add Animals",
                "-c",
                "add Cats --parent Animals",
                "-c",
                "add Lion --parent Cats",
                "-c",
                "set Lion DETAIL_1 LKD",
                "-c",
                "get --level 3 --contains LKD",
                "-c",
                "find LKD --in detail --level 3",
            ]
        )
        self.assertEqual(code, 0)
        env = envs[5]
        self.assertEqual(set(env), ENVELOPE_KEYS)
        self.assertEqual(set(env["status"]), STATUS_KEYS)
        self.assertEqual(env["command"], "get")
        self.assertTrue(env["ok"])
        self.assertEqual(
            set(env["result"]),
            {"ids", "not_found", "truncated", "returned", "total", "hierarchy", "level", "under"},
        )
        _assert_envelope(self, envs[6], "find")
        self.assertIn("level", envs[6]["result"]["hits"][0])

    def test_save_after_new_is_usage(self):
        code, envs = _run(["-c", "new", "-c", "save"], check=False)
        self.assertEqual(code, 1)
        _assert_envelope(self, envs[1], "save", ok=False)
        self.assertEqual(envs[1]["error"]["code"], "usage")

    def test_viewing_hierarchy_delete(self):
        code, envs = _run(["-c", "new", "-c", "column delete PARENT_1"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[1]["error"]["code"], "viewing_hierarchy")
        self.assertEqual(envs[1]["result"], {})

    def test_last_parent_column(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "column add PARENT_2 --hier",
                "-c",
                "column delete PARENT_2 PARENT_1",
            ],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual(envs[2]["error"]["code"], "last_parent_column")
        self.assertEqual(envs[2]["result"], {})
