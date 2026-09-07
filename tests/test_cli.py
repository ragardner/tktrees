# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import csv
import io
import json
import os
import tempfile
import unittest

from src.cli import parse_line, run_cli
from tests.fixtures import ANIMALS_TABLE, FOREST_TABLE


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


class TestCliProcess(unittest.TestCase):
    def test_process_help_explains_invocation(self):
        out = io.StringIO()
        with self.assertRaises(SystemExit) as raised:
            run_cli(["TKTREES.pyw", "cli", "--help"], outfile=out)
        self.assertEqual(raised.exception.code, 0)
        text = out.getvalue()
        self.assertIn("--json", text)
        self.assertIn("-c CMD", text)
        self.assertIn("--script", text)
        self.assertIn("open —", text)

    def test_leftover_argv_is_exit_2(self):
        out = io.StringIO()
        with self.assertRaises(SystemExit) as raised:
            run_cli(["TKTREES.pyw", "cli", "add", "Lion"], outfile=out)
        self.assertEqual(raised.exception.code, 2)

    def test_status_without_document(self):
        code, envs = _run(["-c", "status"])
        self.assertEqual(code, 0)
        self.assertTrue(envs[0]["ok"])
        self.assertEqual(envs[0]["command"], "status")
        self.assertIsNone(envs[0]["status"]["file"])
        self.assertEqual(envs[0]["result"]["ids"], 0)

    def test_add_requires_document(self):
        code, envs = _run(["-c", "add Lion"], check=False)
        self.assertEqual(code, 1)
        self.assertFalse(envs[0]["ok"])
        self.assertEqual(envs[0]["error"]["code"], "no_session")


class TestCliCommands(unittest.TestCase):
    def test_new_add_get_tree(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "add Animals",
                "-c",
                "add Cats --parent Animals",
                "-c",
                "get Cats",
                "-c",
                "tree --under Animals --depth 1",
            ]
        )
        self.assertEqual(code, 0)
        self.assertTrue(envs[0]["ok"])
        self.assertEqual(envs[1]["result"]["added"][0]["id"], "Animals")
        self.assertEqual(envs[2]["result"]["added"][0]["parent"], "Animals")
        self.assertEqual(envs[3]["result"]["id"], "Cats")
        self.assertEqual(envs[3]["result"]["parents"]["PARENT_1"], "Animals")
        self.assertEqual(envs[4]["result"]["nodes"][0]["children"][0]["id"], "Cats")

    def test_open_csv_and_find(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.csv")
            with open(path, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(ANIMALS_TABLE)
            quoted = json.dumps(path)
            code, envs = _run(
                [
                    "-c",
                    f"open {quoted} --id ID --parents Parent",
                    "-c",
                    "find lion --in any",
                ]
            )
            self.assertEqual(code, 0)
            self.assertTrue(envs[0]["ok"], envs[0])
            self.assertEqual(envs[0]["command"], "open")
            self.assertEqual(envs[0]["result"]["ids"], 3)
            self.assertGreaterEqual(envs[1]["result"]["total"], 1)

    def test_multi_add_is_atomic(self):
        code, envs = _run(
            ["-c", "new", "-c", "add A", "-c", "add B A --parent A"],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertFalse(envs[2]["ok"])
        self.assertEqual(envs[2]["error"]["code"], "already_in_hierarchy")
        code, envs = _run(["-c", "new", "-c", "add A", "-c", "add B C --parent A", "-c", "get B"])
        self.assertEqual(code, 0)
        self.assertEqual(envs[3]["result"]["id"], "B")

    def test_multi_add_logs_one_change(self):
        code, envs = _run(
            [
                "-c",
                "new",
                "-c",
                "add Animals",
                "-c",
                "add Cats Dogs --parent Animals",
                "-c",
                "changelog --limit 1",
                "-c",
                "undo",
                "-c",
                "get Cats",
            ],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual(envs[3]["result"]["entries"][0]["type"], "Add 2 IDs")
        self.assertEqual(envs[3]["result"]["entries"][0]["n"], 2)
        self.assertEqual(envs[4]["result"]["undone"], "full sheet")
        self.assertEqual(envs[5]["error"]["code"], "id_not_found")

    def test_tree_under_missing_and_get_all_missing(self):
        code, envs = _run(
            ["-c", "new", "-c", "add A", "-c", "tree --under Nope", "-c", "get Missing Also"],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual(envs[2]["error"]["code"], "id_not_found")
        # -c stops on first failure unless keep-going
        self.assertEqual(len(envs), 3)

    def test_keep_going_and_unknown(self):
        code, envs = _run(["--keep-going", "-c", "nope", "-c", "status"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[0]["error"]["code"], "usage")
        self.assertTrue(envs[1]["ok"])
        self.assertEqual(envs[1]["command"], "status")

    def test_file_flag_opens_then_runs_c(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.xlsx")
            from src.session import Session

            s = Session()
            s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
            self.assertTrue(s.save_path(path, overwrite=True)["ok"])
            code, envs = _run(["--file", path, "-c", "get Lion"])
            self.assertEqual(code, 0)
            self.assertEqual(envs[0]["command"], "open")
            self.assertTrue(envs[0]["ok"], envs[0])
            self.assertEqual(envs[0]["result"]["ids"], 3)
            self.assertEqual(envs[1]["result"]["id"], "Lion")

    def test_file_missing_stops_batch(self):
        code, envs = _run(["--file", "/no/such/file.xlsx", "-c", "status"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[0]["error"]["code"], "file_not_found")
        self.assertEqual(len(envs), 1)

    def test_file_missing_keep_going_runs_c(self):
        code, envs = _run(
            ["--keep-going", "--file", "/no/such/file.xlsx", "-c", "status"],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual(envs[0]["error"]["code"], "file_not_found")
        self.assertEqual(envs[1]["command"], "status")
        self.assertTrue(envs[1]["ok"])

    def test_quit_unsaved_and_discard(self):
        code, envs = _run(["-c", "new", "-c", "add A", "-c", "quit"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[-1]["error"]["code"], "unsaved")
        code, envs = _run(["-c", "new", "-c", "add A", "-c", "quit --discard"])
        self.assertEqual(code, 0)
        self.assertTrue(envs[-1]["ok"])

    def test_text_mode_status_line(self):
        out = io.StringIO()
        with self.assertRaises(SystemExit) as raised:
            run_cli(["TKTREES.pyw", "cli", "-c", "status"], outfile=out)
        self.assertEqual(raised.exception.code, 0)
        self.assertTrue(out.getvalue().startswith("ok  status"))

    def test_eof_with_unsaved_exits_1(self):
        out = io.StringIO()
        with self.assertRaises(SystemExit) as raised:
            run_cli(
                ["TKTREES.pyw", "cli", "--json"],
                infile=io.StringIO("new\nadd A\n"),
                outfile=out,
            )
        self.assertEqual(raised.exception.code, 1)
        envs = [json.loads(line) for line in out.getvalue().splitlines() if line]
        self.assertEqual(envs[-1]["error"]["code"], "unsaved")

    def test_format_text_from_json_mode(self):
        out = io.StringIO()
        with self.assertRaises(SystemExit) as raised:
            run_cli(["TKTREES.pyw", "cli", "--json", "-c", "format text"], outfile=out)
        self.assertEqual(raised.exception.code, 0)
        self.assertTrue(out.getvalue().startswith("ok  format"))

    def test_get_query_level_and_contains(self):
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
                "set Lion DETAIL_1 LKD-42",
                "-c",
                "get --level 3 --contains LKD",
                "-c",
                "get --level 3 --contains LKD --ids-only",
                "-c",
                "find LKD --in detail --level 3",
                "-c",
                "get --level 2 --contains LKD",
            ]
        )
        self.assertEqual(code, 0)
        rows = envs[5]["result"]
        self.assertEqual([row["id"] for row in rows["ids"]], ["Lion"])
        self.assertEqual(rows["ids"][0]["level"], 3)
        self.assertEqual(rows["ids"][0]["details"]["DETAIL_1"], "LKD-42")
        self.assertEqual(rows["total"], 1)
        self.assertEqual(envs[6]["result"]["ids"], [{"id": "Lion", "level": 3}])
        self.assertEqual(envs[7]["result"]["hits"][0]["id"], "Lion")
        self.assertEqual(envs[7]["result"]["hits"][0]["level"], 3)
        self.assertEqual(envs[8]["result"]["ids"], [])
        self.assertEqual(envs[8]["result"]["total"], 0)

    def test_get_query_usage_and_under(self):
        code, envs = _run(
            [
                "--keep-going",
                "-c",
                "new",
                "-c",
                "add Animals",
                "-c",
                "add Cats --parent Animals",
                "-c",
                "add Lion --parent Cats",
                "-c",
                "get --under Animals --level 3",
                "-c",
                "get --ids-only",
                "-c",
                "get --level 0",
                "-c",
                "find LKD --level 3 --all-hier",
            ],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual([row["id"] for row in envs[4]["result"]["ids"]], ["Lion"])
        self.assertEqual(envs[5]["error"]["code"], "usage")
        self.assertEqual(envs[6]["error"]["code"], "usage")
        self.assertEqual(envs[7]["error"]["code"], "usage")


class TestCliDispatchFixes(unittest.TestCase):
    def test_save_positional_is_usage(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.csv")
            with open(path, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(ANIMALS_TABLE)
            quoted = json.dumps(path)
            dest = os.path.join(tmp, "out.csv")
            code, envs = _run(
                ["-c", f"open {quoted} --id ID --parents Parent", "-c", f"save {json.dumps(dest)} --overwrite"],
                check=False,
            )
            self.assertEqual(code, 1)
            self.assertEqual(envs[1]["error"]["code"], "usage")
            self.assertFalse(os.path.exists(dest))

    def test_bad_int_flags_are_usage_not_crash(self):
        code, envs = _run(["-c", "new", "-c", "tree --depth foo"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[1]["error"]["code"], "usage")
        code, envs = _run(["-c", "new", "-c", "column add X --at foo"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[1]["error"]["code"], "usage")
        code, envs = _run(["-c", "new", "-c", "column add X --at -1"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[1]["error"]["code"], "usage")

    def test_tree_positional_is_under(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "forest.csv")
            with open(path, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(FOREST_TABLE)
            quoted = json.dumps(path)
            code, envs = _run(
                ["-c", f"open {quoted} --id ID --parents Parent", "-c", "tree Moss", "-c", "tree --under Moss"],
            )
            self.assertEqual(code, 0)
            self.assertEqual(envs[1]["result"]["under"], "Moss")
            self.assertEqual([n["id"] for n in envs[1]["result"]["nodes"]], ["Moss"])
            self.assertEqual(envs[1]["result"]["nodes"], envs[2]["result"]["nodes"])

    def test_compare_unknown_id_column(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.csv")
            with open(path, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(ANIMALS_TABLE)
            quoted = json.dumps(path)
            code, envs = _run(
                ["-c", f"compare {quoted} {quoted} --id-a Nope --parents-a Parent"],
                check=False,
            )
            self.assertEqual(code, 1)
            self.assertEqual(envs[0]["error"]["code"], "unknown_column")

    def test_keep_going_quit_keeps_failure_exit(self):
        code, envs = _run(["--keep-going", "-c", "nope", "-c", "quit --discard"], check=False)
        self.assertEqual(code, 1)
        self.assertFalse(envs[0]["ok"])
        self.assertTrue(envs[1]["ok"])
        self.assertEqual(envs[1]["command"], "quit")

    def test_piped_stdin_failed_command_exits_1(self):
        out = io.StringIO()
        with self.assertRaises(SystemExit) as raised:
            run_cli(["TKTREES.pyw", "cli"], infile=io.StringIO("nope\nstatus\n"), outfile=out)
        self.assertEqual(raised.exception.code, 1)

    def test_delete_without_ids_is_usage(self):
        code, envs = _run(["-c", "new", "-c", "delete"], check=False)
        self.assertEqual(code, 1)
        self.assertEqual(envs[1]["error"]["code"], "usage")

    def test_open_format_out_of_range(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "animals.csv")
            with open(path, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(ANIMALS_TABLE)
            quoted = json.dumps(path)
            code, envs = _run(
                ["-c", f"open {quoted} --id ID --parents Parent --format 8"],
                check=False,
            )
            self.assertEqual(code, 1)
            self.assertEqual(envs[0]["error"]["code"], "usage")

    def test_tag_from_file_rejects_bad_in(self):
        with tempfile.TemporaryDirectory() as tmp:
            terms = os.path.join(tmp, "terms.csv")
            with open(terms, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows([["Lion"]])
            quoted = json.dumps(terms)
            code, envs = _run(
                ["-c", "new", "-c", "add Lion", "-c", f"tag --from-file {quoted} --in nope"],
                check=False,
            )
            self.assertEqual(code, 1)
            self.assertEqual(envs[2]["error"]["code"], "usage")

    def test_parse_line_windows_paths_keep_backslashes(self):
        import src.cli as clip

        orig = clip.os.name
        try:
            clip.os.name = "nt"
            parsed, err = parse_line(r"open C:\temp\file.xlsx")
        finally:
            clip.os.name = orig
        self.assertIsNone(err)
        self.assertEqual(parsed["command"], "open")
        self.assertEqual(parsed["positionals"], [r"C:\temp\file.xlsx"])
