# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import csv
import io
import json
import os
import tempfile
import unittest

from src.cli import run_cli
from tests.fixtures import ANIMALS_TABLE, FOREST_TABLE, TWO_HIER_TABLE


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


def _csv(table):
    folder = tempfile.TemporaryDirectory()
    path = os.path.join(folder.name, "t.csv")
    with open(path, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(table)
    return folder, path


def _open(path, id_="ID", parents="Parent"):
    return ["-c", f"open {json.dumps(path)} --id {id_} --parents {parents}"]


class TestCliMutations(unittest.TestCase):
    def test_delete_one_changelog_and_undo(self):
        folder, path = _csv(ANIMALS_TABLE)
        self.addCleanup(folder.cleanup)
        code, envs = _run(
            [
                *_open(path),
                "-c",
                "delete Lion",
                "-c",
                "changelog --limit 1",
                "-c",
                "undo",
                "-c",
                "get Lion",
            ]
        )
        self.assertEqual(envs[2]["result"]["entries"][0]["type"], "Delete ID")
        self.assertEqual(envs[3]["result"]["undone"], "delete ids")
        self.assertEqual(envs[4]["result"]["id"], "Lion")

    def test_delete_children_then_orphan(self):
        folder, path = _csv(ANIMALS_TABLE)
        self.addCleanup(folder.cleanup)
        code, envs = _run(
            [
                *_open(path),
                "-c",
                "add Cub --parent Lion",
                "-c",
                "delete Lion --children",
                "-c",
                "get Cub",
            ],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual(envs[-1]["error"]["code"], "id_not_found")
        code, envs = _run(
            [
                *_open(path),
                "-c",
                "delete Cats --orphan",
                "-c",
                "get Lion",
            ]
        )
        self.assertEqual(envs[2]["result"]["parents"]["Parent"], "")

    def test_move_copy_into_new_hierarchy(self):
        folder, path = _csv(FOREST_TABLE)
        self.addCleanup(folder.cleanup)
        code, envs = _run(
            [
                *_open(path),
                "-c",
                "move Lion --parent Animals",
                "-c",
                "get Lion",
                "-c",
                "column add H2 --hier",
                "-c",
                "copy Life --hier H2 --top",
                "-c",
                "copy Lion --hier H2 --parent Life",
                "-c",
                "hier H2",
                "-c",
                "get Lion",
            ]
        )
        self.assertEqual(envs[2]["result"]["parents"]["Parent"], "Animals")
        self.assertEqual(envs[7]["result"]["parents"]["H2"], "Life")

    def test_delete_all_hierarchies(self):
        folder, path = _csv(TWO_HIER_TABLE)
        self.addCleanup(folder.cleanup)
        code, envs = _run(
            [
                *_open(path, parents="H1,H2"),
                "-c",
                "delete B --all-hierarchies",
                "-c",
                "get B",
            ],
            check=False,
        )
        self.assertEqual(code, 1)
        self.assertEqual(envs[-1]["error"]["code"], "id_not_found")

    def test_two_hier_delete_current_keeps_other(self):
        folder, path = _csv(TWO_HIER_TABLE)
        self.addCleanup(folder.cleanup)
        code, envs = _run(
            [
                *_open(path, parents="H1,H2"),
                "-c",
                "delete B",
                "-c",
                "get B",
            ]
        )
        self.assertIsNone(envs[2]["result"]["parents"]["H1"])
        self.assertEqual(envs[2]["result"]["parents"]["H2"], "Root")

    def test_import_changes_roundtrip(self):
        folder, path = _csv(ANIMALS_TABLE)
        self.addCleanup(folder.cleanup)
        log_path = os.path.join(folder.name, "log.csv")
        _run(
            [
                *_open(path),
                "-c",
                "rename Lion Leo",
                "-c",
                f"export-changes {json.dumps(log_path)} --overwrite",
            ]
        )
        self.assertTrue(os.path.isfile(log_path))
        code, envs = _run(
            [
                *_open(path),
                "-c",
                f"import-changes {json.dumps(log_path)}",
                "-c",
                "get Leo",
            ]
        )
        self.assertEqual(envs[1]["result"]["applied"], 1)
        self.assertEqual(envs[2]["result"]["id"], "Leo")

    def test_replace_and_get(self):
        folder, path = _csv(ANIMALS_TABLE)
        self.addCleanup(folder.cleanup)
        map_path = os.path.join(folder.name, "map.csv")
        with open(map_path, "w", newline="", encoding="utf-8") as fh:
            csv.writer(fh).writerows([["Lion", "Leo"]])
        code, envs = _run(
            [
                *_open(path),
                "-c",
                f"replace --from-file {json.dumps(map_path)}",
                "-c",
                "get Leo",
            ]
        )
        self.assertGreaterEqual(envs[1]["result"]["cells_changed"], 1)
        self.assertEqual(envs[2]["result"]["id"], "Leo")
