# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import csv
import os
import tempfile
import unittest

from src.api import run_api
from src.session import Session
from tests.fixtures import ANIMALS_TABLE


class TestFlattenUnflatten(unittest.TestCase):
    def test_flatten_then_unflatten_roundtrip_ids(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "animals.csv")
            flat = os.path.join(tmp, "flat.csv")
            back = os.path.join(tmp, "back.csv")
            with open(src, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(ANIMALS_TABLE)
            run_api(
                [
                    "TKTREES.pyw",
                    "flatten",
                    src,
                    flat,
                    "--parents",
                    "1",
                    "--id",
                    "0",
                    "--parent",
                    "1",
                    "--details",
                    "--overwrite",
                ]
            )
            self.assertTrue(os.path.isfile(flat))
            with open(flat, newline="", encoding="utf-8") as fh:
                flat_rows = list(csv.reader(fh))
            self.assertGreaterEqual(len(flat_rows), 2)
            parent_idxs = [str(i) for i, name in enumerate(flat_rows[0]) if name.startswith("Parent_")]
            run_api(
                [
                    "TKTREES.pyw",
                    "unflatten",
                    flat,
                    back,
                    "--parents",
                    ",".join(parent_idxs),
                    "--order",
                    "top-base",
                    "--overwrite",
                ]
            )
            s = Session()
            with open(back, newline="", encoding="utf-8") as fh:
                rows = list(csv.reader(fh))
            # unflatten writes ID/parent; find those columns
            headers = rows[0]
            id_col = next(i for i, h in enumerate(headers) if h.upper() in ("ID",) or i == 0)
            pcols = [i for i, h in enumerate(headers) if "PARENT" in h.upper() or h.upper() == "PARENT"]
            if not pcols:
                pcols = [1]
            out = s.load_table(rows, id_col=id_col, parent_cols=pcols, fmt=0)
            self.assertTrue(out["ok"], out)
            self.assertEqual(sorted(s.nodes), ["animals", "cats", "lion"])

    def test_flatten_refuses_existing_without_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "animals.csv")
            out = os.path.join(tmp, "flat.csv")
            with open(src, "w", newline="", encoding="utf-8") as fh:
                csv.writer(fh).writerows(ANIMALS_TABLE)
            with open(out, "w", encoding="utf-8") as fh:
                fh.write("x")
            with self.assertRaises(SystemExit) as raised:
                run_api(
                    [
                        "TKTREES.pyw",
                        "flatten",
                        src,
                        out,
                        "--parents",
                        "1",
                        "--id",
                        "0",
                        "--parent",
                        "1",
                    ]
                )
            self.assertEqual(raised.exception.code, 1)
