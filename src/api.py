# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import argparse
import difflib
import os
from sys import stderr
from typing import Literal, NoReturn

from tksheet import alpha2idx

from .classes import tk_trees_api
from .constants import software_version_number
from .functions import try_write_error_log

# File-in -> process -> output. Add future batch commands here and in
# build_parser(). First-token match is case-insensitive.
API_COMMANDS = frozenset({"flatten", "unflatten"})

# Reserved first token for a future interactive command line (app-like, not
# file-to-file). Must never fall through to the GUI.
CLI_COMMAND = "cli"

DATA_SUFFIXES = frozenset({".xlsx", ".xls", ".xlsm", ".csv", ".tsv", ".json"})

ORDER_TOP_BASE = "top-base"
ORDER_BASE_TOP = "base-top"
ORDERS = (ORDER_TOP_BASE, ORDER_BASE_TOP)


class _ApiParser(argparse.ArgumentParser):
    def error(self, message: str) -> None:
        self.print_usage(stderr)
        try_write_error_log(f"{self.prog}: error: {message}")
        raise SystemExit(2)


def _fail_startup(message: str) -> NoReturn:
    try_write_error_log(message)
    raise SystemExit(2)


def _is_gui_file_arg(token: str) -> bool:
    if os.path.isfile(token):
        return True
    return os.path.splitext(token)[1].lower() in DATA_SUFFIXES


def classify_invocation(argv: list[str]) -> Literal["gui", "api"]:
    """Decide GUI vs file-to-file API. Unknown or reserved tokens exit 2."""
    if len(argv) < 2:
        return "gui"
    first = argv[1]
    if first.startswith("-"):
        return "api"
    key = first.lower()
    if key in API_COMMANDS:
        argv[1] = key
        return "api"
    if key == CLI_COMMAND:
        _fail_startup(
            "'cli' is reserved for a future interactive command line and is not available yet.\n"
            "This version has the file-to-file API: flatten and unflatten.\n"
            "python TKTREES.pyw --help"
        )
    if _is_gui_file_arg(first):
        return "gui"
    matches = difflib.get_close_matches(key, sorted(API_COMMANDS | {CLI_COMMAND}), n=1, cutoff=0.6)
    hint = f"\nDid you mean '{matches[0]}'?" if matches else ""
    _fail_startup(
        f"Unknown command '{first}'.{hint}\n"
        "\n"
        "flatten and unflatten are the file-to-file API.\n"
        "cli is reserved for a future interactive command line.\n"
        "A data file path, or no arguments, opens the GUI.\n"
        "python TKTREES.pyw --help"
    )


def is_api_invocation(argv: list[str]) -> bool:
    return classify_invocation(argv) == "api"


def _column_index(token: str, kind: str) -> int:
    i = int(token) if token.isdigit() else alpha2idx(token)
    if not isinstance(i, int) or i < 0:
        raise ValueError(f"{kind} column must be a number or letter representing a column, not '{token}'")
    return i


def _arg_column(kind: str):
    def parse(token: str) -> int:
        try:
            return _column_index(token, kind)
        except ValueError as e:
            raise argparse.ArgumentTypeError(str(e)) from None

    return parse


def _arg_parent_columns(value: str) -> list[int]:
    tokens = [part.strip() for part in value.split(",") if part.strip()]
    if not tokens:
        raise argparse.ArgumentTypeError("Missing --parents (comma-separated column indexes or letters)")
    try:
        return sorted(_column_index(token, "Parent") for token in tokens)
    except ValueError as e:
        raise argparse.ArgumentTypeError(str(e)) from None


def build_parser(prog: str = "TKTREES.pyw") -> argparse.ArgumentParser:
    parser = _ApiParser(
        prog=prog,
        description=(
            "TkTrees API. Flatten an ID/parent table, or unflatten a "
            "level-across-columns table back to ID/parent.\n\n"
            "Running this file with no arguments, or with a data file path, opens the GUI. "
            "flatten, unflatten, --help and --version never open the GUI. "
            "cli is reserved for a future interactive command line. "
            "Any other first argument is an error, not the GUI."
        ),
        epilog="""examples:
  python TKTREES.pyw flatten in.xlsx out.xlsx --parents C,D --id A --parent C --input-sheet Sheet1 --output-sheet "New Sheet" --details --justify --overwrite
  python TKTREES.pyw unflatten in.csv out.csv --parents 0,2,4,6 --order top-base --delim tab --overwrite
  python TKTREES.pyw unflatten in.csv out.csv --parents 0,2,4,6 --order base-top --unique --overwrite
""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        allow_abbrev=False,
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"tktrees {software_version_number}",
    )

    shared = argparse.ArgumentParser(add_help=False)
    shared.add_argument(
        "input_filepath",
        metavar="INPUT",
        help="Input file (.xlsx, .xls, .xlsm, .csv, .tsv, .json)",
    )
    shared.add_argument(
        "output_filepath",
        metavar="OUTPUT",
        help="Output file (.csv, .tsv, .xlsx, .json). Suffix case is ignored for type; the filename is used as given",
    )
    shared.add_argument(
        "--parents",
        required=True,
        type=_arg_parent_columns,
        dest="all_parent_column_indexes",
        metavar="COLUMNS",
        help="All parent/hierarchy columns in the file, comma-separated indexes or letters (e.g. 1,2 or C,D)",
    )
    shared.add_argument(
        "--input-sheet",
        dest="input_sheet",
        metavar="NAME",
        help="Input sheet name for xlsx. Default: first sheet",
    )
    shared.add_argument(
        "--output-sheet",
        dest="output_sheet",
        metavar="NAME",
        help="Output sheet name for xlsx. Default: input sheet name, or Sheet1",
    )
    shared.add_argument(
        "--delim",
        dest="csv_delimiter",
        metavar="CHAR",
        help="Output delimiter for csv/tsv. Default: comma, or tab if OUTPUT ends with .tsv. Use tab for a tab. Quote shell characters, e.g. --delim '|'",
    )
    shared.add_argument(
        "--overwrite",
        action="store_true",
        dest="overwrite_file",
        help="Overwrite OUTPUT if it already exists. Without this, the command fails if the file exists",
    )

    sub = parser.add_subparsers(dest="command", title="commands", metavar="COMMAND")
    sub.required = True

    flatten = sub.add_parser(
        "flatten",
        parents=[shared],
        help="ID/parent table to levels across columns",
        description=(
            "Flatten one hierarchy into a row-per-path table with levels across columns. "
            "--parents lists every parent column in the file. --id and --parent say which "
            "ID column and which of those parent columns to flatten."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        allow_abbrev=False,
    )
    flatten.add_argument(
        "--id",
        required=True,
        type=_arg_column("ID"),
        dest="flatten_id_column",
        metavar="COL",
        help="ID column (index or letter). Required",
    )
    flatten.add_argument(
        "--parent",
        required=True,
        type=_arg_column("Parent"),
        dest="flatten_parent_column",
        metavar="COL",
        help="Parent column of the hierarchy to flatten (index or letter). Required",
    )
    flatten.add_argument(
        "--order",
        choices=ORDERS,
        default=ORDER_TOP_BASE,
        dest="order",
        help="left-to-right direction. top-base is root on the left (default). base-top is leaf on the left (GUI Reverse order)",
    )
    flatten.add_argument(
        "--details",
        action="store_true",
        dest="detail_columns",
        help="Include detail columns",
    )
    flatten.add_argument(
        "--justify",
        action="store_true",
        dest="justify_left",
        help="Pack shorter paths to the left; details sit to the right of each ID",
    )
    flatten.add_argument(
        "--index",
        action="store_true",
        dest="add_index",
        help="Add an index column",
    )

    unflatten = sub.add_parser(
        "unflatten",
        parents=[shared],
        help="Levels across columns back to ID/parent",
        description=(
            "Convert a flattened table (levels across columns) back to ID, parent, details. "
            "--order is required so the direction is not guessed. It must match how the "
            "INPUT hierarchy columns run left to right: top-base is root on the left, "
            "base-top is leaf on the left. If it does not match the file, parent links "
            "come out backwards. --parents lists those hierarchy columns (not the detail "
            "columns between them)."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        allow_abbrev=False,
    )
    unflatten.add_argument(
        "--order",
        required=True,
        choices=ORDERS,
        dest="order",
        help="left-to-right direction of INPUT hierarchy columns. top-base is root on the left, base-top is leaf on the left. Required",
    )
    unflatten.add_argument(
        "--unique",
        action="store_true",
        dest="unique",
        help="Each level keeps its own detail columns instead of one shared detail column (GUI Unique Details)",
    )
    return parser


def namespace_to_kwargs(ns: argparse.Namespace) -> dict:
    d = vars(ns).copy()
    d["api_action"] = d.pop("command")
    d["input_filepath"] = os.path.normpath(d["input_filepath"])
    d["output_filepath"] = os.path.normpath(d["output_filepath"])
    if not d.get("csv_delimiter"):
        d.pop("csv_delimiter", None)
        suffix = os.path.splitext(d["output_filepath"])[1].lower()
        if suffix == ".tsv":
            d["csv_delimiter"] = "tab"
    return {k: v for k, v in d.items() if v is not None}


def parse_api_argv(argv: list[str]) -> dict:
    prog = os.path.basename(argv[0]) if argv else "TKTREES.pyw"
    parser = build_parser(prog=prog)
    ns = parser.parse_args(argv[1:])
    return namespace_to_kwargs(ns)


def run_api(argv: list[str]) -> None:
    kwargs = parse_api_argv(argv)
    tk_trees_api(**kwargs)
