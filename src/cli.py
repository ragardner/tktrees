# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import argparse
import contextlib
import json
import os
import re
import shlex
import sys
from dataclasses import dataclass, field

from .api import _fail_startup
from .session import (
    Session,
    _alpha_to_idx,
    compare_files,
    fail,
    ok,
    read_table,
)

NEED_DOCUMENT = frozenset(
    {
        "save",
        "save-as",
        "tree",
        "get",
        "find",
        "columns",
        "changelog",
        "warnings",
        "hier",
        "add",
        "rename",
        "set",
        "move",
        "copy",
        "delete",
        "column add",
        "column rename",
        "column delete",
        "column validate",
        "option set",
        "tag",
        "untag",
        "tags",
        "sort",
        "undo",
        "merge",
        "replace",
        "import-changes",
        "export-changes",
        "export-flat",
    }
)

TWO_WORD = {
    ("column", "add"): "column add",
    ("column", "rename"): "column rename",
    ("column", "delete"): "column delete",
    ("column", "validate"): "column validate",
    ("option", "set"): "option set",
}

FLAG_SPEC = {
    "open": {"sheet": "val", "id": "val", "parents": "val", "format": "val", "discard": "bool"},
    "save": {"overwrite": "bool"},
    "save-as": {"overwrite": "bool", "sheet": "val"},
    "new": {"discard": "bool"},
    "close": {"discard": "bool"},
    "quit": {"discard": "bool"},
    "exit": {"discard": "bool"},
    "tree": {
        "under": "val",
        "depth": "val",
        "no-details": "bool",
        "all": "bool",
        "force": "bool",
        "hier": "val",
    },
    "get": {
        "level": "val",
        "contains": "val",
        "in": "val",
        "exact": "bool",
        "column": "val",
        "under": "val",
        "hier": "val",
        "limit": "val",
        "ids-only": "bool",
    },
    "find": {
        "in": "val",
        "exact": "bool",
        "hier": "val",
        "all-hier": "bool",
        "limit": "val",
        "level": "val",
        "under": "val",
        "column": "val",
    },
    "changelog": {"limit": "val", "session": "bool"},
    "hier": {"list": "bool"},
    "add": {"parent": "val", "before": "val", "after": "val"},
    "move": {
        "parent": "val",
        "top": "bool",
        "hier": "val",
        "from-hier": "val",
        "before": "val",
        "after": "val",
    },
    "copy": {"parent": "val", "top": "bool", "hier": "val"},
    "delete": {
        "children": "bool",
        "all-hierarchies": "bool",
        "orphan": "bool",
        "from-file": "val",
    },
    "column add": {"hier": "bool", "at": "val"},
    "column validate": {"set": "val", "clear": "bool"},
    "tag": {"clear": "bool", "from-file": "val", "in": "val", "exact": "bool"},
    "sort": {"column": "val", "desc": "bool", "tree": "bool", "children": "val"},
    "merge": {
        "id": "val",
        "parents": "val",
        "format": "val",
        "no-add-ids": "bool",
        "no-add-detail-columns": "bool",
        "no-add-parent-columns": "bool",
        "no-overwrite-details": "bool",
        "no-overwrite-parents": "bool",
    },
    "replace": {"from-file": "val"},
    "export-changes": {"session": "bool", "overwrite": "bool"},
    "export-flat": {
        "hier": "val",
        "no-details": "bool",
        "no-justify": "bool",
        "reverse": "bool",
        "index": "bool",
        "remove-end-ids": "val",
        "overwrite": "bool",
    },
    "compare": {"id-a": "val", "parents-a": "val", "id-b": "val", "parents-b": "val"},
}

COMMAND_HELP = {
    "open": """\
open — open a table file

syntax:
  open FILE
  open FILE --sheet NAME|INDEX
  open FILE --id COL --parents COL,COL [--format 0]
  open FILE --parents COL,COL --format 1|2|3|4
  open FILE --format 5|6|7
  open FILE --discard

flags:
  --sheet NAME|INDEX   workbook sheet (0-based index)
  --id COL             ID column name, 0-based index, or letter
  --parents COL,COL    parent columns
  --format 0..7        GUI Data Format dropdown; default 0
  --discard            replace unsaved work

example:
  open animals.csv --id ID --parents Parent

result:
  ids, format, id_column, columns, hierarchies, warnings

GUI: File → Open / Build tree

errors: unsaved, file_not_found, invalid_format, need_columns, need_sheet, program_data_present, unknown_column, usage
""",
    "save": """\
save — save the open file

syntax:
  save
  save --overwrite

flags:
  --overwrite   required if the file already exists (CLI always requires this for an existing path)

example:
  save --overwrite

result:
  file, kind (xlsx|csv|tsv|json)

GUI: File → Save

errors: no_session, usage (after new, use save-as), file_exists, invalid_format
""",
    "save-as": """\
save-as — save to a new path

syntax:
  save-as FILE
  save-as FILE --overwrite
  save-as FILE --sheet NAME

flags:
  --overwrite     replace an existing file
  --sheet NAME    sheet title for xlsx/json

example:
  save-as out.xlsx --overwrite

result:
  file, kind (xlsx|csv|tsv|json)

GUI: File → Save as

errors: no_session, file_exists, invalid_format, usage
""",
    "new": """\
new — empty document (ID, DETAIL_1, PARENT_1)

syntax:
  new
  new --discard

flags:
  --discard   replace unsaved work

example:
  new --discard

result:
  same shape as open (ids 0, format 0)

GUI: File → New

errors: unsaved, usage
""",
    "close": """\
close — unload the document; process stays

syntax:
  close
  close --discard

flags:
  --discard   discard unsaved work

example:
  close --discard

result:
  {}

GUI: no exact twin

errors: unsaved
""",
    "quit": """\
quit — end the process

syntax:
  quit
  quit --discard

flags:
  --discard   discard unsaved work

example:
  quit --discard

result:
  {} then the process ends

GUI: close the window

errors: unsaved
""",
    "exit": """\
exit — end the process (same as quit)

syntax:
  exit
  exit --discard

flags:
  --discard   discard unsaved work

example:
  exit --discard

result:
  {} then the process ends

errors: unsaved
""",
    "status": """\
status — file, unsaved, hierarchy, ids, undo

syntax:
  status

example:
  status

result:
  file, unsaved, hierarchy, ids, undo, warnings_count, columns, hierarchies, options
  Envelope status is only file, unsaved, hierarchy, ids, undo.

GUI: title bar / status bar
""",
    "help": """\
help — this list or help for a command

syntax:
  help
  help COMMAND
  COMMAND --help
  COMMAND -h

example:
  help add

result:
  help (string)

errors: usage (unknown command)
""",
    "format": """\
format — json or text envelopes (sticky for the process)

syntax:
  format
  format json
  format text

example:
  format json

result:
  format (json|text)
""",
    "tree": """\
tree — print the hierarchy

syntax:
  tree
  tree ID
  tree --under ID
  tree --depth N
  tree --no-details
  tree --all
  tree --force
  tree --hier NAME

flags:
  --under ID     start at this ID (a single positional ID is the same)
  --depth N      child levels; default 1; 0 is the start node(s) only
  --no-details   omit the details object
  --all          no depth cap (still 500 nodes unless --force)
  --force        no node cap
  --hier NAME    hierarchy (default: current)

example:
  tree --under Animals --depth 2

result:
  hierarchy, depth, under, truncated, returned, total, nodes

GUI: tree panel

errors: no_session, id_not_found, not_in_hierarchy, unknown_hierarchy
""",
    "get": """\
get — one ID, several IDs, or a query

syntax:
  get ID [ID ...]
  get --level N
  get --contains TERM [--in id|detail|any] [--exact] [--column COL]
  get --level N --contains TERM
  get --under ID [--level N]
  get --level N --contains TERM --ids-only --limit N

flags:
  --level N          tree depth; 1 is a top ID (same as Treeview levels)
  --under ID         only this ID and its descendants
  --contains TERM    not case sensitive; default --in detail
  --in id|detail|any with --contains; default detail
  --exact            whole cell
  --column COL       only this column
  --hier NAME        hierarchy for --level / --under (default: current)
  --limit N          query only; default 50; 0 means no cap
  --ids-only         {id, level} instead of full rows
  --level / --contains / --under may also filter a listed ID list

example:
  get --level 3 --contains LKD

result:
  one name: {id, parents, children, details, tagged}
  two or more: {ids, not_found}
  query: {ids, not_found, truncated, returned, total, hierarchy, level, under}
  query id objects include level
  all missing (by ID, no query flags) → id_not_found

GUI: ID concise view

errors: no_session, id_not_found, not_in_hierarchy, unknown_hierarchy, unknown_column, usage
""",
    "find": """\
find — search cells (not case sensitive)

syntax:
  find TERM
  find TERM --in id|detail|any
  find TERM --exact
  find TERM --hier NAME
  find TERM --all-hier
  find TERM --limit N
  find TERM --level N
  find TERM --under ID
  find TERM --column COL

flags:
  --in id|detail|any   default any
  --exact              whole cell
  --hier NAME          only IDs in that hierarchy
  --all-hier           no hierarchy filter
  --limit N            default 50; 0 means no cap
  --level N            only IDs at this tree depth (1 is a top ID)
  --under ID           only this ID and its descendants
  --column COL         only this column
  --all-hier cannot combine with --level or --under

example:
  find LKD --in detail --level 3

result:
  hits ({id, column, column_index, type, text, exact, level}), truncated, returned, total
  level is null with --all-hier

GUI: Find box

errors: no_session, id_not_found, not_in_hierarchy, unknown_hierarchy, unknown_column, usage
""",
    "columns": """\
columns — list columns

syntax:
  columns

example:
  columns

result:
  columns: [{index, letter, name, role, type}]
  role is id|parent|detail; type is ID|Parent|Text

errors: no_session
""",
    "changelog": """\
changelog — recent changes

syntax:
  changelog
  changelog --limit N
  changelog --session

flags:
  --limit N    default last 20 actions; 0 means no cap
  --session    actions since open/new

example:
  changelog --limit 5

result:
  entries ({at, type, what, old, new, origin, n, rows}), truncated, returned, total

errors: no_session
""",
    "warnings": """\
warnings — last build warnings

syntax:
  warnings

example:
  warnings

result:
  warnings (array of strings)

GUI: View → View build warnings

errors: no_session
""",
    "hier": """\
hier — current parent column

syntax:
  hier
  hier NAME
  hier INDEX
  hier --list

flags:
  --list   all parent columns and the current name

example:
  hier PARENT_1

result:
  get/set: {hierarchy, index, letter}
  --list: {hierarchies, current}

GUI: Hierarchy dropdown

errors: no_session, unknown_hierarchy, usage
""",
    "add": """\
add — add an ID to the current hierarchy

syntax:
  add ID
  add ID --parent PARENT
  add ID --before SIBLING
  add ID --after SIBLING
  add ID [ID ...] --parent PARENT

flags:
  --parent PARENT    child of PARENT; omit for a top ID
  --before SIBLING   tree-order splice; auto-sort must be off
  --after SIBLING    tree-order splice; auto-sort must be off

example:
  add Lion --parent Cats

result:
  added: [{id, parent, hierarchy, created_row}]
  parent is "" for a top ID
  several IDs are all-or-nothing

GUI: Add child / sibling / top

errors: no_session, empty_id, spaces_not_allowed, id_not_found, not_in_hierarchy, already_in_hierarchy, auto_sort_on, not_sibling, usage
""",
    "rename": """\
rename — rename an ID everywhere

syntax:
  rename ID NEW

example:
  rename Lion Cub

result:
  old, new

GUI: Rename ID

errors: no_session, id_not_found, id_exists, empty_id, spaces_not_allowed, usage
""",
    "set": """\
set — set a detail cell

syntax:
  set ID COLUMN VALUE
  set ID COLUMN ""

example:
  set Lion Name "Panthera leo"

result:
  id, column, old, new

GUI: edit a detail cell

errors: no_session, id_not_found, unknown_column, validation, usage (ID/parent columns)
""",
    "move": """\
move — reparent an ID

syntax:
  move ID --parent PARENT
  move ID --top
  move ID --parent PARENT --hier DEST
  move ID --from-hier SRC --parent PARENT --hier DEST
  move ID [ID ...] --parent PARENT
  move ID --before SIBLING
  move ID --after SIBLING

flags:
  --parent PARENT     new parent; "" is not used — use --top
  --top               top of the destination hierarchy
  --hier DEST         destination parent column (default: current)
  --from-hier SRC     source parent column (default: current)
  --before SIBLING    tree-order splice; auto-sort must be off
  --after SIBLING     tree-order splice; auto-sort must be off

example:
  move Lion --parent Animals

result:
  moved: [{id, from_hierarchy, from_parent, to_hierarchy, to_parent}]
  several IDs are all-or-nothing

GUI: Detach + paste

errors: no_session, id_not_found, not_in_hierarchy, already_in_hierarchy, cycle, unknown_hierarchy, auto_sort_on, not_sibling, usage
""",
    "copy": """\
copy — copy an ID into another hierarchy

syntax:
  copy ID --parent PARENT --hier DEST
  copy ID --top --hier DEST
  copy ID [ID ...] --parent PARENT --hier DEST

flags:
  --hier DEST         required destination parent column
  --parent PARENT     new parent
  --top               top of DEST
  no --before/--after

example:
  copy Lion --parent Cats --hier PARENT_2

result:
  copied: [{id, to_hierarchy, to_parent}]
  several IDs are all-or-nothing

GUI: Copy ID + paste in another hierarchy

errors: no_session, id_not_found, already_in_hierarchy, unknown_hierarchy, usage
""",
    "delete": """\
delete — delete IDs (missing names are skipped)

syntax:
  delete ID [ID ...]
  delete ID --children
  delete ID --all-hierarchies
  delete ID --children --all-hierarchies
  delete ID --orphan
  delete ID --orphan --all-hierarchies
  delete --from-file FILE
  delete --from-file FILE --children
  delete --from-file FILE --all-hierarchies

flags:
  --children          delete descendants too
  --all-hierarchies   every parent column
  --orphan            leave children as tops
  --from-file FILE    first column of IDs (no --orphan list form)

example:
  delete Lion --orphan

result:
  requested, removed_entirely, removed_from_hierarchy, promoted, deleted_descendants, orphaned, not_in_hierarchy, not_found
  --from-file also listed, listed_deleted

GUI: the six Delete menu items / Delete IDs using list

errors: no_session, usage
""",
    "column": """\
column — add, rename, delete, or validate columns

syntax:
  column add NAME [--hier] [--at INDEX]
  column rename OLD NEW
  column delete NAME [NAME ...]
  column validate NAME [--set v1,v2,v3|--clear]

example:
  column add Habitat

errors: no_session, unknown_column, last_parent_column, viewing_hierarchy, spaces_not_allowed, validation, usage
""",
    "column add": """\
column add — add a detail or parent column

syntax:
  column add NAME
  column add NAME --hier
  column add NAME --at INDEX

flags:
  --hier      parent column (default: detail Text)
  --at INDEX  0-based insert index (default: end)

example:
  column add Habitat --at 3

result:
  name, index, letter, role, type

GUI: Add detail / Add hierarchy

errors: no_session, spaces_not_allowed, usage
""",
    "column rename": """\
column rename — rename a column

syntax:
  column rename OLD NEW

example:
  column rename Name Label

result:
  old, new, index

GUI: Rename column

errors: no_session, unknown_column, spaces_not_allowed, usage
""",
    "column delete": """\
column delete — delete columns

syntax:
  column delete NAME [NAME ...]

example:
  column delete Habitat

result:
  deleted: [{name, index}], not_found

GUI: Delete column

errors: no_session, last_parent_column, viewing_hierarchy, usage
""",
    "column validate": """\
column validate — get or set a detail column's list

syntax:
  column validate NAME
  column validate NAME --set v1,v2,v3
  column validate NAME --clear

flags:
  --set v1,v2,v3   comma-separated values; "" is prepended if missing
  --clear          empty list; cells are not rewritten

example:
  column validate Name --set a,b

result:
  column, index, validation

GUI: column Validation

errors: no_session, unknown_column, validation, usage
""",
    "option": """\
option — read or set options

syntax:
  option
  option set NAME VALUE

NAME is allow-spaces-ids | allow-spaces-columns | auto-sort | save-program-data
VALUE is on or off

example:
  option set auto-sort off

result:
  allow-spaces-ids, allow-spaces-columns, auto-sort, save-program-data (JSON booleans)

GUI: Settings / Auto-sort tree IDs

errors: no_session (option set only), unknown_option, usage
""",
    "option set": """\
option set — set an option

syntax:
  option set NAME on|off

NAME:
  allow-spaces-ids
  allow-spaces-columns
  auto-sort
  save-program-data

example:
  option set auto-sort off

result:
  the options object

errors: no_session, unknown_option, usage
""",
    "tag": """\
tag — tag IDs (add only, no toggle)

syntax:
  tag ID [ID ...]
  tag --clear
  tag --from-file FILE [--in id|detail|any] [--exact]

flags:
  --clear            empty the tagged set
  --from-file FILE   first column of search terms
  --in id|detail|any default any
  --exact            whole-cell match for --from-file

example:
  tag Lion Cats

result:
  tag: {tagged, already, not_found}
  --clear: {cleared}
  --from-file: {terms_matched, terms_total, ids_tagged}

GUI: tag / Tag IDs using list / clear tags

errors: no_session, usage
""",
    "untag": """\
untag — remove tags

syntax:
  untag ID [ID ...]

example:
  untag Lion

result:
  untagged, not_tagged, not_found

errors: no_session, usage
""",
    "tags": """\
tags — list tagged IDs

syntax:
  tags

example:
  tags

result:
  ids (display names, sort_key order)

errors: no_session
""",
    "sort": """\
sort — sort the sheet or one ID's children

syntax:
  sort --column NAME [--desc]
  sort --tree
  sort --children ID

flags:
  --column NAME   sheet column, A→Z (Z→A with --desc)
  --desc          descending
  --tree          tree-walk row order
  --children ID   sort children of ID in the current hierarchy

example:
  sort --column Name --desc

result:
  --column: {kind: column, column, descending}
  --tree: {kind: tree}
  --children: {kind: children, id}

GUI: Sort sheet / tree walk / sort children

errors: no_session, unknown_column, id_not_found, usage
""",
    "undo": """\
undo — undo the last snapshot (cap 30)

syntax:
  undo

example:
  undo

result:
  undone (snapshot type string), undo (remaining slots)

GUI: Edit → Undo

errors: no_session, nothing_to_undo
""",
    "merge": """\
merge — merge another table into the open session

syntax:
  merge FILE [--id COL] [--parents COL,COL] [--format 0..7]
  merge FILE --no-add-ids --no-add-detail-columns --no-add-parent-columns --no-overwrite-details --no-overwrite-parents

flags:
  --id / --parents / --format   describe FILE, same rules as open
  --no-add-ids
  --no-add-detail-columns
  --no-add-parent-columns
  --no-overwrite-details
  --no-overwrite-parents
  Five GUI checkboxes default on; each --no-* turns one off.

example:
  merge extra.csv --id ID --parents Parent

result:
  ids_added, details_written, parents_written, detail_columns_added, parent_columns_added

GUI: Import → Merge Sheets

errors: no_session, no_changes, need_columns, file_not_found, invalid_format, program_data_present, unknown_column, usage
""",
    "replace": """\
replace — find/replace from a two-column file (not case sensitive)

syntax:
  replace --from-file FILE

flags:
  --from-file FILE   first two columns: find, replace

example:
  replace --from-file mapping.csv

result:
  file, cells_changed

GUI: Replace using mapping

errors: no_session, file_not_found, invalid_format, usage
""",
    "import-changes": """\
import-changes — apply a five-column changelog (no header row)

syntax:
  import-changes FILE

example:
  import-changes changes.csv

result:
  applied, unnecessary, failed, rows: [{index, ok, reason}]
  reason is null | unnecessary | id_missing | column_missing | value_mismatch | validation | type_mismatch | parent_mismatch | parse | other

GUI: Import changes

errors: no_session, invalid_format, file_not_found, usage
""",
    "export-changes": """\
export-changes — write changelog five-tuples (no header)

syntax:
  export-changes FILE [--session] [--overwrite]

flags:
  --session     rows since open/new
  --overwrite   replace an existing file

example:
  export-changes out.csv --overwrite

result:
  file, rows

errors: no_session, file_exists, invalid_format, usage
""",
    "export-flat": """\
export-flat — write a flattened table of this session

syntax:
  export-flat FILE --hier NAME [--no-details] [--no-justify] [--reverse] [--index] [--remove-end-ids N] [--overwrite]

flags:
  --hier NAME         required parent column
  --no-details        omit detail columns (default: all details on)
  --no-justify        do not justify left (default: on)
  --reverse
  --index             add an Index column
  --remove-end-ids N  default 0
  --overwrite

example:
  export-flat out.xlsx --hier Parent --overwrite

result:
  file, rows, cols (including the header row)

GUI: Export flattened sheet

errors: no_session, unknown_hierarchy, file_exists, invalid_format, usage
""",
    "compare": """\
compare — compare two files (does not use the open session)

syntax:
  compare FILE_A FILE_B --id-a COL --parents-a COLS --id-b COL --parents-b COLS
  compare FILE_A FILE_B

A side with program_data must omit that side's column flags.

example:
  compare a.csv b.csv --id-a ID --parents-a Parent --id-b ID --parents-b Parent

result:
  identical, warnings_a, warnings_b, headers, ids_only_a, ids_only_b, parent_diffs, detail_diffs

GUI: Compare sheets

errors: file_not_found, invalid_format, need_columns, program_data_present, unknown_column, usage
""",
}

_CATALOG_SKIP = frozenset({"column", "option set"})
HELP_LINES = [body.split("\n", 1)[0] for key, body in COMMAND_HELP.items() if key not in _CATALOG_SKIP]

PROCESS_HELP = """\
TkTrees CLI — one process, one document.

  python TKTREES.pyw cli
  python TKTREES.pyw cli --json
  python TKTREES.pyw cli --file PATH
  python TKTREES.pyw cli -c CMD [-c CMD ...]
  python TKTREES.pyw cli --script FILE.txt

  --json / -j     JSON envelopes (sticky for the process)
  --file PATH     open PATH, then REPL (or then -c / --script)
  -c CMD          run CMD (repeatable); then exit
  --script FILE   one command per line; # comments; then exit
  --keep-going    with -c / --script, continue after a failed command

Words after `cli` that are not those flags are not a command (exit 2).
Use -c, --script, or the REPL. help COMMAND for syntax and errors.

Commands:
"""


@dataclass
class CliRuntime:
    session: Session = field(default_factory=Session)
    json_mode: bool = False
    keep_going: bool = False
    stop: bool = False
    exit_code: int = 0


def wrap(command: str, out: dict, session: Session) -> dict:
    warnings = list(out.get("warnings") or [])
    result = out.get("result") if out.get("result") is not None else {}
    if isinstance(result, dict) and result.get("truncated") and "truncated" not in warnings:
        warnings.append("truncated")
    if command == "open":
        for w in result.get("warnings") or []:
            if w not in warnings:
                warnings.append(w)
    return {
        "ok": out["ok"],
        "command": command,
        "error": out["error"],
        "warnings": warnings,
        "result": result,
        "status": session.status_dict(),
    }


def _fmt_value(value):
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    return str(value)


def _tree_lines(nodes, indent=0):
    lines = []
    for node in nodes:
        lines.append(("  " * indent) + node["id"])
        lines.extend(_tree_lines(node.get("children") or [], indent + 1))
    return lines


def format_text(env: dict) -> str:
    if not env["ok"]:
        err = env["error"] or {}
        return f"error  {err.get('code', 'usage')}  {err.get('message', '')}\n"
    cmd = env["command"]
    result = env.get("result") or {}
    if cmd == "help":
        return f"ok  help\n{result.get('help', '')}"
    if not result:
        return f"ok  {cmd}\n"
    if cmd == "tree":
        body = "\n".join(_tree_lines(result.get("nodes") or []))
        line = f"ok  tree  hierarchy={result.get('hierarchy')}  returned={result.get('returned')}"
        return f"{line}\n{body}\n" if body else f"{line}\n"
    if cmd == "get" and "ids" in result:
        names = []
        for item in result.get("ids") or []:
            if isinstance(item, dict):
                names.append(item.get("id", ""))
            else:
                names.append(str(item))
        line = f"ok  get  returned={result.get('returned', len(names))}"
        if "total" in result:
            line += f"  total={result['total']}"
        body = "\n".join(names)
        return f"{line}\n{body}\n" if body else f"{line}\n"
    parts = [f"ok  {cmd}"]
    for key, value in result.items():
        if isinstance(value, (list, dict)):
            continue
        parts.append(f"{key}={_fmt_value(value)}")
    return "  ".join(parts) + "\n"


def print_env(rt: CliRuntime, env: dict, outfile) -> None:
    if rt.json_mode:
        outfile.write(json.dumps(env, ensure_ascii=False, separators=(",", ":")) + "\n")
    else:
        outfile.write(format_text(env))
    outfile.flush()


def usage(message: str) -> dict:
    return fail("usage", message)


def _parse_int(value, *, name, minimum=None, maximum=None):
    try:
        n = int(value)
    except (TypeError, ValueError):
        return usage(f"{name} must be an integer")
    if minimum is not None and n < minimum:
        return usage(f"{name} must be an integer >= {minimum}")
    if maximum is not None and n > maximum:
        return usage(f"{name} must be an integer <= {maximum}")
    return ok({"value": n})


def _resolve_header_token(headers: list[str], token: str):
    t = str(token)
    for i, name in enumerate(headers):
        if name.lower() == t.lower():
            return i
    if t.isdigit():
        i = int(t)
        if 0 <= i < len(headers):
            return i
        return None
    i = _alpha_to_idx(t)
    if i is not None and 0 <= i < len(headers):
        return i
    return None


def _parse_parents(value: str) -> list[str]:
    return [part.strip() for part in value.split(",") if part.strip()]


def _normalize_listed_id(text, allow_spaces: bool) -> str:
    text = "" if text is None else str(text)
    if allow_spaces:
        return text.strip()
    return re.sub(r"[\n\t\s]*", "", text)


def _collect_from_file(path: str, *, mode: str, allow_spaces: bool = False):
    out = read_table(path)
    if not out["ok"]:
        return out
    rows = out["result"]["rows"]
    abs_path = os.path.abspath(path)
    if mode == "mapping":
        mapping = {}
        for row in rows:
            a = row[0] if row else ""
            b = row[1] if len(row) > 1 else ""
            if a or b:
                mapping[str(a).lower()] = "" if b is None else str(b)
        return ok({"mapping": mapping, "path": abs_path})
    names = []
    seen = set()
    for row in rows:
        if not row:
            continue
        text = (
            _normalize_listed_id(row[0], allow_spaces)
            if mode == "ids"
            else ("" if row[0] is None else str(row[0]))
        )
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        names.append(text)
    return ok({"names": names, "path": abs_path})


def parse_line(line: str):
    try:
        tokens = shlex.split(line, posix=os.name != "nt")
    except ValueError as exc:
        return None, usage(str(exc))
    if not tokens:
        return None, None
    first = tokens[0].lower()
    rest = tokens[1:]
    if first == "help" or (rest and rest[0] in ("--help", "-h")):
        command = first
        if first != "help" and rest and rest[0] in ("--help", "-h"):
            return {"command": "help", "positionals": [first], "flags": {}}, None
        return {"command": "help", "positionals": [t.lower() for t in rest], "flags": {}}, None
    command = first
    if rest:
        pair = TWO_WORD.get((first, rest[0].lower()))
        if pair:
            command = pair
            rest = rest[1:]
    spec = FLAG_SPEC.get(command, {})
    positionals = []
    flags = {}
    i = 0
    while i < len(rest):
        tok = rest[i]
        if tok == "--":
            positionals.extend(rest[i + 1 :])
            break
        if tok in ("--help", "-h"):
            return {"command": "help", "positionals": [command], "flags": {}}, None
        if tok.startswith("--"):
            key = tok[2:]
            kind = spec.get(key)
            parsed = {"command": command, "positionals": positionals, "flags": flags}
            if kind is None:
                return parsed, usage(f"Unknown flag --{key}")
            if kind == "bool":
                flags[key] = True
                i += 1
                continue
            if i + 1 >= len(rest) or rest[i + 1].startswith("--"):
                return parsed, usage(f"Missing value for --{key}")
            flags[key] = rest[i + 1]
            i += 2
            continue
        if tok.startswith("-") and tok != "-":
            return (
                {"command": command, "positionals": positionals, "flags": flags},
                usage(f"Unknown flag {tok}"),
            )
        positionals.append(tok)
        i += 1
    return {"command": command, "positionals": positionals, "flags": flags}, None


def _need_doc(rt: CliRuntime, command: str) -> dict | None:
    if command in NEED_DOCUMENT and not rt.session.has_document:
        return fail("no_session", "No document is open")
    return None


def _scope_flags(s, flags, *, default_in="any"):
    in_ = flags.get("in", default_in)
    if in_ not in ("id", "detail", "any"):
        return usage("--in must be id, detail, or any")
    hier = None
    if "hier" in flags:
        hout = s.resolve_hier(flags["hier"])
        if not hout["ok"]:
            return hout
        hier = hout["result"]["index"]
    all_hier = bool(flags.get("all-hier"))
    if all_hier and ("level" in flags or "under" in flags):
        return usage("Cannot combine --all-hier with --level or --under")
    level = None
    if "level" in flags:
        try:
            level = int(flags["level"])
        except (TypeError, ValueError):
            return usage("level must be an integer >= 1")
        if level < 1:
            return usage("level must be an integer >= 1")
    under = None
    if "under" in flags:
        uout = s.resolve_id(flags["under"])
        if not uout["ok"]:
            return uout
        under = uout["result"]["key"]
        h = s.pc if hier is None else hier
        if s.nodes[under].ps[h] is None:
            return fail("not_in_hierarchy", f"{flags['under']} is not in this hierarchy   ")
    column = None
    if "column" in flags:
        col = s.resolve_column(flags["column"])
        if not col["ok"]:
            return col
        column = col["result"]["index"]
    try:
        limit = int(flags.get("limit", 50))
    except (TypeError, ValueError):
        return usage("limit must be an integer")
    if limit < 0:
        return usage("limit must be >= 0")
    contains = flags.get("contains")
    if contains is not None and not str(contains):
        return usage("--contains TERM is empty")
    return ok(
        {
            "in_": in_,
            "hier": hier,
            "all_hier": all_hier,
            "level": level,
            "under": under,
            "column": column,
            "limit": limit,
            "contains": contains,
            "exact": bool(flags.get("exact")),
            "ids_only": bool(flags.get("ids-only")),
        }
    )


def _open_cols(path, flags):
    parsed = _parse_int(flags.get("format", 0), name="format", minimum=0, maximum=7)
    if not parsed["ok"]:
        return usage("format must be an integer 0..7")
    fmt = parsed["result"]["value"]
    id_tok = flags.get("id")
    parents_tok = flags.get("parents")
    peek = read_table(path)
    if not peek["ok"]:
        return peek
    headers = peek["result"]["rows"][0] if peek["result"]["rows"] else []
    id_col = None
    parent_cols = None
    if id_tok is not None:
        id_col = _resolve_header_token(headers, id_tok)
        if id_col is None:
            return fail("unknown_column", f"Unknown column {id_tok}   ")
    if parents_tok is not None:
        parent_cols = []
        for tok in _parse_parents(parents_tok):
            idx = _resolve_header_token(headers, tok)
            if idx is None:
                return fail("unknown_column", f"Unknown column {tok}   ")
            parent_cols.append(idx)
    return ok({"id_col": id_col, "parent_cols": parent_cols, "fmt": fmt})


def dispatch(rt: CliRuntime, parsed: dict) -> dict:
    command = parsed["command"]
    pos = parsed["positionals"]
    flags = parsed["flags"]
    s = rt.session
    blocked = _need_doc(rt, command)
    if blocked is not None:
        return blocked

    if command == "help":
        if not pos:
            return ok({"help": "\n".join(HELP_LINES) + "\n"})
        name = " ".join(pos).lower()
        if name in COMMAND_HELP:
            return ok({"help": COMMAND_HELP[name]})
        return usage(f"Unknown command {name}")
    if command == "format":
        if not pos:
            return ok({"format": "json" if rt.json_mode else "text"})
        mode = pos[0].lower()
        if mode not in ("json", "text"):
            return usage("format json|text")
        rt.json_mode = mode == "json"
        return ok({"format": mode})
    if command == "status":
        st = s.status_dict()
        st["warnings_count"] = len(s.warnings)
        st["columns"] = s.columns_list()["result"]["columns"] if s.has_document else []
        st["hierarchies"] = s.hier_list()["result"]["hierarchies"] if s.has_document else []
        st["options"] = s.options_dict()
        return ok(st)
    if command == "open":
        if len(pos) != 1:
            return usage("open FILE")
        cols = _open_cols(pos[0], flags)
        if not cols["ok"]:
            return cols
        sheet = flags.get("sheet")
        if sheet is not None and str(sheet).isdigit():
            sheet = int(sheet)
        return s.load_path(
            pos[0],
            sheet=sheet,
            id_col=cols["result"]["id_col"],
            parent_cols=cols["result"]["parent_cols"],
            fmt=cols["result"]["fmt"],
            discard=bool(flags.get("discard")),
        )
    if command == "save":
        if pos:
            return usage("save [--overwrite]; use save-as FILE")
        if s.filepath is None:
            return usage("use save-as")
        return s.save_path(s.filepath, overwrite=bool(flags.get("overwrite")))
    if command == "save-as":
        if len(pos) != 1:
            return usage("save-as FILE")
        return s.save_path(pos[0], overwrite=bool(flags.get("overwrite")), sheetname=flags.get("sheet"))
    if command == "new":
        return s.new(discard=bool(flags.get("discard")))
    if command == "close":
        return s.close(discard=bool(flags.get("discard")))
    if command in ("quit", "exit"):
        if s.unsaved and not flags.get("discard"):
            return fail("unsaved", "Unsaved changes")
        rt.stop = True
        return ok({})
    if command == "tree":
        if len(pos) > 1:
            return usage("tree [ID] [--under ID] …")
        if pos and "under" in flags:
            return usage("Use a positional ID or --under, not both")
        under = pos[0] if pos else flags.get("under")
        if flags.get("all"):
            depth = None
        else:
            parsed = _parse_int(flags.get("depth", 1), name="depth", minimum=0)
            if not parsed["ok"]:
                return parsed
            depth = parsed["result"]["value"]
        hier = None
        if "hier" in flags:
            hout = s.resolve_hier(flags["hier"])
            if not hout["ok"]:
                return hout
            hier = hout["result"]["index"]
        return s.tree_dump(
            under=under,
            depth=depth,
            details=not flags.get("no-details"),
            hier=hier,
            force=bool(flags.get("force")),
        )
    if command == "get":
        is_query = any(k in flags for k in ("level", "contains", "under"))
        if not is_query:
            if flags:
                return usage("get query flags require --level, --contains, or --under")
            if not pos:
                return usage("get ID [ID ...] | get --level N | get --contains TERM | get --under ID")
            return s.get_ids(pos)
        if any(k in flags for k in ("in", "exact", "column")) and "contains" not in flags:
            return usage("--in, --exact, and --column require --contains")
        scoped = _scope_flags(s, flags, default_in="detail")
        if not scoped["ok"]:
            return scoped
        r = scoped["result"]
        return s.get_query(
            pos or None,
            level=r["level"],
            contains=r["contains"],
            in_=r["in_"],
            exact=r["exact"],
            column=r["column"],
            under=r["under"],
            hier=r["hier"],
            limit=r["limit"],
            ids_only=r["ids_only"],
        )
    if command == "find":
        if len(pos) != 1:
            return usage("find TERM")
        scoped = _scope_flags(s, flags, default_in="any")
        if not scoped["ok"]:
            return scoped
        r = scoped["result"]
        return s.find(
            pos[0],
            in_=r["in_"],
            exact=r["exact"],
            hier=r["hier"],
            all_hier=r["all_hier"],
            limit=r["limit"],
            level=r["level"],
            under=r["under"],
            column=r["column"],
        )
    if command == "columns":
        return s.columns_list()
    if command == "changelog":
        parsed = _parse_int(flags.get("limit", 20), name="limit", minimum=0)
        if not parsed["ok"]:
            return parsed
        return s.changelog_list(limit=parsed["result"]["value"], session=bool(flags.get("session")))
    if command == "warnings":
        return s.warnings_list()
    if command == "hier":
        if flags.get("list"):
            return s.hier_list()
        if not pos:
            return s.hier_get()
        if len(pos) != 1:
            return usage("hier NAME|INDEX")
        return s.set_hier(pos[0])
    if command == "add":
        if not pos:
            return usage("add ID")
        parent = flags.get("parent", "")
        before = flags.get("before")
        after = flags.get("after")
        if len(pos) == 1:
            return s.add(pos[0], parent, before=before, after=after)
        from .changelog import ChangeBuilder

        s.snapshot_sheet()
        b = ChangeBuilder()
        added = []
        for name in pos:
            out = s.add(name, parent, snapshot=False, before=before, after=after)
            if not out["ok"]:
                s.restore_snapshot(s.vs.pop())
                return out
            added.extend(out["result"]["added"])
            rec = out["result"]["added"][0]
            parent_text = rec["parent"] if rec["parent"] else "n/a - Top ID"
            hier = rec["hierarchy"]
            col = next(i for i, h in enumerate(s.headers) if h.name == hier)
            b.add(
                "Add ID",
                f"Name: {rec['id']} Parent: {parent_text} column #{col + 1} named: {hier}",
                "",
                "",
            )
        b.summary(f"Add {len(pos)} IDs")
        s.commit_change(b)
        return ok({"added": added})
    if command == "rename":
        if len(pos) != 2:
            return usage("rename ID NEW")
        return s.rename(pos[0], pos[1])
    if command == "set":
        if len(pos) != 3:
            return usage("set ID COLUMN VALUE")
        col = s.resolve_column(pos[1])
        if not col["ok"]:
            return col
        idx = col["result"]["index"]
        if idx == s.ic or idx in s.hiers:
            return usage("Use rename or move for ID and parent columns")
        return s.set_detail(pos[0], idx, pos[2])
    if command == "move":
        if not pos:
            return usage("move ID")
        hier = flags.get("hier")
        from_hier = flags.get("from-hier")
        if hier is not None:
            hout = s.resolve_hier(hier)
            if not hout["ok"]:
                return hout
            hier = hout["result"]["index"]
        if from_hier is not None:
            fout = s.resolve_hier(from_hier)
            if not fout["ok"]:
                return fout
            from_hier = fout["result"]["index"]
        if len(pos) == 1:
            return s.move(
                pos[0],
                parent=flags.get("parent"),
                top=bool(flags.get("top")),
                hier=hier,
                from_hier=from_hier,
                before=flags.get("before"),
                after=flags.get("after"),
            )
        from .changelog import ChangeBuilder

        s.snapshot_sheet()
        b = ChangeBuilder()
        moved = []
        for name in pos:
            out = s.move(
                name,
                parent=flags.get("parent"),
                top=bool(flags.get("top")),
                hier=hier,
                from_hier=from_hier,
                before=flags.get("before"),
                after=flags.get("after"),
                snapshot=False,
            )
            if not out["ok"]:
                s.restore_snapshot(s.vs.pop())
                return out
            moved.extend(out["result"]["moved"])
            rec = out["result"]["moved"][0]
            from_col = next(i for i, h in enumerate(s.headers) if h.name == rec["from_hierarchy"]) + 1
            to_col = next(i for i, h in enumerate(s.headers) if h.name == rec["to_hierarchy"]) + 1
            old_p = rec["from_parent"] if rec["from_parent"] else "n/a - Top ID"
            new_p = rec["to_parent"] if rec["to_parent"] else "n/a - Top ID"
            b.add(
                "Cut and paste ID",
                rec["id"],
                f"Old parent: {old_p} old column #{from_col} named: {rec['from_hierarchy']}",
                f"New parent: {new_p} new column #{to_col} named: {rec['to_hierarchy']}",
            )
        b.summary(f"Cut and paste {len(pos)} IDs")
        s.commit_change(b)
        return ok({"moved": moved})
    if command == "copy":
        if not pos or "hier" not in flags:
            return usage("copy ID --hier DEST")
        hout = s.resolve_hier(flags["hier"])
        if not hout["ok"]:
            return hout
        hier = hout["result"]["index"]
        if len(pos) == 1:
            return s.copy(pos[0], parent=flags.get("parent"), top=bool(flags.get("top")), hier=hier)
        from .changelog import ChangeBuilder

        s.snapshot_sheet()
        b = ChangeBuilder()
        copied = []
        for name in pos:
            out = s.copy(name, parent=flags.get("parent"), top=bool(flags.get("top")), hier=hier, snapshot=False)
            if not out["ok"]:
                s.restore_snapshot(s.vs.pop())
                return out
            copied.extend(out["result"]["copied"])
            rec = out["result"]["copied"][0]
            from_col = next(i for i, h in enumerate(s.headers) if h.name == rec["from_hierarchy"]) + 1
            to_col = next(i for i, h in enumerate(s.headers) if h.name == rec["to_hierarchy"]) + 1
            new_p = rec["to_parent"] if rec["to_parent"] else "n/a - Top ID"
            b.add(
                "Copy and paste ID",
                rec["id"],
                f"From column #{from_col} named: {rec['from_hierarchy']}",
                f"New parent: {new_p} new column #{to_col} named: {rec['to_hierarchy']}",
            )
        b.summary(f"Copy and paste {len(pos)} IDs")
        s.commit_change(b)
        return ok({"copied": copied})
    if command == "delete":
        if flags.get("children") and flags.get("orphan"):
            return usage("Cannot combine --children and --orphan")
        if not pos and "from-file" not in flags:
            return usage("delete ID [ID ...] | delete --from-file FILE")
        names = list(pos)
        listed = None
        listed_deleted = None
        if "from-file" in flags:
            if flags.get("orphan") or (flags.get("children") and flags.get("all-hierarchies")):
                return usage("List delete has no --orphan form and no --children --all-hierarchies form")
            collected = _collect_from_file(
                flags["from-file"],
                mode="ids",
                allow_spaces=bool(s.allow_spaces_ids_var),
            )
            if not collected["ok"]:
                return collected
            names = collected["result"]["names"]
            listed = len(names)
            listed_deleted = 0
            for name in names:
                ik = name.lower()
                if ik not in s.nodes:
                    continue
                if not flags.get("all-hierarchies") and s.nodes[ik].ps[s.pc] is None:
                    continue
                listed_deleted += 1
        out = s.delete(
            names,
            children=bool(flags.get("children")),
            all_hierarchies=bool(flags.get("all-hierarchies")),
            orphan=bool(flags.get("orphan")),
        )
        if listed is not None and out["ok"]:
            out["result"]["listed"] = listed
            out["result"]["listed_deleted"] = listed_deleted
        return out
    if command == "column add":
        if len(pos) != 1:
            return usage("column add NAME")
        if any(h.name.lower() == pos[0].lower() for h in s.headers):
            return usage(f"Column {pos[0]} already exists.")
        if "at" in flags:
            parsed = _parse_int(flags["at"], name="at", minimum=0, maximum=len(s.headers))
            if not parsed["ok"]:
                return usage(f"at must be an integer 0..{len(s.headers)}")
            at = parsed["result"]["value"]
        else:
            at = len(s.headers)
        if flags.get("hier"):
            return s.add_hier_col(at, pos[0])
        return s.add_col(at, pos[0], "Text")
    if command == "column rename":
        if len(pos) != 2:
            return usage("column rename OLD NEW")
        col = s.resolve_column(pos[0])
        if not col["ok"]:
            return col
        return s.rename_col(col["result"]["index"], pos[1])
    if command == "column delete":
        if not pos:
            return usage("column delete NAME [NAME ...]")
        cols = []
        not_found = []
        for name in pos:
            col = s.resolve_column(name)
            if not col["ok"]:
                not_found.append(name)
            else:
                cols.append(col["result"]["index"])
        if not cols:
            return ok({"deleted": [], "not_found": not_found})
        for idx in cols:
            if idx == s.ic:
                return usage("Cannot delete the ID column")
            if idx == s.pc:
                return fail("viewing_hierarchy", "Cannot delete the hierarchy you are viewing   ")
            if idx in s.hiers and len([h for h in s.hiers if h not in cols]) < 1:
                return fail("last_parent_column", "Cannot delete the last parent column   ")
        out = s.del_cols(cols)
        if out["ok"]:
            out["result"]["not_found"] = not_found
        return out
    if command == "column validate":
        if len(pos) != 1:
            return usage("column validate NAME")
        col = s.resolve_column(pos[0])
        if not col["ok"]:
            return col
        idx = col["result"]["index"]
        if flags.get("clear"):
            return s.set_validation(idx, [])
        if "set" in flags:
            values = list(flags["set"].split(","))
            return s.set_validation(idx, values)
        return ok(
            {
                "column": s.headers[idx].name,
                "index": idx,
                "validation": list(s.headers[idx].validation),
            }
        )
    if command == "option":
        return ok(s.options_dict())
    if command == "option set":
        if len(pos) != 2:
            return usage("option set NAME VALUE")
        return s.set_option(pos[0], pos[1])
    if command == "tag":
        if flags.get("clear"):
            return s.clear_tags()
        if "from-file" in flags:
            in_ = flags.get("in", "any")
            if in_ not in ("id", "detail", "any"):
                return usage("--in must be id, detail, or any")
            collected = _collect_from_file(flags["from-file"], mode="terms")
            if not collected["ok"]:
                return collected
            return s.tag_from_terms(
                collected["result"]["names"],
                in_=in_,
                exact=bool(flags.get("exact")),
            )
        if not pos:
            return usage("tag ID [ID ...]")
        return s.tag(pos)
    if command == "untag":
        if not pos:
            return usage("untag ID [ID ...]")
        return s.untag(pos)
    if command == "tags":
        return s.tags_list()
    if command == "sort":
        kinds = [k for k in ("column", "tree", "children") if k in flags]
        if len(kinds) != 1:
            return usage("sort --column NAME | --tree | --children ID")
        if "tree" in flags:
            return s.sort_sheet_walk()
        if "children" in flags:
            return s.sort_children(flags["children"])
        return s.sort_sheet(flags["column"], "DESCENDING" if flags.get("desc") else "ASCENDING")
    if command == "undo":
        return s.undo()
    if command == "merge":
        if len(pos) != 1:
            return usage("merge FILE")
        cols = _open_cols(pos[0], flags)
        if not cols["ok"]:
            return cols
        table = read_table(pos[0])
        if not table["ok"]:
            return table
        parsed = _parse_int(flags.get("format", 0), name="format", minimum=0, maximum=7)
        if not parsed["ok"]:
            return usage("format must be an integer 0..7")
        return s.merge_from_rows(
            table["result"]["rows"],
            fmt=parsed["result"]["value"],
            id_col=cols["result"]["id_col"],
            parent_cols=cols["result"]["parent_cols"],
            add_ids=not flags.get("no-add-ids"),
            add_dcols=not flags.get("no-add-detail-columns"),
            add_pcols=not flags.get("no-add-parent-columns"),
            overwrite_details=not flags.get("no-overwrite-details"),
            overwrite_parents=not flags.get("no-overwrite-parents"),
            file_opened=os.path.abspath(pos[0]),
        )
    if command == "replace":
        if "from-file" not in flags:
            return usage("replace --from-file FILE")
        collected = _collect_from_file(flags["from-file"], mode="mapping")
        if not collected["ok"]:
            return collected
        out = s.replace_mapping(collected["result"]["mapping"])
        if out["ok"]:
            out["result"]["file"] = collected["result"]["path"]
        return out
    if command == "import-changes":
        if len(pos) != 1:
            return usage("import-changes FILE")
        table = read_table(pos[0])
        if not table["ok"]:
            return table
        return s.import_changes(table["result"]["rows"], file_opened=pos[0])
    if command == "export-changes":
        if len(pos) != 1:
            return usage("export-changes FILE")
        return s.export_changelog(pos[0], session=bool(flags.get("session")), overwrite=bool(flags.get("overwrite")))
    if command == "export-flat":
        if len(pos) != 1 or "hier" not in flags:
            return usage("export-flat FILE --hier NAME")
        parsed = _parse_int(flags.get("remove-end-ids", 0), name="remove-end-ids", minimum=0)
        if not parsed["ok"]:
            return parsed
        return s.export_flat(
            pos[0],
            hier=flags["hier"],
            details=not flags.get("no-details"),
            justify=not flags.get("no-justify"),
            reverse=bool(flags.get("reverse")),
            index=bool(flags.get("index")),
            remove_end_ids=parsed["result"]["value"],
            overwrite=bool(flags.get("overwrite")),
        )
    if command == "compare":
        if len(pos) != 2:
            return usage("compare FILE_A FILE_B")
        id_a = flags.get("id-a")
        id_b = flags.get("id-b")
        parents_a = _parse_parents(flags["parents-a"]) if "parents-a" in flags else None
        parents_b = _parse_parents(flags["parents-b"]) if "parents-b" in flags else None
        if id_a is not None or parents_a is not None:
            peek = read_table(pos[0])
            if not peek["ok"]:
                return peek
            headers = peek["result"]["rows"][0] if peek["result"]["rows"] else []
            if id_a is not None:
                resolved_id = _resolve_header_token(headers, id_a)
                if resolved_id is None:
                    return fail("unknown_column", f"Unknown column {id_a}   ")
                id_a = resolved_id
            if parents_a is not None:
                resolved = [_resolve_header_token(headers, tok) for tok in parents_a]
                if any(v is None for v in resolved):
                    return fail("unknown_column", "Unknown column   ")
                parents_a = resolved
        if id_b is not None or parents_b is not None:
            peek = read_table(pos[1])
            if not peek["ok"]:
                return peek
            headers = peek["result"]["rows"][0] if peek["result"]["rows"] else []
            if id_b is not None:
                resolved_id = _resolve_header_token(headers, id_b)
                if resolved_id is None:
                    return fail("unknown_column", f"Unknown column {id_b}   ")
                id_b = resolved_id
            if parents_b is not None:
                resolved = [_resolve_header_token(headers, tok) for tok in parents_b]
                if any(v is None for v in resolved):
                    return fail("unknown_column", "Unknown column   ")
                parents_b = resolved
        return compare_files(pos[0], pos[1], id_a=id_a, parents_a=parents_a, id_b=id_b, parents_b=parents_b)
    return usage(f"Unknown command {command}")


def run_one(rt: CliRuntime, line: str, outfile) -> dict:
    parsed, err = parse_line(line)
    if err is not None:
        cmd = parsed["command"] if parsed else "usage"
        env = wrap(cmd, err, rt.session)
        print_env(rt, env, outfile)
        return env
    if parsed is None:
        return {
            "ok": True,
            "command": "",
            "error": None,
            "warnings": [],
            "result": {},
            "status": rt.session.status_dict(),
        }
    try:
        out = dispatch(rt, parsed)
    except (ValueError, IndexError, TypeError) as exc:
        out = usage(str(exc))
    env = wrap(parsed["command"], out, rt.session)
    print_env(rt, env, outfile)
    return env


def _script_lines(path: str) -> list[str]:
    with open(path, "r", encoding="utf-8") as fh:
        lines = []
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            lines.append(line)
        return lines


def _process_parser():
    parser = argparse.ArgumentParser(prog="TKTREES.pyw cli", add_help=False)
    parser.add_argument("--json", "-j", action="store_true")
    parser.add_argument("--file")
    parser.add_argument("-c", action="append", default=[], dest="commands")
    parser.add_argument("--script")
    parser.add_argument("--keep-going", action="store_true")
    parser.add_argument("-h", "--help", action="store_true")
    return parser


def run_cli(argv, *, infile=None, outfile=None):
    infile = sys.stdin if infile is None else infile
    outfile = sys.stdout if outfile is None else outfile
    parser = _process_parser()
    args, leftover = parser.parse_known_args(argv[2:])
    if leftover:
        _fail_startup(
            "Leftover arguments after cli are not an implicit command.\n"
            "Use -c CMD or --script FILE.\n"
            "python TKTREES.pyw cli --help"
        )
    rt = CliRuntime(json_mode=bool(args.json), keep_going=bool(args.keep_going))
    if args.help:
        env = wrap("help", ok({"help": PROCESS_HELP + "\n".join(HELP_LINES) + "\n"}), rt.session)
        print_env(rt, env, outfile)
        raise SystemExit(0)

    batch = list(args.commands)
    if args.script:
        try:
            batch = _script_lines(args.script) + batch
        except OSError as exc:
            env = wrap("usage", fail("file_not_found", str(exc)), rt.session)
            print_env(rt, env, outfile)
            raise SystemExit(1) from None

    if args.file:
        env = run_one(rt, f"open {shlex.quote(args.file)}", outfile)
        if not env["ok"] and batch:
            if not rt.keep_going:
                raise SystemExit(1)
            rt.exit_code = 1
    if batch or args.script is not None:
        for line in batch:
            env = run_one(rt, line, outfile)
            if rt.stop:
                raise SystemExit(rt.exit_code)
            if not env["ok"]:
                rt.exit_code = 1
                if not rt.keep_going:
                    raise SystemExit(1)
        raise SystemExit(rt.exit_code)

    with contextlib.suppress(Exception):
        import readline  # noqa: F401
    while True:
        try:
            if infile is sys.stdin:
                line = input("> ")
            else:
                line = infile.readline()
                if not line:
                    raise EOFError
                line = line.rstrip("\n")
        except EOFError:
            if rt.session.unsaved:
                env = wrap("quit", fail("unsaved", "Unsaved changes"), rt.session)
                print_env(rt, env, outfile)
                interactive = infile is sys.stdin and sys.stdin.isatty()
                if not interactive:
                    raise SystemExit(1) from None
                continue
            break
        except KeyboardInterrupt:
            if rt.session.unsaved:
                env = wrap("quit", fail("unsaved", "Unsaved changes"), rt.session)
                print_env(rt, env, outfile)
                interactive = infile is sys.stdin and sys.stdin.isatty()
                if not interactive:
                    raise SystemExit(1) from None
                continue
            outfile.write("\n")
            break
        line = (line or "").strip()
        if not line:
            continue
        env = run_one(rt, line, outfile)
        if rt.stop:
            raise SystemExit(rt.exit_code)
        if not env["ok"]:
            rt.exit_code = 1
            interactive = infile is sys.stdin and sys.stdin.isatty()
            if not interactive and not rt.keep_going:
                raise SystemExit(1)
    raise SystemExit(rt.exit_code)
