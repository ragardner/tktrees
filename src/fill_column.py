# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Fill a detail column from a template of {tokens}."""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass

from .classes import Node

_TOKEN_RE = re.compile(r"\{([^{}]+)\}")
_ASCII_DIGITS = re.compile(r"[0-9]+")

RESERVED_WORDS = frozenset(
    {
        "id",
        "parent",
        "grandparent",
        "great-grandparent",
        "great-great-grandparent",
        "root",
        "label",
        "depth",
        "path",
    }
)

WHO_CHOICES = (
    "This ID",
    "Parent",
    "Grandparent",
    "Great-grandparent",
    "Great-great-grandparent",
    "Root",
)

_WHO_TOKEN = {
    "This ID": "id",
    "Parent": "parent",
    "Grandparent": "grandparent",
    "Great-grandparent": "great-grandparent",
    "Great-great-grandparent": "great-great-grandparent",
    "Root": "root",
}

_WALK_STEPS = {
    "parent": 1,
    "grandparent": 2,
    "great-grandparent": 3,
    "great-great-grandparent": 4,
}

TREE_FIELDS = ("Label", "Depth", "Path")

_TREE_FIELD_TOKEN = {
    "Label": "label",
    "Depth": "depth",
    "Path": "path",
}

ID_COLUMN_CHOICE = "(ID)"


@dataclass(frozen=True)
class _Token:
    kind: str
    steps: int = 0
    col: int | None = None


CELL_CHOICES = (
    "Only where this column is empty",
    "Only where this column already has text",
    "Every row",
)
CELL_EMPTINESS = {
    "Only where this column is empty": "empty",
    "Only where this column already has text": "not_empty",
    "Every row": "all",
}
TREE_PLACE_CHOICES = (
    "Rows with or without children",
    "Only rows without children",
    "Only rows with children",
)
TREE_PLACE = {
    "Rows with or without children": "any",
    "Only rows without children": "leaf",
    "Only rows with children": "has_children",
}


@dataclass
class FillFilters:
    emptiness: str = "all"
    tree_place: str = "any"
    in_hierarchy: bool = False
    depth: int | None = None
    tagged: bool = False
    selected: bool = False
    descendants: bool = False
    same_level: bool = False
    match_col: int | None = None
    starts_with: str | None = None
    contains: str | None = None
    ends_with: str | None = None
    exact: str | None = None


CHOOSE_MATCH_COLUMN = "Choose a column to match"


def match_text(raw: str | None) -> str | None:
    # A blank pattern matches every cell, so it is not a test.
    if raw is None:
        return None
    text = raw.strip()
    return text or None


def _has_match(filters: FillFilters) -> bool:
    return any(
        match_text(text) is not None
        for text in (filters.exact, filters.starts_with, filters.contains, filters.ends_with)
    )


def _cell_matches(value: str, filters: FillFilters) -> bool:
    folded = value.lower()
    exact = match_text(filters.exact)
    if exact is not None and folded != exact.lower():
        return False
    starts = match_text(filters.starts_with)
    if starts is not None and not folded.startswith(starts.lower()):
        return False
    contains = match_text(filters.contains)
    if contains is not None and contains.lower() not in folded:
        return False
    ends = match_text(filters.ends_with)
    return not (ends is not None and not folded.endswith(ends.lower()))


def ascii_int(spec: str) -> int | None:
    if _ASCII_DIGITS.fullmatch(spec):
        return int(spec)
    return None


def _body_for(who: str, ref: str) -> str:
    if who == "id":
        return ref
    return f"{who}.{ref}"


def _reads_this_column(who: str, ref: str, idx: int, headers: list[str]) -> bool:
    # Add keeps a header name only when parsing that token returns this column.
    pieces, err = parse_template("{" + _body_for(who, ref) + "}", headers)
    if err or len(pieces) != 1 or not isinstance(pieces[0], _Token):
        return False
    token = pieces[0]
    if who == "id":
        return token.kind == "col" and token.col == idx
    if who == "root":
        return token.kind == "root" and token.col == idx
    return token.kind == "walk" and token.steps == _WALK_STEPS[who] and token.col == idx


def inserted_column_token(who_label: str, column_choice: str, headers: list[str]) -> str:
    who = _WHO_TOKEN[who_label]
    # "(ID)" means the ID word, unless a real header has that name.
    if column_choice == ID_COLUMN_CHOICE and column_choice not in headers:
        return who
    idx = next(i for i, name in enumerate(headers) if name == column_choice)
    name = headers[idx]
    ref = name if _reads_this_column(who, name, idx, headers) else str(idx + 1)
    return _body_for(who, ref)


def inserted_tree_field(field_label: str) -> str:
    return _TREE_FIELD_TOKEN[field_label]


def _resolve_column(spec: str, headers: list[str]) -> tuple[int | None, str | None]:
    n = ascii_int(spec)
    if n is not None:
        if n < 1:
            return None, "Column numbers start at 1"
        if n > len(headers):
            return None, f"Column {n} is past the end of the sheet"
        return n - 1, None
    key = spec.lower()
    for i, name in enumerate(headers):
        if name.lower() == key:
            return i, None
    return None, f"Unknown column: {spec}"


def _classify(body: str, headers: list[str]) -> tuple[_Token | None, str | None]:
    body = body.strip()
    if not body:
        return None, "Empty placeholder"
    if "." in body:
        left, right = body.split(".", 1)
        left_l = left.strip().lower()
        right = right.strip()
        if left_l in _WALK_STEPS or left_l in ("id", "root"):
            if not right:
                return None, f"Missing column after {left.strip()}"
            col, err = _resolve_column(right, headers)
            if err or col is None:
                return None, err
            if left_l == "id":
                return _Token("col", col=col), None
            if left_l == "root":
                return _Token("root", col=col), None
            return _Token("walk", steps=_WALK_STEPS[left_l], col=col), None
    low = body.lower()
    if low == "id":
        return _Token("id"), None
    if low == "label":
        return _Token("label"), None
    if low == "depth":
        return _Token("depth"), None
    if low == "path":
        return _Token("path"), None
    if low == "root":
        return _Token("root"), None
    if low in _WALK_STEPS:
        return _Token("walk", steps=_WALK_STEPS[low]), None
    col, err = _resolve_column(body, headers)
    if err or col is None:
        return None, err
    return _Token("col", col=col), None


def parse_template(template: str, headers: list[str]) -> tuple[list[str | _Token], str | None]:
    pieces: list[str | _Token] = []
    pos = 0
    for match in _TOKEN_RE.finditer(template):
        if match.start() > pos:
            pieces.append(template[pos : match.start()])
        token, err = _classify(match.group(1), headers)
        if err or token is None:
            return [], err or "Unknown token"
        pieces.append(token)
        pos = match.end()
    if pos < len(template):
        pieces.append(template[pos:])
    return pieces, None


def _ancestor(nodes: dict[str, Node], ik: str, h: int, steps: int) -> str | None:
    current = ik
    seen: set[str] = set()
    for _ in range(steps):
        if current in seen or current not in nodes:
            return None
        seen.add(current)
        parent = nodes[current].ps[h]
        if not parent:
            return None
        current = parent
    return current


def _root(nodes: dict[str, Node], ik: str, h: int) -> str | None:
    current = ik
    seen: set[str] = set()
    while current not in seen and current in nodes:
        seen.add(current)
        parent = nodes[current].ps[h]
        if not parent:
            return current
        current = parent
    return None


def _depth(nodes: dict[str, Node], ik: str, h: int) -> int:
    level = 1
    current = ik
    seen: set[str] = set()
    while current in nodes and current not in seen:
        seen.add(current)
        parent = nodes[current].ps[h]
        if not parent:
            break
        level += 1
        current = parent
    return level


def _path(nodes: dict[str, Node], ik: str, h: int) -> str:
    names: list[str] = []
    current = ik
    seen: set[str] = set()
    while current in nodes and current not in seen:
        seen.add(current)
        names.append(nodes[current].name)
        parent = nodes[current].ps[h]
        if not parent:
            break
        current = parent
    names.reverse()
    return " > ".join(names)


def _person(token: _Token, nodes: dict[str, Node], ik: str, h: int) -> str | None:
    if token.kind == "root":
        return _root(nodes, ik, h)
    if token.kind == "walk":
        return _ancestor(nodes, ik, h, token.steps)
    return ik


def _at(row: list[str] | None, col: int | None) -> str:
    if row is None or col is None or col < 0 or col >= len(row):
        return ""
    value = row[col]
    return value if isinstance(value, str) else f"{value}"


def _render(
    pieces: list[str | _Token],
    nodes: dict[str, Node],
    ik: str,
    h: int,
    cells: dict[str, list[str]],
    label_col: int,
) -> str:
    out: list[str] = []
    row = cells.get(ik)
    for part in pieces:
        if isinstance(part, str):
            out.append(part)
            continue
        if part.kind == "id":
            out.append(nodes[ik].name)
            continue
        if part.kind == "label":
            out.append(_at(row, label_col))
            continue
        if part.kind == "depth":
            out.append(str(_depth(nodes, ik, h)))
            continue
        if part.kind == "path":
            out.append(_path(nodes, ik, h))
            continue
        if part.kind == "col":
            out.append(_at(row, part.col))
            continue
        person = _person(part, nodes, ik, h)
        if person is None or person not in nodes:
            out.append("")
            continue
        if part.col is None:
            out.append(nodes[person].name)
        else:
            out.append(_at(cells.get(person), part.col))
    return "".join(out)


def _descendants(nodes: dict[str, Node], seeds: Iterable[str], h: int) -> set[str]:
    found: set[str] = set()
    stack: list[str] = []
    for ik in seeds:
        node = nodes.get(ik)
        if node is not None:
            stack.extend(node.cn[h])
    while stack:
        current = stack.pop()
        if current in found:
            continue
        found.add(current)
        node = nodes.get(current)
        if node is not None:
            stack.extend(node.cn[h])
    return found


def _same_level(nodes: dict[str, Node], seeds: Iterable[str], h: int) -> set[str]:
    levels: set[int] = set()
    for ik in seeds:
        node = nodes.get(ik)
        if node is None or node.ps[h] is None:
            continue
        levels.add(_depth(nodes, ik, h))
    if not levels:
        return set()
    return {ik for ik, node in nodes.items() if node.ps[h] is not None and _depth(nodes, ik, h) in levels}


def _passes(
    ik: str,
    old: str,
    nodes: dict[str, Node],
    filter_h: int,
    filters: FillFilters,
    tagged: set[str],
    selected: set[str],
    descendant_ids: set[str],
    level_ids: set[str],
    match_cell: str,
) -> bool:
    if filters.emptiness == "empty" and old != "":
        return False
    if filters.emptiness == "not_empty" and old == "":
        return False
    if not _cell_matches(match_cell, filters):
        return False
    if filters.tagged and ik not in tagged:
        return False
    if filters.selected and ik not in selected:
        return False
    node = nodes[ik]
    if filters.in_hierarchy and node.ps[filter_h] is None:
        return False
    if filters.tree_place == "leaf" and not (node.ps[filter_h] is not None and not node.cn[filter_h]):
        return False
    if filters.tree_place == "has_children" and not node.cn[filter_h]:
        return False
    if filters.depth is not None and (node.ps[filter_h] is None or _depth(nodes, ik, filter_h) != filters.depth):
        return False
    if filters.descendants and ik not in descendant_ids:
        return False
    return not (filters.same_level and ik not in level_ids)


def initial_fill_column(
    detail_names: list[str],
    selected_type: str | None,
    selected_name: str | None,
) -> str | None:
    # A column selection of a detail column wins. Otherwise the leftmost detail column.
    if not detail_names:
        return None
    if selected_type == "columns" and selected_name in detail_names:
        return selected_name
    return detail_names[0]


def plan_fill(
    rows: list[list[str]],
    ic: int,
    target_col: int,
    nodes: dict[str, Node],
    headers: list[str],
    fill_h: int,
    filter_h: int,
    label_col: int,
    template: str,
    filters: FillFilters,
    tagged: set[str] | None = None,
    selected: set[str] | None = None,
) -> tuple[list[tuple[int, str, str]], str | None]:
    pieces, err = parse_template(template, headers)
    if err:
        return [], err
    if _has_match(filters):
        col = filters.match_col
        if col is None or col < 0 or col >= len(headers):
            return [], CHOOSE_MATCH_COLUMN
    if tagged is None:
        tagged = set()
    if selected is None:
        selected = set()
    cells: dict[str, list[str]] = {}
    order: list[tuple[int, str]] = []
    for rn, row in enumerate(rows):
        ik = row[ic].lower()
        if ik not in nodes:
            continue
        cells[ik] = list(row)
        order.append((rn, ik))
    descendant_ids = _descendants(nodes, selected, filter_h) if filters.descendants else set()
    level_ids = _same_level(nodes, selected, filter_h) if filters.same_level else set()
    changes: list[tuple[int, str, str]] = []
    for rn, ik in order:
        old = _at(cells[ik], target_col)
        if not _passes(
            ik,
            old,
            nodes,
            filter_h,
            filters,
            tagged,
            selected,
            descendant_ids,
            level_ids,
            _at(cells[ik], filters.match_col),
        ):
            continue
        new = _render(pieces, nodes, ik, fill_h, cells, label_col)
        if new != old:
            changes.append((rn, old, new))
    return changes, None
