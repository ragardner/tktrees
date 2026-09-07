# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Grouped changelog: one Change is one undoable action.

Display and export still use five columns. Pipes live only in
ChangeRow.display_type, baked at commit. Popup/export iterate the
existing ChangeRow objects; they do not concatenate Type strings.
"""

from __future__ import annotations

ORIGIN_USER = "user"
ORIGIN_IMPORT = "import"
ORIGIN_MERGE = "merge"

EMPTY = ""

_ORIGIN_PREFIX = {
    ORIGIN_USER: "",
    ORIGIN_IMPORT: "Imported change | ",
    ORIGIN_MERGE: "Merge | ",
}

# Historical member Type prefixes. Used by regroup of old 5-tuple logs
# and by Session write shims. Undo does not scan these.
_MEMBER_PREFIXES = (
    "Merge | ",
    "Imported change |",
    "Edit cell |",
    "Delete ID from all hierarchies |",
    "Delete ID |",
    "Delete ID + all children |",
    "Delete ID + all children from all hierarchies |",
    "Cut and paste ID + children |",
    "Copy and paste ID |",
    "Copy and paste ID + children |",
    "Cut and paste ID |",
)


def is_member_type(typ: str) -> bool:
    return typ.startswith(_MEMBER_PREFIXES) or typ.endswith(("|", "| "))


def strip_type(typ: str) -> tuple[str, str]:
    """Return (origin, kind) from a five-column Type value."""
    if typ.startswith("Imported change |"):
        rest = typ.split("Imported change |", 1)[-1].strip()
        rest = rest[:-2].rstrip() if rest.endswith(" |") else rest
        return ORIGIN_IMPORT, rest
    if typ.startswith("Merge |"):
        rest = typ.split("Merge |", 1)[-1].strip()
        rest = rest[:-2].rstrip() if rest.endswith(" |") else rest
        return ORIGIN_MERGE, rest
    if typ.endswith(" |"):
        return ORIGIN_USER, typ[:-2]
    if typ.endswith("|"):
        return ORIGIN_USER, typ[:-1].rstrip()
    return ORIGIN_USER, typ


def member_display_type(kind: str, origin: str, emit_summary: bool) -> str:
    """Bake the Type column once. Callers store the result on ChangeRow."""
    if origin == ORIGIN_USER:
        if emit_summary:
            return kind + " |"
        return kind
    return _ORIGIN_PREFIX[origin] + kind


def emits_summary(origin: str, n_members: int, force_singular: bool = False) -> bool:
    if force_singular:
        return False
    if origin in (ORIGIN_IMPORT, ORIGIN_MERGE):
        return True
    return n_members > 1


class ChangeRow:
    """One five-column display line. Date comes from Change.at."""

    __slots__ = ("change", "display_type", "kind", "new", "old", "what")

    def __init__(self, kind, what=EMPTY, old=EMPTY, new=EMPTY, display_type=None, change=None):
        self.kind = kind
        self.what = what if what is not None else EMPTY
        self.old = old if old is not None else EMPTY
        self.new = new if new is not None else EMPTY
        self.display_type = kind if display_type is None else display_type
        self.change = change

    def __len__(self):
        return 5

    def __getitem__(self, i):
        if i == 0:
            return self.change.at if self.change is not None else EMPTY
        if i == 1:
            return self.display_type
        if i == 2:
            return self.what
        if i == 3:
            return self.old
        if i == 4:
            return self.new
        raise IndexError(i)

    def __iter__(self):
        yield self.change.at if self.change is not None else EMPTY
        yield self.display_type
        yield self.what
        yield self.old
        yield self.new

    def to_dict(self):
        return {
            "kind": self.kind,
            "what": self.what,
            "old": self.old,
            "new": self.new,
            "display_type": self.display_type,
        }

    @classmethod
    def from_dict(cls, d):
        return cls(
            d["kind"],
            d.get("what", EMPTY),
            d.get("old", EMPTY),
            d.get("new", EMPTY),
            display_type=d.get("display_type"),
        )

    def __eq__(self, other):
        return isinstance(other, ChangeRow) and (
            self.kind,
            self.what,
            self.old,
            self.new,
            self.display_type,
        ) == (
            other.kind,
            other.what,
            other.old,
            other.new,
            other.display_type,
        )

    def __repr__(self):
        return (
            f"ChangeRow(kind={self.kind!r}, display_type={self.display_type!r}, "
            f"what={self.what!r}, old={self.old!r}, new={self.new!r})"
        )


class Change:
    """One undoable action. One entry in session.changelog."""

    __slots__ = ("at", "has_summary", "label", "origin", "rows")

    def __init__(self, at, rows, origin=ORIGIN_USER, label=None, has_summary=False):
        if not rows:
            raise ValueError("Change.rows must be non-empty")
        self.at = at
        self.origin = origin
        self.rows = list(rows)
        self.has_summary = bool(has_summary)
        self.label = label if label is not None else self.rows[0].kind
        self._bind()

    def _bind(self):
        for row in self.rows:
            row.change = self

    def add_row(self, kind, what=EMPTY, old=EMPTY, new=EMPTY, display_type=None):
        row = ChangeRow(kind, what, old, new, display_type=display_type, change=self)
        self.rows.append(row)
        return row

    @property
    def kind(self):
        return self.rows[0].kind

    @property
    def n(self):
        n = len(self.rows)
        return n - 1 if self.has_summary and n else n

    @property
    def members(self):
        return self.rows[:-1] if self.has_summary else self.rows

    def to_dict(self):
        # Seven keys so older TkTrees `len(first) > 5` wipes the log.
        return {
            "at": self.at,
            "kind": self.kind,
            "origin": self.origin,
            "label": self.label,
            "rows": [r.to_dict() for r in self.rows],
            "has_summary": self.has_summary,
            "n": self.n,
        }

    @classmethod
    def from_dict(cls, d):
        rows = [ChangeRow.from_dict(r) for r in d["rows"]]
        if not rows:
            raise ValueError("empty rows")
        return cls(
            at=d["at"],
            rows=rows,
            origin=d.get("origin", ORIGIN_USER),
            label=d.get("label"),
            has_summary=d.get("has_summary", False),
        )

    def __eq__(self, other):
        return isinstance(other, Change) and (
            self.at,
            self.origin,
            self.label,
            self.has_summary,
            self.rows,
        ) == (
            other.at,
            other.origin,
            other.label,
            other.has_summary,
            other.rows,
        )

    def __repr__(self):
        return (
            f"Change(at={self.at!r}, kind={self.kind!r}, origin={self.origin!r}, "
            f"label={self.label!r}, n={self.n}, has_summary={self.has_summary}, "
            f"rows={self.rows!r})"
        )


class ChangeBuilder:
    __slots__ = (
        "force_singular",
        "label",
        "origin",
        "rows",
        "summary_new",
        "summary_old",
        "summary_what",
    )

    def __init__(self, origin=ORIGIN_USER, label=None):
        self.origin = origin
        self.rows = []
        self.label = label
        self.summary_what = EMPTY
        self.summary_old = EMPTY
        self.summary_new = EMPTY
        self.force_singular = False

    def add(self, kind, what=EMPTY, old=EMPTY, new=EMPTY):
        self.rows.append(ChangeRow(kind, what, old, new))
        return self

    def summary(self, label, what=EMPTY, old=EMPTY, new=EMPTY):
        self.label = label
        self.summary_what = what
        self.summary_old = old
        self.summary_new = new
        return self

    def finish(self, at) -> Change:
        if not self.rows:
            raise ValueError("ChangeBuilder.finish requires at least one row")
        n_members = len(self.rows)
        emit = emits_summary(self.origin, n_members, self.force_singular)
        for row in self.rows:
            baked = member_display_type(row.kind, self.origin, emit)
            if not emit and baked == row.kind:
                row.display_type = row.kind
            else:
                row.display_type = baked
        label = self.label
        rows = list(self.rows)
        if emit:
            if label is None:
                label = rows[0].kind
            rows.append(
                ChangeRow(
                    kind=label,
                    what=self.summary_what,
                    old=self.summary_old,
                    new=self.summary_new,
                    display_type=label,
                )
            )
        elif label is None:
            label = rows[0].kind
        elif not emit:
            # singular rewrite: Type column is the label, same object when equal
            if label == rows[0].kind:
                rows[0].display_type = rows[0].kind
                rows[0].kind = label
            else:
                rows[0].display_type = label
                rows[0].kind = label
        return Change(at=at, rows=rows, origin=self.origin, label=label, has_summary=emit)


def flatten_change(ch: Change) -> list[tuple]:
    """Project one Change to historical five-tuples (same strings, new tuples)."""
    return [tuple(row) for row in ch.rows]


def flatten_changelog(changes) -> tuple[list[tuple], list[int]]:
    rows = []
    index = []
    for i, ch in enumerate(changes):
        for row in ch.rows:
            rows.append(tuple(row))
            index.append(i)
    return rows, index


def display_rows(changes) -> list[ChangeRow]:
    """List of existing ChangeRow objects for the popup. No new cell strings."""
    rows = []
    for ch in changes:
        rows.extend(ch.rows)
    return rows


def _change_from_buf(buf, origin, summary):
    at = buf[0][0]
    members = [
        ChangeRow(kind=kind, what=what, old=old, new=new, display_type=typ)
        for _at, _origin, kind, typ, what, old, new in buf
    ]
    if summary is None:
        label = members[0].kind
        has_summary = origin in (ORIGIN_IMPORT, ORIGIN_MERGE)
        if has_summary:
            members.append(ChangeRow(kind=label, display_type=label))
        return Change(at=at, rows=members, origin=origin, label=label, has_summary=has_summary)
    _sat, styp, swhat, sold, snew = summary
    members.append(ChangeRow(kind=styp, what=swhat, old=sold, new=snew, display_type=styp))
    return Change(at=at, rows=members, origin=origin, label=styp, has_summary=True)


def regroup_tuples(rows) -> list[Change]:
    groups = []
    buf = []
    buf_origin = ORIGIN_USER
    for row in rows:
        at, typ, what, old, new = (list(row) + [EMPTY, EMPTY, EMPTY, EMPTY, EMPTY])[:5]
        if is_member_type(typ):
            origin, kind = strip_type(typ)
            if buf and origin != buf_origin:
                groups.append(_change_from_buf(buf, buf_origin, summary=None))
                buf = []
            if not buf:
                buf_origin = origin
            buf.append((at, origin, kind, typ, what, old, new))
        else:
            if buf:
                groups.append(_change_from_buf(buf, buf_origin, summary=(at, typ, what, old, new)))
                buf = []
            else:
                groups.append(
                    Change(
                        at=at,
                        rows=[ChangeRow(kind=typ, what=what, old=old, new=new, display_type=typ)],
                        origin=ORIGIN_USER,
                        label=typ,
                        has_summary=False,
                    )
                )
    if buf:
        groups.append(_change_from_buf(buf, buf_origin, summary=None))
    return groups


def load_changelog(raw, *, warnings=None) -> list:
    if not raw:
        return []
    first = raw[0]
    try:
        if isinstance(first, dict) and "rows" in first:
            return [Change.from_dict(x) for x in raw]
        if isinstance(first, (list, tuple)) and len(first) == 5:
            return regroup_tuples(raw)
    except (KeyError, TypeError, ValueError, AttributeError, IndexError):
        if warnings is not None:
            warnings.append(" - Changelog was reset (unrecognized format)")
        return []
    if warnings is not None:
        warnings.append(" - Changelog was reset (unrecognized format)")
    return []
