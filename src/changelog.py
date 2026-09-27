# SPDX-License-Identifier: AGPL-3.0-only
# Copyright © R. A. Gardner

"""One changelog action, shown as seven-column lines.

A Change is one undo. Its rows are the lines in the changelog window.
Pipes are written into the Type cell when the action is committed.
Saved files stay a list of seven-cell rows. Loading regroups those
rows by the Type prefixes. A new action stores one local ISO-8601
datetime and every line of that action shows it. A line loaded from
an older file keeps the date text it was saved with.
"""

from __future__ import annotations

import datetime

ORIGIN_USER = "user"
ORIGIN_IMPORT = "import"
ORIGIN_MERGE = "merge"

EMPTY = ""

_ORIGIN_PREFIX = {
    ORIGIN_USER: "",
    ORIGIN_IMPORT: "Imported change | ",
    ORIGIN_MERGE: "Merge | ",
}

# Type prefixes that mark a line as part of a larger action.
# Same list Tree_Editor.prev_change used before actions were grouped.
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


def action_timestamp() -> str:
    return datetime.datetime.now().replace(microsecond=0).isoformat()


def is_member_type(typ: str) -> bool:
    return typ.startswith(_MEMBER_PREFIXES) or typ.endswith(("|", "| "))


def strip_type(typ: str) -> tuple[str, str]:
    """Return (origin, kind) from a Type cell."""
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


def _cell(value) -> str:
    if value is None:
        return EMPTY
    return value if isinstance(value, str) else f"{value}"


def row_stamp(row: ChangeRow) -> str:
    if row.at_text:
        return row.at_text
    if row.change is not None:
        return row.change.at
    return EMPTY


class ChangeRow:
    """One seven-column line. The date comes from the action unless at_text is set."""

    __slots__ = (
        "at_text",
        "change",
        "display_type",
        "from_col",
        "idn",
        "kind",
        "new",
        "old",
        "to_col",
    )

    def __init__(
        self,
        kind,
        idn=EMPTY,
        old=EMPTY,
        new=EMPTY,
        from_col=EMPTY,
        to_col=EMPTY,
        display_type=None,
        at_text=EMPTY,
        change=None,
    ):
        self.kind = kind
        self.idn = _cell(idn)
        self.old = _cell(old)
        self.new = _cell(new)
        self.from_col = _cell(from_col)
        self.to_col = _cell(to_col)
        self.display_type = kind if display_type is None else display_type
        self.at_text = _cell(at_text)
        self.change = change

    def __len__(self):
        return 7

    def __getitem__(self, i):
        if i == 0:
            return row_stamp(self)
        if i == 1:
            return self.display_type
        if i == 2:
            return self.idn
        if i == 3:
            return self.old
        if i == 4:
            return self.new
        if i == 5:
            return self.from_col
        if i == 6:
            return self.to_col
        raise IndexError(i)

    def __iter__(self):
        yield row_stamp(self)
        yield self.display_type
        yield self.idn
        yield self.old
        yield self.new
        yield self.from_col
        yield self.to_col

    def __eq__(self, other):
        return isinstance(other, ChangeRow) and tuple(self) == tuple(other) and self.kind == other.kind

    def __repr__(self):
        return (
            f"ChangeRow(kind={self.kind!r}, display_type={self.display_type!r}, "
            f"idn={self.idn!r}, old={self.old!r}, new={self.new!r}, "
            f"from_col={self.from_col!r}, to_col={self.to_col!r})"
        )

    @classmethod
    def from_dict(cls, d):
        # `what` is the cell name used by the action records written on the
        # cli branch. `idn` is the same cell in this seven-column log.
        return cls(
            d["kind"],
            d.get("idn", d.get("what", EMPTY)),
            d.get("old", EMPTY),
            d.get("new", EMPTY),
            d.get("from_col", EMPTY),
            d.get("to_col", EMPTY),
            display_type=d.get("display_type"),
            at_text=d.get("at_text", EMPTY),
        )


class Change:
    """One undoable action."""

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

    def first_stamp(self) -> str:
        return row_stamp(self.rows[0])

    def last_stamp(self) -> str:
        return row_stamp(self.rows[-1])

    def __repr__(self):
        return (
            f"Change(at={self.at!r}, kind={self.kind!r}, origin={self.origin!r}, "
            f"label={self.label!r}, n={self.n}, has_summary={self.has_summary})"
        )

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


class ChangeBuilder:
    __slots__ = ("label", "origin", "rows", "summary")

    def __init__(self, origin=ORIGIN_USER):
        self.origin = origin
        self.rows: list[ChangeRow] = []
        self.label = None
        self.summary = None

    def add(self, kind, idn=EMPTY, old=EMPTY, new=EMPTY, from_col=EMPTY, to_col=EMPTY):
        self.rows.append(ChangeRow(kind, idn, old, new, from_col, to_col))
        return self

    def set_summary(self, label, idn=EMPTY, old=EMPTY, new=EMPTY, from_col=EMPTY, to_col=EMPTY):
        self.label = label
        self.summary = (idn, old, new, from_col, to_col)
        return self

    def finish(self, at, *, plain=False, singular=None) -> Change:
        if not self.rows:
            raise ValueError("ChangeBuilder.finish requires at least one row")
        rows = list(self.rows)
        if singular is not None:
            # Same lines the old changelog_singular wrote: earlier rows keep
            # their pipe, and only the last Type cell is replaced.
            for row in rows[:-1]:
                row.display_type = member_display_type(row.kind, self.origin, True)
            rows[-1].kind = singular
            rows[-1].display_type = singular
            return Change(at=at, rows=rows, origin=self.origin, label=singular, has_summary=False)
        if plain:
            for row in rows:
                row.display_type = row.kind
            return Change(at=at, rows=rows, origin=self.origin, label=rows[0].kind, has_summary=False)
        emit = emits_summary(self.origin, len(rows), False)
        for row in rows:
            row.display_type = member_display_type(row.kind, self.origin, emit)
        label = self.label
        if emit:
            if label is None:
                label = rows[0].kind
            idn, old, new, from_col, to_col = self.summary or (EMPTY, EMPTY, EMPTY, EMPTY, EMPTY)
            rows.append(
                ChangeRow(
                    kind=label,
                    idn=idn,
                    old=old,
                    new=new,
                    from_col=from_col,
                    to_col=to_col,
                    display_type=label,
                )
            )
        elif label is None:
            label = rows[0].kind
        return Change(at=at, rows=rows, origin=self.origin, label=label, has_summary=emit)


def flatten_changelog(changes) -> list[tuple]:
    return [tuple(row) for ch in changes for row in ch.rows]


def display_rows(changes) -> list[ChangeRow]:
    rows = []
    for ch in changes:
        rows.extend(ch.rows)
    return rows


def _change_from_buf(buf, origin, summary):
    at = buf[0][0]
    members = [
        ChangeRow(
            kind=kind,
            idn=idn,
            old=old,
            new=new,
            from_col=from_col,
            to_col=to_col,
            display_type=typ,
            at_text=at_text,
        )
        for at_text, _origin, kind, typ, idn, old, new, from_col, to_col in buf
    ]
    if summary is None:
        label = members[0].kind
        has_summary = origin in (ORIGIN_IMPORT, ORIGIN_MERGE)
        if has_summary:
            members.append(ChangeRow(kind=label, display_type=label, at_text=at))
        return Change(at=at, rows=members, origin=origin, label=label, has_summary=has_summary)
    sat, styp, sidn, sold, snew, sfrom, sto = summary
    members.append(
        ChangeRow(
            kind=styp,
            idn=sidn,
            old=sold,
            new=snew,
            from_col=sfrom,
            to_col=sto,
            display_type=styp,
            at_text=sat,
        )
    )
    return Change(at=at, rows=members, origin=origin, label=styp, has_summary=True)


def _seven(row) -> list:
    cells = [_cell(c) for c in list(row)[:7]]
    if len(cells) < 7:
        cells.extend([EMPTY] * (7 - len(cells)))
    return cells


def regroup_tuples(rows) -> list[Change]:
    groups = []
    buf = []
    buf_origin = ORIGIN_USER
    for row in rows:
        at, typ, idn, old, new, from_col, to_col = _seven(row)
        if is_member_type(typ):
            origin, kind = strip_type(typ)
            if buf and origin != buf_origin:
                groups.append(_change_from_buf(buf, buf_origin, summary=None))
                buf = []
            if not buf:
                buf_origin = origin
            buf.append((at, origin, kind, typ, idn, old, new, from_col, to_col))
        else:
            if buf:
                groups.append(
                    _change_from_buf(
                        buf,
                        buf_origin,
                        summary=(at, typ, idn, old, new, from_col, to_col),
                    )
                )
                buf = []
            else:
                groups.append(
                    Change(
                        at=at,
                        rows=[
                            ChangeRow(
                                kind=typ,
                                idn=idn,
                                old=old,
                                new=new,
                                from_col=from_col,
                                to_col=to_col,
                                display_type=typ,
                                at_text=at,
                            )
                        ],
                        origin=ORIGIN_USER,
                        label=typ,
                        has_summary=False,
                    )
                )
    if buf:
        groups.append(_change_from_buf(buf, buf_origin, summary=None))
    return groups


def load_changelog(raw, warnings=None) -> list[Change]:
    """Read a stored changelog.

    Three shapes show up in files:

    - Seven-cell rows, the current save.
    - Five-cell rows from before the From/To columns. Short rows are padded.
      This is the load that stopped wiping those files.
    - One dict per action, with a ``rows`` list. That is what the cli branch
      writes. From/To are empty when the dict does not have them.

    Anything else resets the log and, when ``warnings`` is given, records why.
    """
    if not raw:
        return []
    first = raw[0]
    try:
        if isinstance(first, dict) and "rows" in first:
            return [Change.from_dict(x) for x in raw]
        if isinstance(first, (list, tuple)):
            return regroup_tuples(raw)
    except (KeyError, TypeError, ValueError, AttributeError, IndexError):
        if warnings is not None:
            warnings.append(" - Changelog was reset (unrecognized format)")
        return []
    if warnings is not None:
        warnings.append(" - Changelog was reset (unrecognized format)")
    return []


class ChangelogLog:
    """The editor's changelog. len() is the number of actions."""

    def __init__(self, on_unsaved=None, now=None):
        self.changes: list[Change] = []
        self.opened_at = 0
        self._pending: ChangeBuilder | None = None
        self.on_unsaved = on_unsaved or (lambda: None)
        self.now = now or action_timestamp

    def __len__(self):
        return len(self.changes)

    def __getitem__(self, i):
        return self.changes[i]

    def __bool__(self):
        return bool(self.changes)

    def clear(self):
        self.changes = []
        self._pending = None
        self.opened_at = 0

    def mark_opened(self):
        self._pending = None
        self.opened_at = len(self.changes)

    def load(self, raw, warnings=None):
        self._pending = None
        self.changes = load_changelog(raw, warnings)

    def flatten(self) -> list[tuple]:
        return flatten_changelog(self.changes)

    def display_rows(self) -> list[ChangeRow]:
        return display_rows(self.changes)

    def session_rows(self) -> list[tuple]:
        return flatten_changelog(self.changes[self.opened_at :])

    def last_type(self) -> str:
        if not self.changes:
            return ""
        return self.changes[-1].rows[-1].display_type

    def pop(self):
        self._pending = None
        if self.changes:
            self.changes.pop()

    def prune_through(self, up_to: int):
        removed = up_to + 1
        del self.changes[:removed]
        self.opened_at = max(0, self.opened_at - removed)

    def restore_front(self, changes, opened_at: int):
        self._pending = None
        self.changes = list(changes) + self.changes
        self.opened_at = opened_at

    def _stamp(self) -> str:
        return self.now()

    def _commit(self, builder: ChangeBuilder, *, plain=False, singular=None) -> Change:
        ch = builder.finish(self._stamp(), plain=plain, singular=singular)
        self.changes.append(ch)
        self._pending = None
        self.on_unsaved()
        return ch

    def _pending_add(self, change, idn, old, new, from_col, to_col):
        origin, kind = strip_type(change)
        if self._pending is None:
            self._pending = ChangeBuilder(origin=origin)
        elif self._pending.origin != origin:
            self._commit(self._pending)
            self._pending = ChangeBuilder(origin=origin)
        self._pending.add(kind, idn, old, new, from_col, to_col)

    def append(self, change, id_="", old="", new="", from_col="", to_col=""):
        if self._pending is not None:
            if is_member_type(change):
                self._pending_add(change, id_, old, new, from_col, to_col)
                return
            self._pending.set_summary(change, id_, old, new, from_col, to_col)
            self._commit(self._pending)
            return
        if is_member_type(change):
            self._pending_add(change, id_, old, new, from_col, to_col)
            return
        builder = ChangeBuilder()
        builder.add(change, id_, old, new, from_col, to_col)
        self._commit(builder)

    def append_no_unsaved(self, change, id_="", old="", new="", from_col="", to_col=""):
        self._pending_add(change, id_, old, new, from_col, to_col)

    def singular(self, text):
        if self._pending is None or not self._pending.rows:
            if self.changes and self.changes[-1].rows:
                row = self.changes[-1].rows[-1]
                row.kind = text
                row.display_type = text
                self.changes[-1].label = text
            self._pending = None
            self.on_unsaved()
            return
        self._commit(self._pending, singular=text)

    def finish_plain(self):
        if self._pending is None or not self._pending.rows:
            self._pending = None
            return None
        return self._commit(self._pending, plain=True)
