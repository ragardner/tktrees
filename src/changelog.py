# SPDX-License-Identifier: AGPL-3.0-only
# Copyright © R. A. Gardner

"""The changelog.

One action is one undo. Its lines are the seven cells shown in the
changelog window: date and time, type, ID, old value, new value,
From column, To column. The type text is stored as it was passed in.
"""

import datetime
import getpass

# A line whose type starts with one of these belongs to a larger action.
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


def action_timestamp():
    return datetime.datetime.now().replace(microsecond=0).isoformat()


def os_username():
    """Operating-system login, or "" when it cannot be read."""
    try:
        user = getpass.getuser()
    except OSError:
        return ""
    if not isinstance(user, str):
        return ""
    user = user.strip()
    if not user or any(ch in user for ch in "\r\n\t"):
        return ""
    return user


def changelog_stamp(include_user=False):
    """Local ISO-8601 time. With include_user, the OS login is appended."""
    stamp = action_timestamp()
    if not include_user:
        return stamp
    user = os_username()
    if not user:
        return stamp
    return f"{stamp} {user}"


def _cell(value):
    if value is None:
        return ""
    return value if isinstance(value, str) else f"{value}"


def _is_member(typ):
    return typ.startswith(_MEMBER_PREFIXES) or typ.endswith(("|", "| "))


def _group(typ):
    """Which edit a type belongs to. Used when a new line arrives mid-edit, and when loading."""
    if typ.startswith("Imported change |"):
        return "import"
    if typ.startswith("Merge |"):
        return "merge"
    return "user"


def _seven(row):
    cells = [_cell(c) for c in list(row)[:7]]
    if len(cells) < 7:
        cells.extend([""] * (7 - len(cells)))
    return tuple(cells)


class Action:
    """One undo. rows are the seven-cell lines."""

    __slots__ = ("has_summary", "rows")

    def __init__(self, rows, has_summary=False):
        if not rows:
            raise ValueError("empty action")
        self.rows = list(rows)
        self.has_summary = bool(has_summary)

    @property
    def n(self):
        n = len(self.rows)
        return n - 1 if self.has_summary and n else n

    def first_stamp(self):
        return self.rows[0][0]

    def last_stamp(self):
        return self.rows[-1][0]


class Changelog:
    """actions are finished undos. pending is the edit that has not finished."""

    def __init__(self, on_unsaved=None, now=None):
        self.actions = []
        self.pending = []
        self.opened_at = 0
        self.on_unsaved = on_unsaved or (lambda: None)
        self.now = now or action_timestamp

    def __len__(self):
        return len(self.actions)

    def __getitem__(self, i):
        return self.actions[i]

    def __bool__(self):
        return bool(self.actions)

    def clear(self):
        self.actions = []
        self.pending = []
        self.opened_at = 0

    def mark_opened(self):
        self.pending = []
        self.opened_at = len(self.actions)

    def load(self, raw, warnings=None):
        self.pending = []
        self.actions = load_changelog(raw, warnings)

    def flatten(self):
        return [row for action in self.actions for row in action.rows]

    def display_rows(self):
        return self.flatten()

    def session_rows(self):
        return [row for action in self.actions[self.opened_at :] for row in action.rows]

    def last_type(self):
        if not self.actions:
            return ""
        return self.actions[-1].rows[-1][1]

    def action_index(self, row_index):
        seen = 0
        for i, action in enumerate(self.actions):
            seen += len(action.rows)
            if row_index < seen:
                return i
        raise IndexError(row_index)

    def pop(self):
        self.pending = []
        if self.actions:
            self.actions.pop()

    def prune_through(self, up_to):
        removed = up_to + 1
        del self.actions[:removed]
        self.opened_at = max(0, self.opened_at - removed)

    def restore_front(self, actions, opened_at):
        self.pending = []
        self.actions = list(actions) + self.actions
        self.opened_at = opened_at

    def _line(self, typ, id_, old, new, from_col, to_col):
        return ("", _cell(typ), _cell(id_), _cell(old), _cell(new), _cell(from_col), _cell(to_col))

    def _add_pending(self, typ, id_, old, new, from_col, to_col):
        group = _group(typ)
        if self.pending and _group(self.pending[0][1]) != group:
            self._finish(summary=False, interrupted=True)
        self.pending.append(self._line(typ, id_, old, new, from_col, to_col))

    def _finish(self, summary=False, singular=None, interrupted=False):
        if not self.pending:
            self.pending = []
            return None
        rows = self.pending
        self.pending = []
        has_summary = False
        if singular is not None:
            last = rows[-1]
            rows[-1] = (last[0], singular, last[2], last[3], last[4], last[5], last[6])
        elif summary:
            has_summary = True
        elif interrupted:
            # The next line belongs to a different edit. Close this one first.
            group = _group(rows[0][1])
            several = group in ("import", "merge") or len(rows) > 1
            rows = [_retitle(row, group, several) for row in rows]
            if several:
                rows = rows + [self._line(_bare(rows[0][1]), "", "", "", "", "")]
                has_summary = True
        at = self.now()
        stamped = tuple((at,) + row[1:] for row in rows)
        action = Action(stamped, has_summary)
        self.actions.append(action)
        self.on_unsaved()
        return action

    def append(self, change, id_="", old="", new="", from_col="", to_col=""):
        if self.pending:
            if _is_member(change):
                self._add_pending(change, id_, old, new, from_col, to_col)
                return
            self.pending.append(self._line(change, id_, old, new, from_col, to_col))
            self._finish(summary=True)
            return
        if _is_member(change):
            self._add_pending(change, id_, old, new, from_col, to_col)
            return
        self.pending.append(self._line(change, id_, old, new, from_col, to_col))
        self._finish()

    def append_no_unsaved(self, change, id_="", old="", new="", from_col="", to_col=""):
        self._add_pending(change, id_, old, new, from_col, to_col)

    def singular(self, text):
        if not self.pending:
            if self.actions:
                action = self.actions[-1]
                last = action.rows[-1]
                action.rows[-1] = (last[0], text, last[2], last[3], last[4], last[5], last[6])
            self.on_unsaved()
            return
        self._finish(singular=text)

    def finish_plain(self):
        if not self.pending:
            return None
        return self._finish()


def _retitle(row, group, several):
    typ = _bare(row[1])
    if group == "import":
        typ = "Imported change | " + typ
    elif group == "merge":
        typ = "Merge | " + typ
    elif several:
        typ = typ + " |"
    return (row[0], typ, row[2], row[3], row[4], row[5], row[6])


def _bare(typ):
    """Type with the import/merge prefix and a trailing | removed."""
    if typ.startswith("Imported change |"):
        typ = typ.split("Imported change |", 1)[-1].strip()
    elif typ.startswith("Merge |"):
        typ = typ.split("Merge |", 1)[-1].strip()
    if typ.endswith(" |"):
        typ = typ[:-2]
    elif typ.endswith("|"):
        typ = typ[:-1].rstrip()
    return typ


def _action_from_record(d):
    at = d["at"]
    rows = []
    for record in d["rows"]:
        typ = record.get("display_type")
        if typ is None:
            typ = record["kind"]
        date = record.get("at_text") or at
        rows.append(
            (
                _cell(date),
                _cell(typ),
                _cell(record.get("idn", record.get("what", ""))),
                _cell(record.get("old", "")),
                _cell(record.get("new", "")),
                _cell(record.get("from_col", "")),
                _cell(record.get("to_col", "")),
            )
        )
    if not rows:
        raise ValueError("empty rows")
    return Action(rows, d.get("has_summary", False))


def _regroup(rows):
    actions = []
    buf = []
    buf_group = "user"

    def close(summary):
        nonlocal buf
        has_summary = False
        if summary is None and buf_group in ("import", "merge"):
            buf.append((buf[0][0], _bare(buf[0][1]), "", "", "", "", ""))
            has_summary = True
        elif summary is not None:
            buf.append(summary)
            has_summary = True
        actions.append(Action(buf, has_summary))
        buf = []

    for row in rows:
        cells = _seven(row)
        typ = cells[1]
        if _is_member(typ):
            group = _group(typ)
            if buf and group != buf_group:
                close(None)
            if not buf:
                buf_group = group
            buf.append(cells)
        elif buf:
            close(cells)
        else:
            actions.append(Action([cells], False))
    if buf:
        close(None)
    return actions


def load_changelog(raw, warnings=None):
    if not raw:
        return []
    try:
        first = raw[0]
        if isinstance(first, dict) and "rows" in first:
            return [_action_from_record(item) for item in raw]
        if isinstance(first, (list, tuple)):
            return _regroup(raw)
    except (KeyError, TypeError, ValueError, AttributeError, IndexError):
        if warnings is not None:
            warnings.append(" - Changelog was reset (unrecognized format)")
        return []
    if warnings is not None:
        warnings.append(" - Changelog was reset (unrecognized format)")
    return []
