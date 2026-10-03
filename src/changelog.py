# SPDX-License-Identifier: AGPL-3.0-only
# Copyright © ragardner

"""The changelog.

A group is the lines recorded for one edit. The time is stored once, on the
group. rows() builds the seven-column rows for the Changelog window and for
an export. App data stores the groups. The undo stack is separate.
"""

import datetime
import getpass


def action_timestamp():
    return datetime.datetime.now().replace(microsecond=0).isoformat()


def os_username():
    """Operating-system login, or "" when it cannot be read."""
    try:
        user = getpass.getuser()
    except Exception:
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


class Line:
    """One recorded edit. type is stored as it was passed in."""

    __slots__ = ("from_col", "id", "new", "old", "to_col", "type")

    def __init__(self, type, id_="", old="", new="", from_col="", to_col=""):
        self.type = _cell(type)
        self.id = _cell(id_)
        self.old = _cell(old)
        self.new = _cell(new)
        self.from_col = _cell(from_col)
        self.to_col = _cell(to_col)

    def as_list(self):
        return [self.type, self.id, self.old, self.new, self.from_col, self.to_col]


class Group:
    """lines are the edits. summary is the closing line, when there is one."""

    __slots__ = ("lines", "summary", "time")

    def __init__(self, time, lines, summary=None):
        if not lines:
            raise ValueError("empty group")
        self.time = _cell(time)
        self.lines = list(lines)
        self.summary = summary

    @property
    def has_summary(self):
        return self.summary is not None

    @property
    def n(self):
        return len(self.lines)

    def first_stamp(self):
        return self.time

    def last_stamp(self):
        return self.time

    def rows(self):
        out = [_seven(self.time, line) for line in self.lines]
        if self.summary is not None:
            out.append(_seven(self.time, self.summary))
        return out

    def to_save(self):
        item = {"time": self.time, "lines": [line.as_list() for line in self.lines]}
        if self.summary is not None:
            item["summary"] = self.summary.as_list()
        return item


def _seven(time, line):
    return (time, line.type, line.id, line.old, line.new, line.from_col, line.to_col)


class Changelog:
    """actions are finished groups. pending is the edit that has not finished."""

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

    def rows(self):
        return [row for group in self.actions for row in group.rows()]

    def display_rows(self):
        return self.rows()

    def session_rows(self):
        return [row for group in self.actions[self.opened_at :] for row in group.rows()]

    def to_save(self):
        return [group.to_save() for group in self.actions]

    def last_type(self):
        if not self.actions:
            return ""
        group = self.actions[-1]
        if group.summary is not None:
            return group.summary.type
        return group.lines[-1].type

    def action_index(self, row_index):
        seen = 0
        for i, group in enumerate(self.actions):
            seen += len(group.lines)
            if group.summary is not None:
                seen += 1
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
        return Line(typ, id_, old, new, from_col, to_col)

    def _close(self, summary=None, singular=None, lines=None):
        if lines is None:
            if not self.pending:
                return None
            lines = self.pending
            self.pending = []
        if singular is not None:
            lines[-1].type = _cell(singular)
            summary = None
        if not lines:
            return None
        group = Group(self.now(), lines, summary)
        self.actions.append(group)
        self.on_unsaved()
        return group

    def seal(self):
        """Close lines left open by an edit that did not finish.

        append() stores its line as the summary when lines are still
        pending, so a later edit calls this before it logs or snapshots.
        """
        if not self.pending:
            return None
        return self._close()

    def append(self, change, id_="", old="", new="", from_col="", to_col=""):
        line = self._line(change, id_, old, new, from_col, to_col)
        if self.pending:
            self._close(summary=line)
            return
        self._close(lines=[line])

    def append_no_unsaved(self, change, id_="", old="", new="", from_col="", to_col=""):
        self.pending.append(self._line(change, id_, old, new, from_col, to_col))

    def singular(self, text):
        text = _cell(text)
        if not self.pending:
            if self.actions:
                group = self.actions[-1]
                if group.summary is not None:
                    group.summary.type = text
                else:
                    group.lines[-1].type = text
            self.on_unsaved()
            return
        self._close(singular=text)

    def finish_plain(self):
        return self.seal()


# Types that belong to a larger edit in a changelog saved by the previous code.
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


def _is_member(typ):
    return typ.startswith(_MEMBER_PREFIXES) or typ.endswith(("|", "| "))


def _group(typ):
    if typ.startswith("Imported change |"):
        return "import"
    if typ.startswith("Merge |"):
        return "merge"
    return "user"


def _pad7(row):
    cells = [_cell(c) for c in list(row)[:7]]
    if len(cells) < 7:
        cells.extend([""] * (7 - len(cells)))
    return tuple(cells)


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


def _line_from_six(row):
    cells = [_cell(c) for c in list(row)[:6]]
    if len(cells) < 6:
        cells.extend([""] * (6 - len(cells)))
    return Line(*cells)


def _group_from_saved_rows(rows, has_summary):
    """One old seven-cell action. The group time is the first row's time."""
    summary = None
    lines_src = rows
    if has_summary and len(rows) > 1:
        lines_src = rows[:-1]
        last = rows[-1]
        summary = Line(last[1], last[2], last[3], last[4], last[5], last[6])
    if not lines_src:
        raise ValueError("empty group")
    lines = [Line(r[1], r[2], r[3], r[4], r[5], r[6]) for r in lines_src]
    return Group(rows[0][0], lines, summary)


def _group_from_record(d):
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
    return _group_from_saved_rows(rows, d.get("has_summary", False))


def _group_from_new(d):
    lines = [_line_from_six(row) for row in d["lines"]]
    if not lines:
        raise ValueError("empty group")
    summary = _line_from_six(d["summary"]) if d.get("summary") is not None else None
    return Group(d.get("time", ""), lines, summary)


def _regroup(rows):
    """Turn a previous save, a flat list of seven-cell rows, into groups.

    The type text is kept, including a trailing | and an import or merge
    prefix. An import or merge run with no closing row gets one, as before.
    """
    found = []
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
        found.append(_group_from_saved_rows(buf, has_summary))
        buf = []

    for row in rows:
        cells = _pad7(row)
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
            found.append(_group_from_saved_rows([cells], False))
    if buf:
        close(None)
    return found


def load_changelog(raw, warnings=None):
    if not raw:
        return []
    try:
        first = raw[0]
        if isinstance(first, dict) and "lines" in first:
            return [_group_from_new(item) for item in raw]
        if isinstance(first, dict) and "rows" in first:
            return [_group_from_record(item) for item in raw]
        if isinstance(first, (list, tuple)):
            return _regroup(raw)
    except (KeyError, TypeError, ValueError, AttributeError, IndexError):
        if warnings is not None:
            warnings.append(" - Changelog was reset (unrecognized format)")
        return []
    if warnings is not None:
        warnings.append(" - Changelog was reset (unrecognized format)")
    return []
