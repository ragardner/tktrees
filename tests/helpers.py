# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Shared tree / changelog checks used by Session, GUI, and CLI tests."""

from __future__ import annotations

from src.changelog import Change, ChangeRow
from src.session import Session
from tests.fixtures import ANIMALS_TABLE, FOREST_TABLE, TWO_HIER_TABLE


def load_table(rows, id_col=0, parent_cols=None):
    if parent_cols is None:
        parent_cols = [1]
    s = Session()
    out = s.load_table(rows, id_col=id_col, parent_cols=parent_cols, fmt=0)
    assert out["ok"], out
    return s


def animals():
    return load_table(ANIMALS_TABLE)


def forest():
    return load_table(FOREST_TABLE)


def two_hier():
    return load_table(TWO_HIER_TABLE, parent_cols=[1, 2])


def names(s, iids):
    return [s.nodes[i].name for i in iids]


def children_of(s, parent, h=None):
    if h is None:
        h = s.pc
    return names(s, s.nodes[parent.lower()].cn[h])


def tops_of(s, h=None):
    saved = s.pc
    if h is None:
        h = s.pc
    try:
        s.pc = h
        return names(s, s.top_iids())
    finally:
        s.pc = saved


def expected_sorted_cn(s, cn, h):
    return s.sort_node_cn(list(cn), h)


def expected_tops(s, h):
    wc = []
    woc = []
    for iid, node in s.nodes.items():
        if node.ps[h] == "":
            if node.cn[h]:
                wc.append(iid)
            else:
                woc.append(iid)
    from src.functions import sort_key

    return [s.nodes[i].name for i in sorted(wc, key=sort_key) + sorted(woc, key=sort_key)]


def assert_changelog_sound(test, s, note=""):
    msg = note or "changelog"
    for i, ch in enumerate(s.changelog):
        test.assertIsInstance(ch, Change, f"{msg} [{i}] not a Change")
        test.assertTrue(ch.rows, f"{msg} [{i}] empty rows")
        test.assertIsInstance(ch.label, str, f"{msg} [{i}] label")
        test.assertGreaterEqual(ch.n, 1, f"{msg} [{i}] n")
        for j, row in enumerate(ch.rows):
            test.assertIsInstance(row, ChangeRow, f"{msg} [{i}.{j}]")
            test.assertIs(row.change, ch, f"{msg} [{i}.{j}] backpointer")
            test.assertIsInstance(row.display_type, str, f"{msg} [{i}.{j}] display_type")
            test.assertEqual(len(row), 5, f"{msg} [{i}.{j}] len")
            test.assertEqual(row[0], ch.at, f"{msg} [{i}.{j}] date")
            test.assertEqual(row[1], row.display_type, f"{msg} [{i}.{j}] type cell")


def assert_one_action(test, s, n_log, n_vs, note="", label=None):
    msg = note or "action"
    test.assertEqual(len(s.changelog), n_log + 1, f"{msg} changelog count")
    test.assertEqual(len(s.vs), n_vs + 1, f"{msg} snapshot count")
    if label is not None:
        test.assertEqual(s.changelog[-1].label, label, f"{msg} label")
    assert_changelog_sound(test, s, msg)


def assert_sound(test, s, note="", check_order=True):
    """Pointers, rows, tops, unique children, and (if auto-sort) child order."""
    msg = note or "tree"
    test.assertEqual(len(s.rns), len(s.data), msg)
    test.assertEqual(set(s.rns), {row[s.ic].lower() for row in s.data if row[s.ic]}, msg)
    for iid, rn in s.rns.items():
        test.assertEqual(s.data[rn][s.ic].lower(), iid, msg)
        test.assertIn(iid, s.nodes, msg)
        test.assertEqual(len(s.data[rn]), s.row_len, msg)
    for iid, node in s.nodes.items():
        test.assertIn(iid, s.rns, msg)
        test.assertEqual(set(node.ps), set(s.hiers), msg)
        test.assertEqual(set(node.cn), set(s.hiers), msg)
        for h in s.hiers:
            pk = node.ps[h]
            cn = node.cn[h]
            test.assertEqual(len(cn), len(set(cn)), f"{msg} duplicate child {iid} h={h}")
            if pk is None:
                for other in s.nodes.values():
                    test.assertNotIn(iid, other.cn[h], msg)
                if not s.auto_sort_nodes_bool and h in s.topnodes_order:
                    test.assertNotIn(iid, s.topnodes_order[h], msg)
                continue
            if pk == "":
                test.assertEqual(s.data[s.rns[iid]][h], "", msg)
            else:
                test.assertIn(pk, s.nodes, msg)
                test.assertIn(iid, s.nodes[pk].cn[h], msg)
                test.assertEqual(s.data[s.rns[iid]][h], s.nodes[pk].name, msg)
            for child in cn:
                test.assertIn(child, s.nodes, msg)
                test.assertEqual(s.nodes[child].ps[h], iid, msg)
            actual_children = {oid for oid, on in s.nodes.items() if on.ps[h] == iid}
            test.assertEqual(set(cn), actual_children, f"{msg} cn set {iid} h={h}")
            if check_order and s.auto_sort_nodes_bool and cn:
                test.assertEqual(
                    list(cn),
                    expected_sorted_cn(s, cn, h),
                    f"{msg} cn order {s.nodes[iid].name} h={h}",
                )
    for h in s.hiers:
        actual_tops = {iid for iid, n in s.nodes.items() if n.ps[h] == ""}
        if not s.auto_sort_nodes_bool:
            test.assertIn(h, s.topnodes_order, f"{msg} missing topnodes {h}")
            test.assertEqual(set(s.topnodes_order[h]), actual_tops, f"{msg} topnodes {h}")
            test.assertEqual(len(s.topnodes_order[h]), len(set(s.topnodes_order[h])), msg)
        if check_order and s.auto_sort_nodes_bool:
            test.assertEqual(tops_of(s, h), expected_tops(s, h), f"{msg} top order h={h}")
            test.assertEqual(set(tops_of(s, h)), {s.nodes[i].name for i in actual_tops}, f"{msg} top set h={h}")
    assert_changelog_sound(test, s, msg)
