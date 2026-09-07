# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

from __future__ import annotations

import unittest

from src.toplevels import (
    Add_Child_Or_Sibling_Id_Popup,
    Add_Top_Id_Popup,
    new_id_and_treeview_label,
)


class _FakeEntry:
    def __init__(self, value):
        self.value = value

    def get_my_value(self):
        return self.value


class _FakePopup:
    def __init__(self, allow_spaces, tv_label_col, ic, id_value, label_value):
        self.C = type(
            "Editor",
            (),
            {
                "allow_spaces_ids_var": allow_spaces,
                "tv_label_col": tv_label_col,
                "ic": ic,
            },
        )()
        self.id_name_display = _FakeEntry(id_value)
        self.id_tv_display = _FakeEntry(label_value)
        self.destroyed = False

    def destroy(self):
        self.destroyed = True


class TestNewIdAndTreeviewLabel(unittest.TestCase):
    def test_keeps_label_spaces_when_label_is_not_id(self):
        result, label = new_id_and_treeview_label(False, False, "New ID", "Panthera tigris")
        self.assertEqual(result, "NewID")
        self.assertEqual(label, "Panthera tigris")

    def test_label_follows_stripped_id_when_label_is_id(self):
        result, label = new_id_and_treeview_label(False, True, "New ID", "ignored")
        self.assertEqual(result, "NewID")
        self.assertEqual(label, "NewID")

    def test_keeps_id_spaces_when_allowed(self):
        result, label = new_id_and_treeview_label(True, False, "New ID", "Panthera tigris")
        self.assertEqual(result, "New ID")
        self.assertEqual(label, "Panthera tigris")


class TestAddIdPopupConfirm(unittest.TestCase):
    def _confirm(self, popup_cls, **kwargs):
        popup = _FakePopup(**kwargs)
        popup_cls.confirm(popup)
        self.assertTrue(popup.destroyed)
        return popup

    def test_top_id_popup_does_not_strip_label_spaces(self):
        popup = self._confirm(
            Add_Top_Id_Popup,
            allow_spaces=False,
            tv_label_col=1,
            ic=0,
            id_value="New ID",
            label_value="Panthera tigris",
        )
        self.assertEqual(popup.result, "NewID")
        self.assertEqual(popup.id_label, "Panthera tigris")

    def test_child_popup_does_not_strip_label_spaces(self):
        popup = self._confirm(
            Add_Child_Or_Sibling_Id_Popup,
            allow_spaces=False,
            tv_label_col=1,
            ic=0,
            id_value="New ID",
            label_value="Panthera tigris",
        )
        self.assertEqual(popup.result, "NewID")
        self.assertEqual(popup.id_label, "Panthera tigris")

    def test_label_matches_id_when_treeview_label_is_id_column(self):
        popup = self._confirm(
            Add_Top_Id_Popup,
            allow_spaces=False,
            tv_label_col=0,
            ic=0,
            id_value="New ID",
            label_value="ignored",
        )
        self.assertEqual(popup.result, "NewID")
        self.assertEqual(popup.id_label, "NewID")
