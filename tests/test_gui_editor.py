# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Drive Tree_Editor through Session wrappers. Skipped when Tk has no display."""

from __future__ import annotations

import csv
import os
import tempfile
import unittest
from contextlib import suppress

from openpyxl import Workbook

from src.changelog import display_rows
from src.session import Session
from tests.fixtures import ANIMALS_TABLE
from tests.helpers import assert_sound, children_of

FMT5 = [
    ["Animals", "All animals"],
    ["", "Cats", "Cat family"],
    ["", "", "Lion", "Lion"],
]
FMT7 = [
    ["L1", "L2", "L3", "Name"],
    ["Animals", "", "", "All animals"],
    ["", "Cats", "", "Cat family"],
    ["", "", "Lion", "Lion"],
]
_GUI_SNAP_KEYS = (
    "tv_label_col",
    "saved_info",
    "mirror_bool",
    "sheet_column_alignments",
    "sheet_col_positions",
    "sheet_row_positions",
)


def _tk_works():
    try:
        import tkinter

        root = tkinter.Tk()
        root.withdraw()
        root.update_idletasks()
        root.destroy()
        return True
    except Exception:
        return False


def _silence(appmod, te, widgets, toplevels):
    class Silent:
        def __init__(self, *a, **k):
            self.boolean = True
            self.result = None

    for mod in (appmod, te, widgets, toplevels):
        for name in (
            "First_Start_Popup",
            "Error",
            "Text_Popup",
            "Ask_Confirm",
            "Ask_Confirm_Quit",
            "Help_Popup",
        ):
            if hasattr(mod, name):
                setattr(mod, name, Silent)
    appmod.AppGUI.save_cfg = lambda self, *a, **k: None


def _close(app):
    with suppress(Exception):
        app.unbind("<Configure>")
    with suppress(Exception):
        app.withdraw()
    with suppress(Exception):
        app.destroy()


def _ids(ed):
    return [row[ed.ic] for row in ed.sheet.MT.data]


def _parent(ed, name):
    node = ed.nodes[name.lower()]
    p = node.ps[ed.pc]
    if p in (None, ""):
        return p
    return ed.nodes[p].name


_TK_OK = _tk_works()


@unittest.skipUnless(_TK_OK or os.environ.get("REQUIRE_GUI_TESTS"), "Tk display required")
class TestGuiEditor(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not _TK_OK:
            raise RuntimeError("Tk display required (REQUIRE_GUI_TESTS is set)")
        import src.app as appmod
        import src.tree_editor as te
        from src import toplevels, widgets

        _silence(appmod, te, widgets, toplevels)
        cls.appmod = appmod

    def _app(self):
        app = self.appmod.AppGUI(["TKTREES.pyw"])
        app.configsettings["First GUI start"] = False
        app.unbind("<Configure>")
        app.withdraw()
        app.update_idletasks()
        self.addCleanup(_close, app)
        return app, app.frames["tree_edit"]

    def _tree(self, ed, *pairs):
        for name, parent in pairs:
            self.assertTrue(ed.add(name, parent, errors=False), msg=name)
        ed.redo_tree_display()

    def _assert_one_list(self, ed, where):
        self.assertIs(
            ed.session.data,
            ed.sheet.MT.data,
            msg=f"{where}: session.data is not sheet.MT.data",
        )

    def _set_col_width(self, ed, col, width):
        widths = list(ed.sheet.get_column_widths())
        widths[col] = width
        ed.sheet.set_column_widths(widths)
        return [int(w) for w in ed.sheet.get_column_widths()]

    def _assert_gui_snapshot(self, ed, where):
        self.assertTrue(ed.vs, msg=f"{where}: no snapshot")
        rd = ed.vs[-1]["required_data"]
        missing = [k for k in _GUI_SNAP_KEYS if k not in rd]
        self.assertEqual(missing, [], msg=f"{where}: snapshot missing {missing}")

    def _tree_walk_ids(self, ed):
        return [row[ed.ic] for row in ed.tree.data]

    def _assert_tree_matches_nodes(self, ed, where):
        self.assertEqual(
            [name.lower() for name in self._tree_walk_ids(ed)],
            list(ed.pc_iids()),
            msg=f"{where}: tree walk != pc_iids",
        )
        self.assertEqual(sorted(ed.tree.RI.rns), sorted(n for n in ed.nodes if ed.nodes[n].ps[ed.pc] is not None))

    def _set_format(self, app, fmt):
        cs = app.frames["column_selection"]
        cs.data_format_selector.format_dropdown.current(fmt)
        cs.data_format_selector.dropdown_select()
        return cs

    def _open_path(self, path):
        app = self.appmod.AppGUI(["TKTREES.pyw", path])
        app.configsettings["First GUI start"] = False
        app.unbind("<Configure>")
        app.withdraw()
        app.update_idletasks()
        self.addCleanup(_close, app)
        return app, app.frames["tree_edit"]

    def _build_from_rows(self, app, ed, rows, fmt):
        ed.set_records([list(r) for r in rows])
        ncols = max(map(len, rows), default=0)
        cs = app.frames["column_selection"]
        cs.populate(list(map(str, range(1, ncols + 1))))
        self._set_format(app, fmt)
        cs.try_to_build_tree()
        app.update_idletasks()
        return cs

    def test_session_and_sheet_stay_the_same_list(self):
        _app, ed = self._app()
        self._assert_one_list(ed, "new")
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Dogs", "Animals"),
        )
        self._assert_one_list(ed, "after add")
        ed.sort_sheet("ID", "ASCENDING")
        self._assert_one_list(ed, "after column sort")
        self.assertEqual(_ids(ed), ["Animals", "Cats", "Dogs", "Lion"])
        ed.sort_sheet_walk()
        self._assert_one_list(ed, "after tree-walk sort")
        self.assertEqual(_ids(ed), ["Animals", "Cats", "Lion", "Dogs"])
        ed.undo()
        self._assert_one_list(ed, "after undo tree-walk sort")
        ed.del_id(["dogs"])
        self._assert_one_list(ed, "after delete")
        ed.rebuild_tree()
        self._assert_one_list(ed, "after rebuild_tree")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])

    def test_in_file_ops_keep_column_widths(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Dogs", "Animals"),
        )
        custom = 317
        self._set_col_width(ed, 0, custom)
        self.assertEqual(int(ed.sheet.get_column_widths()[0]), custom)

        heights = list(ed.sheet.get_row_heights())
        lion_h = 64
        heights[ed.rns["lion"]] = lion_h
        ed.sheet.set_row_heights(heights)

        ed.sort_sheet_walk()
        self._assert_one_list(ed, "after tree-walk sort")
        self.assertEqual(int(ed.sheet.get_column_widths()[0]), custom)
        self.assertEqual(int(ed.sheet.get_row_heights()[ed.rns["lion"]]), lion_h)

        ed.rebuild_tree()
        self._assert_one_list(ed, "after rebuild_tree")
        self.assertEqual(int(ed.sheet.get_column_widths()[0]), custom)

    def test_undo_add_removes_id_and_keeps_one_list(self):
        _app, ed = self._app()
        self.assertTrue(ed.add("Animals", "", errors=False))
        self.assertEqual(ed.vs[-1]["row"].get("added_or_changed"), "added")
        self.assertIsNotNone(ed.vs[-1]["row"].get("rn"))
        self._assert_gui_snapshot(ed, "after add")
        self._assert_one_list(ed, "after add")
        ed.undo()
        self.assertNotIn("animals", ed.nodes)
        self.assertEqual(ed.sheet.MT.data, [])
        self._assert_one_list(ed, "after undo add")

    def test_add_child_with_treeview_label_is_one_changelog_action(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"))
        ed.tv_label_col = 1
        ed.selected_ID = "Cats"
        n_log = len(ed.changelog)
        n_vs = len(ed.vs)

        class FakePopup:
            def __init__(self, *a, **k):
                self.result = "Tiger"
                self.id_label = "Panthera tigris"

        import src.tree_editor as te

        orig = te.Add_Child_Or_Sibling_Id_Popup
        te.Add_Child_Or_Sibling_Id_Popup = FakePopup
        try:
            ed.add_child_node()
        finally:
            te.Add_Child_Or_Sibling_Id_Popup = orig

        self.assertIn("tiger", ed.nodes)
        self.assertEqual(ed.sheet.MT.data[ed.rns["tiger"]][1], "Panthera tigris")
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(len(ed.vs), n_vs + 1)
        ch = ed.changelog[-1]
        self.assertEqual(ch.label, "Add ID")
        self.assertEqual([r.display_type for r in ch.rows], ["Add ID", "Edit cell"])
        self.assertEqual(ch.rows[1].new, "Panthera tigris")
        self._assert_one_list(ed, "after add with label")

        ed.undo()
        self.assertNotIn("tiger", ed.nodes)
        self.assertEqual(len(ed.changelog), n_log)
        self.assertEqual(len(ed.vs), n_vs)
        self._assert_one_list(ed, "after undo add with label")

    def test_edit_cell_rebuild_writes_failed_id_edit(self):
        app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        self.assertFalse(ed.change_ID_name("Lion", "Cats", errors=False))
        rn = ed.rns["lion"]
        n_log = len(ed.changelog)
        ed.edit_cell_rebuild(rn, ed.ic, "Cats")
        self.assertFalse(any(row[ed.ic] == "Lion" for row in ed.sheet.MT.data))
        self._assert_one_list(ed, "after id rebuild")
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(ed.changelog[-1].label, "Edit cell")
        self.assertTrue(app.unsaved_changes)
        self.assertTrue(app.title().endswith("*"))

    def test_add_does_not_pad_thousands_of_blank_rows(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        self.assertEqual(len(ed.sheet.MT.data), 3)
        self.assertEqual(_ids(ed), ["Animals", "Cats", "Lion"])
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self.assertEqual([row[ed.ic] for row in ed.tree.data], ["Animals", "Cats", "Lion"])

    def test_rename_detail_sort_move_delete_undo(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Tiger", "Cats"),
            ("Dogs", "Animals"),
        )
        self.assertTrue(ed.change_ID_name("Tiger", "Tigers", errors=False))
        self.assertIn("tigers", ed.nodes)
        self.assertNotIn("tiger", ed.nodes)

        rn = ed.rns["lion"]
        ed.snapshot_ctrl_x_v_del_key()
        ed.vs[-1]["cells"][(rn, 1)] = f"{ed.sheet.MT.data[rn][1]}"
        ed.edit_cell_single(rn, 1, "The lion")
        self.assertEqual(ed.sheet.MT.data[ed.rns["lion"]][1], "The lion")

        before_sort = [list(row) for row in ed.sheet.MT.data]
        ed.sort_sheet("ID", "ASCENDING")
        ed.undo()
        self.assertEqual([list(row) for row in ed.sheet.MT.data], before_sort)

        ed.snapshot_paste_id()
        self.assertTrue(ed.cut_paste("Lion", "Cats", ed.pc, "Dogs", errors=False))
        self.assertEqual(_parent(ed, "Lion"), "Dogs")
        ed.undo()
        self.assertEqual(_parent(ed, "Lion"), "Cats")

        ed.del_id(["dogs"])
        self.assertNotIn("dogs", ed.nodes)
        ed.undo()
        self.assertIn("dogs", ed.nodes)

    def test_six_delete_variants(self):
        _app, ed = self._app()

        def fresh():
            ed.reset_tree(False)
            ed.session.new(discard=True)
            ed.set_records(ed.session.data)
            ed.tv_label_col = 0
            ed.populate()
            self._tree(
                ed,
                ("Animals", ""),
                ("Cats", "Animals"),
                ("Lion", "Cats"),
                ("Tiger", "Cats"),
                ("Dogs", "Animals"),
                ("Wolf", "Dogs"),
            )
            return ed

        fresh()
        ed.del_id(["wolf"])
        self.assertNotIn("wolf", ed.nodes)
        self.assertIn("dogs", ed.nodes)

        ed = fresh()
        ed.del_id(["cats"])
        self.assertNotIn("cats", ed.nodes)
        self.assertEqual(_parent(ed, "Lion"), "Animals")
        self.assertEqual(_parent(ed, "Tiger"), "Animals")

        ed = fresh()
        ed.selected_ID = "cats"
        ed.del_id_orphan()
        self.assertNotIn("cats", ed.nodes)
        self.assertEqual(_parent(ed, "Lion"), "")
        self.assertEqual(_parent(ed, "Tiger"), "")

        ed = fresh()
        ed.del_id_children(["cats"])
        self.assertNotIn("cats", ed.nodes)
        self.assertNotIn("lion", ed.nodes)
        self.assertNotIn("tiger", ed.nodes)
        self.assertIn("wolf", ed.nodes)

        ed = fresh()
        ed.del_id_all(["dogs"])
        self.assertNotIn("dogs", ed.nodes)
        self.assertEqual(_parent(ed, "Wolf"), "Animals")

        ed = fresh()
        ed.selected_ID = "cats"
        ed.del_id_all_orphan()
        self.assertNotIn("cats", ed.nodes)
        self.assertEqual(_parent(ed, "Lion"), "")
        self.assertEqual(_parent(ed, "Tiger"), "")

        ed = fresh()
        ed.del_id_children_all(["cats"])
        self.assertNotIn("cats", ed.nodes)
        self.assertNotIn("lion", ed.nodes)
        self.assertIn("dogs", ed.nodes)

    def test_csv_build_tree_matches_file(self):
        app, ed = self._app()
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = os.path.join(folder.name, "animals.csv")
        with open(path, "w", newline="", encoding="utf-8") as fh:
            csv.writer(fh).writerows(ANIMALS_TABLE)
        app.open_dict["filepath"] = path
        app.created_new = False
        app.load_from_file()
        if not hasattr(ed, "new_sheet"):
            ed.new_sheet = []
        app.frames["column_selection"].try_to_build_tree()
        app.update_idletasks()
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self.assertEqual(len(ed.sheet.MT.data), 3)

    def test_copy_into_new_hierarchy_and_columns(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        src_hier = ed.pc
        ed.add_hier_col(len(ed.headers), "PARENT_2")
        self.assertIn("PARENT_2", [h.name for h in ed.headers])
        ed.pc = ed.hiers[-1]
        ed.snapshot_paste_id()
        self.assertTrue(ed.copy_paste("Lion", src_hier, "", errors=False))
        self.assertEqual(ed.nodes["lion"].ps[ed.pc], "")
        self.assertEqual(ed.nodes["lion"].ps[src_hier], "cats")

        ed.rename_col(1, "Name")
        self.assertEqual(ed.headers[1].name, "Name")
        n = len(ed.headers)
        ed.add_col(n, "NOTE", "Text")
        self.assertEqual(ed.headers[-1].name, "NOTE")
        self.assertEqual(len(ed.sheet.MT.data[0]), len(ed.headers))

    def test_cli_xlsx_opens_in_gui_with_program_data(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = os.path.join(folder.name, "animals.xlsx")
        s = Session()
        s.new()
        s.add("Animals", "")
        s.add("Cats", "Animals")
        s.add("Lion", "Cats")
        s.tag(["Lion"])
        s.set_option("auto-sort", "off")
        out = s.save_path(path, overwrite=True)
        self.assertTrue(out["ok"], out)

        app, ed = self._app()
        app.open_dict["filepath"] = path
        app.created_new = False
        app.load_from_file()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self.assertEqual(ed.nodes["lion"].ps[ed.pc], "cats")
        self.assertIn("lion", ed.tagged_ids)
        self.assertFalse(ed.auto_sort_nodes_bool)
        self.assertTrue(any(ch.label == "Add ID" for ch in ed.changelog))
        self.assertEqual([row[ed.ic] for row in ed.tree.data], ["Animals", "Cats", "Lion"])
        self.assertEqual(len(ed.sheet.MT.data), 3)

    def test_cli_json_opens_in_gui_with_program_data(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = os.path.join(folder.name, "animals.json")
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        self.assertTrue(s.save_path(path, overwrite=True)["ok"])
        app, ed = self._app()
        app.open_dict["filepath"] = path
        app.created_new = False
        app.load_from_file()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self._assert_one_list(ed, "json program_data")

    def test_cut_paste_children_copy_all_tag_sort_validation(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Tiger", "Cats"),
            ("Dogs", "Animals"),
        )
        ed.snapshot_paste_id()
        self.assertTrue(ed.cut_paste_children("Cats", "Dogs", ed.pc, errors=False))
        self.assertEqual(_parent(ed, "Lion"), "Dogs")
        self.assertEqual(_parent(ed, "Tiger"), "Dogs")
        self.assertEqual(ed.nodes["cats"].cn[ed.pc], [])
        self._assert_one_list(ed, "after cut_paste_children")

        src = ed.pc
        ed.add_hier_col(len(ed.headers), "PARENT_2")
        ed.pc = ed.hiers[-1]
        ed.snapshot_paste_id()
        self.assertTrue(ed.copy_paste_all("Animals", src, "", errors=False))
        self.assertEqual(ed.nodes["animals"].ps[ed.pc], "")
        self.assertEqual(ed.nodes["cats"].ps[ed.pc], "animals")

        ed.tag_ids(selection=["lion"], toggle=False, do_tree=False)
        self.assertIn("lion", ed.tagged_ids)
        ed.toggle_sort_all_nodes(False)
        self.assertFalse(ed.auto_sort_nodes_bool)
        ed.toggle_sort_all_nodes(True)
        self.assertTrue(ed.auto_sort_nodes_bool)
        out = ed.session.set_validation(1, ["a", "b"])
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["result"]["validation"][0], "")
        self._assert_one_list(ed, "after tag/sort/validation")

    def test_parent_cell_edit_reparents(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"), ("Dogs", "Animals"))
        ed.snapshot_paste_id()
        self.assertTrue(ed.cut_paste_edit_cell("Lion", "Cats", ed.pc, "Dogs"))
        self.assertEqual(_parent(ed, "Lion"), "Dogs")
        self._assert_one_list(ed, "after parent cell edit")

    def test_new_document_defaults_and_identity(self):
        _app, ed = self._app()
        self._assert_one_list(ed, "new document")
        self.assertEqual([h.name for h in ed.headers], ["ID", "DETAIL_1", "PARENT_1"])
        self.assertEqual(ed.ic, 0)
        self.assertEqual(ed.pc, 2)
        self.assertEqual(ed.hiers, [2])
        self.assertEqual(ed.tv_label_col, 0)
        self.assertEqual(ed.sheet.MT.data, [])
        self.assertEqual(ed.nodes, {})
        self.assertEqual(_app.current_frame, "tree_edit")
        self.assertEqual(_app.open_dict["filepath"], "New sheet")

    def test_gui_snapshots_include_view_fields(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        self._assert_gui_snapshot(ed, "after add")
        ed.sort_sheet("ID", "ASCENDING")
        self._assert_gui_snapshot(ed, "after sort")
        ed.del_id(["lion"])
        self._assert_gui_snapshot(ed, "after delete")
        ed.add_col(len(ed.headers), "NOTE", "Text")
        self._assert_gui_snapshot(ed, "after add col")

    def test_treeview_follows_sheet_after_mutations(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Dogs", "Animals"),
        )
        self._assert_tree_matches_nodes(ed, "after add")
        self.assertEqual(self._tree_walk_ids(ed), ["Animals", "Cats", "Lion", "Dogs"])
        ed.del_id(["cats"])
        self._assert_tree_matches_nodes(ed, "after delete")
        self.assertIn("Lion", self._tree_walk_ids(ed))
        self.assertNotIn("Cats", self._tree_walk_ids(ed))
        ed.undo()
        ed.redo_tree_display()
        self._assert_tree_matches_nodes(ed, "after undo")
        self._assert_one_list(ed, "after undo treeview")

    def test_undo_column_drag_to_start_keeps_ids(self):
        _app, ed = self._app()
        self._tree(ed, ("001", ""), ("002", "001"), ("003", "002"))
        self.assertEqual(ed.ic, 0)
        self.assertEqual(ed.pc, 2)
        self.assertIn("001", ed.rns)
        n_cols = len(ed.headers)
        last = n_cols - 1
        ed.snapshot_begin_drag_cols()
        _data_idxs, _disp_idxs, event_data = ed.sheet.mapping_move_columns(
            {last: 0},
            undo=False,
            emit_event=False,
            redraw=False,
        )
        ed.snapshot_drag_cols(event_data)
        self._assert_one_list(ed, "after column drag")
        self.assertEqual(ed.ic, 1, msg="ID column should shift right when last col moves to start")
        self.assertIn("001", ed.rns)
        self.assertEqual(_ids(ed), ["001", "002", "003"])
        ed.undo()
        self._assert_one_list(ed, "after undo column drag")
        self.assertIn("001", ed.rns)
        self.assertEqual(ed.ic, 0)
        self.assertEqual(ed.pc, 2)
        self.assertEqual(_ids(ed), ["001", "002", "003"])
        self.assertEqual(_parent(ed, "002"), "001")
        self._assert_tree_matches_nodes(ed, "after undo column drag")

    def test_undo_tree_column_drag_to_start_keeps_ids(self):
        _app, ed = self._app()
        self._tree(ed, ("001", ""), ("002", "001"), ("003", "002"))
        last = len(ed.headers) - 1
        ed.snapshot_begin_drag_cols()
        _data_idxs, _disp_idxs, event_data = ed.tree.mapping_move_columns(
            {last: 0},
            undo=False,
            emit_event=False,
            redraw=False,
        )
        orig_has_focus = ed.tree.has_focus
        ed.tree.has_focus = lambda: True
        try:
            ed.snapshot_drag_cols(event_data)
        finally:
            ed.tree.has_focus = orig_has_focus
        self._assert_one_list(ed, "after tree column drag")
        self.assertEqual(ed.ic, 1)
        self.assertIn("001", ed.rns)
        ed.undo()
        self._assert_one_list(ed, "after undo tree column drag")
        self.assertIn("001", ed.rns)
        self.assertEqual(ed.ic, 0)
        self.assertEqual(_ids(ed), ["001", "002", "003"])
        self.assertEqual(_parent(ed, "002"), "001")
        self._assert_tree_matches_nodes(ed, "after undo tree column drag")

    def test_undo_row_drag_keeps_ids(self):
        _app, ed = self._app()
        self._tree(ed, ("001", ""), ("002", "001"), ("003", "002"))
        ed.snapshot_begin_drag_rows()
        _data_idxs, _disp_idxs, event_data = ed.sheet.mapping_move_rows(
            {2: 0},
            undo=False,
            emit_event=False,
            redraw=False,
        )
        ed.snapshot_drag_rows(event_data)
        self._assert_one_list(ed, "after row drag")
        self.assertIn("001", ed.rns)
        ed.undo()
        self._assert_one_list(ed, "after undo row drag")
        self.assertIn("001", ed.rns)
        self.assertEqual(sorted(_ids(ed)), ["001", "002", "003"])
        self.assertEqual(_parent(ed, "002"), "001")
        self._assert_tree_matches_nodes(ed, "after undo row drag")

    def test_add_delete_column_undo_keeps_sheet_and_tree(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        n = len(ed.headers)
        ed.add_col(n, "NOTE", "Text")
        self.assertEqual(ed.headers[-1].name, "NOTE")
        self.assertEqual(len(ed.sheet.MT.data[0]), len(ed.headers))
        self._assert_one_list(ed, "after add col")
        ed.redo_tree_display()
        self.assertEqual(len(ed.tree.data[0]), len(ed.headers))
        ed.undo()
        self.assertEqual(len(ed.headers), n)
        self.assertEqual(len(ed.sheet.MT.data[0]), n)
        self._assert_one_list(ed, "after undo add col")
        self._assert_tree_matches_nodes(ed, "after undo add col")

        ed.add_hier_col(len(ed.headers), "PARENT_2")
        self.assertIn("PARENT_2", [h.name for h in ed.headers])
        note = next((i for i, h in enumerate(ed.headers) if h.name == "DETAIL_1"), None)
        self.assertIsNotNone(note)
        ed.del_cols([note])
        self.assertNotIn("DETAIL_1", [h.name for h in ed.headers])
        self._assert_one_list(ed, "after del col")
        ed.undo()
        self.assertIn("DETAIL_1", [h.name for h in ed.headers])
        self._assert_one_list(ed, "after undo del col")

    def test_gui_cut_paste_child_undo_uses_view_snapshot(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Dogs", "Animals"),
        )
        ed.selected_ID = "lion"
        ed.cut_ids(["lion"])
        ed.selected_ID = "dogs"
        ed.paste_cut_child()
        self.assertEqual(_parent(ed, "Lion"), "Dogs")
        self._assert_gui_snapshot(ed, "after paste_cut_child")
        self._assert_one_list(ed, "after paste_cut_child")
        self._assert_tree_matches_nodes(ed, "after paste_cut_child")
        ed.undo()
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "after undo paste_cut_child")
        self._assert_tree_matches_nodes(ed, "after undo paste_cut_child")

    def test_gui_save_xlsx_json_csv_reopen(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        ed.tag_ids(selection=["lion"], toggle=False, do_tree=False)
        ed.toggle_sort_all_nodes(False)
        xlsx = os.path.join(folder.name, "animals.xlsx")
        json_path = os.path.join(folder.name, "animals.json")
        csv_path = os.path.join(folder.name, "animals.csv")
        self.assertTrue(ed.save_workbook(xlsx, "Sheet1"))
        self.assertTrue(ed.save_json(json_path))
        self.assertTrue(ed.save_csv(csv_path))

        ed.reset_tree()
        app.open_dict["filepath"] = xlsx
        app.created_new = False
        app.load_from_file()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self.assertIn("lion", ed.tagged_ids)
        self.assertFalse(ed.auto_sort_nodes_bool)
        self._assert_one_list(ed, "xlsx gui save reopen")
        self._assert_tree_matches_nodes(ed, "xlsx gui save reopen")

        ed.reset_tree()
        app.open_dict["filepath"] = json_path
        app.created_new = False
        app.load_from_file()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertIn("lion", ed.tagged_ids)
        self._assert_one_list(ed, "json gui save reopen")

        ed.reset_tree()
        app.open_dict["filepath"] = csv_path
        app.created_new = False
        app.load_from_file()
        self.assertEqual(app.current_frame, "column_selection")
        cs = app.frames["column_selection"]
        if cs.selector.get_id_col() is None:
            cs.selector.set_id_col(0)
        if not cs.selector.get_par_cols():
            cs.selector.set_par_cols([1])
        cs.try_to_build_tree()
        app.update_idletasks()
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "csv gui save reopen")

    def test_xlsx_without_program_data_column_selection(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = os.path.join(folder.name, "plain.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        for row in ANIMALS_TABLE:
            ws.append(row)
        wb.save(path)
        app, ed = self._open_path(path)
        self.assertEqual(app.current_frame, "column_selection")
        cs = app.frames["column_selection"]
        if cs.selector.get_id_col() is None:
            cs.selector.set_id_col(0)
        if not cs.selector.get_par_cols():
            cs.selector.set_par_cols([1])
        cs.try_to_build_tree()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "xlsx no program_data")
        self._assert_tree_matches_nodes(ed, "xlsx no program_data")

    def test_column_selection_indented_formats(self):
        app, ed = self._app()
        self._build_from_rows(app, ed, FMT5, 5)
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "fmt5")
        self._assert_tree_matches_nodes(ed, "fmt5")

        ed.reset_tree()
        self._build_from_rows(app, ed, FMT7, 7)
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "fmt7")

    def test_column_selection_flattened_format(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        s = Session()
        s.load_table(ANIMALS_TABLE, id_col=0, parent_cols=[1], fmt=0)
        path = os.path.join(folder.name, "flat.csv")
        self.assertTrue(s.export_flat(path, hier="Parent", overwrite=True)["ok"])
        app, ed = self._app()
        app.open_dict["filepath"] = path
        app.created_new = False
        app.load_from_file()
        cs = self._set_format(app, 1)
        if not cs.flattened_selector.get_par_cols():
            with open(path, newline="", encoding="utf-8") as fh:
                header = next(csv.reader(fh))
            hier_cols = [i for i, name in enumerate(header) if name.lower().startswith("parent")]
            cs.flattened_selector.set_par_cols(hier_cols)
        cs.try_to_build_tree()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "fmt1")
        self._assert_tree_matches_nodes(ed, "fmt1")

    def test_json_without_program_data_column_selection(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        path = os.path.join(folder.name, "plain.json")
        app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        ed.save_json_with_program_data = False
        self.assertTrue(ed.save_json(path))
        ed.reset_tree()
        app.open_dict["filepath"] = path
        app.created_new = False
        app.load_from_file()
        self.assertEqual(app.current_frame, "column_selection")
        cs = app.frames["column_selection"]
        if cs.selector.get_id_col() is None:
            cs.selector.set_id_col(0)
        if not cs.selector.get_par_cols():
            cs.selector.set_par_cols([1])
        cs.try_to_build_tree()
        app.update_idletasks()
        self.assertEqual(app.current_frame, "tree_edit")
        self.assertEqual(sorted(ed.nodes), ["animals", "cats", "lion"])
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_one_list(ed, "json no program_data")

    def test_delete_promote_sorts_and_tree_matches(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Tiger", "Cats"),
            ("Dogs", "Animals"),
            ("Aardvark", "Animals"),
        )
        n_log, n_vs = len(ed.changelog), len(ed.vs)
        ed.del_id(["cats"])
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(len(ed.vs), n_vs + 1)
        # Dogs has no children here, so all remaining Animals children are leaves.
        self.assertEqual(children_of(ed.session, "Animals"), ["Aardvark", "Dogs", "Lion", "Tiger"])
        self.assertEqual(_parent(ed, "Lion"), "Animals")
        self._assert_tree_matches_nodes(ed, "after promote")
        assert_sound(self, ed.session, "gui-promote")
        ed.undo()
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        self._assert_tree_matches_nodes(ed, "undo promote")
        assert_sound(self, ed.session, "gui-undo-promote")

    def test_six_deletes_stay_sound_and_one_action(self):
        _app, ed = self._app()

        def fresh():
            ed.reset_tree(False)
            ed.session.new(discard=True)
            ed.set_records(ed.session.data)
            ed.tv_label_col = 0
            ed.populate()
            self._tree(
                ed,
                ("Animals", ""),
                ("Cats", "Animals"),
                ("Lion", "Cats"),
                ("Tiger", "Cats"),
                ("Dogs", "Animals"),
                ("Wolf", "Dogs"),
            )
            return len(ed.changelog), len(ed.vs)

        n_log, n_vs = fresh()
        ed.del_id(["wolf"])
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(len(ed.vs), n_vs + 1)
        assert_sound(self, ed.session, "del-id")

        n_log, n_vs = fresh()
        ed.del_id_orphan()
        # no selected_ID — should no-op
        self.assertEqual(len(ed.changelog), n_log)
        ed.selected_ID = "Cats"
        ed.del_id_orphan()
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(_parent(ed, "Lion"), "")
        assert_sound(self, ed.session, "orphan")

        n_log, n_vs = fresh()
        ed.del_id_children(["cats"])
        self.assertNotIn("lion", ed.nodes)
        self.assertEqual(len(ed.changelog), n_log + 1)
        assert_sound(self, ed.session, "del-children")

        n_log, n_vs = fresh()
        ed.del_id_all(["dogs"])
        self.assertEqual(_parent(ed, "Wolf"), "Animals")
        self.assertEqual(len(ed.changelog), n_log + 1)
        assert_sound(self, ed.session, "del-all")

        n_log, n_vs = fresh()
        ed.selected_ID = "Cats"
        ed.del_id_all_orphan()
        self.assertEqual(_parent(ed, "Lion"), "")
        assert_sound(self, ed.session, "all-orphan")

        n_log, n_vs = fresh()
        ed.del_id_children_all(["cats"])
        self.assertNotIn("lion", ed.nodes)
        assert_sound(self, ed.session, "children-all")

    def test_copy_paste_is_one_changelog_action(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"), ("Dogs", "Animals"))
        src = ed.pc
        ed.add_hier_col(len(ed.headers), "H2")
        ed.pc = ed.hiers[-1]
        self.assertTrue(ed.add("Animals", "", errors=False))
        n_log, n_vs = len(ed.changelog), len(ed.vs)
        ed.copied = [{"id": "lion", "hier": src}, {"id": "dogs", "hier": src}]
        ed.selected_ID = "Animals"
        ed.paste_copied_child()
        self.assertEqual(ed.nodes["lion"].ps[ed.pc], "animals")
        self.assertEqual(ed.nodes["dogs"].ps[ed.pc], "animals")
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(len(ed.vs), n_vs + 1)
        ch = ed.changelog[-1]
        self.assertEqual(ch.n, 2)
        self.assertTrue(ch.has_summary)
        self.assertEqual(ch.label, "Copy and paste 2 IDs")
        assert_sound(self, ed.session, "gui-copy-paste")
        ed.undo()
        self.assertIsNone(ed.nodes["lion"].ps[ed.pc])
        assert_sound(self, ed.session, "gui-undo-copy")

    def test_prune_changelog_and_undo(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        n = len(ed.changelog)
        self.assertGreaterEqual(n, 3)
        first_what = ed.changelog[0].rows[0].what
        kept_what = ed.changelog[-1].rows[0].what
        ed.snapshot_prune_changelog(0)
        self.assertNotEqual(ed.changelog[0].rows[0].what, first_what)
        self.assertEqual(ed.changelog[-2].rows[0].what, kept_what)
        self.assertEqual(ed.changelog[-1].label, "Pruned changelog")
        flat = display_rows(ed.changelog)
        self.assertTrue(all(row.change is not None for row in flat))
        ed.undo()
        self.assertEqual(ed.changelog[0].rows[0].what, first_what)
        self._assert_one_list(ed, "after prune undo")

    def test_two_hier_delete_keeps_other_and_tree(self):
        _app, ed = self._app()
        self._tree(ed, ("Root", ""), ("A", "Root"), ("B", "A"), ("C", "Root"))
        ed.add_hier_col(len(ed.headers), "H2")
        ed.pc = ed.hiers[-1]
        self.assertTrue(ed.add("Root", "", errors=False))
        self.assertTrue(ed.copy_paste("B", ed.hiers[0], "Root", errors=False))
        ed.pc = ed.hiers[0]
        ed.del_id(["b"])
        self.assertIsNone(ed.nodes["b"].ps[ed.hiers[0]])
        self.assertEqual(ed.nodes["b"].ps[ed.hiers[-1]], "root")
        assert_sound(self, ed.session, "gui-two-hier")
        self._assert_tree_matches_nodes(ed, "after one-hier delete")

    def test_cut_paste_child_is_one_changelog_action(self):
        _app, ed = self._app()
        self._tree(
            ed,
            ("Animals", ""),
            ("Cats", "Animals"),
            ("Lion", "Cats"),
            ("Dogs", "Animals"),
        )
        n_log, n_vs = len(ed.changelog), len(ed.vs)
        ed.cut = [{"id": "lion", "parent": "cats", "hier": ed.pc}]
        ed.selected_ID = "Dogs"
        ed.paste_cut_child()
        self.assertEqual(_parent(ed, "Lion"), "Dogs")
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(len(ed.vs), n_vs + 1)
        self.assertEqual(ed.changelog[-1].label, "Cut and paste ID")
        assert_sound(self, ed.session, "gui-cut-paste")
        ed.undo()
        self.assertEqual(_parent(ed, "Lion"), "Cats")
        assert_sound(self, ed.session, "gui-undo-cut")

    def test_merge_sheets_uses_session_and_stays_one_list(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"), ("Lion", "Cats"))
        ed.show_warnings = lambda *a, **k: None
        n_log, n_vs = len(ed.changelog), len(ed.vs)

        class FakeMerge:
            result = True
            format_selector_current = 0
            ic = 0
            pcols = [2]
            flattened_pcols = []
            add_new_ids = True
            add_new_dcols = False
            add_new_pcols = False
            overwrite_details = True
            overwrite_parents = False
            file_opened = "extra.csv"
            sheet_opened = "n/a"
            row_len = 3

        ed.new_sheet = [
            ["ID", "DETAIL_1", "PARENT_1"],
            ["Tiger", "Panthera tigris", "Cats"],
            ["Lion", "Leo", "Cats"],
        ]
        custom = 317
        self._set_col_width(ed, 0, custom)
        ed.merge_sheets(popup_=FakeMerge())
        self.assertIn("tiger", ed.nodes)
        self.assertEqual(_parent(ed, "Tiger"), "Cats")
        self.assertEqual(int(ed.sheet.get_column_widths()[0]), custom)
        self.assertEqual(ed.sheet.MT.data[ed.rns["lion"]][1], "Leo")
        self.assertEqual(len(ed.changelog), n_log + 1)
        self.assertEqual(len(ed.vs), n_vs + 1)
        self.assertEqual(ed.changelog[-1].origin, "merge")
        self._assert_one_list(ed, "after merge")
        assert_sound(self, ed.session, "gui-merge")
        ed.undo()
        self.assertNotIn("tiger", ed.nodes)
        self.assertNotEqual(ed.sheet.MT.data[ed.rns["lion"]][1], "Leo")
        self._assert_one_list(ed, "after merge undo")

    def test_display_rows_share_strings_with_changelog(self):
        _app, ed = self._app()
        self._tree(ed, ("Animals", ""), ("Cats", "Animals"))
        flat = display_rows(ed.changelog)
        self.assertGreaterEqual(len(flat), 2)
        self.assertIs(flat[0], ed.changelog[0].rows[0])
        self.assertIs(flat[0][2], ed.changelog[0].rows[0].what)

