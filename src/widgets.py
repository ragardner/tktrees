# SPDX-License-Identifier: GPL-3.0-only
# Copyright © R. A. Gardner

import datetime
import os
import re
import tkinter as tk
from itertools import islice
from tkinter import filedialog, ttk
from typing import Literal

from tksheet import (
    Sheet,
)

from . import toplevels
from .classes import (
    Header,
    TreeBuilder,
)
from .constants import (
    BF,
    EF,
    EFB,
    ERR_ASK_FNT,
    TF,
    checked_icon,
    ctrl_button,
    menu_kwargs,
    rc_button,
    sheet_header_font,
    std_font_size,
    themes,
    unchecked_icon,
)
from .functions import (
    equalize_sublist_lens,
    isreal,
)


class Column_Selection(tk.Frame):
    def __init__(self, parent, C):
        tk.Frame.__init__(self, parent)
        self.C = C
        self.parent_cols = []
        self.rowlen = 0
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.flattened_choices = FlattenedToggleAndOrder(self, command=self.flattened_mode_toggle)
        self.flattened_choices.grid(row=1, column=0, padx=20, pady=(10, 5), sticky="wnse")
        self.flattened_selector = Flattened_Column_Selector(self)
        self.selector = Id_Parent_Column_Selector(self)
        self.selector.grid(row=1, column=0, sticky="wnse")
        self.sheetdisplay = Sheet(
            self,
            theme=self.C.theme,
            expand_sheet_if_paste_too_big=True,
            header_font=sheet_header_font,
            outline_thickness=1,
        )
        self.selector.link_sheet(self.sheetdisplay)
        self.flattened_selector.link_sheet(self.sheetdisplay, self.flattened_choices)
        self.sheetdisplay.enable_bindings("all", "ctrl_select")
        self.sheetdisplay.extra_bindings(
            [
                ("begin_edit_cell", self.begin_edit),
                ("end_edit_cell", self.end_edit),
            ]
        )
        self.sheetdisplay.bind("<<SheetModified>>", self.sheet_modified)
        self.sheetdisplay.headers(newheaders=0)
        self.sheetdisplay.grid(row=0, column=1, rowspan=3, sticky="nswe")

        self.cont_ = Button(
            self,
            text="Build tree ",
            style="TF.Std.TButton",
            command=self.try_to_build_tree,
        )
        self.cont_.grid(row=2, column=0, sticky="w", padx=20, pady=(10, 50))

        self.flattened_selector.grid(row=0, column=0, pady=(0, 9), sticky="nswe")
        self.selector.grid_forget()
        self.selector.grid(row=0, column=0, sticky="nswe")
        self.flattened_selector.grid_forget()

    def flattened_mode_toggle(self):
        if self.flattened_choices.flattened:
            self.flattened_selector.grid(row=0, column=0, pady=(0, 9), sticky="nswe")
            self.selector.grid_forget()
        else:
            self.selector.grid(row=0, column=0, sticky="nswe")
            self.flattened_selector.grid_forget()

    def reset_selectors(self, event=None):
        idcol = self.selector.get_id_col()
        parcols = self.selector.get_par_cols()
        ancparcols = self.flattened_selector.get_par_cols()
        self.selector.set_columns([h for h in self.sheetdisplay.data[0]] if self.sheetdisplay.data else [])
        self.flattened_selector.set_columns([h for h in self.sheetdisplay.data[0]] if self.sheetdisplay.data else [])
        try:
            if idcol is not None and self.sheetdisplay.get_sheet_data():
                self.selector.set_id_col(idcol)
        except Exception:
            pass
        try:
            if parcols and self.sheetdisplay.get_sheet_data():
                self.selector.set_par_cols(parcols)
        except Exception:
            pass
        try:
            if ancparcols and self.sheetdisplay.get_sheet_data():
                self.flattened_selector.set_par_cols(ancparcols)
        except Exception:
            pass

    def end_edit(self, event=None):
        self.bind("<Escape>", self.cancel)

    def begin_edit(self, event=None):
        self.unbind("<Escape>")
        return event.value

    def sheet_modified(self, event):
        if "move" in event.eventname:
            self.selector.set_columns(self.sheetdisplay.data[0])
            self.flattened_selector.set_columns(self.sheetdisplay.data[0])
            self.selector.detect_id_col()
            self.selector.detect_par_cols()
        else:
            self.reset_selectors()
        self.sheetdisplay.focus_set()

    def enable_widgets(self):
        self.selector.enable_me()
        self.flattened_selector.enable_me()
        self.flattened_choices.enable_me()
        self.cont_.config(state="normal")
        self.sheetdisplay.basic_bindings(True)
        self.sheetdisplay.enable_bindings("all", "ctrl_select")
        self.sheetdisplay.extra_bindings(
            [
                ("begin_edit_cell", self.begin_edit),
                ("end_edit_cell", self.end_edit),
            ]
        )
        self.sheetdisplay.bind("<<SheetModified>>", self.sheet_modified)

    def disable_widgets(self):
        self.selector.disable_me()
        self.flattened_selector.disable_me()
        self.flattened_choices.disable_me()
        self.cont_.config(state="disabled")
        self.sheetdisplay.basic_bindings(False)
        self.sheetdisplay.disable_bindings()
        self.sheetdisplay.extra_bindings()
        self.sheetdisplay.unbind("<<SheetModified>>")

    def populate(self, columns):
        self.sheetdisplay.deselect("all")
        self.rowlen = len(columns)
        self.selector.set_columns([h for h in self.C.frames.tree_edit.sheet.data[0]])
        self.flattened_selector.set_columns([h for h in self.C.frames.tree_edit.sheet.data[0]])
        self.C.frames.tree_edit.sheet.MT.data = self.sheetdisplay.set_sheet_data(
            data=self.C.frames.tree_edit.sheet.MT.data,
            redraw=True,
        )
        self.sheetdisplay.headers(newheaders=0)
        if len(self.C.frames.tree_edit.sheet.data) < 3000:
            self.sheetdisplay.set_all_cell_sizes_to_text()
        self.selector.detect_id_col()
        self.selector.detect_par_cols()
        self.flattened_selector.detect_par_cols()
        self.C.show_frame("column_selection")

    def try_to_build_tree(self):
        flattened = self.flattened_choices.flattened
        if flattened:
            order = self.flattened_choices.order
            hier_cols = self.flattened_selector.get_par_cols()
        else:
            hier_cols = list(self.selector.get_par_cols())
            idcol = self.selector.get_id_col()
            if idcol is None or idcol in hier_cols:
                return
        if not hier_cols:
            return
        self.C.status_bar.change_text("Loading...   ")
        self.C.disable_at_start()
        self.C.frames.tree_edit.sheet.MT.data = self.sheetdisplay.get_sheet_data()
        equalize_sublist_lens(self.C.frames.tree_edit.sheet.MT.data, self.rowlen)
        if flattened:
            (
                self.C.frames.tree_edit.sheet.MT.data,
                self.rowlen,
                idcol,
                hier_cols,
            ) = TreeBuilder().convert_flattened_to_normal(
                data=self.C.frames.tree_edit.sheet.MT.data,
                hier_cols=hier_cols,
                rowlen=self.rowlen,
                order=order,
                warnings=self.C.frames.tree_edit.warnings,
            )
        self.C.frames.tree_edit.ic = idcol
        self.C.frames.tree_edit.hiers = hier_cols
        self.C.frames.tree_edit.pc = hier_cols[0]
        self.C.frames.tree_edit.row_len = int(self.rowlen)
        self.C.frames.tree_edit.headers = [
            Header(name, type_="ID" if i == idcol else "Parent" if i in hier_cols else "Text Detail")
            for i, name in enumerate(
                self.C.frames.tree_edit.fix_headers(self.C.frames.tree_edit.sheet.MT.data.pop(0), self.rowlen)
            )
        ]
        (
            self.C.frames.tree_edit.sheet.MT.data,
            self.C.frames.tree_edit.nodes,
            self.C.frames.tree_edit.warnings,
        ) = TreeBuilder().build(
            input_sheet=self.C.frames.tree_edit.sheet.MT.data,
            output_sheet=self.C.frames.tree_edit.new_sheet,
            row_len=self.C.frames.tree_edit.row_len,
            ic=self.C.frames.tree_edit.ic,
            hiers=self.C.frames.tree_edit.hiers,
            nodes=self.C.frames.tree_edit.nodes,
            warnings=self.C.frames.tree_edit.warnings,
            strip=not self.C.frames.tree_edit.allow_spaces_ids_var.get(),
        )
        self.C.frames.tree_edit.populate()
        self.C.frames.tree_edit.show_warnings(str(self.C.open_dict["filepath"]), str(self.C.open_dict["sheet"]))


class Workbook_Sheet_Selection(tk.Frame):
    def __init__(self, parent, C):
        tk.Frame.__init__(self, parent)
        self.C = C
        self.columnconfigure(1, weight=1)
        self.sheets_label = Label(self, text="Workbook sheets:", font=EF, theme=self.C.theme)
        self.sheets_label.grid(row=0, column=0, padx=10, pady=(10, 20), sticky="e")
        self.sheet_select = Ez_Dropdown(self, TF)
        self.sheet_select.bind("<<ComboboxSelected>>", lambda focus: self.focus_set())
        self.sheet_select.grid(row=0, column=1, padx=(0, 20), pady=(10, 20), sticky="nswe")
        self.run_with_sheet = Button(
            self,
            text="Read data",
            style="TF.Std.TButton",
            command=self.cont,
        )
        self.run_with_sheet.grid(row=1, column=0, padx=10, sticky="w")

    def enable_widgets(self):
        self.sheet_select.config(state="readonly")
        self.run_with_sheet.config(state="normal")

    def disable_widgets(self):
        self.sheet_select.config(state="disabled")
        self.run_with_sheet.config(state="disabled")

    def updatesheets(self, sheets):
        self.run_with_sheet.config(state="normal")
        self.run_with_sheet.update_idletasks()
        self.sheet_select.set_my_value(sheets[0])
        self.sheet_select["values"] = sheets

    def cont(self):
        self.C.disable_at_start()
        self.C.open_dict["sheet"] = self.sheet_select.get_my_value()
        self.C.wb_sheet_has_been_selected(self.sheet_select.get_my_value())


class Id_Parent_Column_Selector(tk.Frame):
    def __init__(self, parent, headers=[[]], show_disp_1=True, show_disp_2=True, theme="dark", expand=False):
        tk.Frame.__init__(
            self,
            parent,
            background=themes[theme].top_left_bg,
            highlightbackground=themes[theme].table_fg,
            highlightthickness=0,
        )
        self.grid_propagate(False)
        self.grid_rowconfigure(1, weight=1)
        if show_disp_1:
            self.grid_columnconfigure(0, weight=1, uniform="x")
        if show_disp_2:
            self.grid_columnconfigure(1, weight=1, uniform="x")
        self.C = parent
        self.sheet = None
        self.headers = headers
        self.id_col = None
        self.par_cols = set()
        self.id_col_display = Readonly_Entry_With_Scrollbar(self, font=EFB, theme=theme)
        self.id_col_display.set_my_value("   ID column:   ")
        if show_disp_1:
            self.id_col_display.grid(row=0, column=0, sticky="nswe")
        self.id_col_selection = Sheet(
            self,
            height=280 if not expand else None,
            width=250 if not expand else None,
            theme=theme,
            show_selected_cells_border=False,
            show_horizontal_grid=False,
            align="w",
            header_align="w",
            row_index_align="w",
            table_selected_cells_bg="#0078d7",
            table_selected_box_cells_fg="#0078d7",
            table_selected_cells_fg="white",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            header_selected_cells_bg="#0078d7",
            index_selected_cells_bg="#0078d7",
            header_font=sheet_header_font,
            column_width=170,
            row_index_width=60,
            headers=["SELECT ID"],
            default_row_index="letters",
        )
        self.id_col_selection.data_reference(newdataref=self.headers)
        self.id_col_selection.enable_bindings(("single", "column_width_resize", "double_click_column_resize"))
        self.id_col_selection.extra_bindings(
            [("cell_select", self.id_col_selection_B1), ("deselect", self.id_col_selection_B1)]
        )
        if show_disp_1:
            self.id_col_selection.grid(row=1, column=0, sticky="nswe")
        self.par_col_selection = Sheet(
            self,
            height=280 if not expand else None,
            width=250 if not expand else None,
            theme=theme,
            align="w",
            show_selected_cells_border=False,
            show_horizontal_grid=False,
            header_align="w",
            row_index_align="w",
            table_selected_cells_bg="#79A158",
            table_selected_box_cells_fg="#79A158",
            table_selected_cells_fg="white",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            header_selected_cells_bg="#79A158",
            index_selected_cells_bg="#79A158",
            header_font=sheet_header_font,
            column_width=170,
            row_index_width=60,
            headers=["SELECT PARENTS"],
            default_row_index="letters",
        )
        self.par_col_selection.data_reference(newdataref=self.headers)
        self.par_col_selection.extra_bindings(
            [("cell_select", self.par_col_selection_B1), ("deselect", self.par_col_deselection_B1)]
        )
        self.par_col_selection.enable_bindings(("toggle", "column_width_resize", "double_click_column_resize"))
        if show_disp_2:
            self.par_col_selection.grid(row=1, column=1, sticky="nswe")
        self.par_col_display = Readonly_Entry_With_Scrollbar(self, font=EFB, theme=theme)
        self.par_col_display.set_my_value("   Parent columns:   ")
        if show_disp_2:
            self.par_col_display.grid(row=0, column=1, sticky="nswe")
        self.detect_id_col_button = Button(
            self, text="Detect ID column", style="BF.Std.TButton", command=self.detect_id_col
        )
        if show_disp_1:
            self.detect_id_col_button.grid(row=2, column=0, padx=2, pady=2, sticky="ns")
        self.detect_par_cols_button = Button(
            self, text="Detect parent columns", style="BF.Std.TButton", command=self.detect_par_cols
        )
        if show_disp_2:
            self.detect_par_cols_button.grid(row=2, column=1, padx=2, pady=2, sticky="ns")

    def link_sheet(self, sheet):
        self.sheet = sheet

    def reset_size(self, width=500, height=350):
        self.config(width=width, height=height)
        self.update_idletasks()
        self.par_col_selection.refresh()
        self.id_col_selection.refresh()

    def set_columns(self, columns):
        self.clear_displays()
        self.set_par_cols([])
        self.headers = [[h] for h in columns]
        self.id_col_selection.data_reference(newdataref=self.headers, redraw=True)
        self.par_col_selection.data_reference(newdataref=self.headers, redraw=True)

    def disable_me(self):
        self.id_col_selection.basic_bindings(enable=False)
        self.par_col_selection.basic_bindings(enable=False)
        self.detect_id_col_button.config(state="disabled")
        self.detect_par_cols_button.config(state="disabled")

    def enable_me(self):
        self.id_col_selection.basic_bindings(enable=True)
        self.par_col_selection.basic_bindings(enable=True)
        self.detect_id_col_button.config(state="normal")
        self.detect_par_cols_button.config(state="normal")

    def detect_id_col(self):
        if len(self.id_col_selection.MT.data) <= 1:
            return
        for i, e in enumerate(self.headers):
            if not e:
                continue
            x = e[0].lower().strip()
            if x == "id" or x.startswith("id"):
                self.set_id_col(i)
                return
        if self.sheet is not None and len(self.sheet.data) > 1:
            for c in range(self.sheet.total_columns()):
                if (
                    c not in self.par_cols
                    and any(
                        r[c] != f"{n}" for r, n in zip(islice(self.sheet.data, 1, None), range(len(self.sheet.data)))
                    )
                    and all(r[c].rstrip() for r in islice(self.sheet.data, 1, None) if len(r) > c)
                ):
                    self.set_id_col(c)
                    break

    def detect_par_cols(self):
        if len(self.par_col_selection.MT.data) <= 1:
            return
        parent_cols = []
        for i, e in enumerate(self.headers):
            if not e:
                continue
            x = e[0].lower().strip()
            if x.startswith("parent"):
                parent_cols.append(i)
        if parent_cols:
            self.set_par_cols(parent_cols)
        elif self.sheet is not None and len(self.sheet.data) > 1:
            if not isinstance(self.id_col, int):
                self.detect_id_col()
                if not isinstance(self.id_col, int):
                    return
            ids = {r[self.id_col].lower().rstrip() for r in self.sheet.data if len(r) > self.id_col}
            ids.add("")
            for c in range(self.sheet.total_columns()):
                if (
                    c != self.id_col
                    and any(r[c].rstrip() for r in islice(self.sheet.data, 1, None) if len(r) > c)
                    and all(r[c].lower() in ids for r in islice(self.sheet.data, 1, None) if len(r) > c)
                ):
                    parent_cols.append(c)
            if parent_cols:
                self.set_par_cols(parent_cols)

    def id_col_selection_B1(self, event=None):
        if event:
            self.id_col = tuple(tup[0] for tup in self.id_col_selection.get_selected_cells())
            if self.id_col:
                self.id_col = self.id_col[0]
                self.id_col_display.set_my_value(f"   ID column:   {self.id_col + 1}")
            else:
                self.id_col = None
                self.id_col_display.set_my_value("   ID column:   ")

    def par_col_selection_B1(self, event=None):
        if event:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value(
                "   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)])
            )

    def par_col_deselection_B1(self, event=None):
        if event:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value(
                "   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)])
            )

    def clear_displays(self):
        self.headers = [[]]
        self.id_col = None
        self.id_col_selection.deselect("all")
        self.id_col_display.set_my_value("   ID column:   ")
        self.par_cols = set()
        self.par_col_selection.deselect("all")
        self.par_col_display.set_my_value("   Parent columns:   ")
        self.id_col_selection.set_sheet_data()
        self.par_col_selection.set_sheet_data()

    def set_id_col(self, col):
        self.id_col = col
        self.id_col_selection.deselect("all")
        self.id_col_selection.refresh()
        self.id_col_selection.select_cell(row=col, column=0, redraw=True)
        self.id_col_selection.see(row=col, column=0)
        self.id_col_display.set_my_value(f"   ID column:   {col + 1}")

    def set_par_cols(self, cols):
        self.par_col_selection.deselect("all")
        self.par_col_selection.refresh()
        if cols:
            self.par_cols = set(cols)
            for r in cols:
                self.par_col_selection.toggle_select_cell(r, 0, redraw=False)
            self.par_col_selection.see(row=cols[0], column=0, redraw=False)
            self.par_col_selection.redraw()
            self.par_col_display.set_my_value(
                "   Parent columns:   " + ", ".join([f"{n}" for n in sorted(p + 1 for p in self.par_cols)])
            )

    def get_id_col(self):
        return self.id_col

    def get_par_cols(self):
        return sorted(self.par_cols)

    def change_theme(self, theme="dark"):
        self.id_col_selection.change_theme(theme, redraw=False)
        self.par_col_selection.change_theme(theme, redraw=False)
        self.id_col_selection.set_options(
            table_selected_cells_bg="#0078d7",
            table_selected_box_cells_fg="#0078d7",
            table_selected_cells_fg="white",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            header_selected_cells_bg="#0078d7",
            index_selected_cells_bg="#0078d7",
        )
        self.par_col_selection.set_options(
            table_selected_cells_bg="#79A158",
            table_selected_box_cells_fg="#79A158",
            table_selected_cells_fg="white",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            header_selected_cells_bg="#79A158",
            index_selected_cells_bg="#79A158",
        )
        self.config(background=themes[theme].top_left_bg, highlightbackground=themes[theme].table_fg)
        self.id_col_display.my_entry.config(
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )
        self.par_col_display.my_entry.config(
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )


class FlattenedToggleAndOrder(tk.Frame):
    def __init__(self, parent, command, theme="dark"):
        tk.Frame.__init__(self, parent, background=themes[theme].top_left_bg)
        self.C = parent
        self.extra_func = command
        self._flattened = False
        self.order_dropdown = Ez_Dropdown(self, font=EFB, width_=28)
        self.order_dropdown.bind("<<ComboboxSelected>>", self.dropdown_select)
        self.order_dropdown["values"] = [
            "Sheet is NOT in flattened format",
            "Flattened - Left → Right is Top → Base",
            "Flattened - Left → Right is Base → Top",
        ]
        self.order_dropdown.set_my_value("Sheet is NOT in flattened format")
        self.order_dropdown.grid(row=0, column=0, sticky="nswe")
        self.select_mode(func=False)

    def disable_me(self, func=False):
        self.order_dropdown.config(state="disabled")

    def enable_me(self, func=False):
        self.order_dropdown.config(state="readonly")

    def select_mode(self, event=None, func=True):
        if func:
            self.extra_func()

    def dropdown_select(self, event=None):
        if self.order_dropdown.get_my_value() == "Sheet is NOT in flattened format":
            self._flattened = False
        else:
            self._flattened = True
        self.select_mode(func=True)

    @property
    def flattened(self):
        return self._flattened

    @property
    def order(self):
        return self.order_dropdown.get_my_value()

    def change_theme(self, theme="dark"):
        self.config(bg=themes[theme].top_left_bg)


class Flattened_Column_Selector(tk.Frame):
    def __init__(self, parent, headers=[[]], theme="dark"):
        tk.Frame.__init__(self, parent, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.C = parent
        self.sheet = None
        self.flattened_choices = None
        self.headers = headers
        self.par_cols = set()
        self.par_col_display = Readonly_Entry_With_Scrollbar(self, font=EFB, theme=theme)
        self.par_col_display.set_my_value("   Hierarchy columns:   ")
        self.par_col_display.grid(row=0, column=0, sticky="nswe")
        self.par_col_selection = Sheet(
            self,
            width=500,
            height=300,
            theme=theme,
            align="w",
            show_selected_cells_border=False,
            header_align="w",
            row_index_align="w",
            table_selected_cells_bg="#79A158",
            table_selected_box_cells_fg="#79A158",
            header_selected_cells_bg="#79A158",
            index_selected_cells_bg="#79A158",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            table_selected_cells_fg="white",
            header_font=sheet_header_font,
            column_width=350,
            row_index_width=60,
            headers=["SELECT ALL HIERARCHY COLUMNS"],
            default_row_index="letters",
        )
        self.par_col_selection.data_reference(newdataref=self.headers)
        self.par_col_selection.extra_bindings(
            [("cell_select", self.par_col_selection_B1), ("deselect", self.par_col_deselection_B1)]
        )
        self.par_col_selection.enable_bindings(("toggle", "column_width_resize", "double_click_column_resize"))
        self.par_col_selection.grid(row=1, column=0, sticky="nswe")

    def link_sheet(self, sheet, choices):
        self.sheet = sheet
        self.flattened_choices = choices

    def set_columns(self, columns):
        self.clear_displays()
        self.set_par_cols([])
        self.headers = [[h] for h in columns]
        self.par_col_selection.data_reference(newdataref=self.headers, redraw=True)

    def disable_me(self):
        self.par_col_selection.basic_bindings(enable=False)
        self.par_col_selection.extra_bindings([("cell_select", None), ("deselect", None)])

    def enable_me(self):
        self.par_col_selection.basic_bindings(enable=True)
        self.par_col_selection.extra_bindings(
            [("cell_select", self.par_col_selection_B1), ("deselect", self.par_col_deselection_B1)]
        )

    def detect_par_cols(self):
        parent_cols = []
        for i, e in enumerate(self.headers):
            if not e:
                continue
            x = e[0].lower().strip()
            if x.startswith("parent"):
                parent_cols.append(i)
        if parent_cols:
            self.set_par_cols(parent_cols)

    def par_col_selection_B1(self, event=None):
        if event:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value(
                "   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)])
            )

    def par_col_deselection_B1(self, event=None):
        if event:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value(
                "   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)])
            )

    def clear_displays(self):
        self.headers = [[]]
        self.par_col_selection.data_reference(newdataref=[[]], redraw=True)
        self.par_cols = set()
        self.par_col_selection.deselect("all")
        self.par_col_display.set_my_value("   Hierarchy columns:   ")

    def set_par_cols(self, cols):
        self.par_col_selection.deselect("all")
        self.par_col_selection.refresh()
        if cols:
            self.par_cols = set(cols)
            for r in cols:
                self.par_col_selection.toggle_select_cell(r, 0, redraw=False)
            self.par_col_selection.see(row=cols[0], column=0)
            self.par_col_display.set_my_value(
                "   Hierarchy columns:   " + ",".join([str(n) for n in sorted(p + 1 for p in self.par_cols)])
            )
        self.par_col_selection.refresh()

    def get_par_cols(self):
        return sorted(self.par_cols)

    def change_theme(self, theme="dark"):
        self.config(bg=themes[theme].top_left_bg)
        self.par_col_selection.change_theme(theme)
        self.par_col_selection.set_options(
            table_selected_cells_bg="#79A158",
            table_selected_box_cells_fg="#79A158",
            header_selected_cells_bg="#79A158",
            index_selected_cells_bg="#79A158",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            table_selected_cells_fg="white",
        )
        self.par_col_display.my_entry.config(
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )


class Single_Column_Selector(tk.Frame):
    def __init__(self, parent, headers=[[]], width=250, height=350, theme="dark"):
        tk.Frame.__init__(self, parent, width=width, height=height, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.C = parent
        self.headers = headers
        self.col = None
        self.col_display = Readonly_Entry_With_Scrollbar(self, font=EFB, theme=theme)
        self.col_display.set_my_value("  Column:   ")
        self.col_display.grid(row=0, column=0, sticky="nswe")
        self.col_selection = Sheet(
            self,
            theme=theme,
            show_selected_cells_border=False,
            align="w",
            header_align="w",
            row_index_align="w",
            table_selected_cells_bg="#79A158",
            table_selected_box_cells_fg="#79A158",
            header_selected_cells_bg="#79A158",
            index_selected_cells_bg="#79A158",
            header_selected_cells_fg="white",
            index_selected_cells_fg="white",
            table_selected_cells_fg="white",
            header_font=sheet_header_font,
            column_width=180,
            row_index_width=50,
            headers=["SELECT A COLUMN"],
            default_row_index="letters",
        )
        self.col_selection.data_reference(newdataref=self.headers)
        self.col_selection.extra_bindings([("cell_select", self.col_selection_B1), ("deselect", self.col_deselect)])
        self.col_selection.enable_bindings(("single", "column_width_resize", "double_click_column_resize"))
        self.col_selection.grid(row=1, column=0, sticky="nswe")

    def set_columns(self, columns):
        self.clear_displays()
        self.headers = [[h] for h in columns]
        self.col_selection.data_reference(newdataref=self.headers, redraw=True)

    def disable_me(self):
        self.col_selection.basic_bindings(enable=False)

    def enable_me(self):
        self.col_selection.basic_bindings(enable=True)

    def col_selection_B1(self, event=None):
        if event:
            self.col = event.selected.row
            self.col_display.set_my_value(f"   Column:   {self.col + 1}")

    def col_deselect(self, event=None):
        if event:
            self.col = None
            self.col_display.set_my_value("  Column:   ")

    def par_col_deselection_B1(self, event=None):
        if event:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value(
                f"   Parent columns:   {', '.join(str(n) for n in sorted(p + 1 for p in self.par_cols))}"
            )

    def clear_displays(self):
        self.headers = [[]]
        self.col_selection.data_reference(newdataref=[[]], redraw=True)
        self.col = 0
        self.col_selection.deselect("all")
        self.col_display.set_my_value("   Hierarchy columns:   ")

    def set_col(self, col=None):
        if col is not None:
            self.col = int(col)
            self.col_selection.deselect("all")
            self.col_selection.select_cell(col, 0, redraw=False)
            self.col_selection.see(row=col, column=0)
            self.col_selection.refresh()
            self.col_display.set_my_value(f"   Column:   {col + 1}")

    def get_col(self):
        return int(self.col)


class X_Checkbutton(ttk.Button):
    def __init__(
        self, parent, text="", style="Std.TButton", command=None, state="normal", checked=False, compound="right"
    ):
        Button.__init__(self, parent, text=text, style=style, command=command, state=state)
        self.image_compound = compound
        self.on_icon = tk.PhotoImage(format="gif", data=checked_icon)
        self.off_icon = tk.PhotoImage(format="gif", data=unchecked_icon)
        self.checked = checked
        if checked:
            self.config(image=self.on_icon, compound=compound)
        else:
            self.config(image=self.off_icon, compound=compound)
        self.bind("<1>", self.B1)

    def set_checked(self, state="toggle"):
        if state == "toggle":
            self.checked = not self.checked
            if self.checked:
                self.config(image=self.on_icon, compound=self.image_compound)
            else:
                self.config(image=self.off_icon, compound=self.image_compound)
        elif state:
            self.checked = True
            self.config(image=self.on_icon, compound=self.image_compound)
        elif not state:
            self.checked = False
            self.config(image=self.off_icon, compound=self.image_compound)

    def get_checked(self):
        return bool(self.checked)

    def B1(self, event):
        x = str(self["state"])
        if "normal" in x:
            self.checked = not self.checked
            if self.checked:
                self.config(image=self.on_icon, compound=self.image_compound)
            else:
                self.config(image=self.off_icon, compound=self.image_compound)
        self.update_idletasks()

    def change_text(self, text):
        self.config(text=text)
        self.update_idletasks()


class Auto_Add_Condition_Num_Frame(tk.Frame):
    def __init__(self, parent, col_sel, sheet, theme="dark"):
        tk.Frame.__init__(self, parent, height=200, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.C = parent
        self.col_sel = col_sel
        self.sheet_ref = sheet
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(3, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.min_label = Label(self, text="Min:", font=EFB, theme=theme)
        self.min_label.grid(row=0, column=0, sticky="nswe", padx=(20, 10), pady=(5, 0))
        self.min_entry = Numerical_Entry_With_Scrollbar(self, theme=theme)
        self.min_entry.grid(row=0, column=1, sticky="nswe", padx=(20, 0), pady=(20, 0))
        self.max_label = Label(self, text="Max:", font=EFB, theme=theme)
        self.max_label.grid(row=0, column=2, sticky="nswe", padx=(10, 10), pady=(5, 0))
        self.max_entry = Numerical_Entry_With_Scrollbar(self, theme=theme)
        self.max_entry.grid(row=0, column=3, sticky="nswe", padx=(10, 20), pady=(20, 0))
        self.get_col_min = Button(self, text="Get column minimum", style="EF.Std.TButton", command=self.get_col_min_val)
        self.get_col_min.grid(row=1, column=1, sticky="nswe", padx=(20, 0))
        self.get_col_max = Button(self, text="Get column maximum", style="EF.Std.TButton", command=self.get_col_max_val)
        self.get_col_max.grid(row=1, column=3, sticky="nswe", padx=(10, 20))
        self.asc_desc_dropdown = Ez_Dropdown(self, font=EF)
        self.asc_desc_dropdown["values"] = ("ASCENDING", "DESCENDING")
        self.asc_desc_dropdown.set_my_value("ASCENDING")
        self.asc_desc_dropdown.grid(row=2, column=1, sticky="nswe", padx=(20, 0))
        self.button_frame = Frame(self, theme=theme)
        self.button_frame.grid(row=3, column=0, columnspan=4, sticky="nswe")
        self.button_frame.grid_columnconfigure(0, weight=1, uniform="x")
        self.button_frame.grid_columnconfigure(1, weight=1, uniform="x")
        self.button_frame.grid_rowconfigure(0, weight=1)
        self.confirm_button = Button(
            self.button_frame, text="Save conditions", style="EF.Std.TButton", command=self.confirm
        )
        self.confirm_button.grid(row=0, column=0, sticky="nswe", padx=(20, 10), pady=(20, 20))
        self.cancel_button = Button(self.button_frame, text="Cancel", style="EF.Std.TButton", command=self.cancel)
        self.cancel_button.grid(row=0, column=1, sticky="nswe", padx=(10, 20), pady=(20, 20))
        self.result = False
        self.min_val = ""
        self.max_val = ""
        self.order = ""
        self.min_entry.place_cursor()

    def get_col_min_val(self, event=None):
        c = self.col_sel
        try:
            self.min_entry.set_my_value(str(min(float(r[c]) for r in self.sheet_ref if isreal(r[c]))))
        except Exception:
            pass

    def get_col_max_val(self, event=None):
        c = self.col_sel
        try:
            self.max_entry.set_my_value(str(max(float(r[c]) for r in self.sheet_ref if isreal(r[c]))))
        except Exception:
            pass

    def confirm(self, event=None):
        self.result = True
        self.min_val = self.min_entry.get_my_value()
        self.max_val = self.max_entry.get_my_value()
        self.order = self.asc_desc_dropdown.get_my_value()
        self.destroy()

    def cancel(self, event=None):
        self.destroy()


class Auto_Add_Condition_Date_Frame(tk.Frame):
    def __init__(self, parent, col_sel, sheet, DATE_FORM, theme="dark"):
        tk.Frame.__init__(self, parent, height=225, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.col_sel = col_sel
        self.sheet_ref = sheet
        self.DATE_FORM = DATE_FORM
        if DATE_FORM == "%d/%m/%Y":
            label_form = "DD/MM/YYYY"
        elif DATE_FORM == "%m/%d/%Y":
            label_form = "MM/DD/YYYY"
        elif DATE_FORM == "%Y/%m/%d":
            label_form = "YYYY/MM/DD"
        else:
            self.DATE_FORM = "DD/MM/YYYY"
            label_form = "DD/MM/YYYY"
        self.min_label = Label(self, text="Min  " + label_form, font=EFB, theme=theme)
        self.min_label.grid(row=0, column=0, sticky="nswe", padx=(20, 10), pady=(5, 0))
        self.min_entry = Date_Entry(self, DATE_FORM, theme=theme)
        self.min_entry.grid(row=0, column=1, sticky="nswe", padx=(20, 0), pady=(20, 0))
        self.max_label = Label(self, text="Max  " + label_form, font=EFB, theme=theme)
        self.max_label.grid(row=0, column=2, sticky="nswe", padx=(10, 10), pady=(5, 0))
        self.max_entry = Date_Entry(self, DATE_FORM, theme=theme)
        self.max_entry.grid(row=0, column=3, sticky="nswe", padx=(10, 20), pady=(20, 0))
        self.get_col_min = Button(self, text="Get column minimum", style="EF.Std.TButton", command=self.get_col_min_val)
        self.get_col_min.grid(row=1, column=1, sticky="nswe", padx=(20, 0))
        self.get_col_max = Button(self, text="Get column maximum", style="EF.Std.TButton", command=self.get_col_max_val)
        self.get_col_max.grid(row=1, column=3, sticky="nswe", padx=(10, 20))
        self.asc_desc_dropdown = Ez_Dropdown(self, font=EF)
        self.asc_desc_dropdown["values"] = ("ASCENDING", "DESCENDING")
        self.asc_desc_dropdown.set_my_value("ASCENDING")
        self.asc_desc_dropdown.grid(row=2, column=1, sticky="nswe", padx=(20, 0))
        self.button_frame = Frame(self, theme=theme)
        self.button_frame.grid(row=3, column=0, columnspan=4, sticky="nswe")
        self.button_frame.grid_columnconfigure(0, weight=1, uniform="x")
        self.button_frame.grid_columnconfigure(1, weight=1, uniform="x")
        self.button_frame.grid_rowconfigure(0, weight=1)
        self.confirm_button = Button(
            self.button_frame, text="Save conditions", style="EF.Std.TButton", command=self.confirm
        )
        self.confirm_button.grid(row=0, column=0, sticky="nswe", padx=(20, 10), pady=(20, 20))
        self.cancel_button = Button(self.button_frame, text="Cancel", style="EF.Std.TButton", command=self.cancel)
        self.cancel_button.grid(row=0, column=1, sticky="nswe", padx=(10, 20), pady=(20, 20))
        self.result = False
        self.min_val = ""
        self.max_val = ""
        self.order = ""
        self.min_entry.place_cursor()

    def detect_date_form(self, date):
        forms = []
        for form in ("%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d"):
            try:
                datetime.datetime.strptime(date, form).date()
                forms.append(form)
            except Exception:
                pass
        if len(forms) == 1:
            return forms[0]
        return False

    def get_col_min_val(self, event=None):
        c = self.col_sel
        try:
            self.min_entry.set_my_value(
                datetime.datetime.strftime(
                    min(
                        datetime.datetime.strptime(r[c], self.DATE_FORM)
                        for r in self.sheet_ref
                        if self.detect_date_form(r[c]) == self.DATE_FORM
                    ),
                    self.DATE_FORM,
                )
            )
        except Exception:
            pass

    def get_col_max_val(self, event=None):
        c = self.col_sel
        try:
            self.max_entry.set_my_value(
                datetime.datetime.strftime(
                    max(
                        datetime.datetime.strptime(r[c], self.DATE_FORM)
                        for r in self.sheet_ref
                        if self.detect_date_form(r[c]) == self.DATE_FORM
                    ),
                    self.DATE_FORM,
                )
            )
        except Exception:
            pass

    def confirm(self, event=None):
        self.result = True
        self.min_val = self.min_entry.get_my_value()
        self.max_val = self.max_entry.get_my_value()
        self.order = self.asc_desc_dropdown.get_my_value()
        self.destroy()

    def cancel(self, event=None):
        self.destroy()


class Edit_Condition_Frame(tk.Frame):
    def __init__(
        self, parent, condition, colors, color=None, coltype="Text Detail", confirm_text="Save condition", theme="dark"
    ):
        tk.Frame.__init__(self, parent, height=160, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        if coltype in ("ID", "Parent", "Text Detail"):
            self.if_cell_label = Label(self, text="If cell text is exactly:", font=EFB, theme=theme)
        else:
            self.if_cell_label = Label(self, text="If cell value is:", font=EFB, theme=theme)
        self.if_cell_label.grid(row=0, column=0, sticky="nswe", padx=(20, 10), pady=(20, 40))
        self.condition_display = Condition_Entry_With_Scrollbar(self, coltype=coltype, theme=theme)
        self.condition_display.set_my_value(condition)
        self.condition_display.grid(row=0, column=1, sticky="nswe", pady=(20, 20), padx=(0, 0))
        self.color_dropdown = Ez_Dropdown(self, EF)
        self.color_dropdown["values"] = colors
        if color is None:
            self.color_dropdown.set_my_value(colors[0])
        else:
            self.color_dropdown.set_my_value(color)
        self.color_dropdown.grid(row=0, column=2, sticky="nswe", pady=(20, 20), padx=(0, 20))
        self.button_frame = Frame(self, theme=theme)
        self.button_frame.grid(row=1, column=0, columnspan=3, sticky="e")
        self.button_frame.grid_rowconfigure(0, weight=1)
        self.confirm_button = Button(self.button_frame, text=confirm_text, style="EF.Std.TButton", command=self.confirm)
        self.confirm_button.grid(row=0, column=0, sticky="e", padx=(20, 10), pady=(0, 20))
        self.cancel_button = Button(self.button_frame, text="Cancel", style="EF.Std.TButton", command=self.cancel)
        self.cancel_button.grid(row=0, column=1, sticky="e", padx=(10, 20), pady=(0, 20))
        self.result = False
        self.new_condition = ""
        self.color_dropdown.bind("<<ComboboxSelected>>", self.disable_cancel)
        self.condition_display.place_cursor()

    def disable_cancel(self, event=None):
        try:
            self.cancel_button.config(state="disabled")
            self.after(300, self.enable_cancel)
        except Exception:
            pass

    def enable_cancel(self, event=None):
        self.cancel_button.config(state="normal")

    def confirm(self, event=None):
        self.result = True
        self.new_condition = self.condition_display.get_my_value()
        self.color = self.color_dropdown.get_my_value()
        self.destroy()

    def cancel(self, event=None):
        self.destroy()


class Condition_Entry_With_Scrollbar(tk.Frame):
    def __init__(self, parent, coltype="Text Detail", theme="dark"):
        tk.Frame.__init__(self, parent)
        self.config(bg=themes[theme].top_left_bg)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.my_entry = Condition_Normal_Entry(self, font=EF, coltype=coltype, theme=theme)
        self.my_entry.grid(row=0, column=0, sticky="nswe")
        self.my_scrollbar = Scrollbar(self, self.my_entry.xview, "horizontal", self.my_entry)
        self.my_scrollbar.grid(row=1, column=0, sticky="ew")

    def change_my_state(self, state, event=None):
        self.my_entry.config(state=state)

    def place_cursor(self, event=None):
        self.my_entry.focus_set()

    def get_my_value(self, event=None):
        return self.my_entry.get()

    def set_my_value(self, val, event=None):
        self.my_entry.set_my_value(val)


class Condition_Normal_Entry(tk.Entry):
    def __init__(self, parent, font, coltype="Text Detail", width_=None, theme="dark"):
        tk.Entry.__init__(
            self,
            parent,
            font=font,
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )
        if width_:
            self.config(width=width_)
        self.coltype = coltype
        if self.coltype not in ("ID", "Parent", "Text Detail"):
            self.validate_text = True
        else:
            self.validate_text = False
        if coltype != "Date Detail":
            self.allowed_chars = {
                "a",
                "n",
                "d",
                "o",
                "r",
                "A",
                "N",
                "D",
                "O",
                "R",
                "0",
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "!",
                ">",
                "<",
                "=",
                " ",
                "-",
                ".",
                "C",
                "c",
            }
        else:
            self.allowed_chars = {
                "a",
                "n",
                "d",
                "o",
                "r",
                "A",
                "N",
                "D",
                "O",
                "R",
                "0",
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "!",
                ">",
                "<",
                "=",
                " ",
                "/",
                "C",
                "c",
                "-",
            }
        self.sv = tk.StringVar()
        self.config(textvariable=self.sv)
        self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        self.rc_popup_menu = tk.Menu(self, tearoff=0, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all", accelerator="Ctrl+A", command=self.select_all, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut", accelerator="Ctrl+X", command=self.cut, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy", accelerator="Ctrl+C", command=self.copy, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste", accelerator="Ctrl+V", command=self.paste, **menu_kwargs)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)
        self.bind(f"<{ctrl_button}-a>", self.select_all)
        self.bind(f"<{ctrl_button}-A>", self.select_all)
        self.set_my_value(" ")

    def validate_(self, sv):
        if self.validate_text:
            self.sv.set("".join([c.lower() for c in self.sv.get().replace("  ", " ") if c in self.allowed_chars]))

    def rc(self, event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event: object = None) -> Literal["break"]:
        self.select_range(0, "end")
        return "break"

    def cut(self, event=None):
        self.event_generate(f"<{ctrl_button}-x>")
        return "break"

    def copy(self, event=None):
        self.event_generate(f"<{ctrl_button}-c>")
        return "break"

    def paste(self, event=None):
        self.event_generate(f"<{ctrl_button}-v>")
        return "break"

    def set_my_value(self, newvalue):
        self.delete(0, "end")
        self.insert(0, str(newvalue))

    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)

    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(rc_button)


class Askconfirm_Frame(tk.Frame):
    def __init__(
        self,
        parent,
        action,
        confirm_text="Confirm",
        cancel_text="Cancel",
        bgcolor="green",
        fgcolor="white",
        theme="dark",
    ):
        tk.Frame.__init__(self, parent, height=150, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.action_display = Readonly_Entry_With_Scrollbar(self, theme=theme)
        self.action_display.my_entry.config(font=ERR_ASK_FNT)
        self.action_display.set_my_value(action)
        self.action_display.grid(row=0, column=0, sticky="nswe", pady=(20, 5), padx=(20, 20))
        self.button_frame = Frame(self, theme=theme)
        self.button_frame.grid(row=1, column=0, sticky="e")
        self.button_frame.grid_rowconfigure(0, weight=1)
        self.confirm_button = Button(self.button_frame, text=confirm_text, style="BF.Std.TButton", command=self.confirm)
        self.confirm_button.grid(row=0, column=0, sticky="e", padx=(20, 10), pady=20)
        self.cancel_button = Button(self.button_frame, text=cancel_text, style="BF.Std.TButton", command=self.cancel)
        self.cancel_button.grid(row=0, column=1, sticky="e", padx=(10, 20), pady=20)
        self.boolean = False
        self.action_display.place_cursor()

    def confirm(self, event=None):
        self.boolean = True
        self.destroy()

    def cancel(self, event=None):
        self.destroy()


class Error_Frame(tk.Frame):
    def __init__(self, parent, msg, theme="dark"):
        tk.Frame.__init__(self, parent, height=150, bg=themes[theme].top_left_bg)
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.errorlabel = Label(self, text="Error\nmessage:", font=EF, theme=theme)
        self.errorlabel.config(background="red", foreground="white")
        self.errorlabel.grid(row=0, column=0, sticky="nswe", padx=20)
        self.error_display = Readonly_Entry_With_Scrollbar(self, theme=theme)
        self.error_display.my_entry.config(font=ERR_ASK_FNT)
        self.error_display.set_my_value(msg)
        self.error_display.grid(row=0, column=1, sticky="nswe", pady=(20, 5), padx=(0, 20))
        self.confirm_button = Button(self, text="Okay", style="BF.Std.TButton", command=self.confirm)
        self.confirm_button.grid(row=1, column=1, sticky="e", padx=20, pady=(10, 20))

    def confirm(self, event=None):
        self.destroy()


class Working_Text(tk.Text):
    def __init__(
        self,
        parent,
        wrap,
        font=("Calibri", std_font_size, "normal"),
        theme="dark",
        use_entry_bg=True,
        override_bg=None,
        bold=False,
        highlightthickness=1,
    ):
        tk.Text.__init__(
            self,
            parent,
            wrap=wrap,
            font=font if not bold else ("Calibri", 13, "bold"),
            spacing1=5,
            spacing2=5,
            highlightthickness=1 if use_entry_bg else 0,
            relief="flat",
        )
        self.config(
            bg=themes[theme].table_bg if use_entry_bg else themes[theme].top_left_bg,
            fg=themes[theme].table_fg if use_entry_bg else themes[theme].table_fg,
            insertbackground=themes[theme].table_fg if use_entry_bg else themes[theme].table_fg,
        )
        if override_bg is not None:
            self.config(bg=override_bg)
        self.rc_popup_menu = tk.Menu(self, tearoff=0, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all", accelerator="Ctrl+A", command=self.select_all, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut", accelerator="Ctrl+X", command=self.cut, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy", accelerator="Ctrl+C", command=self.copy, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste", accelerator="Ctrl+V", command=self.paste, **menu_kwargs)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)
        self.bind(f"<{ctrl_button}-a>", self.select_all)
        self.bind(f"<{ctrl_button}-A>", self.select_all)

    #     self.font = font
    #     em = 20
    #     default_size = std_font_size
    #     bold_font = (self.font[0], self.font[1], "bold")
    #     italic_font = (self.font[0], self.font[1], "italic")
    #     # Small subset of markdown. Just enough to make text look nice.
    #     self.tag_configure("**", font=bold_font)
    #     self.tag_configure("*", font=italic_font)
    #     self.tag_configure("_", font=italic_font)
    #     self.tag_chars = "*_"
    #     self.tag_char_re = re.compile(r"[*_]")

    #     max_heading = 3
    #     for i in range(1, max_heading + 1):
    #         header_font = (self.font[0], int(default_size * i + 3), "bold")
    #         self.tag_configure(
    #             "#" * (max_heading - i), font=header_font, spacing3=default_size
    #         )
    #     self.tag_configure("bullet", lmargin1=em, lmargin2=20)
    #     self.tag_configure("numbered", lmargin1=em, lmargin2=20)
    #     self.numbered_index = 1

    # def insert_markdown(self, text: str):
    #     for line in text.split("\n"):
    #         if line == "":
    #             # Blank lines reset numbering
    #             self.numbered_index = 1
    #             self.insert("end", line)

    #         elif line.startswith("#"):
    #             tag = re.match(r"(#+) (.*)", line)
    #             line = tag.group(2)
    #             self.insert("end", line, tag.group(1))

    #         elif line.startswith("* "):
    #             line = line[2:]
    #             self.insert("end", f"\u2022 {text}", "bullet")

    #         elif line.startswith("1. "):
    #             line = line[2:]
    #             self.insert("end", f"{self.numbered_index}. {text}", "numbered")
    #             self.numbered_index += 1

    #         elif not self.tag_char_re.search(line):
    #             self.insert("end", line)

    #         else:
    #             tag = None
    #             accumulated = []
    #             skip_next = False
    #             for i, c in enumerate(line):
    #                 if skip_next:
    #                     skip_next = False
    #                     continue
    #                 if c in self.tag_chars and (not tag or c == tag[0]):
    #                     if tag:
    #                         self.insert("end", "".join(accumulated), tag)
    #                         accumulated = []
    #                         tag = None
    #                     else:
    #                         self.insert("end", "".join(accumulated))
    #                         accumulated = []
    #                         tag = c
    #                         next_i = i + 1
    #                         if len(line) > next_i and line[next_i] == tag:
    #                             tag = line[i : next_i + 1]
    #                             skip_next = True

    #                 else:
    #                     accumulated.append(c)
    #             self.insert("end", "".join(accumulated), tag)

    #         self.insert("end", "\n")

    def rc(self, event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event: object = None) -> Literal["break"]:
        self.tag_add(tk.SEL, "1.0", tk.END)
        self.mark_set(tk.INSERT, tk.END)
        return "break"

    def cut(self, event=None):
        self.event_generate(f"<{ctrl_button}-x>")
        return "break"

    def copy(self, event=None):
        self.event_generate(f"<{ctrl_button}-c>")
        return "break"

    def paste(self, event=None):
        self.event_generate(f"<{ctrl_button}-v>")
        return "break"


class Display_Text(tk.Frame):
    def __init__(self, parent, text="", theme="dark", bold=False):
        tk.Frame.__init__(self, parent, bg=themes[theme].top_left_bg)
        self.C = parent
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.textbox = Working_Text(self, wrap="word", theme=theme, use_entry_bg=False, bold=bold)
        self.textbox.config(highlightbackground=themes[theme].top_left_bg, highlightcolor=themes[theme].top_left_bg)
        self.textbox.grid_propagate(False)
        self.grid_propagate(False)
        self.yscrollb = Scrollbar(self, self.textbox.yview, "vertical", self.textbox)
        self.textbox.delete(1.0, "end")
        self.textbox.insert(1.0, text)
        self.textbox.config(state="disabled", relief="flat")
        self.textbox.grid(row=0, column=0, sticky="nswe")
        self.yscrollb.grid(row=0, column=1, sticky="ns")

    def place_cursor(self, index=None):
        if not index:
            self.textbox.focus_set()

    def get_my_value(self):
        return self.textbox.get("1.0", "end")

    def set_my_value(self, value):
        self.textbox.config(state="normal")
        self.textbox.delete(1.0, "end")
        self.textbox.insert(1.0, value)
        self.textbox.config(state="readonly")

    def change_my_state(self, new_state):
        self.current_state = new_state
        self.textbox.config(state=self.current_state)

    def change_my_width(self, new_width):
        self.textbox.config(width=new_width)

    def change_my_height(self, new_height):
        self.textbox.config(height=new_height)


class Wrapped_Text_With_Find_And_Yscroll(tk.Frame):
    def __init__(self, parent, text, current_state, height=None, theme="dark"):
        tk.Frame.__init__(self, parent, bg=themes[theme].top_left_bg)
        self.C = parent
        self.theme = theme
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.current_state = current_state
        self.word = ""
        self.find_results = []
        self.results_number = 0
        self.addon = ""
        self.find_frame = Frame(self, theme=theme)
        self.find_frame.grid(row=0, column=0, columnspan=2, sticky="nswe")
        self.save_text_button = Button(self.find_frame, text="Save as", style="Std.TButton", command=self.save_text)
        self.save_text_button.pack(side="left", fill="x")
        self.find_button = Button(self.find_frame, text="Find: ", style="Std.TButton", command=self.find)
        self.find_button.pack(side="left", fill="x")
        self.find_window = Normal_Entry(self.find_frame, font=BF, theme=theme)
        self.find_window.bind("<Return>", self.find)
        self.find_window.pack(side="left", fill="x", expand=True)
        self.find_reset_button = Button(self.find_frame, text="X", style="Std.TButton", command=self.find_reset)
        self.find_reset_button.pack(side="left", fill="x")
        self.find_results_label = Label(self.find_frame, "0/0", BF, theme=theme)
        self.find_results_label.pack(side="left", fill="x")
        self.find_up_button = Button(self.find_frame, text="▲", style="Std.TButton", command=self.find_up)
        self.find_up_button.pack(side="left", fill="x")
        self.find_down_button = Button(self.find_frame, text="▼", style="Std.TButton", command=self.find_down)
        self.find_down_button.pack(side="left", fill="x")
        self.textbox = Working_Text(self, wrap="word", theme=theme)
        if height:
            self.textbox.config(height=height)
        self.yscrollb = Scrollbar(self, self.textbox.yview, "vertical", self.textbox)
        self.textbox.delete(1.0, "end")
        self.textbox.insert(1.0, text)
        self.textbox.config(state=self.current_state)
        self.textbox.grid(row=1, column=0, sticky="nswe")
        self.yscrollb.grid(row=1, column=1, sticky="ns")

    def place_cursor(self, index=None):
        if not index:
            self.textbox.focus_set()

    def get_my_value(self):
        return self.textbox.get("1.0", "end")

    def set_my_value(self, value):
        self.textbox.config(state="normal")
        self.textbox.delete(1.0, "end")
        self.textbox.insert(1.0, value)
        self.textbox.config(state=self.current_state)

    def change_my_state(self, new_state):
        self.current_state = new_state
        self.textbox.config(state=self.current_state)

    def change_my_width(self, new_width):
        self.textbox.config(width=new_width)

    def change_my_height(self, new_height):
        self.textbox.config(height=new_height)

    def save_text(self):
        newfile = filedialog.asksaveasfilename(
            parent=self,
            title="Save text on popup window",
            filetypes=[("Text File", ".txt"), ("CSV File", ".csv")],
            defaultextension=".txt",
            confirmoverwrite=True,
        )
        if not newfile:
            return
        newfile = os.path.normpath(newfile)
        if not newfile.lower().endswith((".csv", ".txt")):
            toplevels.Error(self.C, "Can only save .csv/.txt files", theme=self.theme)
            return
        self.save_text_button.change_text("Saving...")
        try:
            with open(newfile, "w") as fh:
                fh.write(self.textbox.get("1.0", "end"))  # remove last newline? [:-2]
        except Exception:
            toplevels.Error(self.C, "Error saving file", theme=self.theme)
            self.save_text_button.change_text("Save as")
            return
        self.save_text_button.change_text("Save as")

    def find(self, event=None):
        self.find_reset(True)
        self.word = self.find_window.get()
        if not self.word:
            return
        self.addon = f"+{len(self.word)}c"
        start = "1.0"
        while start:
            start = self.textbox.search(self.word, index=start, nocase=1, stopindex="end")
            if start:
                end = start + self.addon
                self.find_results.append(start)
                self.textbox.tag_add("i", start, end)
                start = end
        if self.find_results:
            self.textbox.tag_config("i", background="Yellow")
            self.find_results_label.config(text=f"1/{len(self.find_results)}")
            self.textbox.tag_add(
                "c",
                self.find_results[self.results_number],
                self.find_results[self.results_number] + self.addon,
            )
            self.textbox.tag_config("c", background="Orange")
            self.textbox.see(self.find_results[self.results_number])

    def find_up(self, event=None):
        if not self.find_results or len(self.find_results) == 1:
            return
        self.textbox.tag_remove(
            "c",
            self.find_results[self.results_number],
            self.find_results[self.results_number] + self.addon,
        )
        if self.results_number == 0:
            self.results_number = len(self.find_results) - 1
        else:
            self.results_number -= 1
        self.find_results_label.config(text=f"{self.results_number + 1}/{len(self.find_results)}")
        self.textbox.tag_add(
            "c",
            self.find_results[self.results_number],
            self.find_results[self.results_number] + self.addon,
        )
        self.textbox.tag_config("c", background="Orange")
        self.textbox.see(self.find_results[self.results_number])

    def find_down(self, event=None):
        if not self.find_results or len(self.find_results) == 1:
            return
        self.textbox.tag_remove(
            "c",
            self.find_results[self.results_number],
            self.find_results[self.results_number] + self.addon,
        )
        if self.results_number == len(self.find_results) - 1:
            self.results_number = 0
        else:
            self.results_number += 1
        self.find_results_label.config(text=f"{self.results_number + 1}/{len(self.find_results)}")
        self.textbox.tag_add(
            "c",
            self.find_results[self.results_number],
            self.find_results[self.results_number] + self.addon,
        )
        self.textbox.tag_config("c", background="Orange")
        self.textbox.see(self.find_results[self.results_number])

    def find_reset(self, newfind=False):
        self.find_results = []
        self.results_number = 0
        self.addon = ""
        if not newfind:
            self.find_window.delete(0, "end")
        for tag in self.textbox.tag_names():
            self.textbox.tag_delete(tag)
        self.find_results_label.config(text="0/0")


class Scrollbar(ttk.Scrollbar):
    def __init__(self, parent, command, orient, widget):
        ttk.Scrollbar.__init__(self, parent, command=command, orient=orient)
        self.orient = orient
        self.widget = widget
        if self.orient == "vertical":
            self.widget.configure(yscrollcommand=self.set)
        elif self.orient == "horizontal":
            self.widget.configure(xscrollcommand=self.set)


class Readonly_Entry(tk.Entry):
    def __init__(
        self,
        parent,
        font,
        width_=None,
        theme="dark",
        use_status_fg=False,
        outline=1,
    ):
        tk.Entry.__init__(
            self,
            parent,
            font=font,
            state="readonly",
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_selected_box_cells_fg if use_status_fg else themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_selected_box_cells_fg if use_status_fg else themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
            highlightthickness=outline,
            relief="flat",
        )
        if width_:
            self.config(width=width_)
        self.use_status_fg = use_status_fg
        self.rc_popup_menu = tk.Menu(self, tearoff=0, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all", accelerator="Ctrl+A", command=self.select_all, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut", accelerator="Ctrl+X", command=self.cut, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy", accelerator="Ctrl+C", command=self.copy, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste", accelerator="Ctrl+V", command=self.paste, **menu_kwargs)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)

    def rc(self, event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event: object = None) -> Literal["break"]:
        self.select_range(0, "end")
        return "break"

    def cut(self, event=None):
        self.event_generate(f"<{ctrl_button}-x>")
        return "break"

    def copy(self, event=None):
        self.event_generate(f"<{ctrl_button}-c>")
        return "break"

    def paste(self, event=None):
        self.event_generate(f"<{ctrl_button}-v>")
        return "break"

    def set_my_value(self, newvalue):
        self.config(state="normal")
        self.delete(0, "end")
        self.insert(0, str(newvalue))
        self.config(state="readonly")

    def change_theme(self, theme="dark"):
        self.config(
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_selected_box_cells_fg if self.use_status_fg else themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_selected_box_cells_fg
            if self.use_status_fg
            else themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )


class Normal_Entry(tk.Entry):
    def __init__(self, parent, font, width_=None, relief="sunken", border=1, textvariable=None, theme="dark"):
        tk.Entry.__init__(
            self,
            parent,
            font=font,
            relief=relief,
            border=border,
            textvariable=textvariable,
            background=themes[theme].table_bg,
            foreground=themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )
        if width_:
            self.config(width=width_)
        if textvariable:
            self.config(textvariable=textvariable)
        self.rc_popup_menu = tk.Menu(self, tearoff=0, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all", accelerator="Ctrl+A", command=self.select_all, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut", accelerator="Ctrl+X", command=self.cut, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy", accelerator="Ctrl+C", command=self.copy, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste", accelerator="Ctrl+V", command=self.paste, **menu_kwargs)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)
        self.bind(f"<{ctrl_button}-a>", self.select_all)
        self.bind(f"<{ctrl_button}-A>", self.select_all)

    def rc(self, event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event: object = None) -> Literal["break"]:
        self.select_range(0, "end")
        return "break"

    def cut(self, event=None):
        self.event_generate(f"<{ctrl_button}-x>")
        return "break"

    def copy(self, event=None):
        self.event_generate(f"<{ctrl_button}-c>")
        return "break"

    def paste(self, event=None):
        self.event_generate(f"<{ctrl_button}-v>")
        return "break"

    def set_my_value(self, newvalue):
        self.delete(0, "end")
        self.insert(0, str(newvalue))

    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)

    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(rc_button)


class Readonly_Entry_With_Scrollbar(tk.Frame):
    def __init__(self, parent, font=EF, theme="dark", use_status_fg=False):
        tk.Frame.__init__(self, parent, bg=themes[theme].top_left_bg)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.my_entry = Readonly_Entry(self, font=font, theme=theme, use_status_fg=use_status_fg)
        self.my_entry.grid(row=0, column=0, sticky="nswe")
        self.my_scrollbar = Scrollbar(self, self.my_entry.xview, "horizontal", self.my_entry)
        self.my_scrollbar.grid(row=1, column=0, sticky="ew")

    def change_my_state(self, state, event=None):
        self.my_entry.config(state=state)

    def place_cursor(self, event=None):
        self.my_entry.focus_set()

    def get_my_value(self, event=None):
        return self.my_entry.get()

    def set_my_value(self, val, event=None):
        self.my_entry.set_my_value(val)

    def change_text(self, text=""):
        self.my_entry.set_my_value(text)

    def change_theme(self, theme="dark"):
        self.config(bg=themes[theme].top_left_bg)
        self.my_entry.change_theme(theme)


class Entry_With_Scrollbar(tk.Frame):
    def __init__(self, parent, theme="dark"):
        tk.Frame.__init__(self, parent, bg=themes[theme].top_left_bg)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.my_entry = Normal_Entry(self, font=EF, theme=theme)
        self.my_entry.grid(row=0, column=0, sticky="nswe")
        self.my_scrollbar = Scrollbar(self, self.my_entry.xview, "horizontal", self.my_entry)
        self.my_scrollbar.grid(row=1, column=0, sticky="ew")

    def change_my_state(self, state, event=None):
        self.my_entry.config(state=state)

    def place_cursor(self, event=None):
        self.my_entry.focus_set()

    def get_my_value(self, event=None):
        return self.my_entry.get()

    def set_my_value(self, val, event=None):
        self.my_entry.set_my_value(val)


class Ez_Dropdown(ttk.Combobox):
    def __init__(self, parent, font, width_=None):
        self.displayed = tk.StringVar()
        ttk.Combobox.__init__(self, parent, font=font, state="readonly", textvariable=self.displayed)
        if width_:
            self.config(width=width_)

    def get_my_value(self, event=None):
        return self.displayed.get()

    def set_my_value(self, value, event=None):
        self.displayed.set(value)


class Numerical_Entry_With_Scrollbar(tk.Frame):
    def __init__(self, parent, theme="dark"):
        tk.Frame.__init__(self, parent, background=themes[theme].top_left_bg)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.my_entry = Numerical_Normal_Entry(self, font=EF, theme=theme)
        self.my_entry.grid(row=0, column=0, sticky="nswe")
        self.my_scrollbar = Scrollbar(self, self.my_entry.xview, "horizontal", self.my_entry)
        self.my_scrollbar.grid(row=1, column=0, sticky="ew")

    def change_my_state(self, state, event=None):
        self.my_entry.config(state=state)

    def place_cursor(self, event=None):
        self.my_entry.focus_set()

    def get_my_value(self, event=None):
        return self.my_entry.get()

    def set_my_value(self, val, event=None):
        self.my_entry.set_my_value(val)


class Numerical_Normal_Entry(tk.Entry):
    def __init__(self, parent, font, width_=None, theme="dark"):
        tk.Entry.__init__(
            self,
            parent,
            font=font,
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_fg,
            disabledbackground=themes[theme].top_left_bg,
            disabledforeground=themes[theme].table_fg,
            insertbackground=themes[theme].table_fg,
            readonlybackground=themes[theme].top_left_bg,
        )
        if width_:
            self.config(width=width_)
        self.allowed_chars = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."}
        self.sv = tk.StringVar()
        self.config(textvariable=self.sv)
        self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        self.rc_popup_menu = tk.Menu(self, tearoff=0, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all", accelerator="Ctrl+A", command=self.select_all, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut", accelerator="Ctrl+X", command=self.cut, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy", accelerator="Ctrl+C", command=self.copy, **menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste", accelerator="Ctrl+V", command=self.paste, **menu_kwargs)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)
        self.bind(f"<{ctrl_button}-a>", self.select_all)
        self.bind(f"<{ctrl_button}-A>", self.select_all)

    def validate_(self, sv):
        x = self.sv.get()
        dotidx = [i for i, c in enumerate(x) if c == "."]
        if dotidx:
            if len(dotidx) > 1:
                x = x[: dotidx[1]]
        if x.startswith("."):
            x = x[1:]
        if x.startswith("-"):
            self.sv.set("-" + "".join([c for c in x if c in self.allowed_chars]))
        else:
            self.sv.set("".join([c for c in x if c in self.allowed_chars]))

    def rc(self, event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def select_all(self, event: object = None) -> Literal["break"]:
        self.select_range(0, "end")
        return "break"

    def cut(self, event=None):
        self.event_generate(f"<{ctrl_button}-x>")
        return "break"

    def copy(self, event=None):
        self.event_generate(f"<{ctrl_button}-c>")
        return "break"

    def paste(self, event=None):
        self.event_generate(f"<{ctrl_button}-v>")
        return "break"

    def set_my_value(self, newvalue):
        self.delete(0, "end")
        self.insert(0, str(newvalue))

    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_button, self.rc)

    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(rc_button)


class Date_Entry(tk.Frame):
    def __init__(self, parent, DATE_FORM, theme="dark"):
        tk.Frame.__init__(
            self,
            parent,
            relief="flat",
            bg=themes[theme].top_left_bg,
            highlightbackground=themes[theme].table_fg,
            highlightthickness=2,
            border=2,
        )
        self.C = parent

        self.allowed_chars = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        self.sv_1 = tk.StringVar()
        self.sv_2 = tk.StringVar()
        self.sv_3 = tk.StringVar()
        self.entry_1 = Normal_Entry(
            self, font=("Calibri", 30, "bold"), width_=4, relief="flat", border=0, textvariable=self.sv_1, theme=theme
        )
        self.sep = "/" if "/" in DATE_FORM else "-"
        self.label_1 = tk.Label(
            self, font=("Calibri", 30, "bold"), text=self.sep, background=themes[theme].top_left_bg, relief="flat"
        )

        self.entry_2 = Normal_Entry(
            self, font=("Calibri", 30, "bold"), width_=2, relief="flat", border=0, textvariable=self.sv_2, theme=theme
        )

        self.label_2 = tk.Label(
            self, font=("Calibri", 30, "bold"), text=self.sep, background=themes[theme].top_left_bg, relief="flat"
        )

        self.entry_3 = Normal_Entry(
            self, font=("Calibri", 30, "bold"), width_=2, relief="flat", border=0, textvariable=self.sv_3, theme=theme
        )

        self.sv_1.trace("w", lambda name, index, mode, sv=self.sv_1: self.validate_1(self.sv_1))
        self.sv_2.trace("w", lambda name, index, mode, sv=self.sv_2: self.validate_2(self.sv_2))
        self.sv_3.trace("w", lambda name, index, mode, sv=self.sv_3: self.validate_3(self.sv_3))
        self.entry_1.bind("<BackSpace>", self.e1_back)
        self.entry_2.bind("<BackSpace>", self.e2_back)
        self.entry_3.bind("<BackSpace>", self.e3_back)

        if DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            self.entry_3.pack(side="left")
            self.label_2.pack(side="left")
            self.entry_2.pack(side="left")
            self.label_1.pack(side="left")
            self.entry_1.pack(side="left")
            self.entries = [self.entry_3, self.entry_2, self.entry_1]
        elif DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            self.entry_1.pack(side="left")
            self.label_1.pack(side="left")
            self.entry_2.pack(side="left")
            self.label_2.pack(side="left")
            self.entry_3.pack(side="left")
            self.entries = [self.entry_1, self.entry_2, self.entry_3]
        elif DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            self.entry_2.pack(side="left")
            self.label_2.pack(side="left")
            self.entry_3.pack(side="left")
            self.label_1.pack(side="left")
            self.entry_1.pack(side="left")
            self.entries = [self.entry_2, self.entry_3, self.entry_1]

        self.DATE_FORM = DATE_FORM

    def e1_back(self, event):
        x = self.sv_1.get()
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            if not x:
                self.entry_2.icursor(2)
                self.entry_2.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_1.set(x)
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_1.set(x)
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            if not x:
                self.entry_3.icursor(2)
                self.entry_3.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_1.set(x)
        return "break"

    def e2_back(self, event):
        x = self.sv_2.get()
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            if not x:
                self.entry_3.icursor(2)
                self.entry_3.focus_set()
                return "break"
            elif len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_2.set(x)
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            if not x:
                self.entry_1.icursor(4)
                self.entry_1.focus_set()
                return "break"
            elif len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_2.set(x)
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_2.set(x)
        return "break"

    def e3_back(self, event):
        x = self.sv_3.get()
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_3.set(x)
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            if not x:
                self.entry_2.icursor(2)
                self.entry_2.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_3.set(x)
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            if not x:
                self.entry_2.icursor(2)
                self.entry_2.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_3.set(x)
        return "break"

    def validate_1(self, sv):
        year = []
        for i, c in enumerate(self.sv_1.get()):
            if c in self.allowed_chars:
                year.append(c)
            if i > 3:
                break
        year = "".join(year)
        if len(year) > 4:
            year = year[:4]
        self.entry_1.set_my_value(year)
        if len(year) == 4:
            if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
                pass
            if self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
                self.entry_2.focus_set()
            elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
                pass

    def validate_2(self, sv):
        month = []
        for i, c in enumerate(self.sv_2.get()):
            if c in self.allowed_chars:
                month.append(c)
            if i > 1:
                break
        move_cursor = False
        if len(month) > 2:
            month = month[:2]
        if not month:
            self.entry_2.set_my_value("")
            return
        if len(month) == 1 and int(month[0]) > 1:
            month = ["0", month[0]]
            move_cursor = True
        elif len(month) == 1 and int(month[0]) <= 1:
            return
        e0 = int(month[0])
        e1 = int(month[1])
        if e0 > 1:
            month[0] = "1"
            int(month[0])
        if e1 > 2 and e0 > 0:
            month[0] = "0"
            month[1] = str(e1)
        self.entry_2.set_my_value("".join(month))
        if move_cursor:
            self.entry_2.icursor(2)
        if len(month) == 2:
            if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
                self.entry_1.focus_set()
            if self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
                self.entry_3.focus_set()
            elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
                self.entry_3.focus_set()

    def validate_3(self, sv):
        day = []
        for i, c in enumerate(self.sv_3.get()):
            if c in self.allowed_chars:
                day.append(c)
            if i > 1:
                break
        move_cursor = False
        if len(day) > 2:
            day = day[:2]
        if not day:
            self.entry_3.set_my_value("")
            return
        if len(day) == 1 and int(day[0]) > 3:
            day = ["0", day[0]]
            move_cursor = True
        elif len(day) == 1 and int(day[0]) <= 3:
            return
        e0 = int(day[0])
        e1 = int(day[1])
        if e0 > 3:
            day[0] = "1"
            int(day[0])
        if e0 >= 3 and e1 > 1:
            day[0] = "0"
            day[1] = str(e1)
        self.entry_3.set_my_value("".join(day))
        if move_cursor:
            self.entry_3.icursor(2)
        if len(day) == 2:
            if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
                self.entry_2.focus_set()
            if self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
                pass
            elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
                self.entry_1.focus_set()

    def get_my_value(self):
        return self.sep.join([e.get() for e in self.entries])

    def set_my_value(self, date):
        date = re.split("|".join(map(re.escape, ("/", "-"))), date)
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            try:
                self.sv_1.set(date[2])
            except Exception:
                pass
            try:
                self.sv_2.set(date[1])
            except Exception:
                pass
            try:
                self.sv_3.set(date[0])
            except Exception:
                pass
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                self.sv_1.set(date[0])
            except Exception:
                pass
            try:
                self.sv_2.set(date[1])
            except Exception:
                pass
            try:
                self.sv_3.set(date[2])
            except Exception:
                pass
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            try:
                self.sv_1.set(date[2])
            except Exception:
                pass
            try:
                self.sv_2.set(date[0])
            except Exception:
                pass
            try:
                self.sv_3.set(date[1])
            except Exception:
                pass

    def place_cursor(self, index=0):
        self.entries[index].focus()


class Frame(tk.Frame):
    def __init__(self, parent, background="white", highlightbackground="white", highlightthickness=0, theme="dark"):
        tk.Frame.__init__(
            self,
            parent,
            background=themes[theme].top_left_bg,
            highlightbackground=highlightbackground,
            highlightthickness=highlightthickness,
        )


class Button(ttk.Button):
    def __init__(self, parent, style="Std.TButton", text="", command=None, state="normal", underline=-1):
        ttk.Button.__init__(self, parent, style=style, text=text, command=command, state=state, underline=underline)

    def change_text(self, text):
        self.config(text=text)
        self.update_idletasks()


class Status_Bar(tk.Label):
    def __init__(self, parent, text, theme="dark", font=("Calibri", std_font_size, "normal")):
        tk.Label.__init__(
            self,
            parent,
            text=text,
            font=font,
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_selected_box_cells_fg,
            anchor="w",
        )
        self.text = text
        self._font = font

    def change_text(self, text="", font=None):
        self.config(text=text, font=self._font if font is None else font)
        self.text = text
        self.update_idletasks()


class Label(tk.Label):
    def __init__(self, parent, text, font, theme="dark", anchor="center"):
        tk.Label.__init__(
            self,
            parent,
            text=text,
            font=font,
            background=themes[theme].top_left_bg,
            foreground=themes[theme].table_fg,
            anchor=anchor,
        )

    def change_text(self, text):
        self.config(text=text)
        self.update_idletasks()


class Display_Label(tk.Label):
    def __init__(self, parent, text, font, theme="dark"):
        tk.Label.__init__(self, parent, background=themes[theme].top_left_bg, text=text, font=font)
        self.config(anchor="w")

    def change_text(self, text):
        self.config(text=text)
        self.update_idletasks()
