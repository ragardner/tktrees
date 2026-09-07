# SPDX-License-Identifier: AGPL-3.0-only
# Copyright © R. A. Gardner

from __future__ import annotations

import csv
import json
import os
import pickle
import re
import tkinter as tk
import zlib
from bisect import bisect_left
from collections import defaultdict, deque
from collections.abc import Generator, Iterator, Sequence
from contextlib import suppress
from itertools import cycle, islice, repeat
from math import floor
from operator import attrgetter
from tkinter import filedialog, font, ttk
from typing import Literal

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from tksheet import (
    ICON_ADD,
    ICON_CLEAR,
    ICON_COPY,
    ICON_CUT,
    ICON_DEL,
    ICON_EDIT,
    ICON_PASTE,
    ICON_REDO,
    ICON_SELECT_ALL,
    ICON_SORT_ASC,
    ICON_SORT_DESC,
    ICON_UNDO,
    DotDict,
    Highlight,
    Sheet,
    is_contiguous,
    move_elements_by_mapping,
    push_n,
)
from tksheet import (
    num2alpha as _n2a,
)

from .changelog import display_rows, flatten_changelog
from .classes import (
    Header,
    SearchResult,
    TreeBuilder,
)
from .constants import (
    BF,
    EF,
    align_c_icon,
    align_e_icon,
    align_w_icon,
    changelog_header,
    ctrl_button,
    ctrl_rc_press,
    letters_icon,
    menu_kwargs,
    rc_button,
    rc_motion,
    rc_press,
    rc_release,
    remove_nrt,
    right_icon,
    search_icon,
    sheet_bindings,
    sheet_header_font,
    software_version_number,
    tag_icon,
    themes,
    tree_bindings,
    tv_lvls_colors,
    warnings_header,
)
from .functions import (
    convert_old_xl_to_xlsx,
    create_cell_align_selector_menu,
    csv_str_x_data,
    dict_x_b32,
    equalize_sublist_lens,
    frame_w_to_nchars,
    full_sheet_to_dict,
    increment_file_version,
    json_to_sheet,
    level_to_color,
    new_info_storage,
    new_saved_info,
    path_numbers,
    path_without_numbers,
    process_search_results,
    search_results_max_column_chars,
    str_io_csv_writer,
    to_clipboard,
    try_remove,
    xlsx_changelog_header,
)
from .session import Session
from .toplevels import (
    Add_Child_Or_Sibling_Id_Popup,
    Add_Detail_Column_Popup,
    Add_Hierarchy_Column_Popup,
    Add_Top_Id_Popup,
    Ask_Confirm,
    Changelog_Popup,
    Delete_Ids_Using_List_Popup,
    Edit_Conditional_Formatting_Popup,
    Edit_Detail_Text_Popup,
    Edit_Validation_Popup,
    Enter_Sheet_Name_Popup,
    Error,
    Export_Flattened_Popup,
    Get_Clipboard_Data_Popup,
    Merge_Sheets_Popup,
    Post_Import_Changes_Popup,
    Rename_Column_Popup,
    Rename_Id_Popup,
    Replace_Popup,
    Save_New_Version_Error_Popup,
    Save_New_Version_Postsave_Popup,
    Save_New_Version_Presave_Popup,
    Settings_Popup,
    Sort_Sheet_Popup,
    Tag_Ids_Using_List_Popup,
    Text_Popup,
    Treeview_Id_Finder,
    View_Id_Popup,
)
from .widgets import (
    Button,
    Ez_Dropdown,
    Frame,
    Normal_Entry,
)

# DEFAULT SETTING FOR SAVING WITH PROGRAM DATA
save_xlsx_and_json_with_program_data = True


class _Fwd:
    def __init__(self, name: str):
        self.name = name

    def __get__(self, obj, owner):
        if obj is None:
            return self
        return getattr(obj.session, self.name)

    def __set__(self, obj, value):
        setattr(obj.session, self.name, value)


class SheetOps:
    def __init__(self, editor: Tree_Editor):
        self.e = editor

    def insert_rows(self, rows, idx=None):
        # tksheet treats a string idx as an Excel letter (alpha2idx("end") == 3747).
        # None means append, same as insert_row() on main.
        self.e.sheet.insert_rows(
            rows=rows,
            idx=idx,
            undo=False,
            emit_event=False,
            redraw=False,
            create_selections=False,
        )

    def delete_rows(self, idxs):
        self.e.sheet.del_rows(idxs, undo=False, emit_event=False, redraw=False)

    def insert_cols(self, idx, n=1):
        if self.e.sheet.MT.data:
            kw = {
                "idx": idx,
                "undo": False,
                "emit_event": False,
                "redraw": False,
                "create_selections": False,
                "add_row_heights": False,
            }
            self.e.sheet.insert_columns(n, **kw)
            self.e.tree.insert_columns(n, **kw)
        else:
            self.e.tree.insert_column_positions(idx=idx, widths=n)
            self.e.sheet.insert_column_positions(idx=idx, widths=n)

    def delete_cols(self, idxs):
        self.e.sheet.del_columns(idxs, undo=False, emit_event=False, redraw=False)
        self.e.tree.del_columns(idxs, undo=False, emit_event=False, redraw=False)


class Tree_Editor(tk.Frame):
    nodes = _Fwd("nodes")
    rns = _Fwd("rns")
    headers = _Fwd("headers")
    ic = _Fwd("ic")
    pc = _Fwd("pc")
    hiers = _Fwd("hiers")
    row_len = _Fwd("row_len")
    changelog = _Fwd("changelog")
    changelog_at_open = _Fwd("changelog_at_open")
    warnings = _Fwd("warnings")
    tagged_ids = _Fwd("tagged_ids")
    topnodes_order = _Fwd("topnodes_order")
    auto_sort_nodes_bool = _Fwd("auto_sort_nodes_bool")
    allow_spaces_ids_var = _Fwd("allow_spaces_ids_var")
    allow_spaces_columns_var = _Fwd("allow_spaces_columns_var")
    vs = _Fwd("vs")
    refresh_rows = _Fwd("refresh_rows")
    sort_later_dct = _Fwd("sort_later_dct")

    @property
    def data(self):
        return self.session.data

    @data.setter
    def data(self, rows):
        self.set_records(rows)

    def __init__(self, parent, C):
        tk.Frame.__init__(self, parent)
        self.C = C
        self.session = Session()
        # try:
        #     self.monitor_scale = self.C.call("tk", "scaling")
        # except Exception:
        #     self.monitor_scale = 1
        self.l_frame_proportion = 0.50
        self.last_width = 0
        self.last_height = 0
        self.currently_adjusting_divider = False
        self.tree_has_focus = True
        self.sheet_has_focus = False
        self.levels = defaultdict(list)
        self.treecolsel = 0
        self.tv_label_col = 0
        self.reset_tree_drag_vars()
        self.rc_iid = None
        self.row_cut_updated = False
        self.mirror_sels_disabler = False
        self.date_split_regex = "|".join(map(re.escape, ("/", "-")))
        self.find_popup = None
        self.fixed_font_w = font.nametofont("TkFixedFont").measure("0")

        self.auto_resize_indexes = True
        self.mirror_var = False
        self.save_xlsx_with_program_data = bool(save_xlsx_and_json_with_program_data)
        self.save_json_with_program_data = bool(save_xlsx_and_json_with_program_data)
        self.save_xlsx_with_changelog = False
        self.save_xlsx_with_treeview = False
        self.save_xlsx_with_flattened = False
        self.xlsx_flattened_detail_columns = True
        self.xlsx_flattened_justify = True
        self.xlsx_flattened_reverse_order = False
        self.xlsx_flattened_add_index = False
        self.json_format = 1
        self.black_theme_bool = self.C.theme == "black"
        self.dark_blue_theme_bool = self.C.theme == "dark_blue"
        self.dark_theme_bool = self.C.theme == "dark"
        self.light_green_theme_bool = self.C.theme == "light_green"
        self.light_blue_theme_bool = self.C.theme == "light_blue"
        self.tv_lvls_bool = False

        self.warnings_filepath = ""
        self.warnings_sheet = ""

        # cell alignment menu images
        self.icons = {
            "w": tk.PhotoImage(format="png", data=align_w_icon),
            "c": tk.PhotoImage(format="png", data=align_c_icon),
            "e": tk.PhotoImage(format="png", data=align_e_icon),
            "letters": tk.PhotoImage(format="png", data=letters_icon),
            "tag": tk.PhotoImage(format="png", data=tag_icon),
            "search": tk.PhotoImage(format="png", data=search_icon),
            "right": tk.PhotoImage(format="png", data=right_icon),
            "ICON_ADD": tk.PhotoImage(format="png", data=ICON_ADD),
            "ICON_CLEAR": tk.PhotoImage(format="png", data=ICON_CLEAR),
            "ICON_COPY": tk.PhotoImage(format="png", data=ICON_COPY),
            "ICON_CUT": tk.PhotoImage(format="png", data=ICON_CUT),
            "ICON_DEL": tk.PhotoImage(format="png", data=ICON_DEL),
            "ICON_EDIT": tk.PhotoImage(format="png", data=ICON_EDIT),
            "ICON_PASTE": tk.PhotoImage(format="png", data=ICON_PASTE),
            "ICON_REDO": tk.PhotoImage(format="png", data=ICON_REDO),
            "ICON_SELECT_ALL": tk.PhotoImage(format="png", data=ICON_SELECT_ALL),
            "ICON_SORT_ASC": tk.PhotoImage(format="png", data=ICON_SORT_ASC),
            "ICON_SORT_DESC": tk.PhotoImage(format="png", data=ICON_SORT_DESC),
            "ICON_UNDO": tk.PhotoImage(format="png", data=ICON_UNDO),
        }

        self.C.file.entryconfig("Save", command=self.save_)
        self.C.file.entryconfig(
            "Save as",
            accelerator="Ctrl+Shift+S",
            command=self.save_as,
        )
        self.C.file.entryconfig("Save new version", command=self.save_new_vrsn)
        self.C.file.entryconfig("Settings", command=self.settings)

        self.edit_menu = tk.Menu(self.C.menubar, tearoff=0, **menu_kwargs)
        self.C.menubar.add_cascade(
            label="Edit",
            menu=self.edit_menu,
            state="disabled",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Undo  0/30",
            accelerator="Ctrl+Z",
            state="disabled",
            command=self.undo,
            image=self.icons["ICON_UNDO"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_separator()
        self.copy_clipboard_menu = tk.Menu(self.edit_menu, tearoff=0, **menu_kwargs)
        self.copy_clipboard_menu.add_command(
            label="Copy sheet to clipboard (indent separated)",
            command=self.clipboard_sheet_indent,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.copy_clipboard_menu.add_command(
            label="Copy sheet to clipboard (comma separated)",
            command=self.clipboard_sheet,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.copy_clipboard_menu.add_command(
            label="Copy sheet to clipboard as json",
            command=self.clipboard_sheet_json,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Sort sheet",
            command=self.sort_sheet_choice,
            image=self.icons["ICON_SORT_ASC"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_cascade(
            label="Copy to clipboard",
            menu=self.copy_clipboard_menu,
            image=self.icons["ICON_COPY"],
            compound="left",
            state="normal",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Tag/Untag IDs",
            command=self.tag_ids,
            accelerator="Ctrl+T",
            image=self.icons["tag"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_separator()
        self.edit_menu.add_command(
            label="Tag IDs using list",
            command=self.tag_ids_using_list,
            image=self.icons["tag"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Delete IDs using list",
            command=self.delete_ids_using_list,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Replace using mapping",
            command=self.replace_using_mapping,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_separator()
        self.edit_menu.add_command(
            label="Clear copied/cut",
            command=self.clear_copied_details,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Clear panel selections",
            command=self.remove_selections,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.edit_menu.add_command(
            label="Clear all tagged IDs",
            command=self.clear_tagged_ids,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )

        # view menu
        self.view_menu = tk.Menu(self.C.menubar, tearoff=0, **menu_kwargs)
        self.C.menubar.add_cascade(label="View", menu=self.view_menu, state="disabled", **menu_kwargs)
        self.view_menu.add_command(
            label="View changelog",
            accelerator="Ctrl+L",
            command=self.show_changelog,
            **menu_kwargs,
        )
        self.view_menu.add_command(
            label="View build warnings",
            command=lambda: self.show_warnings(show_regardless=True),
            **menu_kwargs,
        )
        # self.view_menu.add_command(label="View organizational chart", command=self.show_org_chart,**menu_kwargs)
        self.view_menu.add_separator()
        self.view_menu.add_command(
            label="Treeview IDs information",
            command=self.show_ids_full_info_tree,
            **menu_kwargs,
        )
        self.view_menu.add_command(
            label="Sheet IDs information",
            command=self.show_ids_full_info_sheet,
            **menu_kwargs,
        )
        self.view_menu.add_separator()
        self.view_menu.add_command(
            label="Expand ID",
            accelerator="Ctrl+E",
            command=self.expand_id,
            **menu_kwargs,
        )
        self.view_menu.add_command(
            label="Collapse ID",
            accelerator="Ctrl+R",
            command=self.collapse_id,
            **menu_kwargs,
        )
        self.view_menu.add_separator()
        self.view_menu.add_command(
            label="Zoom in",
            accelerator="Ctrl++",
            command=self.zoom_in,
            **menu_kwargs,
        )
        self.view_menu.add_command(
            label="Zoom out",
            accelerator="Ctrl+-",
            command=self.zoom_out,
            **menu_kwargs,
        )
        self.view_menu.add_separator()
        self.adjustable_bool = tk.BooleanVar()
        self.adjustable_bool.set(False)
        self._50_50_bool = tk.BooleanVar()
        self._50_50_bool.set(False)
        self.full_left_bool = tk.BooleanVar()
        self.full_left_bool.set(True)
        self.full_right_bool = tk.BooleanVar()
        self.full_right_bool.set(False)
        self.display_menu = tk.Menu(self.view_menu, tearoff=0, **menu_kwargs)
        self.display_menu.add_checkbutton(
            label="Display Only Tree",
            variable=self.full_left_bool,
            command=self.option_full_left,
            **menu_kwargs,
        )
        self.display_menu.add_checkbutton(
            label="Display Only Sheet",
            variable=self.full_right_bool,
            command=self.option_full_right,
            **menu_kwargs,
        )
        self.display_menu.add_checkbutton(
            label="50/50 Tree/Sheet",
            variable=self._50_50_bool,
            command=self.option_50_50,
            **menu_kwargs,
        )
        self.display_menu.add_checkbutton(
            label="Adjustable Display",
            variable=self.adjustable_bool,
            command=self.option_adjustable,
            **menu_kwargs,
        )
        self.view_menu.add_cascade(
            label="Layout",
            menu=self.display_menu,
            state="normal",
            **menu_kwargs,
        )
        self.view_menu.add_command(
            label="Set all column widths",
            command=self.set_all_col_widths,
            **menu_kwargs,
        )

        # import menu
        self.import_menu = tk.Menu(self.C.menubar, tearoff=0, **menu_kwargs)
        self.C.menubar.add_cascade(
            label="Import",
            menu=self.import_menu,
            state="disabled",
            **menu_kwargs,
        )
        self.import_menu.add_command(
            label="Import changes",
            command=self.import_changes,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.import_menu.add_command(
            label="Paste Clipboard & Overwrite Sheet",
            command=self.get_clipboard_data,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.import_menu.add_command(
            label="Merge Sheets / Add rows",
            command=self.merge_sheets,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )

        # export menu
        self.export_menu = tk.Menu(self.C.menubar, tearoff=0, **menu_kwargs)
        self.C.menubar.add_cascade(
            label="Export",
            menu=self.export_menu,
            state="disabled",
            **menu_kwargs,
        )
        self.export_menu.add_command(
            label="Export specific changes",
            accelerator="Ctrl+L",
            command=lambda: self.show_changelog("specific"),
            **menu_kwargs,
        )
        self.export_menu.add_command(
            label="Export file session changes",
            command=lambda: self.show_changelog("sheet"),
            **menu_kwargs,
        )
        self.export_menu.add_command(
            label="Export all changes",
            command=lambda: self.show_changelog("all"),
            **menu_kwargs,
        )
        self.export_menu.add_command(
            label="Export flattened sheet",
            command=self.export_flattened,
            **menu_kwargs,
        )

        # help menu
        self.help_menu = tk.Menu(self.C.menubar, tearoff=0, **menu_kwargs)
        self.C.menubar.add_cascade(label="Help", menu=self.help_menu, state="normal", **menu_kwargs)
        self.help_menu.add_command(label="View Help", command=self.C.help_func, **menu_kwargs)
        self.help_menu.add_command(label="View License", command=self.C.license_func, **menu_kwargs)
        self.help_menu.add_command(label="About", command=self.C.about_func, **menu_kwargs)

        # MAIN CANVAS
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.main_canvas = tk.Canvas(self, highlightthickness=0)
        self.main_canvas.grid(row=0, column=0, sticky="nswe")

        # ======================= LEFT FRAME ======================================================
        self.l_frame = tk.Frame(
            self.main_canvas,
            highlightbackground="white",
            highlightcolor="white",
            highlightthickness=2,
        )

        # frames for left frame
        self.btns_tree = Frame(self.l_frame)
        self.btns_tree.pack(side="top", fill="x")
        self.btns_tree.grid_rowconfigure(0, weight=1)
        self.btns_tree.grid_rowconfigure(1, weight=1)
        self.btns_tree.grid_columnconfigure(4, weight=1)
        # self.btns_tree.grid_columnconfigure(4, weight=1)

        self.treeframe = Frame(self.l_frame)
        self.treeframe.pack(side="top", fill="both", expand=True)
        self.treeframe.grid_rowconfigure(0, weight=1)
        self.treeframe.grid_columnconfigure(0, weight=1)

        self.tree = Sheet(
            self.treeframe,
            name="tree",
            header_font=sheet_header_font,
            theme=self.C.theme,
            default_column_width=200,
            treeview=True,
            row_drag_and_drop_perform=False,
            alternate_color="#f6f9fb",
            allow_cell_overflow=True,
            max_undos=0,
        )
        self.tree.grid(row=0, column=0, sticky="nswe")

        # status bar tree
        self.sts_tree = Frame(self.l_frame)
        self.sts_tree.pack(side="top", fill="x")
        self.sts_tree.grid_rowconfigure(0, weight=1)

        # buttons for bottom left frame
        # switch hierarchy dropdown
        self.switch_label = Button(
            self.btns_tree,
            text="Hierarchy:",
            command=self.next_hier,
        )
        self.switch_label.grid(row=0, column=0, sticky="nswe")
        self.switch_displayed = tk.StringVar(self.btns_tree)
        self.switch_displayed.set("")
        self.switch_hier_dropdown = ttk.Combobox(
            self.btns_tree,
            textvariable=self.switch_displayed,
            state="readonly",
            font=BF,
        )
        self.switch_hier_dropdown.grid(row=0, column=1, sticky="nswe")
        self.switch_hier_dropdown.bind("<<ComboboxSelected>>", self.switch_hier)

        # tag ID tree
        self.tree_tag_id_button = Button(self.btns_tree, text="Tag ID:", underline=0, command=self.tag_ids)
        self.tree_tag_id_button.grid(row=1, column=0, ipady=1, sticky="nswe")
        self.tree_tagged_ids_dropdown = Ez_Dropdown(self.btns_tree, font=BF)
        self.tree_tagged_ids_dropdown.grid(row=1, column=1, sticky="nswe")
        self.tree_tagged_ids_dropdown.bind("<<ComboboxSelected>>", self.tree_go_to_tagged_id)

        # buttons for top left frame
        # tree search function
        self.search_displayed = tk.StringVar(self.btns_tree)
        self.search_displayed.set("")
        self.search_button = Button(self.btns_tree, text=" Find:", command=self.search_choice)
        self.search_button.grid(row=0, column=2, sticky="nswe")
        self.search_choice_displayed = tk.StringVar(self.btns_tree)
        self.search_choice_displayed.set("Non-exact")
        self.search_choice_dropdown = ttk.Combobox(
            self.btns_tree,
            textvariable=self.search_choice_displayed,
            state="readonly",
            font=BF,
        )
        self.search_choice_dropdown.config(width=13)
        self.search_choice_dropdown["values"] = [
            "Non-exact",
            "ID non-exact",
            "ID exact",
            "Detail non-exact",
            "Detail exact",
        ]
        self.search_choice_dropdown.grid(row=0, column=3, sticky="nswe")
        self.search_entry = Normal_Entry(self.btns_tree, font=BF, theme="light_blue")
        self.search_entry.grid(row=0, column=4, sticky="nswe")
        self.search_entry.bind("<Return>", self.search_choice)
        self.search_dropdown = ttk.Combobox(
            self.btns_tree,
            textvariable=self.search_displayed,
            state="readonly",
            font=EF,
        )
        self.search_dropdown["values"] = []
        self.search_dropdown.bind("<<ComboboxSelected>>", self.show_search_result)
        self.search_choice_dropdown.bind("<<ComboboxSelected>>", lambda focus: self.search_entry.focus_set())
        self.search_dropdown.grid(row=1, column=2, columnspan=3, sticky="nswe")

        # ======================= RIGHT FRAME ======================================================
        self.r_frame = tk.Frame(
            self.main_canvas,
            highlightbackground="white",
            highlightcolor="white",
            highlightthickness=2,
        )

        # frames for right frame
        self.btns_sheet = Frame(self.r_frame)
        self.btns_sheet.pack(side="top", fill="x")
        self.btns_sheet.grid_rowconfigure(0, weight=1)
        self.btns_sheet.grid_rowconfigure(1, weight=1)
        self.btns_sheet.grid_columnconfigure(3, weight=1)

        self.sheetframe = Frame(self.r_frame)
        self.sheetframe.pack(side="top", fill="both", expand=True)

        self.sheet = Sheet(
            self.sheetframe,
            name="sheet",
            theme=self.C.theme,
            row_index_align="w",
            auto_resize_row_index=True,
            header_font=sheet_header_font,
            allow_cell_overflow=True,
            max_undos=0,
        )
        self.sheet.pack(side="right", fill="both", expand=True)
        self.set_records(self.session.data)
        self.session.ops = SheetOps(self)
        self.session.on_increment_unsaved = self.increment_unsaved
        self.session.on_changelog_row = lambda: None

        # buttons for top right frame
        # tag ID
        self.sheet_tag_id_button = Button(self.btns_sheet, text="↓Tag ID", underline=0, command=self.tag_ids)
        self.sheet_tag_id_button.grid(row=0, column=0, ipady=1, sticky="nswe")
        self.sheet_tagged_ids_dropdown = Ez_Dropdown(self.btns_sheet, font=BF)
        self.sheet_tagged_ids_dropdown.grid(row=1, column=0, sticky="nswe")
        self.sheet_tagged_ids_dropdown.bind("<<ComboboxSelected>>", self.sheet_go_to_tagged_id)

        # sheet search function
        self.sheet_search_displayed = tk.StringVar(self.btns_sheet)
        self.sheet_search_displayed.set("")
        self.sheet_search_button = Button(self.btns_sheet, text="Find:", command=self.sheet_search_choice)
        self.sheet_search_button.grid(row=0, column=1, sticky="nswe")
        self.sheet_search_choice_displayed = tk.StringVar(self.btns_sheet)
        self.sheet_search_choice_displayed.set("Non-exact")
        self.sheet_search_choice_dropdown = ttk.Combobox(
            self.btns_sheet,
            textvariable=self.sheet_search_choice_displayed,
            state="readonly",
            font=BF,
        )
        self.sheet_search_choice_dropdown.config(width=13)
        self.sheet_search_choice_dropdown["values"] = [
            "Non-exact",
            "ID non-exact",
            "ID exact",
            "Detail non-exact",
            "Detail exact",
        ]
        self.sheet_search_choice_dropdown.grid(row=0, column=2, sticky="nswe")
        self.sheet_search_entry = Normal_Entry(self.btns_sheet, font=BF, theme="light_blue")
        self.sheet_search_entry.grid(row=0, column=3, sticky="nswe")
        self.sheet_search_entry.bind("<Return>", self.sheet_search_choice)
        self.sheet_search_dropdown = ttk.Combobox(
            self.btns_sheet,
            textvariable=self.sheet_search_displayed,
            state="readonly",
            font=EF,
        )
        self.sheet_search_dropdown["values"] = []
        self.sheet_search_dropdown.bind("<<ComboboxSelected>>", self.sheet_show_search_result)
        self.sheet_search_choice_dropdown.bind(
            "<<ComboboxSelected>>", lambda focus: self.sheet_search_entry.focus_set()
        )
        self.sheet_search_dropdown.grid(row=1, column=1, columnspan=3, sticky="nswe")

        # RIGHT CLICK MENUS

        # SINGLE CELL MENU - SHEET AND TREE
        self.tree_sheet_rc_menu_single_cell = tk.Menu(self.sheet, tearoff=0, **menu_kwargs)
        self.tree_sheet_rc_menu_single_cell.add_command(
            label="Detail",
            state="disabled",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_cell.add_command(
            label="Edit",
            command=self.tree_sheet_edit_detail,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_cell.add_command(
            label="Cut",
            accelerator="Ctrl+X",
            command=self.cut_key,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_cell.add_command(
            label="Copy",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_cell.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_cell.add_command(
            label="Clear contents",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )

        # MULTI CELL MENU - SHEET AND TREE
        self.tree_sheet_rc_menu_multi_cell = tk.Menu(
            self.sheet,
            tearoff=0,
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_cell.add_command(
            label="Cut",
            accelerator="Ctrl+X",
            command=self.cut_key,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_cell.add_command(
            label="Copy",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_cell.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_cell.add_command(
            label="Clear contents",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )

        # SINGLE COLUMN MENU - SHEET AND TREE
        self.tree_sheet_rc_menu_single_col = tk.Menu(
            self.sheet,
            tearoff=0,
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col_align = create_cell_align_selector_menu(
            parent=self.tree_sheet_rc_menu_single_col,
            command=self.tree_sheet_align,
            menu_kwargs=menu_kwargs,
            icons=self.icons,
        )
        self.tree_sheet_rc_menu_single_col.add_cascade(
            label="Alignment",
            menu=self.tree_sheet_rc_menu_single_col_align,
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Cut",
            accelerator="Ctrl+X",
            command=self.cut_key,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Copy",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Clear contents",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Delete column",
            command=self.del_cols_rc,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_separator()
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Add detail",
            command=self.rc_add_col,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Add hierarchy",
            command=self.rc_add_hier_col,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Rename column",
            command=self.rc_rename_col,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_separator()
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Validation",
            command=self.rc_edit_validation,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Conditional Formatting",
            command=self.rc_edit_formatting,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_separator()
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Sort sheet A → Z",
            command=self.sort_sheet_rc_asc,
            image=self.icons["ICON_SORT_ASC"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Sort sheet Z → A",
            command=self.sort_sheet_rc_desc,
            image=self.icons["ICON_SORT_DESC"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Sort sheet tree walk",
            command=self.sort_sheet_walk,
            image=self.icons["ICON_SORT_ASC"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_single_col.add_separator()
        self.tree_sheet_rc_menu_single_col.add_command(
            label="Set as treeview label",
            command=self.sheet_rc_tv_label,
            image=self.icons["letters"],
            compound="left",
            **menu_kwargs,
        )

        # MULTI COLUMN MENU - SHEET AND TREE
        self.tree_sheet_rc_menu_multi_col = tk.Menu(
            self.sheet,
            tearoff=0,
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_col_align = create_cell_align_selector_menu(
            parent=self.tree_sheet_rc_menu_multi_col,
            command=self.tree_sheet_align,
            menu_kwargs=menu_kwargs,
            icons=self.icons,
        )
        self.tree_sheet_rc_menu_multi_col.add_cascade(
            label="Alignment",
            menu=self.tree_sheet_rc_menu_multi_col_align,
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_col.add_command(
            label="Cut",
            accelerator="Ctrl+X",
            command=self.cut_key,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_col.add_command(
            label="Copy",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_col.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_col.add_command(
            label="Clear contents",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_sheet_rc_menu_multi_col.add_command(
            label="Delete columns",
            command=self.del_cols_rc,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )

        # MULTI ROW MENU - TREE
        self.tree_rc_menu_multi_row = tk.Menu(self.treeframe, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_multi_row_select = self._create_tree_select_menu(self.tree_rc_menu_multi_row)
        self.tree_rc_menu_multi_row.add_cascade(
            label="Select",
            menu=self.tree_rc_menu_multi_row_select,
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_command(
            label="Cut",
            accelerator="Ctrl+X",
            command=self.cut_ids,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_command(
            label="Copy",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_command(
            label="Clipboard",
            command=self.copy_ID_row,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_command(
            label="Clipboard IDs & descendants",
            command=self.copy_ID_children_rows,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_command(
            label="Paste details",
            state="disabled",
            command=self.paste_details,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_separator()
        self.tree_rc_menu_multi_row.add_command(
            label="Tag/Untag IDs",
            accelerator="Ctrl+T",
            command=self.tag_ids,
            image=self.icons["tag"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row.add_separator()
        self.tree_rc_menu_multi_row_del = tk.Menu(self.tree_rc_menu_multi_row, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_multi_row.add_cascade(
            label="Delete",
            menu=self.tree_rc_menu_multi_row_del,
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row_del.add_command(
            label="Clear IDs details",
            command=self.del_all_details,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row_del.add_separator()
        self.tree_rc_menu_multi_row_del.add_command(
            label="Delete IDs",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row_del.add_separator()
        self.tree_rc_menu_multi_row_del.add_command(
            label="Delete IDs all hierarchies",
            command=self.del_id_all,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row_del.add_separator()
        self.tree_rc_menu_multi_row_del.add_command(
            label="Delete IDs + children",
            command=self.del_id_children,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_multi_row_del.add_command(
            label="Delete IDs + children, all hierarchies",
            command=self.del_id_children_all,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )

        # SINGLE ROW MENU - TREE
        self.tree_rc_menu_single_row = tk.Menu(self.treeframe, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_single_row_add = tk.Menu(self.tree_rc_menu_single_row, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_single_row.add_cascade(label="Add", menu=self.tree_rc_menu_single_row_add, **menu_kwargs)
        self.tree_rc_menu_single_row_select = self._create_tree_select_menu(self.tree_rc_menu_single_row)
        self.tree_rc_menu_single_row.add_cascade(
            label="Select",
            menu=self.tree_rc_menu_single_row_select,
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_add.add_command(
            label="Add child",
            command=self.add_child_node,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_add.add_command(
            label="Add sibling",
            command=self.add_sibling_node,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_add.add_command(
            label="Add top ID",
            command=self.add_top_node,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )

        # cut
        self.tree_rc_menu_single_row_cut = tk.Menu(self.tree_rc_menu_single_row, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_single_row_cut.add_command(
            label="Detach ID",
            accelerator="Ctrl+X",
            command=self.cut_ids,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_cut.add_command(
            label="Detach children",
            command=self.cut_children,
            image=self.icons["ICON_CUT"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row.add_cascade(
            label="Cut",
            menu=self.tree_rc_menu_single_row_cut,
            **menu_kwargs,
        )

        # copy
        self.tree_rc_menu_single_row_copy = tk.Menu(self.tree_rc_menu_single_row, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_single_row_copy.add_command(
            label="Copy ID",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_copy.add_command(
            label="Clipboard row",
            command=self.copy_ID_row,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_copy.add_command(
            label="Clipboard ID & descendants",
            command=self.copy_ID_children_rows,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_copy.add_command(
            label="Copy details",
            command=self.copy_details,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row.add_cascade(label="Copy", menu=self.tree_rc_menu_single_row_copy, **menu_kwargs)

        # paste options
        self.tree_rc_menu_single_row_paste = tk.Menu(self.tree_rc_menu_single_row, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_single_row.add_cascade(
            label="Paste",
            state="normal",
            menu=self.tree_rc_menu_single_row_paste,
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_paste.add_command(
            label="Paste IDs as sibling",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_paste.add_command(
            label="Paste IDs as child",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_paste.add_separator()
        self.tree_rc_menu_single_row_paste.add_command(
            label="Paste IDs and children as sibling",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_paste.add_command(
            label="Paste IDs and children as child",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_paste.add_separator()
        self.tree_rc_menu_single_row_paste.add_command(
            label="Attach children",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_paste.add_command(
            label="Paste details",
            state="disabled",
            command=self.paste_details,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )

        # delete options
        self.tree_rc_menu_single_row_del = tk.Menu(self.tree_rc_menu_single_row, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_single_row.add_cascade(label="Delete", menu=self.tree_rc_menu_single_row_del, **menu_kwargs)
        self.tree_rc_menu_single_row_del.add_command(
            label="Clear IDs details",
            command=self.del_all_details,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_del.add_separator()
        self.tree_rc_menu_single_row_del.add_command(
            label="Delete ID",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_del.add_command(
            label="Delete ID, orphan children",
            command=self.del_id_orphan,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_del.add_separator()
        self.tree_rc_menu_single_row_del.add_command(
            label="Delete ID all hierarchies",
            command=self.del_id_all,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_del.add_command(
            label="Delete ID all hierarchies, orphan children",
            command=self.del_id_all_orphan,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_del.add_separator()
        self.tree_rc_menu_single_row_del.add_command(
            label="Delete ID + children",
            command=self.del_id_children,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row_del.add_command(
            label="Delete ID + children, all hierarchies",
            command=self.del_id_children_all,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )

        self.tree_rc_menu_single_row.add_separator()
        self.tree_rc_menu_single_row.add_command(
            label="ID concise view",
            command=self.show_ids_details_tree,
            image=self.icons["search"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row.add_command(
            label="Tag/Untag ID",
            accelerator="Ctrl+T",
            command=self.tag_ids,
            image=self.icons["tag"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_single_row.add_command(
            label="Rename ID",
            command=self.rename_node,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )

        # EMPTY MENU - TREE
        self.tree_rc_menu_empty = tk.Menu(self.treeframe, tearoff=0, **menu_kwargs)
        self.tree_rc_menu_empty.add_command(
            label="Paste IDs",
            state="disabled",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_empty.add_command(
            label="Paste IDs and children",
            state="disabled",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_empty.add_command(
            label="Attach children",
            state="disabled",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_empty.add_separator()
        self.tree_rc_menu_empty.add_command(
            label="Add top ID",
            command=self.add_top_node,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_empty.add_command(
            label="Add rows",
            command=self.add_rows_rc,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_empty.add_command(
            label="Add detail",
            command=self.rc_add_col,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.tree_rc_menu_empty.add_command(
            label="Add hierarchy",
            command=self.rc_add_hier_col,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )

        # SINGLE ROW MENU - SHEET
        self.sheet_rc_menu_single_row = tk.Menu(self.sheet, tearoff=0, **menu_kwargs)
        self.sheet_rc_menu_single_row.add_command(
            label="Tag/Untag ID",
            accelerator="Ctrl+T",
            command=self.tag_ids,
            image=self.icons["tag"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="Go to ID in Treeview",
            command=self.select_id_in_treeview_from_sheet,
            image=self.icons["right"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="ID concise view",
            command=self.show_ids_details_sheet,
            image=self.icons["search"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_separator()
        self.sheet_rc_menu_single_row.add_command(
            label="Clipboard",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_separator()
        self.sheet_rc_menu_single_row.add_command(
            label="Copy details",
            command=self.sheet_copy_details,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="Paste details",
            command=self.sheet_paste_details,
            state="disabled",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="Clear IDs details",
            command=self.sheet_del_all_details,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_separator()
        self.sheet_rc_menu_single_row.add_command(
            label="Add top ID",
            command=self.sheet_add_top_node,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="Insert rows",
            command=lambda: self.add_rows_rc(True),
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_command(
            label="Rename ID",
            command=self.sheet_rename_node,
            image=self.icons["ICON_EDIT"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_single_row.add_separator()
        self.sheet_rc_menu_single_row.add_command(
            label="Del IDs, all hierarchies",
            command=self.del_key,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )

        # MULTI ROW MENU - SHEET
        self.sheet_rc_menu_multi_row = tk.Menu(
            self.sheet,
            tearoff=0,
            **menu_kwargs,
        )
        self.sheet_rc_menu_multi_row.add_command(
            label="Tag/Untag IDs",
            accelerator="Ctrl+T",
            command=self.tag_ids,
            image=self.icons["tag"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_multi_row.add_separator()
        self.sheet_rc_menu_multi_row.add_command(
            label="Clear all details",
            command=self.sheet_del_all_details,
            image=self.icons["ICON_CLEAR"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_multi_row.add_command(
            label="Paste details",
            command=self.sheet_paste_details,
            state="disabled",
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_multi_row.add_separator()
        self.sheet_rc_menu_multi_row.add_command(
            label="Clipboard",
            accelerator="Ctrl+C",
            command=self.copy_key,
            image=self.icons["ICON_COPY"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_multi_row.add_command(
            label="Paste",
            accelerator="Ctrl+V",
            command=self.paste_key,
            image=self.icons["ICON_PASTE"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_multi_row.add_separator()
        self.sheet_rc_menu_multi_row.add_command(
            label="Del IDs, all hierarchies",
            accelerator="Del",
            command=self.del_key,
            image=self.icons["ICON_DEL"],
            compound="left",
            **menu_kwargs,
        )

        # EMPTY MENU - SHEET
        self.sheet_rc_menu_empty = tk.Menu(
            self.sheet,
            tearoff=0,
            **menu_kwargs,
        )
        self.sheet_rc_menu_empty.add_command(
            label="Add top ID",
            command=self.sheet_add_top_node,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_empty.add_command(
            label="Add rows",
            command=self.add_rows_rc,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_empty.add_command(
            label="Add detail",
            command=self.rc_add_col,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )
        self.sheet_rc_menu_empty.add_command(
            label="Add hierarchy",
            command=self.rc_add_hier_col,
            image=self.icons["ICON_ADD"],
            compound="left",
            **menu_kwargs,
        )

        self.l_frame_id = self.main_canvas.create_window(
            (0, 0),
            window=self.l_frame,
            anchor="nw",
            state="normal",
        )
        self.r_frame_id = self.main_canvas.create_window(
            (0, 0),
            window=self.r_frame,
            anchor="nw",
            state="normal",
        )
        self.main_canvas.create_rectangle(0, 1, 0, 1, fill="gray60", outline="", tag="div")

    def set_records(self, rows, *, redraw=False, reset_positions=True):
        self.session.data = rows
        self.sheet.data_reference(
            rows,
            reset_col_positions=reset_positions,
            reset_row_positions=reset_positions,
            redraw=redraw,
        )

    def _adopt_sheet_data(self):
        # tksheet column/row drag assigns a new MT.data list. Session must follow
        # or undo remaps the pre-drag rows and rns loses IDs.
        self.session.data = self.sheet.MT.data

    def _sync_sheet_from_session(self, *, redraw=False):
        # In-file update: do not reset widths/heights (Open/New uses set_records).
        if self.sheet.MT.data is self.session.data:
            return
        self.set_records(self.session.data, redraw=redraw, reset_positions=False)

    def populate(self, program_data=None):
        if program_data:
            self.session.load_program_data(program_data)
            self.set_records(self.session.data)
            self.apply_view_program_data(program_data)
            self.populate_view(from_program_data=True)
        else:
            self._apply_raw_view_defaults()
            self.populate_view(from_program_data=False)

    def _apply_raw_view_defaults(self):
        self.set_headers()
        self.tv_label_col = int(self.ic)
        self.saved_info = new_saved_info(self.hiers)
        self.tree.set_column_widths()
        self.sheet.set_row_heights().set_column_widths()

    def apply_view_program_data(self, d):
        table_align = d.get("sheet_table_align")
        if table_align:
            self.sheet.align(table_align, redraw=False)
            self.tree.align(table_align, redraw=False)
        index_align = d.get("sheet_index_align")
        if index_align:
            self.sheet.row_index_align(index_align, redraw=False)
        header_align = d.get("sheet_header_align")
        if header_align:
            self.sheet.header_align(header_align, redraw=False)
            self.tree.header_align(header_align, redraw=False)
        row_heights = d.get("row_heights")
        column_widths = d.get("column_widths")
        if row_heights:
            self.sheet.set_row_heights(
                row_heights=map(self.sheet.valid_row_height, map(int, row_heights)),
            )
        if column_widths:
            self.sheet.set_column_widths(column_widths=map(int, column_widths))
        if not row_heights:
            self.sheet.set_row_heights()
        if not column_widths:
            self.sheet.set_column_widths()
        saved = d.get("saved_info")
        if saved:
            self.saved_info = DotDict({int(k): v for k, v in saved.items()})
            for dct in self.saved_info.values():
                dct["theights"] = {k: self.tree.valid_row_height(int(v)) for k, v in dct["theights"].items()}
                dct["twidths"] = {k: int(v) for k, v in dct["twidths"].items()}
        else:
            self.saved_info = new_saved_info(self.hiers)
        for c, align in (d.get("sheet_column_alignments") or {}).items():
            self.sheet.align_columns(int(c), align=align, redraw=False)
            self.tree.align_columns(int(c), align=align, redraw=False)
        tv = d.get("tv_label_col")
        self.tv_label_col = int(tv) if tv is not None else int(self.ic)
        self.set_headers()
        self.tag_ids(selection=set(self.tagged_ids), toggle=False, do_tree=False)

    def populate_view(self, *, from_program_data=False):
        self.saved_sheet_row_heights = {}
        self.reset_tree_search_dropdown()
        self.reset_sheet_search_dropdown()
        self.selected_ID = ""
        self.selected_PAR = ""
        self.new_sheet = []
        self.reset_tree_drag_vars()
        self.search_results = []
        self.sheet_search_results = []
        self.sort_later_dct = None
        self.reset_tagged_ids_dropdowns()
        self.C.file.entryconfig("Compare sheets", command=self.compare_from_within_treeframe)
        self.C.file.entryconfig("New", command=self.create_new_from_within_treeframe)
        self.C.file.entryconfig("Open", command=self.open_from_within_treeframe)
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.edit_menu.entryconfig(0, label="Undo  0/30", state="disabled")
        self.tree.unbind("<z>")
        self.tree.unbind("<Z>")
        self.sheet.unbind("<z>")
        self.sheet.unbind("<Z>")
        self.copied_details = {"copied": [], "id": ""}
        self.copied_detail = {"copied": "", "id": ""}
        self.vs = deque(maxlen=30)
        self.cut = []
        self.copied = []
        self.cut_children_dct = {}
        self.sheet.row_index(self.ic)
        self.tree.set_options(show_horizontal_grid=not self.tree.ops.alternate_color)
        self.refresh_formatting()
        self.redo_tree_display()
        self.disable_paste()
        self.refresh_dropdowns()
        if from_program_data:
            self.move_sheet_pos()
            self.move_tree_pos()
        else:
            self.sheet.set_xview(0.0).set_yview(0.0)
            self.tree.set_xview(0.0).set_yview(0.0)
        self.C.show_frame("tree_edit", start=False, msg=self.get_tree_editor_status_bar_text())
        self.C.file.entryconfig("Settings", state="normal")
        self.WINDOW_DIMENSIONS_CHANGED()
        self.focus_tree()

    def settings(self):
        Settings_Popup(self, theme=self.C.theme)

    def change_theme(self, theme="light_green", write=True):
        self.C.theme = theme
        self.dark_blue_theme_bool = self.C.theme == "dark_blue"
        self.black_theme_bool = self.C.theme == "black"
        self.dark_theme_bool = self.C.theme == "dark"
        self.light_green_theme_bool = self.C.theme == "light_green"
        self.light_blue_theme_bool = self.C.theme == "light_blue"
        self.config(bg=themes[theme].top_left_bg)
        self.C.config(bg=themes[theme].top_left_bg)
        self.main_canvas.config(bg=themes[theme].top_left_bg)
        self.l_frame.config(bg=themes[theme].top_left_bg)
        self.treeframe.config(bg=themes[theme].top_left_bg)
        self.sts_tree.config(bg=themes[theme].top_left_bg)
        self.r_frame.config(bg=themes[theme].top_left_bg)
        self.sheetframe.config(bg=themes[theme].top_left_bg)
        # if USER_OS == "darwin":
        # button_kwargs = {
        #     "background": themes["light_green"].table_bg,
        #     "darkcolor": themes["light_green"].table_bg,
        #     "bordercolor": themes["light_green"].table_grid_fg,
        #     "lightcolor": themes["light_green"].table_bg,
        #     "highlightcolor": themes["light_green"].table_bg,
        #     "foreground": themes["light_green"].table_fg,
        #     "borderwidth": 1 if theme.startswith("light") else 0,
        # }
        # self.C.style.configure("Std.TButton", **button_kwargs)
        # self.C.style.configure("EF.Std.TButton", **button_kwargs)
        # self.C.style.configure("TF.Std.TButton", **button_kwargs)
        # self.C.style.configure("STSF.Std.TButton", **button_kwargs)
        # self.C.style.configure("EFB.Std.TButton", **button_kwargs)
        # self.C.style.configure("ERR_ASK_FNT.Std.TButton", **button_kwargs)
        # self.C.style.configure("x_button.Std.TButton", **button_kwargs)
        # for s in (
        #     "Std.TButton",
        #     "EF.Std.TButton",
        #     "TF.Std.TButton",
        #     "STSF.Std.TButton",
        #     "EFB.Std.TButton",
        #     "ERR_ASK_FNT.Std.TButton",
        #     "x_button.Std.TButton",
        # ):
        #     self.C.style.map(
        #         s,
        #         foreground=[
        #             ("!active", themes["light_green"].table_fg),
        #             ("pressed", themes["light_green"].table_fg),
        #             ("active", themes["light_green"].table_fg),
        #         ],
        #         background=[
        #             ("!active", themes["light_green"].table_bg),
        #             ("pressed", themes["light_green"].table_grid_fg),
        #             ("active", "#91c9f7"),
        #         ],
        #     )

        self.search_entry.config(
            background=themes["light_green"].table_bg,
            foreground=themes["light_green"].table_fg,
            disabledbackground=themes["light_green"].table_bg,
            disabledforeground=themes["light_green"].table_fg,
            insertbackground=themes["light_green"].table_fg,
            readonlybackground=themes["light_green"].table_bg,
        )
        self.sheet_search_entry.config(
            background=themes["light_green"].table_bg,
            foreground=themes["light_green"].table_fg,
            disabledbackground=themes["light_green"].table_bg,
            disabledforeground=themes["light_green"].table_fg,
            insertbackground=themes["light_green"].table_fg,
            readonlybackground=themes["light_green"].table_bg,
        )
        self.C.selection_info.change_theme(theme)

        self.btns_tree.config(bg=themes[theme].top_left_bg)
        self.btns_sheet.config(bg=themes[theme].top_left_bg)

        self.C.status_bar.config(bg=themes[theme].top_left_bg, fg=themes[theme].table_selected_box_cells_fg)
        self.C.status_frame.config(bg=themes[theme].top_left_bg)

        self.C.frames["column_selection"].sheet_selector.config(bg=themes[theme].top_left_bg)
        self.C.frames["column_selection"].sheet_selector.sheets_label.config(
            bg=themes[theme].top_left_bg,
            fg=themes[theme].table_fg,
        )
        self.C.frames["column_selection"].config(bg=themes[theme].top_left_bg)
        self.C.frames["column_selection"].data_format_selector.change_theme(theme)
        self.C.frames["column_selection"].selector.change_theme(theme)
        self.C.frames["column_selection"].flattened_selector.change_theme(theme)
        self.C.frames["tree_compare"].sheet_filename1.change_theme(theme)
        self.C.frames["tree_compare"].sheet_filename2.change_theme(theme)
        self.C.frames["tree_compare"].l_frame.config(
            highlightbackground=themes[theme].table_fg, background=themes[theme].top_left_bg
        )
        self.C.frames["tree_compare"].l_frame_btns.config(background=themes[theme].top_left_bg)
        self.C.frames["tree_compare"].r_frame.config(
            highlightbackground=themes[theme].table_fg, background=themes[theme].top_left_bg
        )
        self.C.frames["tree_compare"].r_frame_btns.config(background=themes[theme].top_left_bg)
        self.C.frames["tree_compare"].selector_1.change_theme(theme)
        self.C.frames["tree_compare"].selector_2.change_theme(theme)
        self.C.frames["tree_compare"].sheetdisplay1.change_theme(theme)
        self.C.frames["tree_compare"].sheetdisplay2.change_theme(theme)
        self.C.frames["tree_compare"].file_label1.change_theme(theme)
        self.C.frames["tree_compare"].file_label2.change_theme(theme)
        self.sheet.change_theme(theme)
        self.tree.change_theme(theme)
        self.C.frames["column_selection"].sheetdisplay.change_theme(theme)
        if write:
            self.C.save_cfg()
        self.focus_tree()

    def destroy_find_popup(self, event=None):
        with suppress(Exception):
            self.find_popup.destroy()
        self.find_popup = None

    def reset_tree(self, extra=True):
        self.destroy_find_popup()
        if extra:
            self.C.file.entryconfig("Compare sheets", command=self.C.compare_at_start)
            self.C.file.entryconfig("Open", command=self.C.open_file_at_start)
            self.C.file.entryconfig("New", command=self.C.create_new_at_start)
            self.C.menubar_state("disabled")
            self.bind_or_unbind_save("disabled")
        self.C.unsaved_changes = False
        self.tv_label_col = 0
        self.selected_ID = ""
        self.selected_PAR = ""
        self.rc_iid = None
        self.disable_paste()
        self.search_results = []
        self.sheet_search_results = []
        self.tree.reset()
        self.session.clear()
        self.set_records([], redraw=True)
        self.sheet.deselect("all", redraw=False)
        self.sheet.reset_all_options()
        self.set_headers()
        self.new_sheet = []
        self.sheet.row_index(newindex=self.ic)
        self.C.created_new = False
        self.C.change_app_title(title=None)

    def bind_or_unbind_save(self, save_menu_state: Literal["normal", "save as", "disabled"] | None = None):
        if isinstance(save_menu_state, str):
            self.C.save_menu_state = save_menu_state
        self.C.unbind_class("all", f"<{ctrl_button}-s>")
        self.C.unbind_class("all", f"<{ctrl_button}-S>")
        self.C.unbind_class("all", f"<{ctrl_button}-Shift-S>")
        self.C.unbind_class("all", f"<{ctrl_button}-Shift-s>")
        self.C.file.entryconfig("Save", state="disabled")
        self.C.file.entryconfig("Save as", state="disabled")
        self.C.file.entryconfig("Save new version", state="disabled")
        self.C.file.entryconfig("Settings", state="disabled")
        if self.C.save_menu_state == "normal":
            self.C.file.entryconfig("Save", state="normal")
            self.C.file.entryconfig("Save as", state="normal")
            self.C.file.entryconfig("Save new version", state="normal")
            self.C.bind_class("all", f"<{ctrl_button}-s>", self.save_)
            self.C.bind_class("all", f"<{ctrl_button}-S>", self.save_)
            self.C.bind_class("all", f"<{ctrl_button}-Shift-s>", self.save_as)
            self.C.bind_class("all", f"<{ctrl_button}-Shift-S>", self.save_as)
            self.C.file.entryconfig("Settings", state="normal")
        elif self.C.save_menu_state == "save as":
            self.C.file.entryconfig("Save as", state="normal")
            self.C.bind_class("all", f"<{ctrl_button}-s>", self.save_as)
            self.C.bind_class("all", f"<{ctrl_button}-S>", self.save_as)
            self.C.bind_class("all", f"<{ctrl_button}-Shift-s>", self.save_as)
            self.C.bind_class("all", f"<{ctrl_button}-Shift-S>", self.save_as)
            self.C.file.entryconfig("Settings", state="normal")

    def enable_widgets(self, widgets=True, menubar=True):
        self.C.menubar_state("normal")
        for widget in (self.tree, self.sheet):
            widget.bind(f"<{ctrl_button}-e>", self.expand_id)
            widget.bind(f"<{ctrl_button}-E>", self.expand_id)
            widget.bind(f"<{ctrl_button}-r>", self.collapse_id)
            widget.bind(f"<{ctrl_button}-R>", self.collapse_id)
            widget.bind(f"<{ctrl_button}-z>", self.undo)
            widget.bind(f"<{ctrl_button}-Z>", self.undo)
            widget.bind(f"<{ctrl_button}-l>", self.show_changelog)
            widget.bind(f"<{ctrl_button}-L>", self.show_changelog)
            widget.bind(f"<{ctrl_button}-v>", self.paste_key)
            widget.bind(f"<{ctrl_button}-V>", self.paste_key)
            widget.bind(f"<{ctrl_button}-x>", self.cut_key)
            widget.bind(f"<{ctrl_button}-X>", self.cut_key)
            widget.bind(f"<{ctrl_button}-c>", self.copy_key)
            widget.bind(f"<{ctrl_button}-C>", self.copy_key)
            widget.bind(f"<{ctrl_button}-t>", self.tag_ids)
            widget.bind(f"<{ctrl_button}-T>", self.tag_ids)
            widget.bind("<Delete>", self.del_key)
            widget.bind("<Double-Button-1>", self.tree_sheet_double_left)
            widget.extra_bindings(
                [
                    ("begin_column_header_drag_drop", self.snapshot_begin_drag_cols),
                    ("column_header_drag_drop", self.snapshot_drag_cols),
                ]
            )
        self.sheet.extra_bindings(
            [
                ("begin_row_index_drag_drop", self.snapshot_begin_drag_rows),
                ("row_index_drag_drop", self.snapshot_drag_rows),
            ]
        )
        self.tree.extra_bindings(
            [
                ("begin_row_index_drag_drop", self.begin_tree_drag_drop_ids),
                ("row_index_drag_drop", self.tree_drag_drop_ids),
            ]
        )
        self.sheet.enable_bindings(sheet_bindings).basic_bindings(True)
        self.tree.enable_bindings(tree_bindings).basic_bindings(True)
        self.sheet.bind(rc_release, self.sheet_rc_release)
        self.sheet.bind("<<SheetSelect>>", self.sheet_select_event)
        self.tree.bind("<<SheetSelect>>", self.tree_select_event)
        self.tree.bind(rc_press, self.tree_rc_press)
        self.tree.bind(ctrl_rc_press, lambda e: self.tree_rc_press(e, True))
        # self.tree.bind(rc_motion, self.tree_rc_motion)
        # self.tree.bind(rc_release, self.tree_rc_release)
        self.tree.bind("<FocusIn>", self.tree_focus_enter).bind("<FocusOut>", self.tree_focus_leave)
        self.sheet.bind("<FocusIn>", self.sheet_focus_enter).bind("<FocusOut>", self.sheet_focus_leave)
        self.sheet.bulk_table_edit_validation(self.tree_sheet_edit_table)
        self.tree.bulk_table_edit_validation(self.tree_sheet_edit_table)
        self.sheet_tag_id_button.config(state="normal")
        self.sheet_tagged_ids_dropdown.config(state="readonly")
        self.sheet_tagged_ids_dropdown.bind("<<ComboboxSelected>>", self.sheet_go_to_tagged_id)
        self.tree_tag_id_button.config(state="normal")
        self.tree_tagged_ids_dropdown.config(state="readonly")
        self.tree_tagged_ids_dropdown.bind("<<ComboboxSelected>>", self.tree_go_to_tagged_id)
        self.switch_label.config(state="normal")
        self.search_button.config(state="normal")
        self.search_choice_dropdown.config(state="readonly")
        self.search_choice_dropdown.bind("<<ComboboxSelected>>", lambda focus: self.search_entry.focus_set())
        self.search_entry.config(state="normal")
        self.search_entry.bind("<Return>", self.search_choice)
        self.search_dropdown.config(state="readonly")
        self.search_dropdown.bind("<<ComboboxSelected>>", self.show_search_result)
        self.switch_hier_dropdown.config(state="readonly")
        self.switch_hier_dropdown.bind("<<ComboboxSelected>>", self.switch_hier)
        self.sheet_search_button.config(state="normal")
        self.sheet_search_choice_dropdown.config(state="readonly")
        self.sheet_search_entry.enable_me()
        self.sheet_search_entry.bind("<Return>", self.sheet_search_choice)
        self.sheet_search_dropdown.config(state="readonly")
        self.sheet_search_dropdown.bind("<<ComboboxSelected>>", self.sheet_show_search_result)
        self.sheet_search_choice_dropdown.bind(
            "<<ComboboxSelected>>",
            lambda focus: self.sheet_search_entry.focus_set(),
        )
        self.bind_or_unbind_save("save as" if self.C.created_new else "normal")

    def disable_widgets(self):
        self.C.menubar_state("disabled")
        self.bind_or_unbind_save("disabled")
        for x in (self.tree, self.sheet):
            x.unbind(f"<{ctrl_button}-e>")
            x.unbind(f"<{ctrl_button}-E>")
            x.unbind(f"<{ctrl_button}-r>")
            x.unbind(f"<{ctrl_button}-R>")
            x.unbind(f"<{ctrl_button}-z>")
            x.unbind(f"<{ctrl_button}-Z>")
            x.unbind(f"<{ctrl_button}-l>")
            x.unbind(f"<{ctrl_button}-L>")
            x.unbind(f"<{ctrl_button}-t>")
            x.unbind(f"<{ctrl_button}-T>")
            x.unbind(f"<{ctrl_button}-c>")
            x.unbind(f"<{ctrl_button}-C>")
            x.unbind(f"<{ctrl_button}-v>")
            x.unbind(f"<{ctrl_button}-V>")
            x.unbind(f"<{ctrl_button}-x>")
            x.unbind(f"<{ctrl_button}-X>")
            x.unbind("<Delete>")
            # x.disable_bindings().basic_bindings(False)
            x.unbind("<Double-Button-1>")
            x.unbind("<FocusIn>")
            x.unbind("<FocusOut>")
        self.C.unbind(f"<{ctrl_button}-s>")
        self.C.unbind(f"<{ctrl_button}-S>")
        self.sheet.unbind(rc_button)
        self.sheet.extra_bindings(
            [
                ("row_index_drag_drop", None),
                ("all_select_events", None),
                ("column_header_drag_drop", None),
            ]
        )
        self.tree.unbind(rc_press)
        self.tree.unbind(ctrl_rc_press)
        self.tree.unbind(rc_motion)
        self.tree.unbind(rc_release)
        self.sheet_tag_id_button.config(state="disabled")
        self.sheet_tagged_ids_dropdown.config(state="disabled")
        self.sheet_tagged_ids_dropdown.unbind("<<ComboboxSelected>>")
        self.tree_tag_id_button.config(state="disabled")
        self.tree_tagged_ids_dropdown.config(state="disabled")
        self.tree_tagged_ids_dropdown.unbind("<<ComboboxSelected>>")
        self.switch_label.config(state="disabled")
        self.search_button.config(state="disabled")
        self.search_choice_dropdown.config(state="disabled")
        self.search_choice_dropdown.unbind("<<ComboboxSelected>>")
        self.search_entry.config(state="disabled")
        self.search_entry.unbind("<Return>")
        self.search_dropdown.config(state="disabled")
        self.search_dropdown.unbind("<<ComboboxSelected>>")
        self.switch_hier_dropdown.config(state="disabled")
        self.switch_hier_dropdown.bind("<<ComboboxSelected>>")
        self.sheet_search_button.config(state="disabled")
        self.sheet_search_choice_dropdown.config(state="disabled")
        self.sheet_search_entry.disable_me()
        self.sheet_search_entry.unbind("<Return>")
        self.sheet_search_dropdown.config(state="disabled")
        self.sheet_search_dropdown.unbind("<<ComboboxSelected>>")
        self.sheet_search_choice_dropdown.unbind("<<ComboboxSelected>>")

    def toggle_sort_all_nodes(self, enabled, snapshot=True):
        if enabled and snapshot:
            self.snapshot_auto_sort_nodes()
        self.session.set_option("auto-sort", enabled)
        if enabled:
            self.redo_tree_display()

    def sort_all_children(self):
        self.session.sort_all_children()

    def copy_ID_row(self, event=None):
        selections = self.tree.selection(cells=True)
        if not selections:
            return
        s, writer = str_io_csv_writer(dialect=csv.excel_tab)
        writer.writerow(h.name for h in self.headers)
        writer.writerows(self.sheet.data[self.rns[iid.lower()]] for iid in selections)
        to_clipboard(self.C, s.getvalue().rstrip())

    def copy_ID_children_rows(self, event=None):
        iids = set(self.tree.selection(cells=True))
        if not iids:
            return
        h = int(self.pc)
        tc = set()
        for iid in iids:
            if self.nodes[iid.lower()].ps[h]:
                if all(pk not in iids and pk not in tc for pk in self.check_ps(self.nodes[iid.lower()].ps[h], h)):
                    tc.add(iid)
            elif iid not in tc:
                tc.add(iid)
        s, writer = str_io_csv_writer(dialect=csv.excel_tab)
        writer.writerow(h.name for h in self.headers)
        for iid in sorted(tc, key=lambda x: self.rns[x]):
            iid_lower = iid.lower()
            stack = [iid_lower]
            while stack:
                ik = stack.pop()
                writer.writerow(self.sheet.data[self.rns[ik]])
                children = self.nodes[ik].cn[self.pc]
                if children:
                    stack.extend(reversed(children))
        to_clipboard(self.C, s.getvalue().rstrip())

    def clipboard_sheet(self, event=None):
        s, writer = str_io_csv_writer(dialect=csv.excel)
        writer.writerow(h.name for h in self.headers)
        writer.writerows(self.sheet.data)
        to_clipboard(self.C, s.getvalue().rstrip())

    def clipboard_sheet_indent(self, event=None):
        s, writer = str_io_csv_writer(dialect=csv.excel_tab)
        writer.writerow(h.name for h in self.headers)
        writer.writerows(self.sheet.data)
        to_clipboard(self.C, s.getvalue().rstrip())

    def clipboard_sheet_json(self, event=None):
        to_clipboard(
            self.C,
            json.dumps(
                full_sheet_to_dict(
                    [h.name for h in self.headers],
                    self.sheet.MT.data,
                    include_headers=True,
                    format_=self.json_format,
                )
            ),
        )

    def changelog_singular(self, text):
        self.session.changelog_singular(text)

    def changelog_append(self, change, id_, old, new):
        self.session.changelog_append(change, id_, old, new)

    def changelog_append_no_unsaved(self, change, id_, old, new):
        self.session.changelog_append_no_unsaved(change, id_, old, new)

    def edit_cell_rebuild(self, r, c, value) -> object:
        self.snapshot_ctrl_x_v_del_key_id_par()
        self.changelog_append(
            "Edit cell",
            f"ID: {self.sheet.MT.data[r][self.ic]} column #{c + 1} named: {self.headers[c].name} with type: {self.headers[c].type_}",
            f"{self.sheet.MT.data[r][c]}",
            value,
        )
        self.sheet.MT.data[r][c] = value
        self.rebuild_tree()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        return value

    def edit_cell_single(self, r: int, c: int, value: object) -> None:
        self.session.set_detail(self.sheet.MT.data[r][self.ic], c, value)
        return value

    def edit_cell_multiple(self, r: int, c: int, value: object) -> None:
        self.changelog_append_no_unsaved(
            "Edit cell |",
            f"ID: {self.sheet.MT.data[r][self.ic]} column #{c + 1} named: {self.headers[c].name} with type: {self.headers[c].type_}",
            f"{self.sheet.MT.data[r][c]}",
            value,
        )
        self.sheet.MT.data[r][c] = value
        return value

    def tree_sheet_edit_table(self, event=None):
        if not event:
            return
        if len(event.data) == 1:
            y1, x1 = next(iter(event.data))
            newtext = event.data[(y1, x1)]
            if event.sheetname == "tree":
                y1 = self.rns[self.tree.rowitem(y1, data_index=True)]

            if self.headers[x1].type_ in ("ID", "Parent") and not self.allow_spaces_ids_var:
                newtext = re.sub(r"[\n\t\s]*", "", newtext)

            if newtext != self.sheet.data[y1][x1]:
                ID = self.sheet.data[y1][self.ic]
                ik = ID.lower()

                if self.headers[x1].type_ == "ID":
                    id_ = ID
                    ik = id_.lower()
                    tree_sel = self.tree.selection()
                    if not self.change_ID_name(id_, newtext, errors=False):
                        self.edit_cell_rebuild(y1, x1, newtext)
                        event.data = {}
                        return

                    self.reset_tagged_ids_dropdowns()
                    self.disable_paste()
                    self.rns = {r[self.ic].lower(): i for i, r in enumerate(self.sheet.data)}
                    self.refresh_formatting(rows=self.refresh_rows)
                    self.redo_tree_display()
                    self.refresh_rows = set()
                    if tree_sel:
                        if self.tree.exists(tree_sel[0]):
                            self.tree.scroll_to_item(tree_sel[0])
                            self.tree.selection_set(tree_sel[0])
                        else:
                            self.tree.scroll_to_item(newtext.lower())
                            self.tree.selection_set(newtext.lower())
                    else:
                        self.move_tree_pos()
                    self.sheet.set_cell_size_to_text(y1, x1, only_set_if_too_small=True)
                    self.tree_set_cell_size_to_text(y1, x1)
                    self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

                elif self.headers[x1].type_ == "Parent":
                    self.snapshot_paste_id()
                    oldparent = f"{self.sheet.data[y1][x1]}"
                    tree_sel = self.tree.selection()
                    if self.cut_paste_edit_cell(ID, oldparent, x1, newtext):
                        self.changelog_append(
                            "Cut and paste ID + children" if self.nodes[ik].cn[x1] else "Cut and paste ID",
                            ID,
                            f"Old parent: {oldparent if oldparent else 'n/a - Top ID'} old column #{x1 + 1} named: {self.headers[x1].name}",
                            f"New parent: {newtext if newtext else 'n/a - Top ID'} new column #{x1 + 1} named: {self.headers[x1].name}",
                        )
                        self.refresh_formatting(rows=y1, columns=x1)
                        self.redo_tree_display()
                        self.sheet.set_cell_size_to_text(y1, x1, only_set_if_too_small=True)
                        self.tree_set_cell_size_to_text(y1, x1)
                        if tree_sel:
                            self.tree.scroll_to_item(tree_sel[0])
                            self.tree.selection_set(tree_sel)
                        self.disable_paste()
                        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
                    else:
                        self.vs.pop()
                        self.set_undo_label()
                        self.edit_cell_rebuild(y1, x1, newtext)
                else:
                    if self.detail_is_valid_for_col(x1, newtext):
                        self.snapshot_ctrl_x_v_del_key()
                        self.vs[-1]["cells"][(y1, x1)] = f"{self.sheet.MT.data[y1][x1]}"
                        newtext = self.edit_cell_single(y1, x1, newtext)
                        self.refresh_formatting(rows=y1, columns=x1)
                        self.refresh_tree_item(ID)
                        self.sheet.set_cell_size_to_text(y1, x1, only_set_if_too_small=True)
                        self.tree_set_cell_size_to_text(y1, x1)
                        self.disable_paste()
                        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
                    else:
                        Error(
                            self,
                            "Entered text is not in column validation   ",
                            theme=self.C.theme,
                        )
            loc = event.get("loc")
            key = event.get("key")
            if loc and key in ("Return", "Tab"):
                if event.sheetname == "tree":
                    self.tree.next_cell(*loc, key)
                elif event.sheetname == "sheet":
                    self.sheet.next_cell(*loc, key)

        else:
            self.start_work("Editing table...")
            idcols = set(self.hiers) | {self.ic}
            need_rebuild = any(k[1] in idcols for k in event["data"])
            refresh_rows = set()
            refresh_cols = set()
            edit_ctr = 0
            tree = event.sheetname == "tree"
            if need_rebuild:
                self.snapshot_ctrl_x_v_del_key_id_par()
            else:
                self.snapshot_ctrl_x_v_del_key()
            for (r, c), value in event["data"].items():
                if tree:
                    r = self.rns[self.tree.rowitem(row=r, data_index=True)]
                if (need_rebuild and c in idcols) or (
                    self.detail_is_valid_for_col(c, value) and self.sheet.MT.data[r][c] != value
                ):
                    if not need_rebuild:
                        self.vs[-1]["cells"][(r, c)] = f"{self.sheet.MT.data[r][c]}"
                    self.edit_cell_multiple(r, c, value)
                    refresh_rows.add(r)
                    refresh_cols.add(c)
                    edit_ctr += 1
            self.disable_paste()
            if edit_ctr:
                if need_rebuild:
                    self.rebuild_tree()
                else:
                    self.refresh_formatting(rows=refresh_rows, columns=refresh_cols)
                    for rn in refresh_rows:
                        self.refresh_tree_item(self.sheet.MT.data[rn][self.ic])
                if edit_ctr > 1:
                    self.changelog_append(
                        f"Edit {edit_ctr} cells",
                        "",
                        "",
                        "",
                    )
                else:
                    self.changelog_singular("Edit cell")
                self.redraw_sheets()
                self.stop_work(self.get_tree_editor_status_bar_text())
            else:
                self.vs.pop()
                self.set_undo_label()
                self.redraw_sheets()
                self.stop_work(self.get_tree_editor_status_bar_text())
        event.data = {}

    def tree_set_cell_size_to_text(self, sheet_r, sheet_c):
        if self.tree.exists(self.sheet.data[sheet_r][self.ic].lower()) and self.tree.item_displayed(
            self.sheet.data[sheet_r][self.ic].lower()
        ):
            self.tree.set_cell_size_to_text(
                self.tree.displayed_rows.index(self.tree.itemrow(self.sheet.data[sheet_r][self.ic].lower())),
                sheet_c,
                only_set_if_too_small=True,
            )

    def tag_ids_using_list(self, event=None) -> None:
        Tag_Ids_Using_List_Popup(self, theme=self.C.theme)

    def delete_ids_using_list(self, event=None) -> None:
        Delete_Ids_Using_List_Popup(self, theme=self.C.theme)

    def replace_using_mapping(self, event=None) -> None:
        Replace_Popup(self, theme=self.C.theme)

    def copy_key(self, event: object = None) -> None:
        if self.tree.has_focus():
            iids = tuple(
                self.tree.rowitem(row)
                for box in self.tree.boxes
                for row in range(box.coords.from_r, box.coords.upto_r)
                if box[1] == "rows"
            )
            if iids:
                self.copy_ID(iids=iids)
            else:
                self.tree.copy()
        elif self.sheet.has_focus():
            self.sheet.copy()

    def del_key(self, event: object = None) -> None:
        if self.tree.has_focus():
            iids = [
                self.tree.rowitem(rn)
                for rn in sorted(
                    {
                        row
                        for box in self.tree.boxes
                        for row in range(box.coords.from_r, box.coords.upto_r)
                        if box[1] == "rows"
                    }
                )
            ]
            if iids:
                self.del_id(iids=iids)
            elif self.tree.boxes:
                self.tree.delete()
        elif self.sheet.has_focus():
            if rows := self.sheet.get_selected_rows():
                self.del_id_all(iids=[self.sheet.MT.data[r][self.ic].lower() for r in rows])
            elif self.sheet.boxes:
                self.sheet.delete()

    def rebuild_tree(self, deselect=True, redraw=False):
        if deselect:
            self.sheet.deselect("all", redraw=False)
        self.clear_copied_details()
        self.save_info_get_saved_info()
        self.session.rebuild_identity()
        self._sync_sheet_from_session()
        self.new_sheet = []
        rhs = []
        default_row_height = self.sheet.MT.get_default_row_height()
        for _i, r in enumerate(self.sheet.data):
            ik = r[self.ic].lower()
            if ik in self.saved_sheet_row_heights:
                rhs.append(self.saved_sheet_row_heights[ik])
            else:
                rhs.append(default_row_height)
        self.sheet.set_row_heights(rhs)
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.refresh_rows = set()
        self.refresh_formatting()
        if redraw:
            self.redraw_sheets()
        self.redo_tree_display()
        self.sheet.recreate_all_selection_boxes()

    def cut_key(self, event: object = None) -> None:
        if self.tree.has_focus():
            if iids := tuple(
                self.tree.rowitem(row)
                for box in self.tree.boxes
                for row in range(box.coords.from_r, box.coords.upto_r)
                if box[1] == "rows"
            ):
                self.cut_ids(iids=iids)
            elif self.tree.ctrl_boxes:
                self.tree.cut()
        elif self.sheet.has_focus() and self.sheet.ctrl_boxes:
            self.sheet.cut()

    def paste_key(self, event: object = None) -> None:
        if self.tree.has_focus():
            if event:
                if self.tree.selected:
                    rows = sorted(self.tree.get_selected_rows())
                    if rows and len(rows) == 1 and self.winfo_containing(event.x_root, event.y_root) is not None:
                        self.tree_rc_menu_single_row_paste.tk_popup(self.C.winfo_pointerx(), self.C.winfo_pointery())
                    else:
                        self.tree.paste()
                elif self.winfo_containing(event.x_root, event.y_root) is not None:
                    self.tree_rc_menu_empty.tk_popup(self.C.winfo_pointerx(), self.C.winfo_pointery())
            elif self.tree.selected:
                self.tree.paste()
        elif self.sheet.has_focus() and self.sheet.selected:
            self.sheet.paste()

    def select_id_in_treeview_from_sheet(self, event=None):
        ik = self.sheet.MT.data[self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)[0]][
            self.ic
        ].lower()
        self.go_to_treeview_id_finder(ik)

    def start_work(self, msg="", outside_treeframe=False):
        self.C.working = True
        self.C.save_menu_state = "disabled"
        if not outside_treeframe:
            self.disable_widgets()
        self.C.status_bar.change_text(msg)

    def stop_work(self, msg="", outside_treeframe=False, resume_quit=True):
        self.C.working = False
        self.C.save_menu_state = "normal"
        if resume_quit and self.C.USER_HAS_QUIT:
            self.C.USER_HAS_CLOSED_WINDOW()  # user still wants to quit
            if self.C.USER_HAS_QUIT:
                return
        if not outside_treeframe:
            self.after_idle(self.enable_widgets)
        self.C.status_bar.change_text(msg)

    def hide_frames(self, l_frame=False, r_frame=False, set_dimensions=True):
        if l_frame:
            self.main_canvas.itemconfig(
                self.l_frame_id,
                state="hidden",
            )
        if r_frame:
            self.main_canvas.itemconfig(
                self.r_frame_id,
                state="hidden",
            )
        if set_dimensions:
            self.WINDOW_DIMENSIONS_CHANGED()

    def unhide_frames(self, l_frame=False, r_frame=False, set_dimensions=True):
        if l_frame:
            self.main_canvas.itemconfig(
                self.l_frame_id,
                state="normal",
            )
        if r_frame:
            self.main_canvas.itemconfig(
                self.r_frame_id,
                state="normal",
            )
        if set_dimensions:
            self.WINDOW_DIMENSIONS_CHANGED()

    def ask_continue_unsaved(self):
        if self.C.unsaved_changes:
            confirm = Ask_Confirm(
                self,
                "Discard unsaved changes?",
                theme=self.C.theme,
                button_text="Continue",
            )
            return confirm.boolean
        else:
            return True

    def compare_from_within_treeframe(self):
        if not self.ask_continue_unsaved():
            return
        self.reset_tree()
        self.bind_or_unbind_save("disabled")
        self.C.frames["tree_compare"].populate()

    def open_from_within_treeframe(self, event=None):
        if not self.ask_continue_unsaved():
            return
        fp = filedialog.askopenfilename(parent=self.C, title="Select file")
        if not fp:
            return
        try:
            fp = os.path.normpath(fp)
        except Exception:
            Error(self, "Filepath invalid   ", theme=self.C.theme)
            return
        if not fp.lower().endswith((".json", ".xlsx", ".xls", ".xlsm", ".csv", ".tsv")):
            Error(self, "Please select json/excel/csv   ", theme=self.C.theme)
            return
        self.disable_widgets()
        if os.path.isfile(fp):
            self.C.open_dict["filepath"] = fp
            self.reset_tree()
            self.C.load_from_file()
        else:
            Error(self, "Filepath invalid   ", theme=self.C.theme)
            self.enable_widgets()

    def create_new_from_within_treeframe(self, event=None):
        if not self.ask_continue_unsaved():
            return
        self.reset_tree(False)
        self.session.new(discard=True)
        self.set_records(self.session.data)
        self.tv_label_col = 0
        self.C.created_new = True
        self.C.open_dict["filepath"] = "New sheet"
        self.C.change_app_title(title="New sheet")
        self.C.open_dict["sheet"] = "Sheet1"
        self.warnings_filepath = "n/a - CREATED NEW"
        self.warnings_sheet = "n/a"
        self.populate()

    def enter_divider(self, event):
        if not self.currently_adjusting_divider:
            self.main_canvas.config(cursor="sb_h_double_arrow")

    def leave_divider(self, event):
        if not self.currently_adjusting_divider:
            self.main_canvas.config(cursor="")

    def divider_b1_press(self, event):
        self.currently_adjusting_divider = True

    def divider_b1_motion(self, event):
        if self.currently_adjusting_divider:
            self.l_frame_proportion = float(round(event.x / self.winfo_width(), 2))
            if self.l_frame_proportion < 0.01:
                self.l_frame_proportion = 0.01
            elif self.l_frame_proportion > 0.99:
                self.l_frame_proportion = 0.99
            self.WINDOW_DIMENSIONS_CHANGED(place_left_panel=False)

    def divider_b1_release(self, event):
        self.currently_adjusting_divider = False

    def unhide_adjustable_divider(self):
        self.main_canvas.itemconfig("div", state="normal")
        self.main_canvas.tag_bind("div", "<Enter>", self.enter_divider)
        self.main_canvas.tag_bind("div", "<Leave>", self.leave_divider)
        self.main_canvas.tag_bind("div", "<ButtonPress-1>", self.divider_b1_press)
        self.main_canvas.tag_bind("div", "<B1-Motion>", self.divider_b1_motion)
        self.main_canvas.tag_bind("div", "<ButtonRelease-1>", self.divider_b1_release)

    def hide_adjustable_divider(self):
        self.main_canvas.itemconfig("div", state="hidden")
        self.main_canvas.tag_unbind("div", "<Enter>")
        self.main_canvas.tag_unbind("div", "<Leave>")
        self.main_canvas.tag_unbind("div", "<ButtonPress-1>")
        self.main_canvas.tag_unbind("div", "<B1-Motion>")
        self.main_canvas.tag_unbind("div", "<ButtonRelease-1>")

    def get_display_option(self):
        if self.full_left_bool.get():
            return "left"
        if self.adjustable_bool.get():
            return "adjustable"
        if self.full_right_bool.get():
            return "right"
        if self._50_50_bool.get():
            return "50/50"

    def set_display_option(self, option: Literal["left", "adjustable", "right", "50/50"]) -> None:
        if option == "left":
            self.option_full_left(event="config")
        elif option == "right":
            self.option_full_right(event="config")
        elif option == "adjustable":
            self.option_adjustable(event="config")
        elif option == "50/50":
            self.option_50_50(event="config")
        self.WINDOW_DIMENSIONS_CHANGED()

    def option_adjustable(self, event=None):
        if event is None:
            if not (
                self.full_left_bool.get()
                or self.full_right_bool.get()
                or self._50_50_bool.get()
                or self.adjustable_bool.get()
            ):
                self.adjustable_bool.set(True)
                return
            self.unhide_adjustable_divider()
            if self.full_left_bool.get():
                self.unhide_frames(r_frame=True)
                self.full_left_bool.set(False)
                self.focus_tree()
            elif self.full_right_bool.get():
                self.unhide_frames(l_frame=True)
                self.full_right_bool.set(False)
                self.focus_sheet()
            elif self._50_50_bool.get():
                self._50_50_bool.set(False)
            self.WINDOW_DIMENSIONS_CHANGED()
        elif event == "config":
            self.full_left_bool.set(False)
            self.full_right_bool.set(False)
            self._50_50_bool.set(False)
            self.adjustable_bool.set(True)
            self.unhide_adjustable_divider()
            self.unhide_frames(l_frame=True, r_frame=True)
        else:
            if self.adjustable_bool.get():
                self.adjustable_bool.set(False)
            else:
                self.adjustable_bool.set(True)
            self.option_adjustable()

    def option_50_50(self, event=None):
        if event is None:
            if not (
                self.full_left_bool.get()
                or self.full_right_bool.get()
                or self._50_50_bool.get()
                or self.adjustable_bool.get()
            ):
                self._50_50_bool.set(True)
                return
            if self.full_left_bool.get():
                self.unhide_frames(r_frame=True)
                self.full_left_bool.set(False)
                self.focus_tree()
            elif self.full_right_bool.get():
                self.unhide_frames(l_frame=True)
                self.full_right_bool.set(False)
                self.focus_sheet()
            elif self.adjustable_bool.get():
                self.adjustable_bool.set(False)
            self.hide_adjustable_divider()
            self.WINDOW_DIMENSIONS_CHANGED()
        elif event == "config":
            self.full_left_bool.set(False)
            self.full_right_bool.set(False)
            self._50_50_bool.set(True)
            self.adjustable_bool.set(False)
            self.hide_adjustable_divider()
            self.unhide_frames(l_frame=True, r_frame=True)
        else:
            if self._50_50_bool.get():
                self._50_50_bool.set(False)
            else:
                self._50_50_bool.set(True)
            self.option_50_50()

    def option_full_left(self, event=None):
        if event is None:
            if not (
                self.full_left_bool.get()
                or self.full_right_bool.get()
                or self._50_50_bool.get()
                or self.adjustable_bool.get()
            ):
                self.full_left_bool.set(True)
                return
            if self._50_50_bool.get():
                self._50_50_bool.set(False)
            elif self.full_right_bool.get():
                self.unhide_frames(l_frame=True, set_dimensions=False)
                self.full_right_bool.set(False)
            elif self.adjustable_bool.get():
                self.adjustable_bool.set(False)
            self.hide_adjustable_divider()
            self.hide_frames(r_frame=True)
            self.focus_tree()
        elif event == "config":
            self.full_left_bool.set(True)
            self.full_right_bool.set(False)
            self._50_50_bool.set(False)
            self.adjustable_bool.set(False)
            self.hide_adjustable_divider()
            self.hide_frames(r_frame=True, set_dimensions=False)
            self.unhide_frames(l_frame=True)
            self.focus_tree()
        else:
            if self.full_left_bool.get():
                self.full_left_bool.set(False)
            else:
                self.full_left_bool.set(True)
            self.option_full_left()

    def option_full_right(self, event=None):
        if event is None:
            if not (
                self.full_left_bool.get()
                or self.full_right_bool.get()
                or self._50_50_bool.get()
                or self.adjustable_bool.get()
            ):
                self.full_right_bool.set(True)
                return
            if self._50_50_bool.get():
                self._50_50_bool.set(False)
            elif self.full_left_bool.get():
                self.unhide_frames(r_frame=True, set_dimensions=False)
                self.full_left_bool.set(False)
            elif self.adjustable_bool.get():
                self.adjustable_bool.set(False)
            self.hide_adjustable_divider()
            self.hide_frames(l_frame=True)
            self.focus_sheet()
        elif event == "config":
            self.full_left_bool.set(False)
            self.full_right_bool.set(True)
            self._50_50_bool.set(False)
            self.adjustable_bool.set(False)
            self.hide_adjustable_divider()
            self.hide_frames(l_frame=True, set_dimensions=False)
            self.unhide_frames(r_frame=True)
            self.focus_sheet()
        else:
            if self.full_right_bool.get():
                self.full_right_bool.set(False)
            else:
                self.full_right_bool.set(True)
            self.option_full_right()

    def WINDOW_DIMENSIONS_CHANGED(self, event=None, place_left_panel=True):
        if event is not None:
            if event.height == self.last_height and event.width == self.last_width:
                return
            self.last_width = event.width
            self.last_height = event.height
        if self.C.current_frame == "tree_edit":
            if self.adjustable_bool.get():
                width = self.winfo_width()
                height = self.winfo_height()
                if self.l_frame_proportion == 0.01:
                    l_frame_width = 1
                    l_frame_x = -1
                    r_frame_width = int(width) - 5
                    r_frame_x = 5
                elif self.l_frame_proportion == 0.99:
                    l_frame_width = int(width) - 5
                    l_frame_x = 0
                    r_frame_width = 1
                    r_frame_x = width + 1
                else:
                    l_frame_width = int(width * self.l_frame_proportion)
                    r_frame_width = int(width - l_frame_width) - 5
                    l_frame_x = 0
                    r_frame_x = l_frame_width + 5
                self.main_canvas.itemconfig(self.l_frame_id, width=l_frame_width, height=height)
                self.main_canvas.itemconfig(self.r_frame_id, width=r_frame_width, height=height)
                self.btns_tree.update_idletasks()
                self.btns_sheet.update_idletasks()
                self.main_canvas.update_idletasks()
                self.main_canvas.coords("div", l_frame_x + l_frame_width, 0, l_frame_width + 5, height)
                self.main_canvas.coords(self.l_frame_id, l_frame_x, 0)
                self.main_canvas.coords(self.r_frame_id, r_frame_x, 0)
                self.btns_tree.update_idletasks()
                self.btns_sheet.update_idletasks()
                self.main_canvas.update_idletasks()
            elif self._50_50_bool.get():
                width = floor(self.winfo_width() / 2) - 1
                height = self.winfo_height()
                self.main_canvas.coords(self.l_frame_id, 0, 0)
                self.main_canvas.coords(self.r_frame_id, width + 1, 0)
                self.main_canvas.itemconfig(self.l_frame_id, width=width, height=height)
                self.main_canvas.itemconfig(self.r_frame_id, width=self.winfo_width() - width - 1, height=height)
            elif self.full_left_bool.get():
                self.main_canvas.coords(self.l_frame_id, 0, 0)
                self.main_canvas.itemconfig(
                    self.l_frame_id,
                    width=self.winfo_width(),
                    height=self.winfo_height(),
                )
            elif self.full_right_bool.get():
                self.main_canvas.coords(self.r_frame_id, 0, 0)
                self.main_canvas.itemconfig(
                    self.r_frame_id,
                    width=self.winfo_width(),
                    height=self.winfo_height(),
                )

    def fix_headers(self, headers, row_len, warnings=True):
        return self.session.fix_headers(headers, row_len, warnings=warnings)

    def remove_selections(self, event=None):
        self.sheet.deselect()
        self.tree.deselect()

    def tree_select_event(self, event):
        selected = event.selected
        if selected and self.tree.data:
            iid = self.tree.rowitem(selected.row)
            if isinstance(iid, str):
                self.selected_ID = self.nodes[iid].name
                pariid = self.tree.parent(iid)
                if pariid == "":
                    self.selected_PAR = ""
                else:
                    self.selected_PAR = self.nodes[pariid].name
                if self.mirror_var and not self.mirror_sels_disabler:
                    self.go_to_row()
                self.mirror_sels_disabler = False
        else:
            self.selected_ID = ""
            self.selected_PAR = ""
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.C.selection_info.set_my_value(self.get_tree_selection_info())

    def _create_tree_select_menu(self, parent: tk.Menu) -> tk.Menu:
        menu = tk.Menu(parent, tearoff=0, **menu_kwargs)
        for label, command, icon in (
            ("Siblings", self.tree_select_siblings, "ICON_SELECT_ALL"),
            ("Children", self.tree_select_children, "ICON_SELECT_ALL"),
            ("Descendants", self.tree_select_descendants, "ICON_SELECT_ALL"),
            ("Ancestors", self.tree_select_ancestors, "ICON_SELECT_ALL"),
            ("All at this level", self.tree_select_same_level, "ICON_SELECT_ALL"),
            ("Tagged IDs", self.tree_select_tagged, "tag"),
        ):
            menu.add_command(
                label=label,
                command=command,
                image=self.icons[icon],
                compound="left",
                **menu_kwargs,
            )
        return menu

    def _tree_rc_iid(self) -> str | None:
        iid = self.rc_iid
        if isinstance(iid, str) and self.tree.exists(iid):
            return iid
        return None

    def _tree_select_add(self, iids: Iterator[str] | Sequence[str]) -> None:
        iids = list(iids)
        if not iids:
            return
        self.tree.selection_add(iids)

    def tree_select_siblings(self, event=None):
        iid = self._tree_rc_iid()
        if iid is None:
            return
        self._tree_select_add(s for s in self.tree.get_children(self.tree.parent(iid)) if s != iid)

    def tree_select_children(self, event=None):
        iid = self._tree_rc_iid()
        if iid is None:
            return
        self._tree_select_add(self.tree.get_children(iid))

    def tree_select_descendants(self, event=None):
        iid = self._tree_rc_iid()
        if iid is None:
            return
        self._tree_select_add(self.tree.descendants(iid))

    def tree_select_ancestors(self, event=None):
        iid = self._tree_rc_iid()
        if iid is None:
            return
        ancestors = []
        parent = self.tree.parent(iid)
        while parent:
            ancestors.append(parent)
            parent = self.tree.parent(parent)
        self._tree_select_add(ancestors)

    def tree_select_same_level(self, event=None):
        iid = self._tree_rc_iid()
        if iid is None:
            return
        target = self.get_node_level(self.nodes[iid])
        found = []
        stack = [(top, 1) for top in self.top_iids()]
        while stack:
            current, level = stack.pop()
            if level == target:
                if current != iid:
                    found.append(current)
            elif level < target:
                stack.extend((child, level + 1) for child in self.nodes[current].cn[self.pc])
        self._tree_select_add(found)

    def tree_select_tagged(self, event=None):
        self._tree_select_add(ik for ik in self.tagged_ids if self.tree.exists(ik))

    def sheet_select_event(self, event=None):
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.C.selection_info.set_my_value(self.get_sheet_selection_info())

    def get_tree_selection_info(self):
        count = 0
        _sum = 0.0
        quick_data = self.tree.MT.data
        quick_displayed_rows = self.tree.MT.displayed_rows
        for r, c in self.tree.gen_selected_cells(get_rows=True, get_columns=True):
            count += 1
            try:
                _sum += float(quick_data[quick_displayed_rows[r]][c])
            except Exception:
                continue
        s = f"Count {count} Sum {int(_sum) if _sum.is_integer() else _sum}"
        self.C.selection_info.config(width=len(s))
        return s

    def get_sheet_selection_info(self):
        count = 0
        _sum = 0.0
        quick_data = self.sheet.MT.data
        for r, c in self.sheet.gen_selected_cells(get_rows=True, get_columns=True):
            count += 1
            try:
                _sum += float(quick_data[r][c])
            except Exception:
                continue
        s = f"Count {count} Sum {int(_sum) if _sum.is_integer() else _sum}"
        self.C.selection_info.config(width=len(s))
        return s

    def get_tree_editor_status_bar_text(self):
        if self.tree.selected:
            sels = self.tree.selection()
            box = next(reversed(self.tree.boxes))
            if box.type_ == "rows":
                tree_addition = f"|   Tree {len(sels)} IDs selected   "
            elif box.type_ == "columns":
                tree_addition = f"|   Tree Columns: {_n2a(box.coords.from_c)}:{_n2a(box.coords.upto_c - 1)}   "
            else:
                if box.coords.upto_r - box.coords.from_r == 1 and box.coords.upto_c - box.coords.from_c == 1:
                    tree_addition = f"|   Tree Cells: {_n2a(box.coords.from_c)}{box.coords.from_r + 2}   "
                else:
                    tree_addition = f"|   Tree Cells: {_n2a(box.coords.from_c)}{box.coords.from_r + 2}:{_n2a(box.coords.upto_c - 1)}{box.coords.upto_r + 1}   "
        else:
            tree_addition = ""
        if self.sheet.selected:
            box = next(reversed(self.sheet.boxes))
            if box.type_ == "rows":
                sheet_addition = f"|   Sheet Rows: {box.coords.from_r + 2}:{box.coords.upto_r + 1}   "
            elif box.type_ == "columns":
                sheet_addition = f"|   Sheet Columns: {_n2a(box.coords.from_c)}:{_n2a(box.coords.upto_c - 1)}   "
            else:
                if box.coords.upto_r - box.coords.from_r == 1 and box.coords.upto_c - box.coords.from_c == 1:
                    sheet_addition = f"|   Sheet Cells: {_n2a(box.coords.from_c)}{box.coords.from_r + 2}   "
                else:
                    sheet_addition = f"|   Sheet Cells: {_n2a(box.coords.from_c)}{box.coords.from_r + 2}:{_n2a(box.coords.upto_c - 1)}{box.coords.upto_r + 1}   "
        else:
            sheet_addition = ""
        if self.copied:
            cc_add = (
                f"|   Copied {len(self.copied)} IDs   "
                if len(self.copied) > 3
                else f"|   Copied: {', '.join(self.nodes[dct['id']].name for dct in self.copied)}   "
            )
        elif self.cut:
            cc_add = (
                f"|   Cut {len(self.cut)} IDs   "
                if len(self.cut) > 3
                else f"|   Cut: {', '.join(self.nodes[dct['id']].name for dct in self.cut)}   "
            )
        else:
            cc_add = ""
        if self.changelog:
            end = f"|   Last Edit: {self.changelog[-1].label}   {cc_add}"
        else:
            end = f"|   No Changes Made   {cc_add}"
        return f"{len(self.sheet.MT.data)} IDs   {tree_addition}{sheet_addition}{end}"

    def tree_rc_press(self, event, ctrl=False):
        self.focus_tree()
        row = self.tree.identify_row(event, allow_end=False)
        col = self.tree.identify_column(event, allow_end=False)
        region = self.tree.identify_region(event)
        if region == "header" and isinstance(col, int):
            if not self.tree.column_selected(col):
                if ctrl:
                    self.tree.add_column_selection(col)
                else:
                    self.tree.select_column(col)
        elif region == "table" and isinstance(row, int) and isinstance(col, int):
            if not self.tree.cell_selected(
                row,
                col,
                rows=True,
                columns=True,
            ):
                if ctrl:
                    self.tree.add_cell_selection(row, col)
                else:
                    self.tree.select_cell(row, col)
        elif region == "index" and isinstance(row, int):
            self.rc_iid = self.tree.rowitem(row)
            rows = self.tree.get_selected_rows()
            if row not in rows:
                if ctrl:
                    self.tree.add_row_selection(row)
                else:
                    self.tree.select_row(row)
                rows = self.tree.get_selected_rows()
            if len(rows) == 1:
                iid = self.tree.rowitem(next(iter(rows)))
                self.drag_pc = int(self.pc)
                self.drag_iid = iid
                self.drag_pariid = self.tree.parent(iid)
                self.last_rced = iid
                self.drag_start_index = self.tree.index(iid)
        elif not ctrl:
            self.tree.deselect()
        self.tree_rc_release(event)

    def tree_rc_motion(self, event):
        pass
        # if self.auto_sort_nodes_bool or self.drag_iid is None:
        #     return
        # iid = self.tree.rowitem(self.tree.identify_row(event, allow_end=False))
        # if not iid or iid == self.last_rced:
        #     return
        # selections = self.tree.selection()
        # if not selections or selections[0] != self.drag_iid:
        #     self.tree_drop_iid()
        #     return
        # if not selections or selections[0] != self.drag_iid:
        #     self.tree_drop_iid()
        #     return
        # if self.pc != self.drag_pc or len(selections) > 1:
        #     self.reset_tree_drag_vars()
        #     return
        # if iid:
        #     pariid = self.tree.parent(iid)
        #     if pariid != self.drag_pariid:
        #         return
        #     # try:
        #     move_to_index = self.tree.index(iid)
        #     parik = self.drag_pariid.lower()
        #     if parik:
        #         pop_index = self.nodes[parik].cn[self.pc].index(self.drag_iid)
        #         self.nodes[parik].cn[self.pc].insert(
        #             move_to_index,
        #             self.nodes[parik].cn[self.pc].pop(pop_index),
        #         )
        #     else:
        #         pop_index = self.topnodes_order[self.pc].index(self.drag_iid)
        #         self.topnodes_order[self.pc].insert(
        #             move_to_index,
        #             self.topnodes_order[self.pc].pop(pop_index),
        #         )
        #     self.tree.move(self.drag_iid, self.drag_pariid, move_to_index)
        #     self.tree.selection_set(self.drag_iid)
        #     # except Exception:
        #     #     self.tree_drop_iid()

    def tree_sheet_rc_menu_option_enabler_disabler(self, col: int):
        if col == self.ic or col in self.hiers:
            self.tree_sheet_rc_menu_single_col.entryconfig("Validation", state="disabled")
        else:
            self.tree_sheet_rc_menu_single_col.entryconfig("Validation", state="normal")

    def tree_rc_release(self, event):
        if self.drag_iid is not None:
            self.drag_end_index = self.tree.index(self.drag_iid)
        self.drag_iid = None  # del if enabling rc move
        if self.auto_sort_nodes_bool or self.drag_iid is None or self.drag_end_index == self.drag_start_index:
            row = self.tree.identify_row(event, allow_end=False)
            col = self.tree.identify_column(event, allow_end=False)
            self.tree_sheet_rc_menu_option_enabler_disabler(col)
            region = self.tree.identify_region(event)
            if region == "header":
                if isinstance(col, int):
                    self.treecolsel = col
                    if len(self.tree.get_selected_columns()) > 1:
                        self.tree_sheet_rc_menu_multi_col.tk_popup(event.x_root, event.y_root)
                    else:
                        self.tree_sheet_rc_menu_single_col.tk_popup(event.x_root, event.y_root)
                else:
                    self.tree_rc_menu_empty.tk_popup(event.x_root, event.y_root)
            elif region == "index":
                if self.auto_sort_nodes_bool:
                    with suppress(Exception):
                        self.tree_rc_menu_single_row.delete("Sort children")
                    with suppress(Exception):
                        self.tree_rc_menu_multi_row.delete("Sort children")
                else:
                    if self.tree_rc_menu_single_row.entrycget("end", "label") != "Sort children":
                        self.tree_rc_menu_single_row.add_command(
                            label="Sort children",
                            command=self.tree_sort_children,
                            image=self.icons["ICON_SORT_ASC"],
                            compound="left",
                            **menu_kwargs,
                        )
                    if self.tree_rc_menu_multi_row.entrycget("end", "label") != "Sort children":
                        self.tree_rc_menu_multi_row.add_command(
                            label="Sort children",
                            command=self.tree_sort_children,
                            image=self.icons["ICON_SORT_ASC"],
                            compound="left",
                            **menu_kwargs,
                        )
                self.treecolsel = self.ic
                if isinstance(row, int):
                    if len(self.tree.get_selected_rows()) > 1:
                        self.tree_rc_menu_multi_row.tk_popup(event.x_root, event.y_root)
                    else:
                        self.tree_rc_menu_single_row.tk_popup(event.x_root, event.y_root)
                else:
                    self.tree_rc_menu_empty.tk_popup(event.x_root, event.y_root)
            elif region == "table":
                if isinstance(row, int) and isinstance(col, int):
                    self.treecolsel = col
                    if len(self.tree.get_selected_cells()) > 1:
                        self.tree_sheet_rc_menu_multi_cell.tk_popup(event.x_root, event.y_root)
                    else:
                        self.tree_sheet_rc_menu_single_cell.entryconfig(
                            0,
                            label=self.headers[self.treecolsel].name[:25],
                        )
                        self.tree_sheet_rc_menu_single_cell.tk_popup(event.x_root, event.y_root)
                else:
                    self.tree_rc_menu_empty.tk_popup(event.x_root, event.y_root)
        self.reset_tree_drag_vars()

    def tree_sort_children(self, event=None):
        for iid in self.tree.selection():
            self.nodes[iid].cn[self.pc] = self.sort_node_cn(self.nodes[iid].cn[self.pc], self.pc)
        self.save_info_get_saved_info()
        self.redo_tree_display()

    def reset_tree_drag_vars(self):
        self.drag_pc = None
        self.drag_iid = None
        self.drag_pariid = None
        self.drag_start_index = None
        self.drag_end_index = None
        self.last_rced = None

    def tree_drop_iid(self):
        self.reset_tree_drag_vars()
        self.redo_tree_display()

    def sheet_rc_release(self, event):
        self.focus_sheet()
        row = self.sheet.identify_row(event, allow_end=False)
        col = self.sheet.identify_column(event, allow_end=False)
        self.tree_sheet_rc_menu_option_enabler_disabler(col)
        region = self.sheet.identify_region(event)
        if region == "header":
            if isinstance(col, int):
                self.treecolsel = col
                if len(self.sheet.get_selected_columns()) > 1:
                    self.tree_sheet_rc_menu_multi_col.tk_popup(event.x_root, event.y_root)
                else:
                    self.tree_sheet_rc_menu_single_col.tk_popup(event.x_root, event.y_root)
            else:
                self.sheet_rc_menu_empty.tk_popup(event.x_root, event.y_root)
        elif region == "index":
            self.treecolsel = self.ic
            if isinstance(row, int):
                if len(self.sheet.get_selected_rows()) > 1:
                    self.sheet_rc_menu_multi_row.tk_popup(event.x_root, event.y_root)
                else:
                    self.sheet_rc_menu_single_row.tk_popup(event.x_root, event.y_root)
            else:
                self.sheet_rc_menu_empty.tk_popup(event.x_root, event.y_root)
        elif region == "table":
            if isinstance(row, int) and isinstance(col, int):
                self.treecolsel = col
                if len(self.sheet.get_selected_cells()) > 1:
                    self.tree_sheet_rc_menu_multi_cell.tk_popup(event.x_root, event.y_root)
                else:
                    self.tree_sheet_rc_menu_single_cell.entryconfig(
                        0,
                        label=self.headers[self.treecolsel].name[:25],
                    )
                    self.tree_sheet_rc_menu_single_cell.tk_popup(event.x_root, event.y_root)
            else:
                self.sheet_rc_menu_empty.tk_popup(event.x_root, event.y_root)

    def tree_sheet_double_left(self, event):
        if (
            self.tree.event_widget_is_sheet(event)
            and (column := self.tree.identify_column(event, allow_end=False)) is not None
        ) or (
            self.sheet.event_widget_is_sheet(event)
            and (column := self.sheet.identify_column(event, allow_end=False)) is not None
        ):
            self.treecolsel = column

    def switch_hier(self, event=None, hier: int | None = None):
        if isinstance(hier, int):
            self.switch_displayed.set(self.headers[hier].name)
            index = self.hiers.index(hier)
        else:
            index = self.switch_hier_dropdown.current()
            if self.hiers[index] == self.pc:
                self.focus_tree()
                return
        self.save_info_get_saved_info()
        self.pc = int(self.hiers[index])
        self.tree.close_dropdown()
        self.redo_tree_display()
        self.move_tree_pos()
        self.mirror_sels_disabler = True
        self.refresh_tree_dropdowns()
        self.focus_tree()

    def next_hier(self, event=None):
        if self.pc == self.hiers[-1]:
            self.switch_hier(hier=self.hiers[0])
        else:
            self.switch_hier(hier=self.hiers[self.switch_hier_dropdown.current() + 1])

    def check_cn(self, iid: str, h: int) -> Generator[str]:
        yield from self.session.check_cn(iid, h)

    def check_ps(self, iid: str, h: int) -> Generator[str]:
        yield from self.session.check_ps(iid, h)

    def add(self, ID, parent, insert_row=None, snapshot=True, errors=True):
        if snapshot:
            self.snapshot_add_id()
        out = self.session.add(ID, parent, insert_row=insert_row, snapshot=snapshot)
        if not out["ok"]:
            if snapshot and self.vs and self.vs[-1].get("type") == "add id" and not self.vs[-1]["row"]:
                self.vs.pop()
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            return False
        if snapshot:
            self.refresh_formatting(rows=len(self.sheet.data) - 1 if insert_row is None else insert_row)
        return True

    def change_ID_name(self, ID, new_name, snapshot=True, errors=True):
        ik = ID.lower()
        nnk = new_name.lower()
        if snapshot:
            self.snapshot_rename_id()
        out = self.session.rename(ID, new_name, snapshot=snapshot)
        if not out["ok"]:
            if snapshot and self.vs and self.vs[-1].get("type") == "rename id" and not self.vs[-1]["rows"]:
                self.vs.pop()
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            return False
        if ik in self.saved_info[self.pc].opens:
            self.saved_info[self.pc].opens[nnk] = self.saved_info[self.pc].opens.pop(ik)
        return True

    def cut_paste(
        self,
        ID,
        oldparent,
        hier,
        newparent,
        snapshot=True,
        errors=True,
        sort_later=False,
    ):
        out = self.session.cut_paste(ID, oldparent, hier, newparent, snapshot=snapshot, sort_later=sort_later)
        if not out["ok"]:
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            return False
        return True

    def cut_paste_all(
        self,
        ID,
        oldparent,
        hier,
        newparent,
        snapshot=True,
        errors=True,
        sort_later=False,
    ):
        out = self.session.cut_paste_all(ID, oldparent, hier, newparent, snapshot=snapshot, sort_later=sort_later)
        if not out["ok"]:
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            return False
        return True

    def copy_paste(self, ID, hier, newparent, snapshot=True, errors=True, sort_later=False):
        out = self.session.copy_paste(ID, hier, newparent, snapshot=snapshot, sort_later=sort_later)
        if not out["ok"]:
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            return False
        return True

    def copy_paste_all(self, ID, hier, newparent, snapshot=True, errors=True, sort_later=False):
        out = self.session.copy_paste_all(ID, hier, newparent, snapshot=snapshot, sort_later=sort_later)
        if not out["ok"]:
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            return False
        return True

    def cut_paste_children(self, oldparent, newparent, hier, snapshot=True, errors=True):
        pk = oldparent.lower()
        npk = newparent.lower()
        if pk not in self.nodes:
            if errors:
                Error(self, "ID doesn't exist   ", theme=self.C.theme)
            return
        if not len(self.nodes[pk].cn[hier]):
            if errors:
                Error(
                    self,
                    f"{self.nodes[pk].name} has no children   ",
                    theme=self.C.theme,
                )
            return
        already_in = set()
        if hier != self.pc:
            for ciid in self.nodes[pk].cn[hier]:
                for diid in self.check_cn(ciid, hier):
                    if self.nodes[diid].ps[self.pc] is not None:
                        already_in.add(ciid)
                        break
            if len(already_in) == len(self.nodes[pk].cn[hier]):
                if errors:
                    Error(
                        self,
                        f"Unable to move children, key IDs are already in {self.headers[self.pc].name}   ",
                        theme=self.C.theme,
                    )
                return
        else:
            if any(npk == ck for ck in self.check_cn(pk, hier)):
                if errors:
                    Error(self, "Cannot add ID to same line   ", theme=self.C.theme)
                return False
            if pk == npk:
                if errors:
                    Error(self, "Children already have this parent   ", theme=self.C.theme)
                return False
        if already_in and errors and snapshot:
            confirm = Ask_Confirm(
                self,
                f"Move {oldparent}'s children\nCannot move the following IDs to {self.headers[self.pc].name}, skip?\n{', '.join(already_in)}",
                theme=self.C.theme,
            )
            if not confirm.boolean:
                return False
        out = self.session.cut_paste_children(oldparent, newparent, hier, snapshot=snapshot)
        if not out["ok"]:
            if errors:
                Error(self, out["error"]["message"], theme=self.C.theme)
            msg = out["error"]["message"]
            if out["error"]["code"] == "cycle" or "already have this parent" in msg:
                return False
            return
        return True

    def cut_paste_edit_cell(self, ID, oldparent, hier, newparent, snapshot=True):
        ik = ID.lower()
        pk = oldparent.lower()
        npk = newparent.lower()
        if ik == npk:
            return False
        if npk != "":
            if npk not in self.nodes:
                return False
            if self.nodes[npk].ps[hier] is None:
                return False
            if self.nodes[ik].ps[hier] and npk == self.nodes[ik].ps[hier]:
                return False
        else:
            if self.nodes[ik].ps[hier] == "":
                return False
        if any(npk == ck for ck in self.check_cn(ik, hier)):
            return False
        if oldparent == "" and self.nodes[ik].ps[hier] is None and newparent:
            for ck in self.check_cn(ik, hier):
                if self.nodes[ck].ps[hier] is not None:
                    return False
        self.nodes[ik].ps[hier] = None
        if pk != "":
            self.nodes[pk].cn[hier].remove(ik)
        if npk == "":
            self.nodes[ik].ps[hier] = ""
        else:
            self.nodes[ik].ps[hier] = npk
            self.nodes[npk].cn[hier].append(ik)
            if self.auto_sort_nodes_bool:
                self.nodes[npk].cn[hier] = self.sort_node_cn(self.nodes[npk].cn[hier], hier)
                if self.nodes[npk].ps[hier]:
                    parent_parent_node = self.nodes[self.nodes[npk].ps[hier]]
                    parent_parent_node.cn[hier] = self.sort_node_cn(parent_parent_node.cn[hier], hier)
        if not self.auto_sort_nodes_bool:
            if pk == "":
                try_remove(self.topnodes_order[hier], ik)
            if npk == "":
                self.topnodes_order[hier].append(ik)
        idrow = self.rns[ik]
        if snapshot:
            self.vs[-1]["rows"].append(
                zlib.compress(
                    pickle.dumps(
                        (
                            idrow,
                            hier,
                            self.sheet.MT.data[idrow][hier],
                            hier,
                            self.sheet.MT.data[idrow][hier],
                        )
                    )
                )
            )
        self.sheet.MT.data[idrow][hier] = newparent
        return True

    def _del_id_core(self, name: str, to_del: list[str] | None = None, snapshot: bool = True) -> list[str]:
        return self.session._del_id_core(name, to_del, snapshot)

    def _del_id_all_core(self, name: str, to_del: list[str] | None = None, snapshot: bool = True) -> list[str]:
        return self.session._del_id_all_core(name, to_del, snapshot)

    def _del_id_orphan_core(self, name: str, parent: str, snapshot: bool = True) -> list[str]:
        return self.session._del_id_orphan_core(name, parent, snapshot=snapshot)

    def _del_id_all_orphan_core(self, name: str, snapshot: bool = True) -> list[str]:
        return self.session._del_id_all_orphan_core(name, snapshot=snapshot)

    def get_lvls(self, iid: str, lvl=1):
        # Initialize stack with the initial node at lvl - 1
        stack = [(iid, lvl - 1)]

        while stack:
            # Pop the current node and its level
            current_iid, current_lvl = stack.pop()

            # Get the children of the current node
            children = self.nodes[current_iid].cn[self.pc]

            # The children's level is the next level
            next_lvl = current_lvl + 1

            # Ensure the level exists in self.levels
            if next_lvl not in self.levels:
                self.levels[next_lvl] = []

            # Process each child
            for child in children:
                self.levels[next_lvl].append(child)
                stack.append((child, next_lvl))

    def _del_id_children_core(self, name: str, to_del: list[str] | None = None, snapshot: bool = True) -> list[str]:
        return self.session._del_id_children_core(name, to_del, snapshot)

    def _del_id_children_all_core(self, name: str, to_del: list[str] | None = None, snapshot: bool = True) -> list[str]:
        return self.session._del_id_children_all_core(name, to_del, snapshot)

    def details(self, ik):
        allrows = []
        spaces = " " * 10
        string = f"\n ID:   {self.nodes[ik].name}"
        allrows.append(string)
        allrows.append("\n\n Parents across all hierarchies:")
        for h, p in self.nodes[ik].ps.items():
            allrows.append(f"    Column #{h + 1} {self.headers[h].name}: ")
            if p == "":
                allrows.append(spaces + "Appears as top ID")
            elif p is not None:
                allrows.append(spaces + self.nodes[p].name)
        allrows.append("\n\n Children across all hierarchies:")
        for h in self.nodes[ik].cn:
            allrows.append(f"    Column #{h + 1} {self.headers[h].name}: ")
            for ciid in self.nodes[ik].cn[h]:
                allrows.append(spaces + self.nodes[ciid].name)
        if len(self.hiers) + 1 == self.row_len:
            allrows.append(spaces + "\n\n No detail columns in sheet")
        else:
            idcol_hiers = set(self.hiers) | {self.ic}
            allrows.append("\n\n Details:")
            for index, cell in enumerate(self.sheet.MT.data[self.rns[ik]]):
                if index not in idcol_hiers:
                    allrows.append(f"    Column #{index + 1} {self.headers[index].name}:")
                    allrows.append(spaces + cell)
        return "\n".join(allrows)

    def fix_associate_sort(self, startup=True):
        self.session.associate(startup=startup)

    def fix_associate_sort_edit_cells(self):
        self.session.associate_after_edit()
        self._sync_sheet_from_session()
        return "break"

    def associate(self):
        self.session.repair_empty_hierarchies()

    def sort_node_cn(self, cn: list[str], h: int):
        return self.session.sort_node_cn(cn, h)

    def top_iids(self):
        yield from self.session.top_iids()

    def pc_iids(self) -> Generator[str]:
        top_iter = iter(self.top_iids())
        stack = []
        while True:
            if not stack:
                try:
                    top_iid = next(top_iter)
                    yield top_iid
                    stack.extend(reversed(self.nodes[top_iid].cn[self.pc]))
                except StopIteration:  # If no more top nodes, exit
                    break
            else:
                iid = stack.pop()
                yield iid
                stack.extend(reversed(self.nodes[iid].cn[self.pc]))

    def remake_topnodes_order(self):
        self.session.remake_topnodes_order()

    def gen_sheet_w_headers(self):
        return self.session.gen_sheet_w_headers()

    def check_validation_validity(self, col: int, validation: list[str]) -> str | list[str]:
        return self.session.check_validation_validity(col, validation)

    def apply_validation_to_col(self, col):
        self.session.apply_validation_to_col(col)

    def refresh_formatting(
        self,
        rows: int | Iterator | None = None,
        columns: int | Iterator | None = None,
        dehighlight: bool = False,
        ignore_empty: bool = False,
    ):
        if dehighlight:
            self.sheet.dehighlight_cells(all_=True, redraw=False)

        if rows is None:
            rows = range(len(self.sheet.MT.data))
        elif isinstance(rows, int):
            rows = (rows,)

        if columns is None:
            columns = tuple(range(len(self.headers)))
        elif isinstance(columns, int):
            columns = (columns,)

        if not rows:
            return

        quick_data = self.sheet.MT.data
        for col in columns:
            if ignore_empty and not self.headers[col].formatting:
                continue
            conditions = self.headers[col].formatting
            for rn in rows:
                self.sheet.dehighlight_cells(row=rn, column=col, redraw=False)
                cell = quick_data[rn][col]
                for cond, color in conditions:
                    if cell.lower() == cond.lower():
                        self.sheet.highlight_cells(row=rn, column=col, bg=color, fg="black")
                        break

        self.refresh_rows = set()

    def rc_edit_validation(self, event=None):
        if (col := self.rc_selected_col()) is None:
            return
        popup = Edit_Validation_Popup(
            self,
            self.headers[col].name,
            self.headers[col].validation,
            self.C.theme,
        )
        if popup.new_validation is None:
            return
        if popup.new_validation:
            validation = self.check_validation_validity(col, popup.new_validation)
            if isinstance(validation, str):
                Error(
                    self,
                    f" {validation}     see 'Help' under the 'File' menu for instructions on validation   ",
                    theme=self.C.theme,
                )
                return
        else:
            validation = []
        if validation == self.headers[col].validation:
            return
        self.snapshot_edit_validation(col, validation)
        self.session.set_validation(col, validation, snapshot=False)
        self.refresh_dropdowns()
        self.refresh_formatting(columns=col)
        self.redo_tree_display()
        self.redraw_sheets()

    def rc_edit_formatting(self, event=None):
        if (col := self.rc_selected_col(allow_hiers=True)) is None:
            return
        self.save_info_get_saved_info()
        Edit_Conditional_Formatting_Popup(self, column=col, theme=self.C.theme)
        self.headers[col].formatting = [tup for tup in self.headers[col].formatting if tup[1]]
        self.refresh_formatting(columns=col, ignore_empty=bool(self.headers[col].formatting))
        self.redo_tree_display()
        self.redraw_sheets()

    def rc_selected_col(self, allow_hiers=False):
        widget = self.sheet if self.sheet.has_focus() else self.tree
        col = widget.get_selected_columns()
        if len(col) != 1:
            return
        col = widget.selected.column
        if not allow_hiers and (col == self.ic or col in self.hiers):
            return
        return col

    def is_in_validation(self, validation, text):
        return self.session.is_in_validation(validation, text)

    def detail_is_valid_for_col(self, col, detail):
        return self.session.detail_is_valid_for_col(col, detail)

    def increment_unsaved(self):
        self.C.unsaved_changes = True
        self.set_undo_label()
        self.C.change_app_title(star="add")

    def get_datetime_changelog(self, increment_unsaved=True):
        return self.session.get_datetime_changelog(increment_unsaved=increment_unsaved)

    def rc_rename_col(self, event=None):
        if (col := self.rc_selected_col(allow_hiers=True)) is None:
            return
        if col in self.hiers:
            popup = Rename_Column_Popup(self, self.headers[col].name, "hierarchy", theme=self.C.theme)
        elif col == self.ic:
            popup = Rename_Column_Popup(self, self.headers[col].name, "ID", theme=self.C.theme)
        else:
            popup = Rename_Column_Popup(self, self.headers[col].name, "detail", theme=self.C.theme)
        if not popup.result:
            return
        new_name = popup.result
        new_name_k = new_name.lower()
        if any(new_name_k == h.name.lower() for h in self.headers):
            Error(self, f"Name: {new_name} already exists", theme=self.C.theme)
            return
        self.rename_col(col, new_name)
        self.disable_paste()
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.set_headers()

    def rename_col(self, col, name, snapshot=True):
        if snapshot:
            self.snapshot_rename_col()
        self.session.rename_col(col, name, snapshot=snapshot)
        if snapshot:
            self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def add_hier_col(self, col, name, snapshot=True):
        if snapshot:
            self.snapshot_add_col(col)
        self.tv_label_col = push_n(self.tv_label_col, [col])
        self.saved_info = {push_n(k, [col]): v for k, v in self.saved_info.items()}
        self.session.add_hier_col(col, name, snapshot=snapshot)
        self.saved_info[col] = new_info_storage()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def rc_add_hier_col(self, event=None):
        if (col := self.rc_selected_col(allow_hiers=True)) is None:
            col = len(self.headers)
        popup = Add_Hierarchy_Column_Popup(self, theme=self.C.theme)
        if not popup.result:
            return
        name = popup.result
        namekey = name.lower()
        if any(namekey == h.name.lower() for h in self.headers):
            Error(self, f"Column {name} already exists.", theme=self.C.theme)
            return
        self.add_hier_col(col, name)
        self.disable_paste()
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.set_headers()

    def rc_add_col(self, event=None):
        if (col := self.rc_selected_col(allow_hiers=True)) is None:
            col = len(self.headers)
        popup = Add_Detail_Column_Popup(self, theme=self.C.theme)
        if not popup.result:
            return
        name = popup.result
        namekey = name.lower()
        type_ = popup.type_
        if any(namekey == h.name.lower() for h in self.headers):
            Error(self, f"Column {name} already exists.", theme=self.C.theme)
            return
        self.add_col(col, name, type_)
        self.disable_paste()
        self.set_headers()

    def insert_columns_no_blank_row(self, columns=1, **kwargs):
        if isinstance(columns, int) and columns < 1:
            return
        if self.sheet.MT.data:
            self.tree.insert_columns(columns, **kwargs)
            self.sheet.insert_columns(columns, **kwargs)
            return
        # insert_columns() fabricates a row when there is no data; only add column positions
        n = columns if isinstance(columns, int) else len(columns)
        idx = kwargs.get("idx")
        if idx is None:
            idx = "end"
        self.tree.insert_column_positions(idx=idx, widths=n)
        self.sheet.insert_column_positions(idx=idx, widths=n)

    def add_col(self, col, name, type_, snapshot=True):
        if snapshot:
            self.snapshot_add_col(col)
        self.tv_label_col = push_n(self.tv_label_col, [col])
        self.saved_info = {push_n(k, [col]): v for k, v in self.saved_info.items()}
        self.session.add_col(col, name, type_, snapshot=snapshot)
        if snapshot:
            self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def del_cols(self, cols, snapshot=True):
        if snapshot:
            self.snapshot_del_cols()
        cols_set = set(cols)
        hiers_orig = list(self.hiers)
        if hiers_to_del := list(filter(cols_set.__contains__, reversed(hiers_orig))):
            for col in hiers_to_del:
                del self.saved_info[col]
        self.session.del_cols(cols, snapshot=snapshot)
        if self.tv_label_col == self.ic or self.tv_label_col in cols:
            self.tv_label_col = self.ic
        else:
            self.tv_label_col = (
                self.tv_label_col if not (num := bisect_left(cols, self.tv_label_col)) else self.tv_label_col - num
            )
        self.saved_info = {k if not (num := bisect_left(cols, k)) else k - num: v for k, v in self.saved_info.items()}
        if snapshot:
            self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def del_cols_rc(self, event=None):
        if self.tree.has_focus():
            cols = self.tree.get_selected_columns()
        elif self.sheet.has_focus():
            cols = self.sheet.get_selected_columns()
        focused = self.tree if self.tree.has_focus() else self.sheet
        if self.ic in cols or self.pc in cols:
            Error(
                self,
                "Cannot delete selected columns, they contain either the ID column or the current hierarchy   ",
                theme=self.C.theme,
            )
            return
        cols = sorted(cols)
        self.save_info_get_saved_info()
        self.del_cols(cols)
        self.clear_copied_details()
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.sheet.row_index(newindex=self.ic)
        self.refresh_dropdowns()
        self.redo_tree_display()
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        if focused == self.tree:
            self.focus_tree()
        else:
            self.focus_sheet()

    def cut_cols(self, cols):
        self.cut_columns = cols

    def adjust_hiers_del_cols(self, cols):
        auto_sort_nodes_bool = self.auto_sort_nodes_bool
        self.hiers = [k if not (num := bisect_left(cols, k)) else k - num for k in self.hiers]
        for node in self.nodes.values():
            node.ps = {k if not (num := bisect_left(cols, k)) else k - num: v for k, v in node.ps.items()}
            node.cn = {k if not (num := bisect_left(cols, k)) else k - num: v for k, v in node.cn.items()}
        self.saved_info = {k if not (num := bisect_left(cols, k)) else k - num: v for k, v in self.saved_info.items()}
        if not auto_sort_nodes_bool:
            self.topnodes_order = {
                k if not (num := bisect_left(cols, k)) else k - num: v for k, v in self.topnodes_order.items()
            }

    def adjust_hiers_add_cols(self, cols):
        auto_sort_nodes_bool = self.auto_sort_nodes_bool
        self.hiers = [push_n(k, cols) for k in self.hiers]
        for node in self.nodes.values():
            node.ps = {push_n(k, cols): v for k, v in node.ps.items()}
            node.cn = {push_n(k, cols): v for k, v in node.cn.items()}
        self.saved_info = {push_n(k, cols): v for k, v in self.saved_info.items()}
        if not auto_sort_nodes_bool:
            self.topnodes_order = {push_n(k, cols): v for k, v in self.topnodes_order.items()}

    def refresh_hier_dropdown(self, idx):
        switch_values = [f"{self.headers[h].name}" for h in self.hiers]
        self.switch_hier_dropdown["values"] = switch_values
        self.switch_displayed.set(switch_values[idx])

    def refresh_dropdowns(self):
        self.tree.del_dropdown("A:")
        self.sheet.del_dropdown("A:")
        for c, hdr in enumerate(self.headers):
            if hdr.validation:
                self.sheet.dropdown(
                    _n2a(c),
                    values=hdr.validation,
                    edit_data=False,
                )
                self.tree.dropdown(
                    _n2a(c),
                    values=hdr.validation,
                    edit_data=False,
                )
        self.redraw_sheets()

    def refresh_tree_dropdowns(self):
        self.tree.del_dropdown("A:")
        for c, hdr in enumerate(self.headers):
            if hdr.validation:
                self.tree.dropdown(
                    _n2a(c),
                    values=hdr.validation,
                    edit_data=False,
                )
        self.redraw_sheets()

    def undo(self, event=None):
        if self.C.working or not self.vs:
            return "break"
        self.start_work("Undoing last action...")
        self.C.unsaved_changes = True
        self.C.change_app_title(star="add")
        new_vs = self.vs[-1]
        rd = new_vs["required_data"]
        typ = new_vs["type"]
        self.session.undo()
        self._sync_sheet_from_session()
        self.tv_label_col = rd["tv_label_col"]
        self.mirror_var = rd["mirror_bool"]
        self.saved_info = pickle.loads(rd["saved_info"])
        self.sheet.align_columns(columns=rd["sheet_column_alignments"], redraw=False)
        self.tree.align_columns(columns=rd["sheet_column_alignments"], redraw=False)
        self.reset_tagged_ids_dropdowns()
        self.clear_copied_details()
        if typ.startswith("full"):
            self.warnings_filepath = new_vs.get("og_file", self.warnings_filepath)
            self.warnings_sheet = new_vs.get("og_sheet", self.warnings_sheet)
        self.refresh_formatting(dehighlight=True)
        self.sheet.row_index(newindex=self.ic)
        self.sheet.set_column_widths(new_vs["required_data"]["sheet_col_positions"], canvas_positions=True)
        self.sheet.set_safe_row_heights(new_vs["required_data"]["sheet_row_positions"])
        self.redo_tree_display()
        self.set_headers()
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.rehighlight_tagged_ids()
        self.set_undo_label()
        self.mirror_sels_disabler = True
        self.move_tree_pos()
        if new_vs["required_data"]["sheet_selections"] is not None:
            self.reselect_sheet_sel(new_vs["required_data"]["sheet_selections"])
            self.sheet_select_event()
        self.refresh_dropdowns()
        self.move_sheet_pos()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        del new_vs
        self.stop_work(self.get_tree_editor_status_bar_text())

    def tree_gen_heights_from_saved(self) -> Generator[int]:
        heights_dict = self.saved_info[self.pc].theights
        default_row_height = self.tree.MT.get_default_row_height()
        return (
            heights_dict[iid] if (iid := self.tree.rowitem(r, data_index=True)) in heights_dict else default_row_height
            for r in self.tree.displayed_rows
        )

    def tree_gen_widths_from_saved(self) -> Generator[int]:
        widths_dict = self.saved_info[self.pc].twidths
        default_col_width = self.tree.ops.default_column_width
        return (widths_dict[h.name] if h.name in widths_dict else default_col_width for h in self.headers)

    def save_sheet_row_heights(self) -> None:
        quick_data = self.sheet.MT.data
        self.saved_sheet_row_heights = {
            quick_data[rn][self.ic].lower(): h for rn, h in enumerate(self.sheet.get_row_heights())
        }

    def set_undo_label(self, event=None):
        if self.vs:
            self.edit_menu.entryconfig(0, label=f"Undo {len(self.vs)}/30")
        else:
            self.edit_menu.entryconfig(0, label=f"Undo {len(self.vs)}/30", state="disabled")

    def copy_headers(self):
        return [
            Header(
                f"{h.name}",
                f"{h.type_}",
                [tuple(t) for t in h.formatting],
                h.validation.copy(),
            )
            for h in self.headers
        ]

    def save_info_get_saved_info(self):
        self.saved_info[self.pc] = new_info_storage(
            scrolls=(
                float(self.tree.get_xview()[0]),
                float(self.tree.get_yview()[0]),
                float(self.sheet.get_xview()[0]),
                float(self.sheet.get_yview()[0]),
            ),
            opens=dict.fromkeys(self.tree.tree_get_open()),
            boxes=self.tree.boxes,
            selected=self.tree.selected,
            twidths={self.headers[i].name: width for i, width in enumerate(self.tree.get_column_widths())},
            theights={
                self.tree.rowitem(i, data_index=False): height
                for i, height in enumerate(self.tree.get_safe_row_heights())
            },
        )
        return self.saved_info

    def get_required_snapshot_data(self):
        return {
            "saved_info": pickle.dumps(self.save_info_get_saved_info()),
            "sheet_col_positions": list(self.sheet.get_column_widths(canvas_positions=True)),
            "sheet_row_positions": self.sheet.get_safe_row_heights(),
            "topnodes_order": {k: list(v) for k, v in self.topnodes_order.items()},
            "tv_label_col": int(self.tv_label_col),
            "tagged_ids": set(self.tagged_ids),
            "sheet_column_alignments": dict(self.sheet.get_column_alignments()),
            "headers": self.copy_headers(),
            "ic": int(self.ic),
            "pc": int(self.pc),
            "hiers": list(self.hiers),
            "row_len": int(self.row_len),
            "auto_sort_nodes_bool": bool(self.auto_sort_nodes_bool),
            "mirror_bool": bool(self.mirror_var),
            "nodes": pickle.dumps(self.nodes),
            "focus": self.tree.has_focus(),
            "sheet_selections": self.get_sheet_sel(),
        }

    def snapshot_ctrl_x_v_del_key_id_par(self):
        self.snapshot_chore()
        self.save_sheet_row_heights()
        self.vs.append(
            {
                "type": "ctrl x, v, del key id par",
                "sheet": zlib.compress(pickle.dumps(self.sheet.MT.data)),
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_ctrl_x_v_del_key(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "ctrl x, v, del key",
                "cells": {},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_sheet(self, type_: str = "full sheet"):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": type_,
                "og_file": self.warnings_filepath,
                "og_sheet": self.warnings_sheet,
                "build_warnings": self.warnings,
                "sheet": zlib.compress(pickle.dumps(self.sheet.MT.data)),
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_add_id(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "add id",
                "row": {},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_rename_id(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "rename id",
                "rows": [],
                "ikrow": (),
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_paste_id(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "paste id",
                "rows": [],
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_delete_ids(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "delete ids",
                "rows": {},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_add_col(self, treecolsel):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "add col",
                "treecolsel": int(treecolsel),
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_del_cols(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "del cols",
                "cols": {},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_rename_col(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "rename col",
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_edit_validation(self, col, validation):
        self.snapshot_chore()
        self.changelog_append(
            "Edit validation",
            f"Column #{col + 1} named: {self.headers[col].name} with type: {self.headers[col].type_}",
            f"{','.join(self.headers[col].validation)}",
            f"{','.join(validation)}",
        )
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.vs.append(
            {
                "type": "edit validation",
                "col_num": col,
                "col": zlib.compress(pickle.dumps([r[col] for r in self.sheet.MT.data])),
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_begin_drag_rows(self, event=None):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "drag rows",
                "row_mapping": {},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_drag_rows(self, event_data):
        self.vs[-1]["row_mapping"] = self.sheet.full_move_rows_idxs(event_data["moved"]["rows"]["data"])
        old_locs = ",".join(f"{r}" for r in event_data["moved"]["rows"]["data"])
        new_locs = ",".join(f"{r}" for r in event_data["moved"]["rows"]["data"].values())
        self.changelog_append(
            "Move rows",
            f"{len(event_data['moved']['rows']['data'])} rows",
            f"Old locations: {old_locs}",
            f"New locations: {new_locs}",
        )
        self._adopt_sheet_data()
        self.rns = {r[self.ic].lower(): i for i, r in enumerate(self.sheet.data)}
        self.disable_paste()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.redraw_sheets()

    def begin_tree_drag_drop_ids(self, event=None):
        self.save_info_get_saved_info()

    def get_ids_parent(self, iid) -> str:
        return self.session.get_ids_parent(iid)

    def tree_drag_drop_ids(self, event=None):
        if not event.moved.rows.data:
            return
        self.start_work("Moving IDs...")
        moved_rows = [self.tree.data_r(r) for r in event.moved.rows.displayed]
        as_sibling = []
        index_only = []
        if event.value > max(event.moved.rows.displayed):
            event.value -= 1
        move_to_iid = self.tree.rowitem(event.value)
        new_parent = self.tree.parent(move_to_iid)
        for r in moved_rows:
            iid = self.tree.rowitem(r, data_index=True)
            iid_parent = self.tree.parent(iid)
            if iid_parent == new_parent:
                index_only.append(iid)
            else:
                as_sibling.append(iid)
        successful = []
        if as_sibling:
            self.selected_ID = self.nodes[move_to_iid].name
            self.cut_ids(as_sibling, status_bar=False)
            if new_parent:
                self.selected_PAR = self.nodes[new_parent].name
            else:
                self.selected_PAR = new_parent
            successful = [dct["id"] for dct in self.paste_cut_sibling_all(redo_tree=False)]
            if not successful:
                self.disable_paste()
        if successful and (not self.auto_sort_nodes_bool or index_only):
            index_only += successful
        if index_only:
            if self.auto_sort_nodes_bool:
                self.auto_sort_nodes_bool = False
                self.remake_topnodes_order()
            self.redo_tree_display(selections=False)
            move_to_index = self.tree.index(move_to_iid)
            if (
                not is_contiguous(event.moved.rows.displayed)
                and max(moved_rows) > self.tree.data_r(event.value)
                and min(moved_rows) < self.tree.data_r(event.value)
            ):
                move_to_index -= 1
            if parik := self.get_ids_parent(index_only[0]):
                self.nodes[parik].cn[self.pc].insert(
                    move_to_index,
                    self.nodes[parik].cn[self.pc].pop(self.tree.index(index_only[0])),
                )
            else:
                self.topnodes_order[self.pc].insert(
                    move_to_index,
                    self.topnodes_order[self.pc].pop(self.tree.index(index_only[0])),
                )
            if len(index_only) > 1:
                new_index = move_to_index
                for iid in islice(index_only, 1, None):
                    if parik := self.get_ids_parent(iid):
                        current_index = self.nodes[parik].cn[self.pc].index(iid)
                        if new_index < current_index:
                            new_index += 1
                        self.nodes[parik].cn[self.pc].insert(
                            new_index,
                            self.nodes[parik].cn[self.pc].pop(current_index),
                        )
                    else:
                        current_index = self.topnodes_order[self.pc].index(iid)
                        if new_index < current_index:
                            new_index += 1
                        self.topnodes_order[self.pc].insert(
                            new_index,
                            self.topnodes_order[self.pc].pop(current_index),
                        )
        self.redo_tree_display(selections=False)
        self.redraw_sheets()
        all_iids = index_only + successful
        if all_iids:
            self.tree.scroll_to_item(all_iids[0])
            self.tree.selection_set(all_iids)
        self.stop_work(self.get_tree_editor_status_bar_text())

    def snapshot_begin_drag_cols(self, event=None):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "drag cols",
                "column_mapping": {},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_drag_cols(self, event_data):
        if self.tree.has_focus():
            full_new_idxs = self.tree.full_move_columns_idxs(event_data["moved"]["columns"]["data"])
            self.sheet.mapping_move_columns(
                event_data["moved"]["columns"]["data"],
                event_data["moved"]["columns"]["displayed"],
                undo=False,
            )
        else:
            full_new_idxs = self.sheet.full_move_columns_idxs(event_data["moved"]["columns"]["data"])
            self.tree.mapping_move_columns(
                event_data["moved"]["columns"]["data"],
                event_data["moved"]["columns"]["displayed"],
                undo=False,
            )
        old_locs = ",".join(f"{c}" for c in event_data["moved"]["columns"]["data"])
        new_locs = ",".join(f"{c}" for c in event_data["moved"]["columns"]["data"].values())
        self.changelog_append(
            "Move columns",
            f"{len(event_data['moved']['columns']['data'])} columns",
            f"Old locations: {old_locs}",
            f"New locations: {new_locs}",
        )
        self.ic = full_new_idxs[self.ic]
        self.pc = full_new_idxs[self.pc]
        self.headers = move_elements_by_mapping(
            self.headers,
            event_data["moved"]["columns"]["data"],
        )
        self.set_headers()
        self.hiers = sorted(full_new_idxs[c] for c in self.hiers)
        self.tv_label_col = full_new_idxs[self.tv_label_col]
        for node in self.nodes.values():
            node.cn = {full_new_idxs[k]: v for k, v in node.cn.items()}
            node.ps = {full_new_idxs[k]: v for k, v in node.ps.items()}
        self.saved_info = {full_new_idxs[k]: v for k, v in self.saved_info.items()}
        if not self.auto_sort_nodes_bool:
            self.topnodes_order = {full_new_idxs[k]: v for k, v in self.topnodes_order.items()}
        self.clear_copied_details()
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.sheet.row_index(newindex=self.ic)
        self.vs[-1]["column_mapping"] = dict(zip(full_new_idxs.values(), full_new_idxs))
        self._adopt_sheet_data()
        self.refresh_dropdowns()
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def set_headers(self, tree: bool = True, sheet: bool = True):
        headers = [
            "\n".join((f"{h.name}", f"{i + 1}/{_n2a(i)} {h.name}", f"{h.type_} {h.name}"))
            for i, h in enumerate(self.headers)
        ]
        if sheet:
            self.sheet.headers(
                headers,
                reset_col_positions=False,
                show_headers_if_not_sheet=False,
            )
        if tree:
            self.tree.headers(
                headers.copy(),
                reset_col_positions=False,
                show_headers_if_not_sheet=False,
            )

    def snapshot_sheet_sort(self):
        self.snapshot_chore()
        self.vs.append(
            {
                "type": "sort",
                "ids": {v: k for k, v in self.rns.items()},
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_prune_changelog(self, up_to):
        self.snapshot_chore()
        out = self.session.prune_changelog(up_to)
        if not out["ok"]:
            return
        self.vs.append(
            {
                "type": "prune changelog",
                "rows": out["result"]["removed"],
                "changelog_at_open": out["result"]["changelog_at_open"],
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_auto_sort_nodes(self):
        self.snapshot_chore()
        self.changelog_append(
            "Sort treeview",
            "Alphanumerically sorted order of Treeview IDs",
            "",
            "",
        )
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.vs.append(
            {
                "type": "node sort",
                "required_data": self.get_required_snapshot_data(),
            }
        )

    def snapshot_chore(self):
        self.save_info_get_saved_info()
        self.edit_menu.entryconfig(0, label=f"Undo {len(self.vs)}/30", state="normal")

    def sort_sheet_choice(self):
        popup = Sort_Sheet_Popup(self, [h.name for h in self.headers], theme=self.C.theme)
        if popup.sort_decision["type"] is None:
            return
        if popup.sort_decision["type"] == "by column":
            self.sort_sheet(popup.sort_decision["col"], popup.sort_decision["order"])
        elif popup.sort_decision["type"] == "by tree":
            self.sort_sheet_walk()

    def sort_sheet_rc_asc(self):
        widget = self.tree if self.tree.has_focus() else self.sheet
        if widget.get_selected_columns():
            self.sort_sheet(header=self.headers[widget.selected.column].name, order="ASCENDING")

    def sort_sheet_rc_desc(self):
        widget = self.tree if self.tree.has_focus() else self.sheet
        if widget.get_selected_columns():
            self.sort_sheet(header=self.headers[widget.selected.column].name, order="DESCENDING")

    def sort_sheet(self, header, order, snapshot=True):
        row_heights = self.sheet.get_row_heights()
        old_rns = dict(self.rns)
        if snapshot:
            self.snapshot_sheet_sort()
        self.session.sort_sheet(header, order, snapshot=snapshot)
        nrhs = [row_heights[old_rns[r[self.ic].lower()]] for r in self.sheet.MT.data]
        self.sheet.set_row_heights(nrhs)
        if snapshot:
            self.disable_paste()
            self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
            self.refresh_formatting()
            self.reset_tagged_ids_dropdowns()
            self.rehighlight_tagged_ids()
            self.redraw_sheets()

    def sort_sheet_walk(self, snapshot=True):
        oldrns = self.rns.copy()
        row_heights = self.sheet.get_row_heights()
        if snapshot:
            self.snapshot_sheet_sort()
        self.session.sort_sheet_walk(snapshot=snapshot)
        self._sync_sheet_from_session()
        nrhs = [row_heights[oldrns[r[self.ic].lower()]] for r in self.sheet.MT.data]
        self.sheet.set_row_heights(nrhs)
        if snapshot:
            self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
            self.refresh_formatting()
            self.reset_tagged_ids_dropdowns()
            self.rehighlight_tagged_ids()
            self.disable_paste()
            self.redraw_sheets()

    def search_choice(self, event=None):
        choice = self.search_choice_displayed.get()
        if choice == "Non-exact":
            self.search_for_any(None)
        elif choice == "ID non-exact":
            self.search_for_ID(None, False)
        elif choice == "ID exact":
            self.search_for_ID(None, True)
        elif choice == "Detail non-exact":
            self.search_for_detail(None, False)
        elif choice == "Detail exact":
            self.search_for_detail(None, True)

    def sheet_search_choice(self, event=None):
        choice = self.sheet_search_choice_displayed.get()
        if choice == "Non-exact":
            self.sheet_search_for_any()
        elif choice == "ID non-exact":
            self.sheet_search_for_ID(None, False)
        elif choice == "ID exact":
            self.sheet_search_for_ID(None, True)
        elif choice == "Detail non-exact":
            self.sheet_search_for_detail(None, False)
        elif choice == "Detail exact":
            self.sheet_search_for_detail(None, True)

    def search_for_any(self, find=None):
        if not (search := self.search_entry.get() if find is None else find):
            return
        self.reset_tree_search_dropdown()
        search = search.lower()
        for iid, node in self.nodes.items():
            for i, e in enumerate(self.sheet.MT.data[self.rns[iid]]):
                if search in e.lower():
                    for h, par in node.ps.items():
                        if par is not None:
                            self.search_results.append(
                                SearchResult(
                                    hierarchy=h,
                                    text=(
                                        self.headers[h].name,
                                        node.name,
                                        self.headers[i].name,
                                        re.sub(remove_nrt, "", e),
                                    ),
                                    iid=iid,
                                    column=i,
                                    term=search,
                                    type_=2,
                                    exact=False,
                                )
                            )
        if self.search_results:
            col_chars = frame_w_to_nchars(
                frame_w=self.search_dropdown.winfo_width(),
                fixed_font_w=self.fixed_font_w,
                ncols=4,
            )
            process_search_results(
                self.search_results,
                search_results_max_column_chars(self.search_results, col_chars),
                col_chars,
            )
        self.display_search_results()

    def search_for_ID(self, find=None, exact=False):
        if not (search := self.search_entry.get() if find is None else find):
            return
        self.reset_tree_search_dropdown()
        search = search.lower()
        for iid, node in self.nodes.items():
            if (exact and search == iid) or (not exact and search in iid):
                for h, par in node.ps.items():
                    if par is not None:
                        self.search_results.append(
                            SearchResult(
                                hierarchy=h,
                                text=(
                                    self.headers[h].name,
                                    node.name,
                                ),
                                iid=iid,
                                column=self.ic,
                                term=search,
                                type_=0,
                                exact=exact,
                            )
                        )
        if self.search_results:
            col_chars = frame_w_to_nchars(
                frame_w=self.search_dropdown.winfo_width(),
                fixed_font_w=self.fixed_font_w,
                ncols=2,
            )
            process_search_results(
                self.search_results,
                search_results_max_column_chars(self.search_results, col_chars),
                col_chars,
            )
        self.display_search_results()

    def search_for_detail(self, find=None, exact=False):
        if not (search := self.search_entry.get() if find is None else find):
            return
        self.reset_tree_search_dropdown()
        search = search.lower()
        idcol_hiers = set(self.hiers) | {self.ic}
        for iid, node in self.nodes.items():
            for i, e in enumerate(self.sheet.MT.data[self.rns[iid]]):
                if i not in idcol_hiers and ((exact and search == e.lower()) or (not exact and search in e.lower())):
                    for h, par in node.ps.items():
                        if par is not None:
                            self.search_results.append(
                                SearchResult(
                                    hierarchy=h,
                                    text=(
                                        self.headers[h].name,
                                        node.name,
                                        self.headers[i].name,
                                        re.sub(remove_nrt, "", e),
                                    ),
                                    iid=iid,
                                    column=i,
                                    term=search,
                                    type_=1,
                                    exact=exact,
                                )
                            )
        if self.search_results:
            col_chars = frame_w_to_nchars(
                frame_w=self.search_dropdown.winfo_width(),
                fixed_font_w=self.fixed_font_w,
                ncols=4,
            )
            process_search_results(
                self.search_results,
                search_results_max_column_chars(self.search_results, col_chars),
                col_chars,
            )
        self.display_search_results()

    def display_search_results(self):
        if self.search_results:
            self.search_results.sort(key=attrgetter("hierarchy"))
            self.search_dropdown["values"] = [result.text for result in self.search_results]
            self.search_displayed.set(self.search_results[0].text)
            self.show_search_result()

    def sheet_search_for_any(self, find=None):
        if not (search := self.sheet_search_entry.get() if find is None else find):
            return
        self.reset_sheet_search_dropdown()
        search = search.lower()
        for r in self.sheet.MT.data:
            for i, e in enumerate(r):
                if search in e.lower():
                    self.sheet_search_results.append(
                        SearchResult(
                            hierarchy=self.pc,
                            text=(
                                r[self.ic],
                                self.headers[i].name,
                                re.sub(remove_nrt, "", e),
                            ),
                            iid=r[self.ic].lower(),
                            column=i,
                            term=search,
                            type_=2,
                            exact=False,
                        )
                    )
        if self.sheet_search_results:
            col_chars = frame_w_to_nchars(
                frame_w=self.sheet_search_dropdown.winfo_width(),
                fixed_font_w=self.fixed_font_w,
                ncols=3,
            )
            process_search_results(
                self.sheet_search_results,
                search_results_max_column_chars(self.sheet_search_results, col_chars),
                col_chars,
            )
        self.sheet_display_search_results()

    def sheet_search_for_ID(self, find=None, exact=False):
        if not (search := self.sheet_search_entry.get() if find is None else find):
            return
        self.reset_sheet_search_dropdown()
        search = search.lower()
        for r in self.sheet.MT.data:
            if (exact and search == r[self.ic].lower()) or (not exact and search in r[self.ic].lower()):
                self.sheet_search_results.append(
                    SearchResult(
                        hierarchy=self.pc,
                        text=(r[self.ic],),
                        iid=r[self.ic].lower(),
                        column=self.ic,
                        term=search,
                        type_=0,
                        exact=exact,
                    )
                )
        if self.sheet_search_results:
            col_chars = frame_w_to_nchars(
                frame_w=self.sheet_search_dropdown.winfo_width(),
                fixed_font_w=self.fixed_font_w,
                ncols=1,
            )
            process_search_results(
                self.sheet_search_results,
                search_results_max_column_chars(self.sheet_search_results, col_chars),
                col_chars,
            )
        self.sheet_display_search_results()

    def sheet_search_for_detail(self, find=None, exact=False):
        if not (search := self.sheet_search_entry.get() if find is None else find):
            return
        self.reset_sheet_search_dropdown()
        search = search.lower()
        idcol_hiers = set(self.hiers) | {self.ic}
        for r in self.sheet.MT.data:
            for i, e in enumerate(r):
                if i not in idcol_hiers and ((exact and search == e.lower()) or (not exact and search in e.lower())):
                    self.sheet_search_results.append(
                        SearchResult(
                            hierarchy=self.pc,
                            text=(
                                r[self.ic],
                                self.headers[i].name,
                                re.sub(remove_nrt, "", e),
                            ),
                            iid=r[self.ic].lower(),
                            column=i,
                            term=search,
                            type_=1,
                            exact=exact,
                        )
                    )
        if self.sheet_search_results:
            col_chars = frame_w_to_nchars(
                frame_w=self.sheet_search_dropdown.winfo_width(),
                fixed_font_w=self.fixed_font_w,
                ncols=3,
            )
            process_search_results(
                self.sheet_search_results,
                search_results_max_column_chars(self.sheet_search_results, col_chars),
                col_chars,
            )
        self.sheet_display_search_results()

    def sheet_display_search_results(self):
        if self.sheet_search_results:
            self.sheet_search_dropdown["values"] = [result.text for result in self.sheet_search_results]
            self.sheet_search_displayed.set(self.sheet_search_results[0].text)
            self.sheet_show_search_result()

    def enable_copy_paste(self):
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs as child",
            command=self.paste_copied_child,
            state="normal",
        )
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs as sibling",
            command=self.paste_copied_sibling,
            state="normal",
        )
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs and children as sibling",
            command=self.paste_copied_sibling_all,
            state="normal",
        )
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs and children as child",
            command=self.paste_copied_child_all,
            state="normal",
        )
        self.tree_rc_menu_empty.entryconfig(
            "Paste IDs",
            command=self.paste_copied_empty,
            state="normal",
        )
        self.tree_rc_menu_empty.entryconfig(
            "Paste IDs and children",
            command=self.paste_copied_empty_all,
            state="normal",
        )

    def enable_cut_paste(self):
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs as child",
            command=self.paste_cut_child,
            state="normal",
        )
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs as sibling",
            command=self.paste_cut_sibling,
            state="normal",
        )
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs and children as sibling",
            command=self.paste_cut_sibling_all,
            state="normal",
        )
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Paste IDs and children as child",
            command=self.paste_cut_child_all,
            state="normal",
        )
        self.tree_rc_menu_empty.entryconfig(
            "Paste IDs",
            command=self.paste_cut_empty,
            state="normal",
        )
        self.tree_rc_menu_empty.entryconfig(
            "Paste IDs and children",
            command=self.paste_cut_empty_all,
            state="normal",
        )

    def enable_cut_paste_children(self):
        self.tree_rc_menu_single_row_paste.entryconfig(
            "Attach children",
            state="normal",
            command=self.paste_cut_children,
        )
        self.tree_rc_menu_empty.entryconfig(
            "Attach children",
            state="normal",
            command=self.paste_cut_children_empty,
        )

    def disable_paste(self):
        self.tree_rc_menu_single_row_paste.entryconfig("Paste IDs as child", state="disabled")
        self.tree_rc_menu_single_row_paste.entryconfig("Paste IDs as sibling", state="disabled")
        self.tree_rc_menu_single_row_paste.entryconfig("Paste IDs and children as sibling", state="disabled")
        self.tree_rc_menu_single_row_paste.entryconfig("Paste IDs and children as child", state="disabled")
        self.tree_rc_menu_single_row_paste.entryconfig("Attach children", state="disabled")
        self.tree_rc_menu_empty.entryconfig("Paste IDs", state="disabled")
        self.tree_rc_menu_empty.entryconfig("Paste IDs and children", state="disabled")
        self.tree_rc_menu_empty.entryconfig("Attach children", state="disabled")
        self.cut_columns = None
        self.cut = []
        self.copied = []
        self.cut_children_dct = {}
        return "break"

    def cut_children(self, event=None):
        if not self.selected_ID:
            return
        self.cut_children_dct["id"] = f"{self.selected_ID.lower()}"
        self.cut_children_dct["hier"] = int(self.pc)
        self.enable_cut_paste_children()

    def paste_cut_children(self):
        if not self.cut_children_dct:
            return
        self.snapshot_paste_id()
        if self.nodes[self.cut_children_dct["id"]].cn[self.cut_children_dct["hier"]]:
            see_iid = self.nodes[self.cut_children_dct["id"]].cn[self.cut_children_dct["hier"]][0]
            select_iids = tuple(self.nodes[self.cut_children_dct["id"]].cn[self.cut_children_dct["hier"]])
        else:
            see_iid = ""
            select_iids = ()
        success = self.cut_paste_children(
            self.cut_children_dct["id"], f"{self.selected_ID}", self.cut_children_dct["hier"]
        )
        if not success:
            self.vs.pop()
            self.set_undo_label()
            return
        iid = self.nodes[self.cut_children_dct["id"]].name
        np = self.nodes[self.selected_ID.lower()].name
        self.changelog_append(
            "Cut and paste children",
            "",
            f"Old parent: {iid} old column #{self.cut_children_dct['hier'] + 1} named: {self.headers[self.cut_children_dct['hier']].name}",
            f"New parent: {np} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
        )
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.disable_paste()
        if see_iid:
            self.tree.scroll_to_item(see_iid)
            self.tree.selection_set(select_iids)
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def paste_cut_children_empty(self):
        if not self.cut_children_dct:
            return
        self.snapshot_paste_id()
        if self.nodes[self.cut_children_dct["id"]].cn[self.cut_children_dct["hier"]]:
            see_iid = self.nodes[self.cut_children_dct["id"]].cn[self.cut_children_dct["hier"]][0]
            select_iids = tuple(self.nodes[self.cut_children_dct["id"]].cn[self.cut_children_dct["hier"]])
        else:
            see_iid = ""
            select_iids = ()
        success = self.cut_paste_children(self.cut_children_dct["id"], "", self.cut_children_dct["hier"])
        if not success:
            self.vs.pop()
            self.set_undo_label()
            return
        iid = self.nodes[self.cut_children_dct["id"]].name
        self.changelog_append(
            "Cut and paste children",
            "",
            f"Old parent: {iid} old column #{self.cut_children_dct['hier'] + 1} named: {self.headers[self.cut_children_dct['hier']].name}",
            f"New parent: n/a - No parent new column #{self.pc + 1} named: {self.headers[self.pc].name}",
        )
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.disable_paste()
        if see_iid:
            self.tree.scroll_to_item(see_iid)
            self.tree.selection_set(select_iids)
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def paste_copied_child(self):
        if not self.copied or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.copied)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.copied, 1):
            if self.copy_paste(dct["id"], dct["hier"], self.selected_ID, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.copied)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Copy and paste ID |",
                iid,
                f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: {self.nodes[self.selected_ID.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Copy and paste {len(successful)} IDs",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Copy and paste ID")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_copied_sibling(self):
        if not self.copied or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.copied)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.copied, 1):
            if self.copy_paste(dct["id"], dct["hier"], self.selected_PAR, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.copied)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            if self.selected_PAR == "":
                self.changelog_append_no_unsaved(
                    "Copy and paste ID |",
                    iid,
                    f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
            else:
                self.changelog_append_no_unsaved(
                    "Copy and paste ID |",
                    iid,
                    f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: {self.nodes[self.selected_PAR.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
        if len(successful) > 1:
            self.changelog_append(
                f"Copy and paste {len(successful)} IDs",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Copy and paste ID")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_copied_empty(self):
        if not self.copied:
            return
        self.start_work(f"Pasting {len(self.copied)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.copied, 1):
            if self.copy_paste(dct["id"], dct["hier"], "", sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.copied)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Copy and paste ID |",
                iid,
                f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Copy and paste {len(successful)} IDs",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Copy and paste ID")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_copied_child_all(self):
        if not self.copied or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.copied)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.copied, 1):
            if self.copy_paste_all(dct["id"], dct["hier"], self.selected_ID, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.copied)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Copy and paste ID + children |",
                iid,
                f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: {self.nodes[self.selected_ID.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Copy and paste {len(successful)} IDs + children",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Copy and paste ID + children")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_copied_sibling_all(self):
        if not self.copied or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.copied)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.copied, 1):
            if self.copy_paste_all(dct["id"], dct["hier"], self.selected_PAR, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.copied)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            if self.selected_PAR == "":
                self.changelog_append_no_unsaved(
                    "Copy and paste ID + children |",
                    iid,
                    f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
            else:
                self.changelog_append_no_unsaved(
                    "Copy and paste ID + children |",
                    iid,
                    f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: {self.nodes[self.selected_PAR.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
        if len(successful) > 1:
            self.changelog_append(
                f"Copy and paste {len(successful)} IDs + children",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Copy and paste ID + children")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_copied_empty_all(self):
        if not self.copied:
            return
        self.start_work(f"Pasting {len(self.copied)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.copied, 1):
            if self.copy_paste_all(dct["id"], dct["hier"], "", sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.copied)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Copy and paste ID + children |",
                iid,
                f"From column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Copy and paste {len(successful)} IDs + children",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Copy and paste ID + children")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_cut_child(self):
        if not self.cut or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.cut)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.cut, 1):
            if self.cut_paste(dct["id"], dct["parent"], dct["hier"], self.selected_ID, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.cut)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Cut and paste ID |",
                iid,
                f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: {self.nodes[self.selected_ID.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Cut and paste {len(successful)} IDs",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Cut and paste ID")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_cut_sibling(self):
        if not self.cut or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.cut)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.cut, 1):
            if self.cut_paste(dct["id"], dct["parent"], dct["hier"], self.selected_PAR, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.cut)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            if self.selected_PAR == "":
                self.changelog_append_no_unsaved(
                    "Cut and paste ID |",
                    iid,
                    f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
            else:
                self.changelog_append_no_unsaved(
                    "Cut and paste ID |",
                    iid,
                    f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: {self.nodes[self.selected_PAR.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
        if len(successful) > 1:
            self.changelog_append(
                f"Cut and paste {len(successful)} IDs",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Cut and paste ID")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_cut_empty(self):
        if not self.cut:
            return
        self.start_work(f"Pasting {len(self.cut)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.cut, 1):
            if self.cut_paste(dct["id"], dct["parent"], dct["hier"], "", sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.cut)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Cut and paste ID |",
                iid,
                f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Cut and paste {len(successful)} IDs",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Cut and paste ID")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_cut_child_all(self):
        if not self.cut or not self.selected_ID:
            return
        self.start_work(f"Pasting {len(self.cut)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.cut, 1):
            if self.cut_paste_all(dct["id"], dct["parent"], dct["hier"], self.selected_ID, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.cut)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Cut and paste ID + children |",
                iid,
                f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: {self.nodes[self.selected_ID.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Cut and paste {len(successful)} IDs + children",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Cut and paste ID + children")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def paste_cut_sibling_all(self, redo_tree=True) -> bool:
        successful = []
        if not self.cut or not self.selected_ID:
            return successful
        if redo_tree:
            self.start_work(f"Pasting {len(self.cut)} IDs...")
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.cut, 1):
            if self.cut_paste_all(dct["id"], dct["parent"], dct["hier"], self.selected_PAR, sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.cut)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return successful
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            if self.selected_PAR == "":
                self.changelog_append_no_unsaved(
                    "Cut and paste ID + children |",
                    iid,
                    f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
            else:
                self.changelog_append_no_unsaved(
                    "Cut and paste ID + children |",
                    iid,
                    f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                    f"New parent: {self.nodes[self.selected_PAR.lower()].name} new column #{self.pc + 1} named: {self.headers[self.pc].name}",
                )
        if len(successful) > 1:
            self.changelog_append(
                f"Cut and paste {len(successful)} IDs + children",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Cut and paste ID + children")
        self.refresh_formatting(rows=self.refresh_rows)
        self.refresh_rows = set()
        if redo_tree:
            self.redo_tree_display()
            self.redraw_sheets()
            self.tree.scroll_to_item(iid.lower())
            self.tree.selection_set(tuple(dct["id"] for dct in successful))
            self.disable_paste()
            self.stop_work(self.get_tree_editor_status_bar_text())
        return successful

    def paste_cut_empty_all(self):
        if not self.cut:
            return
        self.start_work(f"Pasting {len(self.cut)} IDs...")
        successful = []
        self.sort_later_dct = None
        self.snapshot_paste_id()
        for i, dct in enumerate(self.cut, 1):
            if self.cut_paste_all(dct["id"], dct["parent"], dct["hier"], "", sort_later=True):
                successful.append(dct)
                if not i % 50:
                    self.C.status_bar.change_text(
                        f"Pasting {len(self.cut)} IDs... attempted: {i} | successful: {len(successful)} "
                    )
                    self.C.update()
        if not successful:
            self.unsuccessful_paste()
            return
        self.session.apply_sort_later()
        for dct in successful:
            iid = self.nodes[dct["id"]].name
            self.changelog_append_no_unsaved(
                "Cut and paste ID + children |",
                iid,
                f"Old parent: {self.nodes[dct['parent']].name if dct['parent'] else 'n/a - Top ID'} old column #{dct['hier'] + 1} named: {self.headers[dct['hier']].name}",
                f"New parent: n/a - Top ID new column #{self.pc + 1} named: {self.headers[self.pc].name}",
            )
        if len(successful) > 1:
            self.changelog_append(
                f"Cut and paste {len(successful)} IDs + children",
                "",
                "",
                "",
            )
        else:
            self.changelog_singular("Cut and paste ID + children")
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(iid.lower())
        self.tree.selection_set(tuple(dct["id"] for dct in successful))
        self.disable_paste()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def unsuccessful_paste(self):
        self.vs.pop()
        self.set_undo_label()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def copy_ID(self, iids: Sequence[str]) -> None | Literal["break"]:
        if not self.selected_ID:
            return
        if self.cut:
            self.cut = []
        self.copied = []
        first_iid = iids[0]
        first_iid_par = self.tree.parent(first_iid)
        h = int(self.pc)
        self.copied.append({"id": first_iid.lower(), "parent": first_iid_par.lower(), "hier": h})
        self.levels = defaultdict(list)
        self.get_par_lvls(h, first_iid.lower())
        first_iid_level = max(self.levels, default=0)
        tr = []
        for iid in islice(iids, 1, None):
            self.levels = defaultdict(list)
            self.get_par_lvls(h, iid.lower())
            iid_level = max(self.levels, default=0)
            if self.tree.parent(iid) == first_iid_par or iid_level == first_iid_level:
                self.copied.append(
                    {
                        "id": iid.lower(),
                        "parent": self.tree.parent(iid).lower(),
                        "hier": h,
                    }
                )
            else:
                tr.append(iid)
        if tr:
            self.tree.selection_remove(tr)
        self.enable_copy_paste()
        self.levels = defaultdict(list)
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        return "break"

    def cut_ids(self, iids: list[str] | None = None, status_bar=True):
        if not self.selected_ID:
            return
        if self.copied:
            self.copied = []
        self.cut = []
        if iids is None:
            iids = self.tree.selection()
        first_iid = iids[0]
        first_iid_par = self.tree.parent(first_iid)
        h = int(self.pc)
        self.cut.append({"id": first_iid.lower(), "parent": first_iid_par.lower(), "hier": h})
        self.levels = defaultdict(list)
        self.get_par_lvls(h, first_iid.lower())
        first_iid_level = max(self.levels, default=0)
        tr = []
        for iid in islice(iids, 1, None):
            self.levels = defaultdict(list)
            self.get_par_lvls(h, iid.lower())
            iid_level = max(self.levels, default=0)
            if self.tree.parent(iid) == first_iid_par or iid_level == first_iid_level:
                self.cut.append(
                    {
                        "id": iid.lower(),
                        "parent": self.tree.parent(iid).lower(),
                        "hier": h,
                    }
                )
            else:
                tr.append(iid)
        if tr:
            self.tree.selection_remove(tr)
        self.enable_cut_paste()
        self.levels = defaultdict(list)
        if status_bar:
            self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        return "break"

    def _set_new_id_treeview_label(self, new_id, new_ik, label):
        if self.tv_label_col == self.ic:
            return
        if not label:
            label = new_id
        col = self.tv_label_col
        old = f"{self.sheet.MT.data[self.rns[new_ik]][col]}"
        self.sheet.MT.data[self.rns[new_ik]][col] = label
        if not self.changelog:
            return
        self.changelog[-1].add_row(
            "Edit cell",
            f"ID: {new_id} column #{col + 1} named: {self.headers[col].name} with type: {self.headers[col].type_}",
            old,
            f"{label}",
        )

    def add_child_node(self):
        if not self.selected_ID:
            return
        sel = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)
        if sel:
            popup = Add_Child_Or_Sibling_Id_Popup(
                self,
                "child",
                self.selected_ID,
                self.sheet.MT.data[sel[0]][self.ic],
                theme=self.C.theme,
            )
        else:
            popup = Add_Child_Or_Sibling_Id_Popup(self, "child", self.selected_ID, None, theme=self.C.theme)
        if not popup.result:
            return
        new_id = popup.result
        new_ik = new_id.lower()
        success = self.add(new_id, self.selected_ID)
        if not success:
            return
        self._set_new_id_treeview_label(new_id, new_ik, popup.id_label)
        self.disable_paste()
        self.redo_tree_display()
        self.refresh_dropdowns()
        self.tree.scroll_to_item(new_ik)
        self.tree.selection_set(new_ik)
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def add_sibling_node(self):
        if not self.selected_ID:
            return
        sel = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)
        if sel:
            popup = Add_Child_Or_Sibling_Id_Popup(
                self,
                "sibling",
                self.selected_PAR,
                self.sheet.MT.data[sel[0]][self.ic],
                theme=self.C.theme,
            )
        else:
            popup = Add_Child_Or_Sibling_Id_Popup(self, "sibling", self.selected_PAR, None, theme=self.C.theme)
        if not popup.result:
            return
        new_id = popup.result
        new_ik = new_id.lower()
        success = self.add(new_id, self.selected_PAR)
        if not success:
            return
        self._set_new_id_treeview_label(new_id, new_ik, popup.id_label)
        self.disable_paste()
        self.redo_tree_display()
        self.refresh_dropdowns()
        self.tree.scroll_to_item(new_ik)
        self.tree.selection_set(new_ik)
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def sheet_add_top_node(self):
        popup = Add_Top_Id_Popup(self, None, theme=self.C.theme)
        if not popup.result:
            return
        new_id = popup.result
        new_ik = new_id.lower()
        if not self.sheet.anything_selected(exclude_columns=True):
            insert_row = len(self.sheet.MT.data)
        else:
            insert_row = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)[0]
        success = self.add(new_id, "", insert_row)
        if not success:
            return
        self._set_new_id_treeview_label(new_id, new_ik, popup.id_label)
        self.disable_paste()
        self.redo_tree_display()
        self.refresh_dropdowns()
        if self.tree.exists(new_ik):
            self.tree.scroll_to_item(new_ik)
            self.tree.selection_set(new_ik)
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def add_top_node(self):
        sel = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)
        if sel:
            popup = Add_Top_Id_Popup(self, self.sheet.MT.data[sel[0]][self.ic], theme=self.C.theme)
        else:
            popup = Add_Top_Id_Popup(self, None, theme=self.C.theme)
        if not popup.result:
            return
        new_id = popup.result
        new_ik = new_id.lower()
        success = self.add(new_id, "")
        if not success:
            return
        self._set_new_id_treeview_label(new_id, new_ik, popup.id_label)
        self.disable_paste()
        self.redo_tree_display()
        self.refresh_dropdowns()
        self.tree.scroll_to_item(new_ik)
        self.tree.selection_set(new_ik)
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def sheet_rename_node(self):
        rn = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)[0]
        id_ = self.sheet.MT.data[rn][self.ic]
        popup = Rename_Id_Popup(self, id_, theme=self.C.theme)
        if not popup.result:
            return
        tree_sel = self.tree.selection()[0] if self.tree.selection() else False
        success = self.change_ID_name(id_, popup.result)
        if not success:
            return
        self.reset_tagged_ids_dropdowns()
        self.disable_paste()
        self.rns = {r[self.ic].lower(): i for i, r in enumerate(self.sheet.data)}
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        if tree_sel:
            try:
                self.tree.scroll_to_item(tree_sel)
                self.tree.selection_set(tree_sel)
            except Exception:
                self.tree.scroll_to_item(popup.result.lower())
                self.tree.selection_set(popup.result.lower())
        else:
            self.move_tree_pos()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def rename_node(self):
        if not self.selected_ID:
            return
        popup = Rename_Id_Popup(self, self.selected_ID, theme=self.C.theme)
        if not popup.result:
            return
        success = self.change_ID_name(self.selected_ID, popup.result)
        if not success:
            return
        self.reset_tagged_ids_dropdowns()
        self.disable_paste()
        self.rns = {r[self.ic].lower(): i for i, r in enumerate(self.sheet.data)}
        self.refresh_formatting(rows=self.refresh_rows)
        self.redo_tree_display()
        self.refresh_rows = set()
        self.redraw_sheets()
        self.tree.scroll_to_item(popup.result.lower())
        self.tree.selection_set(popup.result.lower())
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def del_id(self, iids: Sequence[str]):
        if not iids:
            return
        self.start_work(f"Deleting {len(iids)} IDs")
        self.snapshot_delete_ids()
        self.session.delete(iids)
        self.sheet.deselect("all", redraw=False)
        self.refresh_formatting(rows=(self.rns[iid] for iid in self.refresh_rows if iid in self.rns))
        self.disable_paste()
        self.move_tree_pos()
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redo_tree_display()
        self.redraw_sheets()
        self.focus_tree()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def del_id_all(self, iids: Iterator[str] | None = None) -> None:
        if not iids:
            iids = self.tree.selection()
        self.start_work(f"Deleting {len(iids)} IDs")
        self.snapshot_delete_ids()
        self.session.delete(iids, all_hierarchies=True)
        self.sheet.deselect("all", redraw=False)
        self.refresh_formatting(rows=(self.rns[iid] for iid in self.refresh_rows if iid in self.rns))
        self.disable_paste()
        self.move_tree_pos()
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redo_tree_display()
        self.redraw_sheets()
        self.focus_tree()
        self.stop_work(self.get_tree_editor_status_bar_text())

    def del_id_orphan(self, event=None):
        if not self.selected_ID:
            return
        self.snapshot_delete_ids()
        self.sheet.deselect("all", redraw=False)
        self.disable_paste()
        self.session.delete([self.selected_ID], orphan=True)
        self.refresh_formatting(rows=(self.rns[iid] for iid in self.refresh_rows if iid in self.rns))
        self.move_tree_pos()
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redo_tree_display()
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.focus_tree()

    def _del_id_selection_roots(self, iids: Sequence[str]) -> list[str]:
        selected = []
        seen = set()
        for iid in iids:
            ik = iid.lower()
            if ik in self.nodes and ik not in seen:
                selected.append(ik)
                seen.add(ik)
        selected_set = set(selected)
        roots = []
        for ik in selected:
            p = self.nodes[ik].ps[self.pc]
            while p:
                if p in selected_set:
                    break
                p = self.nodes[p].ps[self.pc] if p in self.nodes else ""
            else:
                roots.append(ik)
        return roots

    def del_id_children(self, iids: Sequence[str] | None = None) -> None:
        if not iids:
            iids = self.tree.selection()
        if not iids and self.selected_ID:
            iids = (self.selected_ID,)
        if not iids:
            return
        self.start_work(f"Deleting {len(iids)} IDs and all children...")
        self.snapshot_delete_ids()
        self.sheet.deselect("all", redraw=False)
        self.disable_paste()
        self.session.delete(iids, children=True)
        self.refresh_formatting(rows=(self.rns[iid] for iid in self.refresh_rows if iid in self.rns))
        self.redo_tree_display()
        self.move_tree_pos()
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redraw_sheets()
        self.stop_work(self.get_tree_editor_status_bar_text())
        self.focus_tree()

    def del_id_children_all(self, iids: Sequence[str] | None = None) -> None:
        if not iids:
            iids = self.tree.selection()
        if not iids and self.selected_ID:
            iids = (self.selected_ID,)
        if not iids:
            return
        self.start_work(f"Deleting {len(iids)} IDs and all children...")
        self.snapshot_delete_ids()
        self.sheet.deselect("all", redraw=False)
        self.disable_paste()
        self.session.delete(iids, children=True, all_hierarchies=True)
        self.refresh_formatting(rows=(self.rns[iid] for iid in self.refresh_rows if iid in self.rns))
        self.redo_tree_display()
        self.move_tree_pos()
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redraw_sheets()
        self.stop_work(self.get_tree_editor_status_bar_text())
        self.focus_tree()

    def del_id_all_orphan(self):
        if not self.selected_ID:
            return
        self.snapshot_delete_ids()
        self.sheet.deselect("all", redraw=False)
        self.disable_paste()
        self.session.delete([self.selected_ID], orphan=True, all_hierarchies=True)
        self.refresh_formatting(rows=(self.rns[iid] for iid in self.refresh_rows if iid in self.rns))
        self.redo_tree_display()
        self.move_tree_pos()
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redraw_sheets()
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        self.focus_tree()

    def reselect_sheet_sel(self, boxes):
        self.sheet.deselect("all")
        if boxes[0]:
            self.sheet.set_currently_selected(boxes[0][0], boxes[0][1])
        for box in boxes[1]:
            r1, c1, r2, c2 = box[0]
            self.sheet.create_selection_box(r1, c1, r2, c2, box[1])

    def get_sheet_sel(self):
        return (
            self.sheet.get_currently_selected(),
            self.sheet.get_all_selection_boxes_with_types(),
        )

    def clear_copied_details(self):
        self.disable_paste()
        self.tree_rc_menu_single_row_paste.entryconfig("Paste details", state="disabled")
        self.sheet_rc_menu_single_row.entryconfig("Paste details", state="disabled")
        self.sheet_rc_menu_multi_row.entryconfig("Paste details", state="disabled")
        self.tree_rc_menu_multi_row.entryconfig("Paste details", state="disabled")
        self.copied_details = {"copied": [], "id": ""}
        return "break"

    def tree_sheet_edit_detail(self):
        if self.tree.has_focus():
            selected = self.tree.selected
            if not selected:
                return
            rn = self.rns[self.tree.rowitem(selected.row)]
        else:
            selected = self.sheet.selected
            if not selected:
                return
            rn = selected.row
        col = selected.column
        currentdetail = self.sheet.MT.data[rn][col]
        heading = self.headers[col].name
        ID = self.sheet.MT.data[rn][self.ic]
        if self.headers[col].type_ in ("ID", "Parent"):
            popup = Edit_Detail_Text_Popup(self, ID, heading, currentdetail, theme=self.C.theme)
        else:
            validation = self.headers[col].validation
            if validation:
                set_value = currentdetail if currentdetail in set(validation) else validation[0]
                popup = Edit_Detail_Text_Popup(
                    self,
                    ID,
                    heading,
                    currentdetail,
                    validation_values=validation,
                    set_value=set_value,
                    theme=self.C.theme,
                )
            else:
                popup = Edit_Detail_Text_Popup(self, ID, heading, currentdetail, theme=self.C.theme)
        if popup.result:
            self.tree_sheet_edit_table(
                event=DotDict(
                    sheetname="sheet",
                    data={(rn, col): popup.saved_string},
                )
            )

    def sheet_copy_details(self):
        rn = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)[0]
        ik = self.sheet.MT.data[rn][self.ic].lower()
        self.copied_details["copied"] = self.sheet.MT.data[rn].copy()
        self.copied_details["id"] = ik
        self.tree_rc_menu_single_row_paste.entryconfig("Paste details", state="normal")
        self.sheet_rc_menu_single_row.entryconfig("Paste details", state="normal")
        self.sheet_rc_menu_multi_row.entryconfig("Paste details", state="normal")
        self.tree_rc_menu_multi_row.entryconfig("Paste details", state="normal")
        s, writer = str_io_csv_writer(dialect=csv.excel_tab)
        writer.writerow(self.sheet.MT.data[rn])
        to_clipboard(widget=self, s=s.getvalue().rstrip())

    def copy_details(self):
        if not self.selected_ID:
            return
        rn = self.rns[self.selected_ID.lower()]
        self.copied_details["copied"] = self.sheet.MT.data[rn].copy()
        self.copied_details["id"] = self.selected_ID.lower()
        self.tree_rc_menu_single_row_paste.entryconfig("Paste details", state="normal")
        self.sheet_rc_menu_single_row.entryconfig("Paste details", state="normal")
        self.sheet_rc_menu_multi_row.entryconfig("Paste details", state="normal")
        self.tree_rc_menu_multi_row.entryconfig("Paste details", state="normal")
        s, writer = str_io_csv_writer(dialect=csv.excel_tab)
        writer.writerow(self.sheet.MT.data[rn])
        to_clipboard(widget=self, s=s.getvalue().rstrip())

    def sheet_paste_details(self):
        idcol_hiers = set(self.hiers) | {self.ic}
        event = DotDict(
            sheetname="sheet",
            data={
                (rn, c): e
                for rn in sorted(self.sheet.get_selected_rows())
                for c, e in enumerate(self.copied_details["copied"])
                if c not in idcol_hiers
            },
        )
        self.tree_sheet_edit_table(event=event)

    def paste_details(self):
        idcol_hiers = set(self.hiers) | {self.ic}
        event = DotDict(
            sheetname="sheet",
            data={
                (self.rns[iid], c): e
                for iid in self.tree.selection()
                for c, e in enumerate(self.copied_details["copied"])
                if c not in idcol_hiers
            },
        )
        self.tree_sheet_edit_table(event=event)

    def sheet_del_all_details(self):
        idcol_hiers = set(self.hiers) | {self.ic}
        event = DotDict(
            sheetname="sheet",
            data={
                (rn, c): ""
                for rn in sorted(self.sheet.get_selected_rows())
                for c in range(len(self.sheet.MT.data[rn]))
                if c not in idcol_hiers
            },
        )
        self.tree_sheet_edit_table(event=event)

    def del_all_details(self):
        idcol_hiers = set(self.hiers) | {self.ic}
        event = DotDict(
            sheetname="sheet",
            data={
                (self.rns[iid], c): ""
                for iid in self.tree.selection()
                for c in range(len(self.sheet.MT.data[self.rns[iid]]))
                if c not in idcol_hiers
            },
        )
        self.tree_sheet_edit_table(event=event)

    def tag_ids(
        self,
        event=None,
        selection: Iterator[str] | None = None,
        toggle: bool = True,
        do_tree: bool = True,
    ):
        if selection is None:
            if self.tree_has_focus:
                selection = self.tree.selection(cells=True)
            elif self.sheet_has_focus:
                selection = (
                    self.sheet.data[r][self.ic].lower() for r in self.sheet.get_selected_rows(get_cells_as_rows=True)
                )
            if not selection:
                return
        for iid in selection:
            if (ik := iid.lower()) not in self.rns:
                self.tagged_ids.discard(ik)
                continue
            rn = self.rns[ik]
            if toggle and ik in self.tagged_ids:
                self.tagged_ids.discard(ik)
                self.sheet.dehighlight_cells(
                    row=rn,
                    canvas="row_index",
                    redraw=False,
                )
                if do_tree and self.tree.exists(ik):
                    self.tree.dehighlight_cells(
                        row=self.tree.itemrow(ik),
                        canvas="row_index",
                        redraw=False,
                    )
            else:
                self.tagged_ids.add(ik)
                self.sheet.highlight_cells(
                    row=rn,
                    bg="orange",
                    fg="black",
                    canvas="row_index",
                    redraw=False,
                )
                if do_tree and self.tree.exists(ik):
                    self.tree.highlight_cells(
                        row=self.tree.itemrow(ik),
                        bg="orange",
                        fg="black",
                        canvas="row_index",
                        redraw=False,
                    )
        self.reset_tagged_ids_dropdowns()
        self.redraw_sheets()

    def tree_sheet_align(self, align):
        boxes = self.sheet.boxes if self.sheet.has_focus() else self.tree.boxes
        for box in boxes:
            if box.type_ == "columns":
                self.sheet.align(
                    f"{_n2a(box.coords.from_c)}:{_n2a(box.coords.upto_c - 1)}",
                    align=align,
                )
                self.tree.align(
                    f"{_n2a(box.coords.from_c)}:{_n2a(box.coords.upto_c - 1)}",
                    align=align,
                )

    def untag_id(self, ik):
        if ik in self.tagged_ids:
            self.tagged_ids.discard(ik)
            self.sheet.dehighlight_cells(row=self.rns[ik], canvas="row_index")
            if self.tree.exists(ik):
                self.tree.dehighlight_rows(self.tree.itemrow(ik))

    def clear_tagged_ids(self, event=None):
        self.session.clear_tags()
        self.reset_tagged_ids_dropdowns()
        self.sheet.dehighlight_cells(canvas="row_index", all_=True, redraw=True)
        self.redo_tree_display()

    def reset_tagged_ids_dropdowns(self, event=None):
        res = []
        for ik in tuple(self.tagged_ids):
            if ik in self.nodes:
                res.append(self.nodes[ik].name)
            else:
                self.tagged_ids.discard(ik)
        self.tree_tagged_ids_dropdown["values"] = res
        self.sheet_tagged_ids_dropdown["values"] = res
        if self.tagged_ids:
            self.sheet_tagged_ids_dropdown.set_my_value(res[0])
            self.tree_tagged_ids_dropdown.set_my_value(res[0])
        else:
            self.sheet_tagged_ids_dropdown.set_my_value("")
            self.tree_tagged_ids_dropdown.set_my_value("")
        return "break"

    def rehighlight_tagged_ids(self, event=None):
        self.sheet.dehighlight_cells(canvas="row_index", all_=True, redraw=False)
        self.tree.dehighlight_cells(canvas="row_index", all_=True, redraw=False)
        for ik in tuple(self.tagged_ids):
            try:
                self.sheet.highlight_cells(
                    row=self.rns[ik],
                    bg="orange",
                    fg="black",
                    canvas="row_index",
                    redraw=False,
                )
            except Exception:
                self.tagged_ids.discard(ik)
        for ik in filter(self.tree.RI.rns.__contains__, self.tagged_ids):
            self.tree.highlight_cells(
                row=self.tree.itemrow(ik),
                bg="orange",
                fg="black",
                canvas="row_index",
                redraw=False,
            )
        return "break"

    def go_to_treeview_id_finder(self, ik: str):
        hs = [self.headers[h].name for h, p in self.nodes[ik].ps.items() if p is not None]
        if len(hs) > 1:
            popup = Treeview_Id_Finder(self, hs, theme=self.C.theme)
            if not popup.GO:
                return
            selected = popup.selected
        else:
            selected = hs[0]
        self.switch_displayed.set(f"{selected}")
        self.switch_hier()
        self.tree.scroll_to_item(ik)
        self.tree.selection_set(ik)

    def tree_go_to_tagged_id(self, event=None):
        if not (ik := self.tree_tagged_ids_dropdown.get_my_value().lower()):
            return
        if ik in self.nodes:
            if self.tree.exists(ik):
                self.tree.scroll_to_item(ik)
                self.tree.selection_set(ik)
            else:
                self.go_to_treeview_id_finder(ik)
        else:
            Error(self, f"{ik} no longer exists and has been removed from tagged ids.", theme=self.C.theme)
            self.discard_tagged_id(ik)

    def discard_tagged_id(self, ik):
        self.tagged_ids.discard(ik)
        self.rns = {r[self.ic].lower(): i for i, r in enumerate(self.sheet.data)}
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.redraw_sheets()

    def sheet_go_to_tagged_id(self, event=None):
        if not (ik := self.sheet_tagged_ids_dropdown.get_my_value().lower()):
            return
        if ik in self.rns:
            self.sheet.select_row(self.rns[ik])
            self.sheet.see(row=self.rns[ik], keep_xscroll=True)
            if self.tree.exists(ik):
                self.tree.scroll_to_item(ik)
                self.tree.selection_set(ik)
        else:
            self.discard_tagged_id(ik)

    def find_next_main(self, event=None):
        if self.find_popup:
            self.find_popup.find_next()
        elif self.search_results and not self.sheet_search_results:
            self.find_next()
        elif self.sheet_search_results and not self.search_results:
            self.sheet_find_next()
        elif self.sheet_search_results and self.search_results and self.tree_has_focus:
            self.find_next()
        elif self.sheet_search_results and self.search_results and self.sheet_has_focus:
            self.sheet_find_next()

    def find_next(self, event=None):
        if self.search_dropdown["values"]:
            idx = self.search_dropdown.current()
            if idx + 1 == len(self.search_dropdown["values"]):
                idx = 0
            else:
                idx += 1
            self.search_dropdown.current(idx)
            self.show_search_result()

    def sheet_find_next(self, event=None):
        if self.sheet_search_dropdown["values"]:
            idx = self.sheet_search_dropdown.current()
            if idx + 1 == len(self.sheet_search_dropdown["values"]):
                idx = 0
            else:
                idx += 1
            self.sheet_search_dropdown.current(idx)
            self.sheet_show_search_result()

    def show_search_result(self, event=None):
        result = self.search_results[self.search_dropdown.current()]
        if not self.check_search_result(result):
            return
        if self.pc != result.hierarchy:
            self.switch_hier(hier=result.hierarchy)
        self.tree.scroll_to_item(result.iid, redraw=True)
        self.tree.selection_set(result.iid)
        self.tree.set_currently_selected(column=result.column, item=self.tree.selected.fill_iid)
        self.tree.see(column=result.column, keep_yscroll=True)
        self.focus_tree()

    def sheet_show_search_result(self, event=None):
        result = self.sheet_search_results[self.sheet_search_dropdown.current()]
        if not self.check_search_result(result):
            return
        self.sheet.select_row(row=self.rns[result.iid])
        self.sheet.see(row=self.rns[result.iid], column=result.column, redraw=True)
        self.sheet.set_currently_selected(column=result.column, item=self.sheet.selected.fill_iid)
        self.focus_sheet()

    def check_search_result(self, result: SearchResult) -> bool:
        result_ok = True
        if result.iid not in self.rns:
            result_ok = False
        else:
            sheet_rn = self.rns[result.iid]
            try:
                sheet_cell = self.sheet.data[sheet_rn][result.column].lower()
                if (
                    (result.hierarchy not in self.hiers)
                    or (result.exact and sheet_cell != result.term)
                    or (not result.exact and result.term not in sheet_cell)
                ):
                    result_ok = False
            except Exception:
                result_ok = False
        if not result_ok:
            Error(
                self,
                "Search result not found, refresh the search. Data may have been modified after searching.",
                theme=self.C.theme,
            )
        return result_ok

    def sheet_rc_tv_label(self, event=None):
        if (
            self.tree.has_focus()
            and self.tree.selected
            and self.tree.selected.type_ == "columns"
            or self.sheet.has_focus()
            and self.sheet.selected
            and self.sheet.selected.type_ == "columns"
        ):
            selected_col = self.tree.selected.column
        if self.headers[selected_col].type_ == "Parent":
            Error(
                self,
                "Cannot select Parent column as Treeview label",
                theme=self.C.theme,
            )
            return
        self.tv_label_col = selected_col
        self.save_info_get_saved_info()
        self.redo_tree_display()

    def set_all_col_widths(self, event=None):
        self.tree.set_all_cell_sizes_to_text()
        self.sheet.set_all_cell_sizes_to_text()

    def toggle_auto_resize_index(self, enabled):
        self.tree.set_options(auto_resize_row_index=enabled)
        self.sheet.set_options(auto_resize_row_index=enabled)
        self.auto_resize_indexes = enabled

    def toggle_mirror(self, enabled, select_row=True):
        self.mirror_var = enabled
        if select_row and self.mirror_var and self.tree.selection():
            self.go_to_row()

    def focus_tree(self):
        self.sheet.focus_set()
        self.sheet_focus_leave()
        self.tree_focus_enter()
        self.tree.focus_set()

    def focus_sheet(self):
        self.tree.focus_set()
        self.sheet_focus_enter()
        self.tree_focus_leave()
        self.sheet.focus_set()

    def tree_focus_leave(self, event=None):
        self.l_frame.config(
            highlightbackground=themes[self.C.theme].table_bg,
            highlightcolor=themes[self.C.theme].table_bg,
        )
        self.l_frame.update_idletasks()

    def tree_focus_enter(self, event=None):
        if self.get_display_option() in ("50/50", "adjustable"):
            self.l_frame.config(
                highlightbackground=themes[self.C.theme].table_selected_box_cells_fg,
                highlightcolor=themes[self.C.theme].table_selected_box_cells_fg,
                highlightthickness=2,
            )
        else:
            self.tree_focus_leave()
            self.l_frame.config(highlightthickness=0)
        self.l_frame.update_idletasks()
        self.tree_has_focus = True
        self.sheet_has_focus = False

    def sheet_focus_leave(self, event=None):
        self.r_frame.config(
            highlightbackground=themes[self.C.theme].table_bg,
            highlightcolor=themes[self.C.theme].table_bg,
        )
        self.r_frame.update_idletasks()

    def sheet_focus_enter(self, event=None):
        if self.get_display_option() in ("50/50", "adjustable"):
            self.r_frame.config(
                highlightbackground=themes[self.C.theme].table_selected_box_cells_fg,
                highlightcolor=themes[self.C.theme].table_selected_box_cells_fg,
                highlightthickness=2,
            )
        else:
            self.sheet_focus_leave()
            self.r_frame.config(highlightthickness=0)
        self.r_frame.update_idletasks()
        self.tree_has_focus = False
        self.sheet_has_focus = True

    def details_focus_set(self, event=None):
        if self.show_ids_details_dropdown.get_my_value() == "Treeview selection":
            self.focus_tree()
        else:
            self.focus_sheet()

    def show_ids_details_sheet(self, event=None):
        sel = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)
        if sel:
            View_Id_Popup(
                self,
                ids_row={"row": self.sheet.MT.data[sel[0]], "rn": sel[0]},
                theme=self.C.theme,
            )

    def show_ids_details_tree(self, event=None):
        if self.selected_ID:
            rn = self.rns[self.selected_ID.lower()]
            View_Id_Popup(
                self,
                ids_row={"row": self.sheet.MT.data[rn], "rn": rn},
                theme=self.C.theme,
            )

    def show_ids_full_info_sheet(self, event=None):
        sel = self.sheet.get_selected_rows(get_cells_as_rows=True, return_tuple=True)
        if sel:
            ik = self.sheet.MT.data[sel[0]][self.ic].lower()
            Text_Popup(self, self.details(ik), theme=self.C.theme)
        else:
            Error(self, "Select an ID in the sheet\nTo display the sheet go to View > Layout", theme=self.C.theme)

    def show_ids_full_info_tree(self, event=None):
        if self.selected_ID:
            Text_Popup(self, self.details(self.selected_ID.lower()), theme=self.C.theme)
        else:
            Error(self, "Select an ID in the tree", theme=self.C.theme)

    def show_warnings(self, filepath=None, sheetname=None, show_regardless=False):
        if filepath and sheetname:
            self.warnings_filepath = filepath
            self.warnings_sheet = sheetname
        top = f"File opened: {self.warnings_filepath}\nSheet opened: {self.warnings_sheet}\n\n"
        if show_regardless:
            if self.warnings:
                Text_Popup(
                    self,
                    f"{top}{warnings_header}\n" + "\n".join(self.warnings),
                    theme=self.C.theme,
                )
            else:
                Text_Popup(
                    self,
                    f"{top}{warnings_header}\n - NO WARNINGS TO DISPLAY - ",
                    theme=self.C.theme,
                )
        else:
            if self.warnings:
                Text_Popup(
                    self,
                    f"{top}{warnings_header}\n" + "\n".join(self.warnings),
                    theme=self.C.theme,
                )

    def show_changelog(self, event=None):
        if not isinstance(event, str) or event == "specific":
            Changelog_Popup(self, theme=self.C.theme)
        else:
            self.start_work("Opened save dialog")
            newfile = filedialog.asksaveasfilename(
                parent=self,
                title="Save changes as",
                filetypes=[
                    ("CSV File", ".csv"),
                    ("TSV File", ".tsv"),
                    ("Excel file", ".xlsx"),
                    ("JSON File", ".json"),
                ],
                defaultextension=".csv",
                confirmoverwrite=True,
            )
            if not newfile:
                self.stop_work()
                return
            newfile = os.path.normpath(newfile)
            if not newfile.lower().endswith((".csv", ".xlsx", ".json", ".tsv")):
                self.stop_work("Can only save .csv/.xlsx/.json file types")
                return
            self.C.status_bar.change_text("Saving changelog...")
            if event == "all":
                rows = display_rows(self.changelog)
            else:
                rows = display_rows(self.changelog[self.changelog_at_open :])
            try:
                if newfile.lower().endswith(".xlsx"):
                    self.C.wb = Workbook(write_only=True)
                    ws = self.C.wb.create_sheet(title="Changelog")
                    ws.append(xlsx_changelog_header(ws))
                    for row in rows:
                        ws.append(e if e else None for e in row)
                    self.C.wb.save(newfile)
                    self.C.try_to_close_workbook()
                elif newfile.lower().endswith((".csv", ".tsv")):
                    with open(newfile, "w", newline="", encoding="utf-8") as fh:
                        writer = csv.writer(
                            fh,
                            dialect=csv.excel_tab if newfile.lower().endswith(".tsv") else csv.excel,
                            lineterminator="\n",
                        )
                        writer.writerow(changelog_header)
                        writer.writerows(rows)
                elif newfile.lower().endswith(".json"):
                    with open(newfile, "w", newline="") as fh:
                        fh.write(
                            json.dumps(
                                full_sheet_to_dict(
                                    changelog_header,
                                    rows,
                                    include_headers=True,
                                    format_=self.json_format,
                                ),
                                indent=4,
                            )
                        )
            except Exception as error_msg:
                self.C.try_to_close_workbook()
                self.stop_work(f"Error saving file: {error_msg}")
                return
            self.stop_work("Success! Changelog saved")

    def go_to_row(self):
        if not self.selected_ID:
            return
        rn, sheet_rn = self.rns[self.selected_ID.lower()], self.sheet.selected
        if not sheet_rn or rn != sheet_rn.row or not self.sheet.cell_visible(sheet_rn.row, sheet_rn.column):
            self.sheet.select_row(rn, redraw=False)
            self.sheet.see(row=rn, keep_xscroll=True, redraw=True)
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())

    def zoom_in(self, event=None):
        self.tree.zoom_in()
        self.sheet.zoom_in()

    def zoom_out(self, event=None):
        self.tree.zoom_out()
        self.sheet.zoom_out()

    def expand_id(self, event=None):
        if current := self.tree.selected:
            selections = self.tree.selection()
            self.tree.tree_open({self.selected_ID.lower()} | set(self.tree.descendants(self.selected_ID.lower())))
            self.tree.selection_set(selections)
            self.tree.selected = current
        else:
            self.tree.tree_open()

    def collapse_id(self, event=None):
        if current := self.tree.selected:
            selections = self.tree.selection()
            current_iid = self.tree.tree_selected
            self.tree.tree_close({self.selected_ID.lower()} | set(self.tree.descendants(self.selected_ID.lower())))
            self.tree.selection_set(set(filter(self.tree.item_displayed, selections)))
            if self.tree.item_displayed(current_iid):
                self.tree.selected = current
        else:
            self.tree.tree_close()

    def move_tree_pos(self):
        self.tree.set_yview(self.saved_info[self.pc].scrolls.treey)
        self.tree.set_xview(self.saved_info[self.pc].scrolls.treex)

    def move_sheet_pos(self):
        self.sheet.set_yview(self.saved_info[self.pc].scrolls.sheety)
        self.sheet.set_xview(self.saved_info[self.pc].scrolls.sheetx)

    def refresh_tree_labels(self):
        if not self.tree.data:
            return
        row_index = self.tree.MT._row_index
        sheet = self.sheet.MT.data
        label_col = self.tv_label_col
        if self.tv_lvls_bool:
            nodes = self.nodes
            for node in row_index:
                label = sheet[self.rns[node.iid]][label_col]
                node.text = f"{self.get_node_level(nodes[node.iid])}. {label}"
        else:
            for node in row_index:
                node.text = sheet[self.rns[node.iid]][label_col]
        self.tree.set_refresh_timer(redraw=True)

    def refresh_tree_item(self, ID):
        iid = ID.lower()
        if self.tree.exists(iid):
            rn = self.rns[iid]
            highlights = {
                (rn, c): self.sheet.MT.cell_options[(rn, c)]["highlight"]
                for c in range(self.row_len)
                if (rn, c) in self.sheet.MT.cell_options and "highlight" in self.sheet.MT.cell_options[(rn, c)]
            }
            tree_row = self.tree.itemrow(iid)
            self.tree.dehighlight_cells(cells=[(tree_row, c) for c in range(self.row_len)])
            if highlights:
                for cell, highlight in highlights.items():
                    self.tree.highlight_cells(tree_row, cell[1], bg=highlight.bg, fg="black")
            r = self.sheet.MT.data[rn]
            if self.tv_lvls_bool:
                self.tree.item(
                    iid,
                    text=f"{self.get_node_level(self.nodes[iid])}. {r[self.tv_label_col]}",
                    values=r,
                )
            else:
                self.tree.item(
                    iid,
                    text=f"{r[self.tv_label_col]}",
                    values=r,
                )

    def redraw_sheets(self):
        self.sheet.set_refresh_timer()
        self.tree.set_refresh_timer()

    def reset_tree_search_dropdown(self):
        self.search_dropdown["values"] = []
        self.search_displayed.set("")
        self.search_results = []

    def reset_sheet_search_dropdown(self):
        self.sheet_search_dropdown["values"] = []
        self.sheet_search_displayed.set("")
        self.sheet_search_results = []

    def get_node_level(self, node, level=1):
        current_node = node
        current_level = level
        while True:
            if not current_node.ps[self.pc]:
                break
            current_node = self.nodes[current_node.ps[self.pc]]
            current_level += 1
        return current_level

    def redo_tree_display(self, selections=True):
        if self.saved_info[self.pc].twidths:
            self.tree.set_column_widths(self.tree_gen_widths_from_saved())
        else:
            self.tree.set_column_widths()
        self.selected_ID = ""
        self.selected_PAR = ""
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        if self.sheet.data:
            open_ids = self.saved_info[self.pc].opens if self.saved_info[self.pc].opens else None
            if self.tv_lvls_bool:
                data = []
                labels = []
                for iid in self.pc_iids():
                    data.append(self.sheet.data[self.rns[iid]])
                    labels.append(
                        f"{self.get_node_level(self.nodes[iid])}. {self.sheet.data[self.rns[iid]][self.tv_label_col]}"
                    )
                self.tree.tree_build(
                    data=data,
                    iid_column=self.ic,
                    parent_column=self.pc,
                    text_column=labels,
                    row_heights=False,
                    open_ids=open_ids,
                    safety=False,
                    ncols=self.row_len,
                    lower=True,
                ).dehighlight_all()
            else:
                self.tree.tree_build(
                    data=[self.sheet.data[self.rns[iid]] for iid in self.pc_iids()],
                    iid_column=self.ic,
                    parent_column=self.pc,
                    text_column=self.tv_label_col,
                    row_heights=False,
                    open_ids=open_ids,
                    safety=False,
                    ncols=self.row_len,
                    lower=True,
                ).dehighlight_all()
        else:
            self.tree.reset(cell_options=False, column_widths=False, header=False, redraw=False)

        if self.saved_info[self.pc].theights:
            self.tree.set_safe_row_heights(self.tree_gen_heights_from_saved())
        else:
            self.tree.set_row_heights()
        if selections:
            try:
                self.tree.boxes = self.saved_info[self.pc].boxes
                self.tree.selected = self.saved_info[self.pc].selected
            except Exception:
                self.saved_info[self.pc].boxes = ()
                self.saved_info[self.pc].selected = ()
        tree_rns = self.tree.RI.rns
        if self.tagged_ids:
            options = self.tree.RI.cell_options
            highlight = Highlight(
                bg="orange",
                fg="black",
                end=False,
            )
            for ik in filter(tree_rns.__contains__, self.tagged_ids):
                options[tree_rns[ik]] = {}
                options[tree_rns[ik]]["highlight"] = highlight
        options = self.tree.MT.cell_options
        sheet = self.sheet.MT.data
        numrows = len(sheet)
        out_of_bounds = []
        for cell, dct in self.sheet.MT.cell_options.items():
            if cell[0] >= numrows or cell[1] >= self.row_len:
                out_of_bounds.append(cell)
            elif "highlight" in dct and (iid := sheet[cell[0]][self.ic].lower()) in tree_rns:
                options[key := (tree_rns[iid], cell[1])] = {}
                options[key]["highlight"] = dct["highlight"]
        for cell in out_of_bounds:
            del self.sheet.MT.cell_options[cell]
        return "break"

    def get_clipboard_data(self, event=None):
        self.start_work("Loading data from clipboard...")
        self.new_sheet = []
        try:
            data = self.C.clipboard_get()
        except Exception as error_msg:
            Error(self, f"Error: {error_msg}", theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        try:
            if data.startswith("{") and data.endswith("}"):
                self.new_sheet = json_to_sheet(json.loads(data))
            else:
                self.new_sheet = csv_str_x_data(data)
        except Exception as error_msg:
            self.new_sheet = []
            self.stop_work(self.get_tree_editor_status_bar_text())
            Error(self, f"Error parsing clipboard data: {error_msg}", theme=self.C.theme)
            return
        if not self.new_sheet:
            self.new_sheet = []
            self.stop_work(self.get_tree_editor_status_bar_text())
            Error(self, "No data found on clipboard", theme=self.C.theme)
            return
        new_row_len = equalize_sublist_lens(self.new_sheet)
        self.C.status_bar.change_text(self.get_tree_editor_status_bar_text())
        popup = Get_Clipboard_Data_Popup(
            self,
            cols=self.new_sheet[0],
            row_len=new_row_len,
            theme=self.C.theme,
        )
        if not popup.result:
            self.new_sheet = []
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        new_row_len = equalize_sublist_lens(self.new_sheet)
        flattened = popup.flattened
        if flattened:
            hier_cols = popup.flattened_pcols
            if not hier_cols:
                return
        self.C.status_bar.change_text("Building tree...")
        self.snapshot_sheet("full overwrite")
        self.C.status_bar.change_text("Loading...   ")
        self.C.disable_at_start()
        self.warnings = []
        fmt = popup.format_selector_current
        if flattened:
            self.new_sheet, self.row_len, self.ic, self.hiers = TreeBuilder().convert_flattened_to_normal(
                data=self.new_sheet,
                hier_cols=hier_cols,
                rowlen=new_row_len,
                fmt=fmt,
                warnings=self.warnings,
            )
        elif fmt == 5:
            self.new_sheet, self.row_len, self.ic, self.hiers = (
                TreeBuilder().convert_indented_tree_detail_adjacent_to_normal(
                    data=self.new_sheet,
                )
            )
        elif fmt == 6:
            self.new_sheet, self.row_len, self.ic, self.hiers = (
                TreeBuilder().convert_indented_tree_details_adjacent_to_normal(
                    data=self.new_sheet,
                )
            )
        elif fmt == 7:
            self.new_sheet, self.row_len, self.ic, self.hiers = (
                TreeBuilder().convert_indented_tree_with_header_to_normal(
                    data=self.new_sheet,
                )
            )
        else:
            self.row_len = new_row_len
            self.ic = popup.ic
            self.hiers = popup.pcols
        self.selected_ID = ""
        self.selected_PAR = ""
        self.pc = int(self.hiers[0])
        self.tv_label_col = self.ic

        new_headers = self.fix_headers(self.new_sheet.pop(0), self.row_len)
        new_headers = [Header(name) for name in new_headers]
        new_headers[self.ic].type_ = "ID"
        for h in self.hiers:
            new_headers[h].type_ = "Parent"
        existing_headers = {h.name: i for i, h in enumerate(self.headers)}
        existing_col_alignments = {
            self.headers[c].name: align for c, align in self.sheet.get_column_alignments().items()
        }

        self.tree.reset()
        self.sheet.reset()

        for h in new_headers:
            if h.name in existing_headers and h.type_ == self.headers[existing_headers[h.name]].type_:
                h.formatting = self.headers[existing_headers[h.name]].formatting
            if h.name in existing_col_alignments:
                self.tree.align_columns(existing_headers[h.name], existing_col_alignments[h.name])
                self.sheet.align_columns(existing_headers[h.name], existing_col_alignments[h.name])

        self.headers = new_headers
        self.set_records(self.new_sheet)
        self.new_sheet = []
        self.saved_info = new_saved_info(self.hiers)
        self.clear_copied_details()
        self.cut = []
        self.copied = []
        self.cut_children_dct = {}
        self.sheet.set_xview(0.0)
        self.sheet.set_yview(0.0)
        self.auto_sort_nodes_bool = True
        built, nodes, warnings = TreeBuilder().build(
            input_sheet=self.sheet.MT.data,
            output_sheet=[],
            row_len=self.row_len,
            ic=self.ic,
            hiers=self.hiers,
            nodes={},
            warnings=self.warnings,
            add_warnings=True,
            strip=not self.allow_spaces_ids_var,
        )
        self.nodes = nodes
        self.warnings = warnings
        self.set_records(built)
        self.new_sheet = []
        self.fix_associate_sort(startup=True)
        self.set_headers()
        self.refresh_hier_dropdown(self.hiers.index(self.pc))
        self.rns = {r[self.ic].lower(): i for i, r in enumerate(self.sheet.data)}
        self.sheet.set_row_heights().set_column_widths().row_index(newindex=self.ic)
        self.tagged_ids = set(filter(self.rns.__contains__, self.tagged_ids))
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.refresh_formatting(dehighlight=True)
        self.redo_tree_display()
        self.refresh_dropdowns()
        self.changelog_append(
            "Overwrite sheet with clipboard data",
            "",
            "",
            "",
        )
        self.stop_work(self.get_tree_editor_status_bar_text())
        self.show_warnings("n/a - Data obtained from clipboard", "n/a")

    def import_changes(self):
        fp = filedialog.askopenfilename(parent=self, title="Select file")
        if not fp:
            return
        self.start_work("Import changes...")
        try:
            fp = os.path.normpath(fp)
        except Exception:
            Error(self, "Filepath invalid   ", theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        if not fp.lower().endswith((".tsv", ".csv", ".xls", ".xlsx", ".xlsm", ".json")):
            Error(self, "Invalid file format   ", theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        if not os.path.isfile(fp):
            Error(self, "Filepath invalid   ", theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        from .session import read_table

        loaded = read_table(fp)
        if not loaded["ok"]:
            Error(self, loaded["error"]["message"], theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        changes = loaded["result"]["rows"]
        if not changes:
            Error(self, "File contains no data   ", theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        self.snapshot_sheet()
        out = self.session.import_changes(changes, file_opened=fp)
        if not out["ok"]:
            if self.vs:
                self.vs.pop()
                self.set_undo_label()
            Error(self, out["error"]["message"], theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())
            return
        self._sync_sheet_from_session()
        self.pc = int(self.hiers[0])
        self.clear_copied_details()
        self.refresh_hier_dropdown(0)
        self.set_headers()
        self.sheet.deselect().set_column_widths().row_index(newindex=self.ic)
        self.reset_tagged_ids_dropdowns()
        self.rehighlight_tagged_ids()
        self.refresh_rows = set()
        self.refresh_formatting()
        self.redo_tree_display()
        self.refresh_dropdowns()
        self.stop_work(self.get_tree_editor_status_bar_text())
        result_rows = out["result"]["rows"]
        successful = [r["ok"] and r["reason"] is None for r in result_rows]
        applicable_changes = {
            "Edit cell",
            "Edit cell |",
            "Move rows",
            "Move columns",
            "Add new hierarchy column",
            "Add new detail column",
            "Delete hierarchy column",
            "Delete detail column",
            "Column rename",
            "Edit validation",
            "Change detail column type",
            "Date format change",
            "Cut and paste ID",
            "Cut and paste ID |",
            "Cut and paste ID + children",
            "Cut and paste ID + children |",
            "Cut and paste children",
            "Copy and paste ID",
            "Copy and paste ID |",
            "Copy and paste ID + children",
            "Copy and paste ID + children |",
            "Add ID",
            "Rename ID",
            "Delete ID",
            "Delete ID |",
            "Delete ID, orphan children",
            "Delete ID + all children",
            "Delete ID + all children |",
            "Delete ID + all children from all hierarchies",
            "Delete ID + all children from all hierarchies |",
            "Delete ID from all hierarchies",
            "Delete ID from all hierarchies |",
            "Delete ID from all hierarchies, orphan children",
            "Sort sheet",
        }
        applicable_changes = applicable_changes | {f"Imported change | {change}" for change in applicable_changes}
        shown = []
        flags = []
        for change, ok_row in zip(changes, successful):
            ctyp = change[1] if len(change) > 1 else ""
            if ctyp in applicable_changes or ctyp.startswith("Merge | "):
                shown.append(change)
                flags.append(ok_row)
        Post_Import_Changes_Popup(self, shown, flags, theme=self.C.theme)
        self.focus_tree()

    def add_rows_rc(self, insert=False):
        self.new_sheet = []
        popup = Merge_Sheets_Popup(self, theme=self.C.theme, add_rows=True)
        if not popup.result:
            self.new_sheet = []
            return
        self.merge_sheets(
            insert_row=min(self.sheet.get_selected_rows()) if insert else len(self.sheet.MT.data),
            popup_=popup,
        )

    def merge_sheets(self, insert_row=None, popup_=None):
        try:
            if popup_ is None:
                self.new_sheet = []
                popup = Merge_Sheets_Popup(self, theme=self.C.theme)
                if not popup.result:
                    self.new_sheet = []
                    return
            else:
                popup = popup_
            self.start_work("Merging sheets...")
            self.snapshot_sheet()
            fmt = popup.format_selector_current
            incoming = [list(r) for r in self.new_sheet]
            id_col = popup.ic if fmt == 0 else None
            if fmt == 0:
                parent_cols = popup.pcols
            elif fmt in (1, 2, 3, 4):
                parent_cols = popup.flattened_pcols
            else:
                parent_cols = None
            out = self.session.merge_from_rows(
                incoming,
                fmt=fmt,
                id_col=id_col,
                parent_cols=parent_cols,
                add_ids=popup.add_new_ids,
                add_dcols=popup.add_new_dcols,
                add_pcols=popup.add_new_pcols,
                overwrite_details=popup.overwrite_details,
                overwrite_parents=popup.overwrite_parents,
                insert_row=insert_row,
                file_opened=popup.file_opened or "",
            )
            self.new_sheet = []
            if not out["ok"]:
                if out["error"]["code"] != "no_changes" and self.vs:
                    self.vs.pop()
                    self.set_undo_label()
                Error(self, out["error"]["message"], theme=self.C.theme)
                self.stop_work(self.get_tree_editor_status_bar_text())
                self.focus_sheet()
                return
            self.clear_copied_details()
            for h in self.hiers:
                if h not in self.saved_info:
                    self.saved_info[h] = new_info_storage()
            self._sync_sheet_from_session()
            self.refresh_hier_dropdown(self.hiers.index(self.pc))
            self.sheet.deselect()
            self.set_headers()
            self.refresh_formatting()
            self.reset_tagged_ids_dropdowns()
            self.rehighlight_tagged_ids()
            self.redo_tree_display()
            self.refresh_dropdowns()
            self.show_warnings("n/a - Data imported from: " + popup.file_opened, popup.sheet_opened)
            self.stop_work(self.get_tree_editor_status_bar_text())
            self.focus_sheet()
            return
        except Exception as error_msg:
            Error(self, f"Error: {error_msg}", theme=self.C.theme)
            self.stop_work(self.get_tree_editor_status_bar_text())

    def get_par_lvls(self, h: int, iid: str, lvl=1):
        current_iid = iid
        current_level = lvl
        while self.nodes[current_iid].ps[h]:
            pid = self.nodes[current_iid].ps[h]
            self.levels[current_level] = self.nodes[pid].name
            current_iid = pid
            current_level += 1

    def export_flattened(self, event=None):
        self.start_work("Flattening sheet...")
        self.new_sheet = []
        Export_Flattened_Popup(self, theme=self.C.theme)
        self.stop_work(self.get_tree_editor_status_bar_text())
        self.new_sheet = []

    def get_save_json(self):
        d = full_sheet_to_dict(
            [h.name for h in self.headers],
            self.sheet.MT.data,
            format_=self.json_format,
        )
        if self.save_json_with_program_data:
            d["version"] = software_version_number
            d["changelog"] = flatten_changelog(self.changelog)[0]
            d["program_data"] = dict_x_b32(self.get_program_data_dict())
        return d

    def get_program_data_dict(self, sheetname="n/a"):
        d = self.session.program_data_dict(sheetname)
        d["row_heights"] = self.sheet.get_safe_row_heights()
        d["column_widths"] = self.sheet.get_column_widths()
        d["sheet_column_alignments"] = self.sheet.get_column_alignments()
        d["sheet_table_align"] = self.sheet.table_align()
        d["sheet_header_align"] = self.sheet.header_align()
        d["sheet_index_align"] = self.sheet.index_align()
        d["saved_info"] = self.save_info_get_saved_info()
        d["tv_label_col"] = self.tv_label_col
        return d

    def jsonify_nodes(self):
        return self.session.jsonify_nodes()

    def nodes_json_x_dict(self, njson: dict, hiers: Sequence[int]) -> dict:
        return self.session.nodes_json_x_dict(njson, hiers)

    def xlsx_chunker(self, seq):
        size = min(len(seq), 32000)
        return (seq[pos : pos + size] for pos in range(0, len(seq), size))

    def write_program_data_to_workbook(self, wb, sheetnames_):
        with suppress(Exception):
            wb.remove(wb["program_data"])
        ws = wb.create_sheet(title="program_data")
        ws.append([f"{software_version_number}"])
        for chunk in self.xlsx_chunker(dict_x_b32(self.get_program_data_dict(sheetnames_[1]))):
            ws.append([chunk])
        ws.sheet_state = "hidden"

    def write_changelog_to_workbook(self, wb, sheetnames_):
        sheetname = sheetnames_[1]
        new_title1 = sheetname + " Changelog"
        for sname in (i for i in wb.sheetnames if "Changelog" in i):
            try:
                wb.remove(wb[sname])
            except Exception:
                continue
        ws = wb.create_sheet(title=new_title1)
        ws.append(xlsx_changelog_header(ws))
        for r in reversed(display_rows(self.changelog)):
            ws.append(e if e else None for e in r)

    def write_flattened_to_workbook(self, wb, sheetnames_):
        sheetname = sheetnames_[1]
        new_title1 = sheetname + " Flattened"
        for sname in (i for i in wb.sheetnames if "flattened" in i or "Flattened" in i):
            try:
                wb.remove(wb[sname])
            except Exception:
                continue
        ws = wb.create_sheet(title=new_title1)
        ws.freeze_panes = "A2"
        self.new_sheet = []
        for r in TreeBuilder().build_flattened(
            input_sheet=self.sheet.MT.data,
            output_sheet=self.new_sheet,
            nodes=self.nodes,
            headers=[f"{hdr.name}" for hdr in self.headers],
            ic=int(self.ic),
            pc=int(self.pc),
            hiers=list(self.hiers),
            detail_columns=self.xlsx_flattened_detail_columns,
            justify_left=self.xlsx_flattened_justify,
            reverse=self.xlsx_flattened_reverse_order,
            add_index=self.xlsx_flattened_add_index,
        ):
            ws.append(e if e else None for e in r)
        self.new_sheet = []

    def write_treeview_to_workbook(self, wb: Workbook, sheetnames_):
        sheetname = sheetnames_[1]
        new_title1 = sheetname + " Treeview"
        for sname in (i for i in wb.sheetnames if "Treeview" in i):
            try:
                wb.remove(wb[sname])
            except Exception:
                continue
        ws = wb.create_sheet(title=new_title1)
        ws.freeze_panes = "A2"
        oldpc = int(self.pc)
        self.levels = defaultdict(list)
        for h in self.hiers:
            for iid, node in self.nodes.items():
                if node.ps[h] and not node.cn[h]:
                    self.get_par_lvls(h, iid)
        maxlvls = max(self.levels, default=0) + 1
        self.xl_tv_detail_cols = tuple(i for i, h in enumerate(self.headers) if h.type_ not in ("ID", "Parent"))
        self.level_colors = tuple(tv_lvls_colors[level_to_color(i)] for i in range(maxlvls + 1))
        cycle_colors = cycle(tv_lvls_colors)

        # Write header row
        row = []
        for lvl in range(1, maxlvls + 1):
            cell = WriteOnlyCell(ws, value=lvl)
            cell.fill = next(cycle_colors)
            row.append(cell)
        for _, hdr in enumerate((h for h in self.headers if h.type_ not in ("ID", "Parent")), start=maxlvls + 1):
            cell = WriteOnlyCell(ws, value=hdr.name)
            row.append(cell)
        ws.append(row)

        # Process hierarchies iteratively
        for h in self.hiers:
            self.pc = int(h)  # Ensure self.pc is an integer as expected elsewhere
            self.hier_disp = f"{self.headers[self.pc].name} - "

            # Initialize stack with top nodes at level 1
            stack = [(iid, 1) for iid in self.top_iids()]

            # Process nodes using the stack
            while stack:
                iid, level = stack.pop()
                # Construct row with indentation
                row = list(repeat(None, level - 1))
                cell = WriteOnlyCell(
                    ws,
                    value=f"{self.hier_disp}{self.sheet.MT.data[self.rns[iid]][self.tv_label_col]}",
                )
                cell.fill = self.level_colors[level - 1]
                row.append(cell)
                # Add None values up to detail columns
                num_nones = maxlvls - level
                row.extend(list(repeat(None, num_nones)))
                # Add detail columns
                for col in self.xl_tv_detail_cols:
                    cell = WriteOnlyCell(ws, value=self.sheet.MT.data[self.rns[iid]][col])
                    row.append(cell)
                ws.append(row)
                # Push children in reverse order to maintain original order
                for ciid in reversed(self.nodes[iid].cn[self.pc]):
                    stack.append((ciid, level + 1))

        # Restore state
        self.pc = int(oldpc)
        self.levels = defaultdict(list)

    def write_additional_sheets_to_workbook(self, new_sheet_name=None):
        if self.save_xlsx_with_flattened:
            self.C.status_bar.change_text("Saving flattened sheet...")
            self.write_flattened_to_workbook(
                self.C.wb,
                (
                    self.C.open_dict["sheet"],
                    self.C.open_dict["sheet"] if new_sheet_name is None else new_sheet_name,
                ),
            )
        if self.save_xlsx_with_changelog:
            self.C.status_bar.change_text("Saving changelog...")
            self.write_changelog_to_workbook(
                self.C.wb,
                (
                    self.C.open_dict["sheet"],
                    self.C.open_dict["sheet"] if new_sheet_name is None else new_sheet_name,
                ),
            )
        if self.save_xlsx_with_treeview:
            self.C.status_bar.change_text("Saving treeview...")
            self.write_treeview_to_workbook(
                self.C.wb,
                (
                    self.C.open_dict["sheet"],
                    self.C.open_dict["sheet"] if new_sheet_name is None else new_sheet_name,
                ),
            )
        if self.save_xlsx_with_program_data:
            self.C.status_bar.change_text("Saving program data...")
            self.write_program_data_to_workbook(
                self.C.wb,
                (
                    self.C.open_dict["sheet"],
                    self.C.open_dict["sheet"] if new_sheet_name is None else new_sheet_name,
                ),
            )
        self.C.status_bar.change_text("Writing file...")

    def save_workbook(self, filepath, sheetname):
        self.C.wb = Workbook(write_only=True)
        ws = self.C.wb.create_sheet(title=sheetname)
        if not self.ic:
            ws.freeze_panes = "B2"
        else:
            ws.freeze_panes = "A2"
        for row in self.gen_sheet_w_headers():
            ws.append(row)
        self.write_additional_sheets_to_workbook(sheetname)
        self.C.wb.active = self.C.wb[sheetname]
        filepath = convert_old_xl_to_xlsx(filepath)
        self.C.wb.save(filepath)
        self.C.open_dict["filepath"] = filepath
        self.C.change_app_title(title=os.path.basename(filepath))
        return True

    def save_csv(self, filepath):
        with open(filepath, "w", newline="", encoding="utf-8") as fh:
            writer = csv.writer(
                fh,
                dialect=csv.excel_tab if filepath.lower().endswith(".tsv") else csv.excel,
                lineterminator="\n",
            )
            writer.writerows(self.gen_sheet_w_headers())
        self.C.open_dict["filepath"] = filepath
        self.C.change_app_title(title=os.path.basename(filepath))
        self.C.open_dict["sheet"] = "Sheet1"
        return True

    def save_json(self, filepath):
        with open(filepath, "w") as fh:
            fh.write(json.dumps(self.get_save_json(), indent=4))
        self.C.open_dict["filepath"] = filepath
        self.C.change_app_title(title=os.path.basename(filepath))
        self.C.open_dict["sheet"] = "Sheet1"
        return True

    def save_(self, event=None, quitting=False):
        if self.C.current_frame != "tree_edit":
            return False
        newfile = os.path.normpath(self.C.open_dict["filepath"])
        self.start_work("Saving... ")
        successful = False
        try:
            if newfile.lower().endswith((".csv", ".tsv")):
                successful = self.save_csv(newfile)
            elif newfile.lower().endswith(".json"):
                successful = self.save_json(newfile)
            elif newfile.lower().endswith((".xlsx", ".xls", ".xlsm")):
                successful = self.save_workbook(newfile, self.C.open_dict["sheet"])
                self.C.try_to_close_workbook()
        except Exception as error_msg:
            Error(self, f"Error: {error_msg}", theme=self.C.theme)
        if successful:
            self.C.created_new = False
            self.bind_or_unbind_save("normal")
            self.C.unsaved_changes = False
        if not (quitting and successful):
            self.stop_work(self.get_tree_editor_status_bar_text(), resume_quit=not quitting)
        return successful

    def save_as(self, event=None, quitting=False):
        if self.C.current_frame != "tree_edit":
            return False
        newfile = filedialog.asksaveasfilename(
            parent=self.C,
            title="Save as",
            filetypes=[
                ("Excel file", ".xlsx"),
                ("JSON file", ".json"),
                ("CSV File (Comma separated values)", ".csv"),
                ("TSV File (Tab separated values)", ".tsv"),
            ],
            defaultextension=".xlsx",
            confirmoverwrite=True,
        )
        if not newfile:
            return False
        newfile = os.path.normpath(newfile)
        if not newfile.lower().endswith((".json", ".csv", ".xlsx", ".tsv")):
            Error(self, "Can only write .json, .xlsx or .csv    ", theme=self.C.theme)
            return False
        self.start_work("Saving... ")
        successful = False
        try:
            if newfile.lower().endswith((".csv", ".tsv")):
                successful = self.save_csv(newfile)
            elif newfile.lower().endswith(".json"):
                successful = self.save_json(newfile)
            elif newfile.lower().endswith(".xlsx"):
                popup = Enter_Sheet_Name_Popup(self, theme=self.C.theme)
                if popup.result and all(
                    reserved not in popup.result.lower()
                    for reserved in (
                        "program_data",
                        " treeview",
                        " changelog",
                        " flattened",
                    )
                ):
                    successful = self.save_workbook(newfile, popup.result)
                else:
                    error_msg = "Enter a sheet name / sheet name must not be equal to 'program data'   "
                    Error(self, error_msg, theme=self.C.theme)
        except Exception as error_msg:
            Error(self, f"Error: {error_msg}", theme=self.C.theme)
        if successful:
            self.C.created_new = False
            self.bind_or_unbind_save("normal")
            self.C.unsaved_changes = False
        if not (quitting and successful):
            self.stop_work(self.get_tree_editor_status_bar_text(), resume_quit=not quitting)
        return successful

    def save_new_vrsn(self, event=None):
        newfile = self.C.open_dict["filepath"]
        folder = os.path.dirname(newfile)
        if not folder:
            folder = os.path.dirname(os.path.abspath(__file__))
        popup = Save_New_Version_Presave_Popup(self, folder, theme=self.C.theme)
        if not popup.result:
            return False
        self.start_work("Saving... ")
        folder = popup.result
        newfile = os.path.join(folder, os.path.basename(newfile))
        if not newfile.lower().endswith((".csv", ".xls", ".tsv", ".xlsx", ".json", ".xlsm")):
            Error(
                self,
                "Error saving file, file extension must be .csv/.xlsx/.json   ",
                theme=self.C.theme,
            )
            self.stop_work(self.get_tree_editor_status_bar_text())
            return False
        newfile_without_numbers = os.path.basename(path_without_numbers(newfile))
        matches = {}
        found_suitable_folder = False
        while not found_suitable_folder:
            try:
                matches = {}
                for file in os.listdir(folder):
                    if (
                        file.lower().endswith((".json", ".xlsx", ".csv", ".xls", ".xlsm", ".tsv"))
                        and path_without_numbers(file) == newfile_without_numbers
                    ):
                        matches[file] = path_numbers(file)
                found_suitable_folder = True
            except Exception:
                popup = Save_New_Version_Error_Popup(self, theme=self.C.theme)
                if popup.result:
                    folder = os.path.normpath(
                        filedialog.askdirectory(
                            parent=self.C,
                            title="Select a folder to save new version in",
                        )
                    )
                    if folder == ".":
                        self.stop_work(self.get_tree_editor_status_bar_text())
                        return False
                    newfile = os.path.join(folder, os.path.basename(newfile))
                else:
                    self.stop_work(self.get_tree_editor_status_bar_text())
                    return False
        if matches:
            latest_num = float("-inf")
            latest_name = None
            for k, v in matches.items():
                if v > latest_num:
                    latest_num = v
                    latest_name = k
            newfile = os.path.join(folder, increment_file_version(latest_name))
        else:
            newfile = increment_file_version(newfile)
        x = 0
        while os.path.isfile(newfile):
            newfile = increment_file_version(newfile)
            x += 1
            if x > 200:
                Error(
                    self,
                    "Error saving file, could not get name for new version   ",
                    theme=self.C.theme,
                )
                self.stop_work(self.get_tree_editor_status_bar_text())
                return False
        successful = False
        try:
            if newfile.lower().endswith((".csv", ".tsv")):
                successful = self.save_csv(newfile)
            elif newfile.lower().endswith(".json"):
                successful = self.save_json(newfile)
            elif newfile.lower().endswith((".xlsx", ".xls", ".xlsm")):
                successful = self.save_workbook(newfile, self.C.open_dict["sheet"])
                self.C.try_to_close_workbook()
        except Exception as error_msg:
            Error(self, f"Error: {error_msg}", theme=self.C.theme)
        if successful:
            self.C.unsaved_changes = False
            popup = Save_New_Version_Postsave_Popup(self, folder, os.path.basename(newfile), theme=self.C.theme)
        self.stop_work(self.get_tree_editor_status_bar_text())
        return successful
