from __future__ import annotations

import re
import tkinter as tk
from collections.abc import Callable
from typing import Any, Literal

from .constants import ctrl_key, rc_binding
from .functions import recursive_bind
from .other_classes import DotDict


class FindWindowTkText(tk.Text):
    """Custom Text widget for the FindWindow class."""

    def __init__(
        self,
        parent: tk.Misc,
    ) -> None:
        super().__init__(
            parent,
            spacing1=0,
            spacing2=1,
            spacing3=0,
            bd=0,
            highlightthickness=0,
            undo=True,
            maxundo=30,
        )
        self.parent = parent
        self.rc_popup_menu = tk.Menu(self, tearoff=0)
        self.bind("<1>", lambda event: self.focus_set())
        self.bind(rc_binding, self.rc)
        self.bind(f"<{ctrl_key}-a>", self.select_all)
        self.bind(f"<{ctrl_key}-A>", self.select_all)
        self.bind("<Delete>", self.delete_key)

    def reset(
        self,
        menu_kwargs: dict,
        sheet_ops: dict,
        font: tuple,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
    ) -> None:
        """Reset the text widget's appearance and menu options."""
        self.config(
            font=font,
            background=bg,
            foreground=fg,
            insertbackground=fg,
            selectbackground=select_bg,
            selectforeground=select_fg,
        )
        self.editor_del_key = sheet_ops.editor_del_key
        self.rc_popup_menu.delete(0, "end")
        self.rc_popup_menu.add_command(
            label=sheet_ops.select_all_label,
            accelerator=sheet_ops.select_all_accelerator,
            command=self.select_all,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.cut_label,
            accelerator=sheet_ops.cut_accelerator,
            command=self.cut,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.copy_label,
            accelerator=sheet_ops.copy_accelerator,
            command=self.copy,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.paste_label,
            accelerator=sheet_ops.paste_accelerator,
            command=self.paste,
            **menu_kwargs,
        )
        self.rc_popup_menu.add_command(
            label=sheet_ops.undo_label,
            accelerator=sheet_ops.undo_accelerator,
            command=self.undo,
            **menu_kwargs,
        )

    def rc(self, event: Any) -> None:
        """Show the right-click popup menu."""
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root, event.y_root)

    def delete_key(self, event: Any = None) -> None:
        """Handle the Delete key based on editor configuration."""
        if self.editor_del_key == "forward":
            return
        elif not self.editor_del_key:
            return "break"
        elif self.editor_del_key == "backward":
            if self.tag_ranges("sel"):
                return
            if self.index("insert") == "1.0":
                return "break"
            self.delete("insert-1c")
            return "break"

    def select_all(self, event: Any = None) -> Literal["break"]:
        """Select all text in the widget."""
        self.tag_add(tk.SEL, "1.0", tk.END)
        self.mark_set(tk.INSERT, tk.END)
        return "break"

    def cut(self, event: Any = None) -> Literal["break"]:
        """Cut selected text."""
        self.event_generate(f"<{ctrl_key}-x>")
        self.event_generate("<KeyRelease>")
        return "break"

    def copy(self, event: Any = None) -> Literal["break"]:
        """Copy selected text."""
        self.event_generate(f"<{ctrl_key}-c>")
        return "break"

    def paste(self, event: Any = None) -> Literal["break"]:
        """Paste text from clipboard."""
        self.event_generate(f"<{ctrl_key}-v>")
        self.event_generate("<KeyRelease>")
        return "break"

    def undo(self, event: Any = None) -> Literal["break"]:
        """Undo the last action."""
        self.event_generate(f"<{ctrl_key}-z>")
        self.event_generate("<KeyRelease>")
        return "break"


class FindWindow(tk.Frame):
    """A frame containing find and replace functionality."""

    def __init__(
        self,
        parent: tk.Misc,
        find_next_func: Callable,
        find_prev_func: Callable,
        close_func: Callable,
        replace_func: Callable,
        replace_all_func: Callable,
        toggle_replace_func: Callable,
    ) -> None:
        super().__init__(
            parent,
            width=0,
            height=0,
            bd=0,
        )
        # Configure grid: column 1 for text widgets, both rows expandable
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(4, uniform="group1")
        self.grid_columnconfigure(5, uniform="group2")
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self.grid_propagate(False)
        self.parent = parent

        # Store functions for use in handle_return
        self.find_next_func = find_next_func
        self.replace_func = replace_func

        # Toggle label to show/hide replace window
        self.toggle_replace = tk.Label(self, text="↓", cursor="hand2", highlightthickness=1)
        self.toggle_replace.grid(row=0, column=0, sticky="ns")
        self.toggle_replace.bind("<Button-1>", self.toggle_replace_window)
        self.toggle_replace.bind("<Enter>", lambda e: self.enter_label(widget=self.toggle_replace))
        self.toggle_replace.bind("<Leave>", lambda e: self.leave_label(widget=self.toggle_replace))
        self.toggle_replace_func = toggle_replace_func

        # Find text widget
        self.tktext = FindWindowTkText(self)
        self.tktext.grid(row=0, column=1, sticky="nswe")

        # Find action labels
        self.find_previous_arrow = tk.Label(self, text="▲", cursor="hand2", highlightthickness=1)
        self.find_previous_arrow.bind("<Button-1>", find_prev_func)
        self.find_previous_arrow.grid(row=0, column=2)

        self.find_next_arrow = tk.Label(self, text="▼", cursor="hand2", highlightthickness=1)
        self.find_next_arrow.bind("<Button-1>", find_next_func)
        self.find_next_arrow.grid(row=0, column=3)

        self.find_in_selection = False
        self.in_selection = tk.Label(self, text="🔎", cursor="hand2", highlightthickness=1)
        self.in_selection.bind("<Button-1>", self.toggle_in_selection)
        self.in_selection.grid(row=0, column=4)

        self.close = tk.Label(self, text="✕", cursor="hand2", highlightthickness=1)
        self.close.bind("<Button-1>", close_func)
        self.close.grid(row=0, column=5, sticky="nswe")

        # Replace text widget, initially hidden
        self.replace_tktext = FindWindowTkText(self)
        self.replace_tktext.grid(row=1, column=1, columnspan=4, sticky="nswe")
        self.replace_tktext.grid_remove()

        # Replace action labels, initially hidden
        self.replace_next = tk.Label(self, text="→", cursor="hand2", highlightthickness=1)
        self.replace_next.bind("<Button-1>", replace_func)
        self.replace_next.grid(row=1, column=4, sticky="nswe")
        self.replace_next.grid_remove()

        self.replace_all = tk.Label(self, text="→*", cursor="hand2", highlightthickness=1)
        self.replace_all.bind("<Button-1>", replace_all_func)
        self.replace_all.grid(row=1, column=5, sticky="nswe")
        self.replace_all.grid_remove()

        # Bind Tab for focus switching
        self.tktext.bind("<Tab>", self.handle_tab)
        self.replace_tktext.bind("<Tab>", self.handle_tab)

        # Bind Return for find/replace actions
        self.tktext.bind("<Return>", self.handle_return)
        self.replace_tktext.bind("<Return>", self.handle_return)

        # Bind hover events for all action labels
        for widget in (
            self.find_previous_arrow,
            self.find_next_arrow,
            self.in_selection,
            self.close,
            self.toggle_replace,
            self.replace_next,
            self.replace_all,
        ):
            widget.bind("<Enter>", lambda e, w=widget: self.enter_label(widget=w))
            widget.bind("<Leave>", lambda e, w=widget: self.leave_label(widget=w))

        # State variables
        self.replace_visible = False
        self.bg = None
        self.fg = None

        # Existing bindings for in-selection toggle
        for b in ("Option", "Alt"):
            for c in ("l", "L"):
                recursive_bind(self, f"<{b}-{c}>", self.toggle_in_selection)

    def handle_tab(self, event):
        """Handle Tab key presses to switch focus between find and replace text widgets."""
        if not self.replace_visible:
            self.toggle_replace_window()
        if event.widget == self.tktext:
            self.replace_tktext.focus_set()
        elif event.widget == self.replace_tktext:
            self.tktext.focus_set()
        return "break"

    def handle_return(self, event):
        """Handle Return key presses to trigger find next or replace next based on focus."""
        if event.widget == self.tktext:
            self.find_next_func()
        elif event.widget == self.replace_tktext:
            self.replace_func()
        return "break"

    def toggle_replace_window(self, event: tk.Misc = None) -> None:
        """Toggle the visibility of the replace window and update the toggle label."""
        if self.replace_visible:
            self.replace_tktext.grid_remove()
            self.replace_next.grid_remove()
            self.replace_all.grid_remove()
            self.toggle_replace.config(text="↓")
            self.toggle_replace.grid(row=0, column=0, rowspan=1, sticky="ns")
            self.replace_visible = False
        else:
            self.replace_tktext.grid(row=1, column=1, columnspan=4, sticky="nswe")
            self.replace_next.grid(row=1, column=4, sticky="nswe")
            self.replace_all.grid(row=1, column=5, sticky="nswe")
            self.toggle_replace.config(text="↑")
            self.toggle_replace.grid(row=0, column=0, rowspan=2, sticky="ns")
            self.replace_visible = True
        self.toggle_replace_func()

    def enter_label(self, widget: tk.Misc) -> None:
        """Highlight label on hover."""
        widget.config(
            highlightbackground=self.fg,
            highlightcolor=self.fg,
        )

    def leave_label(self, widget: tk.Misc) -> None:
        """Remove highlight when not hovering, unless toggled."""
        if widget == self.in_selection and self.find_in_selection:
            return
        widget.config(
            highlightbackground=self.bg,
            highlightcolor=self.fg,
        )

    def toggle_in_selection(self, event: tk.Misc) -> None:
        """Toggle the in-selection state."""
        self.find_in_selection = not self.find_in_selection
        self.enter_label(self.in_selection)
        self.leave_label(self.in_selection)

    def get(self) -> str:
        """Get the find text."""
        return self.tktext.get("1.0", "end-1c")

    def get_replace(self) -> str:
        """Get the replace text."""
        return self.replace_tktext.get("1.0", "end-1c")

    def get_num_lines(self) -> int:
        """Get the number of lines in the find text."""
        return int(self.tktext.index("end-1c").split(".")[0])

    def set_text(self, text: str = "") -> None:
        """Set the find text."""
        self.tktext.delete(1.0, "end")
        self.tktext.insert(1.0, text)

    def reset(
        self,
        border_color: str,
        menu_kwargs: DotDict,
        sheet_ops: DotDict,
        bg: str,
        fg: str,
        select_bg: str,
        select_fg: str,
    ) -> None:
        """Reset styles and configurations."""
        self.bg = bg
        self.fg = fg
        self.tktext.reset(
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            font=menu_kwargs.font,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
        )
        self.replace_tktext.reset(
            menu_kwargs=menu_kwargs,
            sheet_ops=sheet_ops,
            font=menu_kwargs.font,
            bg=bg,
            fg=fg,
            select_bg=select_bg,
            select_fg=select_fg,
        )
        for widget in (
            self.find_previous_arrow,
            self.find_next_arrow,
            self.in_selection,
            self.close,
            self.toggle_replace,
            self.replace_next,
            self.replace_all,
        ):
            widget.config(
                font=menu_kwargs.font,
                bg=bg,
                fg=fg,
                highlightbackground=bg,
                highlightcolor=fg,
            )
        if self.find_in_selection:
            self.enter_label(self.in_selection)
        self.config(
            background=bg,
            highlightbackground=border_color,
            highlightcolor=border_color,
            highlightthickness=1,
        )


def replacer(find: str, replace: str, current: str) -> Callable[[re.Match], str]:
    """
    Creates a replacement function for re.sub with special empty string handling.

    Parameters:
        find (str): String to search for. If empty, behavior varies.
        replace (str): String to replace matches with.
        current (str): Input string where replacements occur.

    Returns:
        Callable[[re.Match], str]: Function taking a match object, returning a str.

    Behavior:
        - If `find` is non-empty, returns `replace` for each match.
        - If `find` is empty:
            - Returns `replace` if `current` is empty.
            - Returns matched string (empty) if `current` is non-empty, no change.
    """

    def _replacer(match: re.Match) -> str:
        if find:
            return replace  # Normal replacement when find is non-empty
        else:
            # Special case when find is empty
            if len(current) == 0:
                return replace  # Return "hello" when current is empty
            else:
                return match.group(0)  # Preserve current when non-empty

    return _replacer
