# TKTREES DOCUMENTATION

TkTrees is an app for managing hierarchy data.

To start the app, use Python 3.9 or newer to run the file named "TKTREES.pyw". tkinter (which usually comes with Python installations on Windows) is required.

After starting the app or opening a file, a treeview shows the items (IDs) on the left, and their respective rows to their right. To change the view, such as to view the underlying table, go to View -> Layout or the File menu then Settings.

TkTrees is licensed under AGPL-3.0 and is the copyright of R. A. Gardner.

- Author: github.com/ragardner
- Source Code: github.com/ragardner/tktrees
- Contact Email: github@ragardner.simplelogin.com

---

# PROGRAM BASICS

If using the program for the first time you can go to the "File" menu or right click in the empty space to get started. File -> Open loads an existing file.

Words used here:

- ID: the item itself (a name or code)
- Parent: the ID it sits under
- Hierarchy: one parent column. Extra parent columns are extra hierarchies for the same IDs
- Detail: extra info on an ID, such as a name or note
- Flattened: levels spread across columns, usually one row per path from top to bottom

A tree is just items nested under other items. For example:

| ID      | Parent  | Name        |
|---------|---------|-------------|
| Animals |         | All animals |
| Cats    | Animals | Cat family  |
| Lion    | Cats    | Lion        |

The examples below use Top, Mid and Base for the same idea, so the file shapes are easy to compare.

Please note:

- Save writes a new file. It does not update the old file in place. If you save over an existing file, that file is replaced. Other sheets, charts, macros, pictures and anything else that was in it are gone. Only what TkTrees writes remains.
- Delete in this hierarchy removes the ID from the hierarchy you are viewing. If that was its last appearance, the ID is removed completely.
- Delete in all hierarchies removes the ID everywhere.
- If you save with app data, the hidden app sheet is what loads next time, not later edits you make in the visible Excel sheet.
- Undo does not survive closing the file. The changelog can be saved; the undo history cannot.

Supported file formats are:

- .xlsx, .xls, .xlsm
- .json JavaScript object notation where the full table is under the key 'records'
- .csv/.tsv (comma or tab delimited)

The following data formats are supported for loading:

| Data Structure                    | Header Requirement   | Details Requirement |
|-----------------------------------|----------------------|---------------------|
| ID, Parent                        | Must have header     | Optional, Unlimited |
| Top -> Base                       | Must have header     | Optional, Unlimited |
| Top -> Base (Unique Details)      | Must have header     | Optional, Unlimited |
| Base -> Top                       | Must have header     | Optional, Unlimited |
| Base -> Top (Unique Details)      | Must have header     | Optional, Unlimited |
| Level-Indent Columns              | Must NOT have header | Optional, Max. 1    |
| Level-Indent Columns Multi-Detail | Must NOT have header | Optional, Unlimited |
| Level-Indent Columns with Header  | Must have header     | Optional, Unlimited |

Unique details means each level keeps its own detail columns, instead of one shared detail column.

**Example: ID, Parent**

| ID    | Parent    | Detail           |
|-------|-----------|------------------|
| Top   |           | Top Description  |
| Mid   | Top       | Mid Description  |
| Base  | Mid       | Base Description |

**Example: Top -> Base**

| Level0 | Description0    | Level1 | Description1    | Level2 | Description2     |
|--------|-----------------|--------|-----------------|--------|------------------|
| Top    | Top Description | Mid    | Mid Description | Base   | Base Description |

**Output Example: Top -> Base with normal Mode ->**

| ID    | Parent    | Description0     |
|-------|-----------|------------------|
| Top   |           | Top Description  |
| Mid   | Top       | Mid Description  |
| Base  | Mid       | Base Description |

**Output Example: Top -> Base with Unique Details Mode ->**

| ID    | Parent    | Description0     | Description1    | Description2     |
|-------|-----------|------------------|-----------------|------------------|
| Top   |           | Top Description  |                 |                  |
| Mid   | Top       |                  | Mid Description |                  |
| Base  | Mid       |                  |                 | Base Description |

**Example: Base -> Top**

| Level2 | Description2     | Level1 | Description1    | Level0 | Description0    |
|--------|------------------|--------|-----------------|--------|-----------------|
| Base   | Base Description | Mid    | Mid Description | Top    | Top Description |

**Example: Level-Indent Columns**

|     |                 |                 |                  |
|-----|-----------------|-----------------|------------------|
| Top | Top Description |                 |                  |
|     | Mid             | Mid Description |                  |
|     |                 | Base            | Base Description |

**Example: Level-Indent Columns Multi-Detail**

|     |                   |                   |                    |                    |
|-----|-------------------|-------------------|--------------------|--------------------|
| Top | Top Description 1 | Top Description 2 |                    |                    |
|     | Mid               | Mid Description 1 | Mid Description 2  |                    |
|     |                   | Base              | Base Description 1 | Base Description 2 |

**Example: Level-Indent Columns with Header**

| Level0 | Level1 | Level2 | Description 1      | Description 2      |
|--------|--------|--------|--------------------|--------------------|
| Top    |        |        | Top Description 1  | Top Description 2  |
|        | Mid    |        | Mid Description 1  | Mid Description 2  |
|        |        | Base   | Base Description 1 | Base Description 2 |

**Notes:**

- Additional settings and data such as the changelog, formatting and column types can be saved with the formats .xlsx and .json.
- There is no limit to the number of characters allowed for headers, details or ID names. To allow spaces in ID/Header names go to File -> Settings on the main menubar while in the Treeview. Details are exempt from this rule.
- Any mistakes in the sheet such as infinite loops of children, IDs appearing in a parent column but not in the ID column and duplications will be corrected upon creating the tree.
- The corrections will not be made to the original sheet unless you choose to save the sheet. Such corrections will appear as warnings when you first view the treeview window.
- Upon opening a file if an ID has no parents or children in any hierarchy it will be placed in the first hierarchy (in order of the columns).

---

# HELPFUL TIPS AND TUTORIALS

#### Adding an ID

To add a single ID in the treeview:

1. Right click an existing ID.
2. Go to Add and choose Add child (under that ID), Add sibling (next to it), or Add top ID (at the top, with no parent).
3. Type the new ID name and confirm.

On the sheet you can also right click a row and choose Add top ID.

To add many IDs at once, see Guides -> Merge sheets.

#### Adding a column

Right click a column header and choose Add detail or Add hierarchy.

- Add detail: extra information such as a name or note. You pick a name for the column.
- Add hierarchy: another parent column, so the same IDs can sit in a second tree.

The new column is inserted where you right clicked, or at the end if you did not right click a header.

Formatting is under Managing Columns.

#### Renaming an ID

Right click the ID in the tree or sheet and choose Rename ID, then type the new name. This changes the ID everywhere it appears.

#### Editing a detail

Double click a cell to edit it. Confirm with the cell empty to clear it. Right click and choose Edit if you want a larger window.

#### Looking at a different hierarchy

Use the Hierarchy dropdown at the top of the tree. Each parent column is a different hierarchy.

#### Finding an ID or detail

Use the Find box at the top of the tree or the sheet. Choose whether to search for an ID or a detail. Tree results are only for the hierarchy you are viewing. Search is not case sensitive.

When the tree or sheet has focus, Ctrl + F opens a find and replace window. This searches the cells you can see, and can also replace text. Ctrl + H shows the replace box if it is hidden.

#### Moving IDs between hierarchies

To move an ID to another hierarchy or add an ID to another hierarchy:

1. Right click on the ID in the treeview panel and go to Cut or Copy and then either Detach ID or Copy ID.
2. Then using the dropdown box labeled "Hierarchy" at the top of the treeview panel select the hierarchy you would like to move / add the ID to.
3. Go to the position or ID where you would like to place the Detached / Copied ID and right click and select a paste option.

To move multiple IDs in one go you can use Shift + Left Click or Ctrl + Left Click to select multiple IDs then use Ctrl + X (Cut) or Ctrl + C (Copy) or Right Click on one of the selected IDs.

#### Moving IDs by drag and drop

You can move IDs that are at the same level as one another around in a specific hierarchy by using the mouse to drag and drop:

1. Selecting the IDs by left clicking and holding the mouse button down.
2. Moving the mouse to drag IDs from their existing locations to a new location.
3. Release the mouse button to drop them.

If any dragged IDs are on different levels from one another then they will not be included in the move.

#### Deleting IDs

- When using Delete on an ID in the sheet panel or Delete in all hierarchies in the treeview panel it will delete an ID completely; across all hierarchies.
- When using any other delete option it will only delete an ID in the currently selected hierarchy. However, if that ID is the last appearance of the ID across all hierarchies then it will completely delete it, just like with Delete in all hierarchies.

#### Deleting a column

- To delete a column right click on the column you wish to delete and select Delete column. Note you cannot delete a parent column if it is the only parent column in the sheet and you cannot delete a parent column if you are currently viewing it.

#### Getting all information on an ID

- An easy way to get an ID's complete information within the sheet, including parents and children across all hierarchies and all details is to select an ID in the treeview or sheet panel and then go to View -> Treeview IDs information or View -> Sheet IDs information.
- You can also get a more concise view of an ID by right clicking on it and selecting ID concise view.

#### Changing the order of IDs in the treeview

To disable automatic ordering of IDs in the treeview go to:

1. The File menu then Settings.
2. Select Auto-sort tree IDs.

You can re-order children by selecting a single row in the tree and dragging using the left mouse button. To move an ID between parents see the above section on "Moving IDs between hierarchies".

---

# MANAGING COLUMNS

Right clicking on columns in the header will show a popup menu with a few column specific options.

Detail columns hold text. You can still restrict a column to a list of allowed values with Validation.

#### Conditional Formatting:

You can add conditional formatting to columns, meaning when certain conditions are met the cells in that column will be filled with a chosen color. You can set a maximum of 35 conditions.

Conditions are text matching: the cell must contain exactly the user input. Conditions are not case sensitive.

---

# GUIDES

#### Changelog

Every change you make is recorded. Open the list with View -> View changelog, Export -> Export specific changes, or Ctrl + L.

The list has five columns: date, type, what was changed, old value (red), new value (green).

From that window you can:

- Export all: save the whole list as .csv, .tsv, .xlsx or .json
- Export selected as: save only the rows you have selected
- Prune up to selected: delete from the start of the list through the selected row. If that row is part of a grouped change (the type ends with |), pruning continues to the end of the group. This can be undone.

Two other export menu items skip the window:

- Export file session changes: only changes made since this file was opened
- Export all changes: the whole list, straight to a file

The changelog can be stored with app data, and you can also save a viewable changelog sheet (see XLSX Files). Undo does not survive closing the file, but the changelog can be saved.

#### Import changes

Import -> Import changes replays a saved changelog on the file you have open. Use this to apply the same edits to another file, or to replay an exported list.

The file must be .csv, .tsv, .xlsx, .xlsm, .xls or .json. For Excel, only the first sheet is read. The table must have exactly five columns, the same as an exported changelog.

Lines that already start with "Imported change |" or "Merge |" are treated as the action after that prefix, so you can export and import the same list again.

Each row is tried on its own. A change is applied only if the sheet still matches what the row expects, for example:

- The column still exists, with the same name and type
- The ID still exists
- For a cell edit, the current value is still the old value in the row
- For a move or delete, the parent is still the parent recorded in the row
- Detail values still pass that column's validation

If the new value is already what the sheet has, that row is counted as unnecessary, not as a failure.

When it finishes, a window lists the rows that were tried. Green applied, red did not. The status line shows how many succeeded. It does not say why a row failed.

The whole import can be undone as one step.

#### Merge sheets

Import -> Merge Sheets / Add rows combines another table with the open file.

You can open a file, paste from the clipboard, or type in the sheet on the right. Pick the same kind of file shape you would when opening a file, and set the ID and parent columns if asked. Opening a file resets the merge sheet.

Options (on unless you turn them off):

- Add new IDs: IDs that are not in the open file are added
- Add new detail columns / Add new parent columns: columns whose names are not already in the open file are added. New detail columns are Text.
- Overwrite details / Overwrite parents: for IDs that exist in both, copy values from columns with the same name

IDs and column names are matched without caring about case. A detail is only written if it is valid for that column's type. After the merge the tree is rebuilt, so parent values on new IDs are applied then.

If nothing applies you get "No applicable changes were made". The merge can be undone.

Right-click insert row uses the same window, already showing the sheet so you can paste extra rows.

Import -> Paste Clipboard & Overwrite Sheet replaces the whole open sheet with clipboard data. That can be undone. It is not a merge.

#### Export flattened sheet

Export -> Export flattened sheet opens a window with levels across columns. The open file is not changed.

Pick which hierarchy (parent column) to flatten. Then:

- Include detail columns
- Justify left
- Reverse order: bottom of the tree on the left, top on the right
- Add index column
- Remove End IDs: drop that many levels from the end of each path

View -> Show Detail Excluder lets you leave some detail columns out.

File -> Save As writes .xlsx, .csv, .tsv or .json. Edit has copy as tab-separated, comma-separated or json.

Saving the main file can also add a flattened sheet. That uses File -> Settings -> xlsx Flatten Settings, not this window. If you are viewing all hierarchies when you save, the first hierarchy is the one written.

#### Tag IDs

Edit -> Tag/Untag IDs, Ctrl + T, or the Tag ID button. Works on the current selection in the tree or the sheet. Tagging the same ID again removes the tag.

Tagged IDs get an orange mark on the row index and appear in the dropdowns at the top. Pick one to jump to it. If that ID is in more than one hierarchy, you choose which.

Tags are stored with app data when you save .xlsx or .json.

Edit -> Clear all tagged IDs cannot be undone.

#### Tag IDs using list

Edit -> Tag IDs using list. One column of terms. Empty cells are ignored.

Find IDs and Find Details control which main-sheet cells are searched. At least one of those stays on. Exact Match requires a full cell match; otherwise the term can appear anywhere in the cell. All matching is not case sensitive. Parent columns are not searched.

Load a file, paste from the clipboard, or type in the mini table. Tag adds matches to tagged IDs (it does not untag). After Tag, terms that found something are green at the top, terms that found nothing are red below. Empty rows stay at the bottom. The status line says how many terms matched and how many IDs were tagged, for example `8/12 terms matched, 15 IDs tagged`.

#### Delete IDs using list

Edit -> Delete IDs using list. One column of IDs. Empty cells are ignored.

Load a file, paste from the clipboard, or type in the mini table. Delete IDs and Delete IDs + children use the hierarchy you are viewing. Delete IDs all hierarchies removes the listed IDs everywhere. There is no list version of delete IDs + children from all hierarchies; that stays on the tree right-click menu, where you can see which descendants you are removing. After a delete, the status line says how many of the listed IDs were deleted, for example `5/10 ids deleted`.

#### Replace using mapping

Edit -> Replace using mapping. Two columns: find (not case sensitive) and replace with. It runs on the whole sheet.

Load a file, paste from the clipboard, or type in the mini table. After you click Replace, the status line says how many cells changed.

#### Save new version

File -> Save new version writes a new file next to the current one (you pick the folder). It looks for other files with the same name and a number on the end, then uses a higher number. If the name has no number, one is added.

This still writes a new file. It does not update an old Excel workbook in place. See XLSX Files.

---

# TREE BUTTONS

In the tree panel:

1. Find: Clicking the find button will attempt to find either an ID or detail.
    - This depends on which is selected in the drop-down box on the right of "Find".
    - The drop-down box below the Find button will display any results found within the CURRENTLY viewed hierarchy.
    - All finds are not case sensitive, including "exact match".
2. Hierarchy: This is the drop-down box where you can select which parent column/hierarchy to view.
3. Tag ID: tags the selection. Tagged IDs show up in the dropdown next to the button. See Guides -> Tag IDs.

In the sheet panel:

- Tagged IDs (Ctrl + T): same tagging as in the tree. See Guides -> Tag IDs.
- Find: Works the same way that the Find button for the Tree panel works except it searches the sheet instead.

---

# TREE FUNCTIONS

By right clicking on an ID in the tree panel you can select various functions. The main functions are Detach, Copy and Delete.

To detach or copy an ID between different hierarchies:

- Right click on the ID and select whichever option you want then switch hierarchy and right click in empty space or on the ID you want to paste the detached/copied ID to as a sibling or child.
- If you want to paste an ID as an ID without a parent right click on a top ID and choose paste as sibling.
- You can also detach all of an ID's children, including grandchildren and so on, and paste them under where you right click.
- Using shift click you can select multiple up or down of an existing selection. Using Ctrl click you can make multiple selections.
- When using the Ctrl X, C and V keys to cut/copy and paste they will work on the selected ID, not on the position where the mouse is hovering (unless pasting over empty space using Ctrl V).
- Cutting and copying using this method will only perform on IDs that are on the same level as the top most (index-wise) ID, after pressing Ctrl X or C it will deselect any selections that were not cut or copied.

ID Deletion:

- Pressing the Delete key on multiple selections will work the same way, except performing a Delete immediately. The delete key uses the typical Delete ID function, not deleting its children.
- In the tree panel there are 5 delete ID options. Delete ID only removes the ID from the hierarchy you're currently viewing IF the ID occurs in another hierarchy, if it does not then it totally removes the ID.
- Del all of ID totally removes the ID. Del ID+children is the same as Delete ID but for every child and child of that child and so on recursively under the selected ID.

Editing cells:

- You can quickly edit a detail by double clicking on the detail/cell you want to edit. To delete a detail press Confirm when editing a detail with the cell empty.
- Right clicking in a cell and selecting edit will pop up a larger window so the text may be easier to view.
- Pasting a detail or details will work between both panels. You can drag and drop rows in the sheet panel to change their order.
- When using drag and drop you can use your mousewheel to scroll down, move the mouse a little after scrolling to cause the selection to move.

---

# TREE COMPARE

File -> Compare sheets. Two panels, left and right. You can mix file types.

1. Open a file on each side. Opening another file on that side resets it.
2. For Excel without app data, pick the sheet and click Load sheet. If the workbook has a program_data sheet, that is used and you skip the sheet picker.
3. Set the ID column and at least one parent column on each side. An ID column cannot also be a parent column.
4. Create Report. You can save the report as .xlsx.

The report can include:

- Warnings from building each tree
- Different ID column index or name
- Parent or detail columns that exist on only one side, or in different positions
- IDs that exist on only one side
- Different parents or details on IDs that exist on both sides

IDs are matched without caring about case. If nothing differs, the header says the sheets are identical.

---

# XLSX FILES

The default save format is .xlsx.

Save writes a new workbook. It does not open the old Excel file and update it. If you save over an existing .xlsx, the whole file is replaced. Other sheets, charts, macros, pictures and anything else that was in that file are gone. Only what TkTrees writes remains.

When saving .xlsx files you can also save program data to keep your changelog, row heights, column widths, formatting, validation, treeview ID order and more. That is stored on a sheet named program_data.

When loading a file saved with program data the sheet and changelog in the program data, not the visible sheet, will take precedence. This means any edits in the viewable sheet will not be loaded.

To disable saving with program data go to File -> Settings -> xlsx save options -> Save with app data.

You can also add a viewable changelog sheet, a tree sheet, and a flattened sheet for the currently viewed hierarchy. If viewing all hierarchies when saving then the first hierarchy will be saved.

When comparing or merging if the workbook contains program data then it will take precedence, else a sheet will need to be selected to load data.

---

# JSON FILES

There are four loadable json formats, with each one the entire sheet is kept under the key "records". However the program will also look for the keys: sheet, data and table. The first format, also the first option under "File -> Settings -> json save options -> json format" is displayed as an example below:

A dictionary of key (column header) and value (list of column cells)
```
{"records":
    {
        "ID":
                    [
                     "ID_1",
                     "ID_2"
                    ],
        "DETAIL_1":
                    [
                     "",
                     ""
                    ],
        "PARENT_1":
                    [
                     "ID_1s_Parent",
                     "ID_2s_Parent"
                    ]
    }
}
```


The second json format option example is displayed below:

A list of dictionaries (rows) where inside each dictionary the key is the header and the
value is the cell
```
{
 "records": [
        {
         "ID":       "ID_1",
         "DETAIL_1": "",
         "PARENT_1": "ID1s_Parent"
         },
        {
         "ID":       "ID1s_Parent",
         "DETAIL_1": "",
         "PARENT_1": ""
         }
            ]
}
```

The third json format option is displayed below:

A list of lists (rows) where each row simply contains values that are the cells
```
{
 "records":
    [
        [
         "ID",
         "DETAIL_1",
         "PARENT_1"
        ],
        [
         "ID_1",
         "",
         "ID_1s_parent"
        ]
    ]
}
```

The fourth json format option is displayed below:

A tab delimited csv stored as a string under the key 'records', this format is really non-
typical so only use it if you really need to.
```
{
 "records":
    "ID\\tDetail-1\\tParent-1\\nID_1\\t\\tID_1s_Parent"
}
```

Program data is only included if Save is used as opposed to Copy to clipboard. It is in the following format:
```
{
    "version": "1.00",
    "records": <full sheet including headers stored here>,
    "changelog": [],
    "program_data": "base32string"
}
```

---

# BUNDLED LIBRARIES

**openpyxl**

- Version: 3.1.5
- Original source: https://foss.heptapod.net/openpyxl
- License: MIT License
- Full license text and conditions: See `openpyxl/LICENCE.rst`
- Authors and copyright holders: See `openpyxl/AUTHORS.rst`
- Note: This library is bundled to handle Excel file operations.

**defusedxml**

- Version: 0.7.1
- Original source: https://github.com/tiran/defusedxml
- License: Python Software Foundation License (PSFL)
- Full license text and conditions: See `defusedxml/LICENSE`
- Copyright: Copyright (c) 2013-2023 by Christian Heimes
- Note: This library is bundled for secure XML parsing.

**tksheet**

- Original source: https://github.com/ragardner/tksheet
- License: MIT License
- Full license text and conditions: See `tksheet/LICENSE.txt`
- Copyright: Copyright (c) ragardner
- Note: This library is bundled for tkinter sheet/table functionality.

# USING THE API

The API reads a file, flattens or unflattens it, and writes a new file. It does not open the GUI.

How `TKTREES.pyw` starts:

- No arguments: GUI
- A data file path (`.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv`, `.json`), or any existing file: GUI
- `flatten` or `unflatten` (also `--help` / `--version`): file-to-file API. Never the GUI
- `cli`: reserved for a future interactive command line. Not available yet. Does not open the GUI
- Anything else, including typos: error, exit `2`, does not open the GUI

```
python TKTREES.pyw --help
python TKTREES.pyw flatten --help
python TKTREES.pyw unflatten --help
```

If any value has spaces, quote it, e.g. `"New Sheet"`.

The input file must be .xlsx, .xls, .xlsm, .csv, .tsv or .json.

The output file must be .csv, .tsv, .xlsx or .json. Suffix case is ignored when deciding the type (`.CSV` is csv). The filename is used as given.

#### Commands

`flatten` turns an ID/parent table into levels across columns.

`unflatten` turns a levels-across-columns table back into ID, parent, details.

#### Required for both

- Input filepath
- Output filepath
- `--parents` every parent/hierarchy column in the file, comma-separated indexes or letters. `0` is the first column. Examples: `--parents 2,3` or `--parents C,D`

#### Required only for `flatten`

- `--id` ID column, e.g. `--id 0` or `--id A`
- `--parent` parent column of the hierarchy to flatten, e.g. `--parent 2` or `--parent C`

#### Required only for `unflatten`

- `--order top-base` or `--order base-top`. You have to pass this; the API will not guess. Look at the input file: `top-base` means the root is on the left, `base-top` means the leaf is on the left. If `--order` does not match the file, every parent comes out the wrong way around.

#### Optional

- `--input-sheet NAME` xlsx input sheet. Default: first sheet
- `--output-sheet NAME` xlsx output sheet. Default: input sheet name, or Sheet1
- `--delim CHAR` delimiter for csv/tsv output. Default: comma, or tab if the output filepath ends with `.tsv`. `--delim` overrides that. Examples: `--delim tab`, `--delim ,`, `--delim "|"`
- `--order top-base` or `--order base-top` for **flatten** output direction. Default `top-base`. `base-top` is the same as GUI Reverse order.

If the delimiter is a shell special character such as `|`, `;`, `>` or `&`, quote it e.g. `--delim "|"`. Without quotes the shell treats those characters as operators and they never reach the program.

Omitted flags are off.

| Flag        | Used for                        | Applies to         |
|-------------|---------------------------------|--------------------|
| --overwrite | Replace output if it exists     | flatten, unflatten |
| --details   | Include detail columns          | flatten            |
| --justify   | Pack shorter paths to the left  | flatten            |
| --index     | Add an index column             | flatten            |
| --order     | top-base (default) or base-top  | flatten            |
| --unique    | One detail column set per level | unflatten          |

`--justify` packs shorter paths to the left, with details to the right of each ID.

Without `--overwrite`, the API fails if the output file already exists (it does not replace it). Errors are written to `TKTREES-ERROR.txt` next to the app, and the process exits `1`. Bad arguments exit `2` and do not open the GUI.

Examples:

Flatten an xlsx hierarchy at column C, output top-base, with details and justify, overwriting if needed:
```
python TKTREES.pyw flatten "input filepath here.xlsx" "output filepath here.xlsx" --parents C,D --id A --parent C --input-sheet Sheet1 --output-sheet "New Sheet" --details --justify --overwrite
```

Unflatten a file whose hierarchy columns are left to right Top -> Base:
```
python TKTREES.pyw unflatten "input filepath here.csv" "output filepath here.csv" --parents 0,2,4,6 --order top-base --delim tab --overwrite
```

Unflatten a file whose hierarchy columns are left to right Base -> Top:
```
python TKTREES.pyw unflatten "input filepath here.csv" "output filepath here.csv" --parents 0,2,4,6 --order base-top --overwrite
```

Unflatten with unique details (each level keeps its own detail columns):
```
python TKTREES.pyw unflatten "input filepath here.csv" "output filepath here.csv" --parents 0,2,4,6 --order top-base --unique --overwrite
```
