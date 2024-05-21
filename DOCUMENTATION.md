# TK-TREES DOCUMENTATION

This document is in markdown format.

Tk-Trees is an app for management of hierarchy data in table format. It was written in the Python programming language and utilizes the following libraries:
 - tkinter, tksheet, openpyxl

Tk-Trees is licensed under GPL-3.0 and is the copyright of R. A. Gardner.
github.com/ragardner
github@ragardner.simplelogin.com

## PROGRAM BASICS

If using the program for the first time you can go to the "File" menu or right click in the left or right panel to get started.

This program is for management of hierarchy based master data which is stored in table format. Supported file formats are:
- .xlsx, .xls, .xlsm
- .json Javascript object notation where the full table is under the key 'records'
- .csv/.tsv (comma or tab delimited)

Any sheets opened with tk-trees must contain a single header row at the top of the sheet.

Additional settings and data such as the changelog, formatting and column types can be saved with the formats .xlsx and .json.

Any changes made when using this program will not be saved unless you choose to do so. This can be done by going to the main menubar and choosing a save option under the 'File' menu while in the Treeview.

Sheets must have an ID column and atleast one parent column, it does not matter in which order. e.g.

```
ID     Parent    Detail
ID1    Par1      Detail 1
ID2    Par2      Detail 2
```

Sheets can have multiple parent columns (hierarchies) and multiple detail columns but must have only one ID column. In the ID column each ID must be unique else they will be renamed with '_DUPLICATE_'.

Sheets can have an unlimited number of parent columns (hierarchies) and an unlimited number of detail columns.

The columns can be in any order and multiple columns of the same type can be separated by other types of columns.

If the headers are not unique they will be renamed with a duplicate number. Any missing headers will have names created for them. Header names are not case sensitive.

There is no limit to the number of characters allowed for headers, details or ID names. To allow spaces in ID/Header names go to Options on the main menubar while in the Treeview. Details are exempt from this rule.

Any mistakes in the sheet such as infinite loops of children, IDs appearing in a parent column but not in the ID column and duplications will be corrected upon creating the tree.

The corrections will not be made to the original sheet unless you choose to save the sheet. Such corrections will appear as warnings when you first view the treeview window.

To display the actual sheet go to View -> Layout.

If an ID has no parents or children in any hierarchy it will be placed in the first hierarchy (in order of the columns). The same will happen if you Undo a change made to anything but details.

## HELPFUL TIPS AND TUTORIALS

#### Moving IDs between hierarchies

To move an ID to another hierarchy or add an ID to another hierarchy you can right click on the ID in the treeview panel and go to Cut or Copy and then either Cut ID or Copy ID. Then using the dropdown box at the top of the treeview panel, labeled "Hierarchy" select the hierarchy you would like to move / add the ID to. Then go to the position or ID where you would like to place the Cut / Copied ID and right click and select a paste option.

To move multiple IDs in one go you can use Shift + Left Click or Ctrl + Left Click to select multiple IDs then use Ctrl + X (Cut) or Ctrl + C (Copy) or Right Click on one of the selected IDs.

#### Moving IDs by drag and drop

You can move IDs that are at the same level as one another around in a specific hierarchy by selecting them and left clicking and holding the mouse button down while moving the mouse to drag IDs from their existing locations to a new location, release the mouse button to drop them.

If any dragged IDs are on different levels from one another then they will not be included in the move.

#### Deleting IDs

When using Delete on an ID in the sheet panel or Delete in all hierarchies in the treeview panel it will Delete an ID completely; across all hierarchies.

When using any other delete option it will only delete an ID in the currently selected hierarchy. However, if that ID is the last appearance of the ID across all hierarchies then it will completely delete it, just like with Delete in all hierarchies.

#### Deleting a column

To delete a column right click on the column you wish to delete and select Delete column. Note you cannot delete a parent column if it is the only parent column in the sheet and you cannot a delete a parent column if you are currently viewing it.

#### Conversion using mapping (find and replace using multiple values in a table)

By right clicking in the sheet while multiple cells are selected you can bring up the find and replace window, you can also access it under the Edit menu. Click on the tab labelled Replace using mapping.

This allows a large scale find and replace using a 2 column table, column 1 is the values to find and column 2 the corresponding values to replace them with.

You can load a .csv, .tsv, excel file or allowed json formats, you can also paste into the mini table from the clipboard.

#### Adding multiple new rows

To add multiple new rows you can use Import -> Merge sheets and then either opening a file or using the clipboard or just using the sheet display in the popup to paste / insert new rows. Right clicking in the header or index will result in a popup box where you can insert a new row or column. You can use Ctrl + V to paste data in, as long as it's in the form of tab delimited text.

#### Getting all information on an ID

An easy way to get an IDs complete information within the sheet, including parents and children across all hierarchies and all details is to select an ID in the treeview or sheet panel and then go to View -> IDs details.

You can also get a more concise view of an ID by right clicking on it and selecting ID concise view.

#### Date column conditional formatting

When entering conditional formatting in Date Detail columns, use forward slash dates e.g. DD/MM/YYYY. This is because hyphens will be interpreted as subtractions. If you want to enter a specific date, for current date use the letters: cd

#### Changing the order of IDs in the treeview

To disable automatic ordering of IDs in the treeview go to the Format menu and select Auto-sort treeview IDs so that it's unchecked. You can then select a row in the treeview and hold down the right mouse button and drag and drop the ID to it's desired position. To move an ID between parents see the above section on "Moving IDs by drag and drop".

## MANAGING COLUMNS

Right clicking on columns in the header will show a popup menu with a few column specific options.

#### Column types:

A detail column can have one of three different types, Text, Numerical and Date. Text details can be any text, Numerical details can be any number and Date details can be either a date one of three formats (YYYY/MM/DD, DD/MM/YYYY, MM/DD/YYYY) or a whole number (integer).

Changing a column type will result in any details, formatting or validation being evaluated and potentially deleted if they do not meet the column types requirements.

#### Conditional Formatting:

You can add conditional formatting to columns, meaning when certain conditions are met the cells in that column will be filled with a chosen color. You can set a maximum of 35 conditions.

To add a condition in the conditional formatting popup right click in the empty space. To manage individual conditions right click on their rows.

For Text Detail columns conditions are limited to text matching, e.g. if the cell contains exactly the user input. Text conditions are not case sensitive.

For Numerical Detail columns the following characters are allowed:
```
c   Column e.g. c5
0-9 Any number
.   Decimal place
-   Negative number
>   Greater than
<   Less than
==  Equal to
>=  Greater than or equal to
<=  Less than or equal to
and Used to add extra condition e.g. > 5 and < 10
or  Used to add extra condition e.g. == 5 or == 6
```

e.g. > 100
e.g. > 100 and < 200

For Date Detail columns the following characters are allowed:
```
c   Column e.g. c5
cd  Current date
0-9 Any number
.   Decimal place
-   Negative number
>   Greater than
<   Less than
==  Equal to
>=  Greater than or equal to
<=  Less than or equal to
and Used to add extra condition e.g. > 5 and < 10
or  Used to add extra condition e.g. == 5 or == 6
```

e.g. > 20/06/2019
e.g. == 100

Conditions must have spaces in between statements.

## MENU BAR

Under the File Menu there are the following options:
- New: Create a new sheet.
- Open: Open a file.
- Compare sheets: This option takes you to a tree comparison window. For more information find the help section "Tree Compare".
- Save: Options are Save (Ctrl + S), Save as, Save as with username-date- time (saves with the users login name, current date and time added to end of the filename) and Save new version (adds one to any detected file of the same name found in the chosen folder).
- Quit: Quits the program.

Under the Edit Menu there are the following options:
- Undo (Ctrl + Z): there are a maximum of 75 undos. Although the changelog can be saved with any filetype other than .csv the changes cannot be undone across saves.
- Copy to clipboard copies the underlying sheet to your computers clipboard to be pasted as a string, Copy as json will use the json format you have selected under the Options menu.
- Tag/Untag IDs tags the selected IDs, tagged IDs will be displayed in a dropdown box at the top of the window so you can find them later.
- Find, replace (Ctrl + F / Ctrl + H). This window also allows a large scale find and replace using a 2 column table, column 1 is the values to find and column 2 the corresponding values to replace them with.
- Find next (Ctrl + G) goes to the next appropriate search result.
- Clear copied/cut clears any copied/cut IDs
- Clear panel selections deselects both the treeview and sheet selections.
- Clear all tagged IDs clears all tagged IDs and the associated drop-down boxes. This is not an Undo-able action.

Please note that when you undo a change not related to details such as copying or deleting an ID any IDs without parents and children in any hierarchy will be placed into the FIRST hierarchy.

Under the Format Menu there are the following options:
- Sort sheet gives you two options for sorting the sheet:
    - Sort by tree: This button sorts the sheet in the order that the IDs occur in the tree.
    - Sort by column: Using this button and the two drop-down boxes to its right you can sort the sheet using a basic natrual sorting order, numbers taking priority.
- Autosort treeview IDs re-sorts the order of all treeview IDs and children, this is on by default but if you disable it the order of IDs will not longer be automatically alphanumerically maintained. You can manually set your own treeview IDs order by uncheck-ing this option and then pressing right click and holding over the desired ID and then dragging it to the desired location.
- Date format switches the date format for the program, it will try to change formats for conditional formatting and details for all date columns.

Under the View Menu there are 9 options and 4 panel configurations:
- View changelog shows an enumerated view of all changes made to the sheet, it is bound to Ctrl + L.
- View build warnings shows all warnings and issues that occurred and were fixed during first construction of the tree.
- Treeview IDs information shows the tree's currently selected IDs full information.
- Sheet IDs information shows the sheet's currently selected IDs full information.
- Expand all opens all IDs in the tree panel so that all children are visible. It is bound to the E key.
- Collapse all closes all IDs in the tree panel so that only the top IDs are visible. It is bound to the R key.
- Zoom in zooms in on both the tree and sheet.
- Zoom out zooms out on both the tree and sheet.
- Save position saves the current scroll position in the tree panel.
- Go to saved scrolls to the previously saved tree panel scroll position.
- Layout gives four choices for viewing the treeview/sheet.
- Set all column widths changes the size of the columns in the tree and sheet panels to be wide enough to show the whole of the widest cell.

Under the Import Menu there are the following options:
- Import changes allows an exported/saved changelog to be imported and the individual changes are then attempted on the currently open sheet. Supported changes are:

```
    Edit cell
    Edit cell |
    Move rows
    Move columns
    Add new hierarchy column
    Add new detail column
    Delete hierarchy column
    Delete detail column
    Column rename
    Edit validation
    Change detail column type
    Date format change
    Cut and paste ID
    Cut and paste ID |
    Cut and paste ID + children
    Cut and paste ID + children |
    Cut and paste children
    Copy and paste ID
    Copy and paste ID |
    Copy and paste ID + children
    Copy and paste ID + children |
    Add ID
    Rename ID
    Delete ID
    Delete ID |
    Delete ID, orphan children
    Delete ID + all children
    Delete ID from all hierarchies
    Delete ID from all hierarchies |
    Delete ID from all hierarchies, orphan children
    Sort sheet
```

You can also recycle the imported changes, importing them again into another file. There are certain things that may stop a change from being imported, for example if the change was made to a column with a different name or number than the column in the open sheet or if an IDs parent is different. Unfortunately at this time it does not tell you why a change has not been imported successfully, this may be improved in a future version.

- Get sheet from clipboard and overwrite allows you to get copied data from your devices clipboard and overwrite all current data. This action can be undone.
- Merge sheets allows you to merge one sheet with another, you have options to overwrite details, parents, add new ids etc. You also can simply add multiple additional rows by pasting into the sheet on the right hand side of the pop-up.

Under the Export Menu there are the following options:
- Export changes gives a view of the changelog and allows saving/exporting of changes.
- Export flattened sheet allows you to add all IDs flattened levels for any hierarchies to a sheet and then gives options for saving as .xlsx or .csv or copying to clipboard.

Under the Options Menu there are the following options:
- Auto-resize row indexes allows width adjustment of the row indexes in the treeview and sheet.
- Auto-select sheet ID toggles auto selecting of a row in the sheet when you select it in the treeview.
- Allow spaces in ID names allows spaces in ID names, disabling this will not remove existing spaces from ID names. This option is saved on a per file basis with the program data.
- Allow spaces in column names allows spaces in column names, disabling this will not remove existing spaces from column names. This option is saved on a per file basis with the program data.
- XLSX save options gives options for what data is saved within .xlsx files.
- JSON output format gives four choices for the format of ALL json output, saving and copying to clipboard. They are explained in more detail under the JSON Help section.
- Set all sheet alignments sets the global cell text alignment for columns in both the treeview and sheet. These global alignments are overridden when setting specific alignments by using the right click popup menu.
- Theme allows you to change the general colours of the app.

## TREE BUTTONS

In the tree panel:

Find: Clicking the find button will attempt to find either an ID or detail (depending on which you has selected in the drop-down box on the right of "Find") with letters, numbers etc. you have typed into the entry box on the right of the drop-down box. The drop-down box below the Find button will display any results found within the CURRENTLY viewed hierarchy. 

All finds are NOT case sensitive, including "exact match".

Hierarchy: This is the drop-down box where you can select which parent column/hierarchy to view.

In the sheet panel:

Tagged IDs (Ctrl + T): Allows you to tag IDs, tagged IDs show up in the dropdown box next to the button and persist through saving as .json and .xlsx.

Find: Works the same way that the Find button for the Tree panel works except it searches the sheet instead.

## TREE FUNCTIONS

By right clicking on an ID in the tree panel you can select various functions. The main functions are Cut, Copy and delete.

To cut or copy an ID between different hierarchies simply right click on the ID and select whichever option you want then switch hierarchy and right click in empty space or on the ID you want to paste the cut/ copied ID to as a sibling or child. If you want to paste an ID as an ID without a parent right click on a top ID and choose paste as sibling.

You can also cut all of an IDs children, including grandchildren and so on, and paste them under where you right click.

Using shift click you can select multiple up or down of an existing selection. Using Ctrl click you can make multiple selections.

When using the Ctrl X, C and V keys to cut/copy and paste they will work on the selected ID, not on the position where the mouse is hovering (unless pasting over empty space using Ctrl V).

Cutting and copying using this method will only perform on IDs that are on the same level as the top most (index-wise) ID, after pressing Ctrl X or C it will deselect any selections that were not cut or copied. 

Pressing the Delete key on multiple selections will work the same way, except performing a Delete immediately. The delete key uses the typical Delete ID function, not deleting its children.

In the tree panel there are 5 delete ID options. Delete ID only removes the ID from the hierarchy you're currently viewing IF the ID occurs in another hierarchy, if it does not then it totally removes the ID. 

Del all of ID totally removes the ID. Del ID+children is the same as Delete ID but for every child and child of that child and so on recursively under the selected ID.

You can quickly edit a detail by double clicking on the detail/cell you want to edit. To delete a detail press Confirm when editing a detail with the cell empty.

Right clicking in a cell and selecting edit will pop up a larger window so the text may be easier to view.

Pasting a detail or details will work between both panels. You can drag and drop rows in the sheet panel to change their order.

When using drag and drop you can use your mousewheel to scroll down, just move your mouse a bit after scrolling to cause the selection to move.

## TREE COMPARE

Accessible from the "File" menu, this window allows comparison of trees and sheets. You can open files using the "Open file" button which will open a file dialog. Once open the file name or path will display next to this button and if the file is an excel file and was opened from the file dialog then you will have to select a sheet from the drop down box next to "Load sheet". 

You can select your sheets ID column and parent column numbers and do the same with the 2nd panel on the right. After you are happy with your selections click the "Create Report" button to compare. A report will be generated and you have the option to save it as a .xlsx file which will have 2 sheets.

You can mix different file types when comparing.

## XLSX FILES

The default save format is .xlsx.

When saving .xlsx files you can also save program data to keep your changelog, row heights, column widths, formatting, validation, treeview ID order and more.

When loading a file saved with program data the sheet and changelog in the program data, not the visible sheet, will take precedent. This means any edits in the viewable sheet will not be loaded.

To disable saving with program data go to Options -> XLSX save options -> Save xlsx with program data.

If choosing to save program data any sheet named "program_data" will be overwritten when saving a workbook.

You can also save a viewable changelog sheet. Sheets with "Changelog" in their name will be overwritten when this option is chosen.

When saving .xlsx files you can also save the flattened format of the currently viewed hierarchy any sheets with "flattened" in their name will be overwritten. If viewing all hierarchies when saving then the first hierarchy will be saved.

When comparing or merging if the workbook contains program data then it will take precedent, else a sheet will need to be selected to load data.

## JSON FILES

There are four loadable json formats, with each one the entire sheet is kept under the key "records". However the program will also look for the keys: sheet, data and table. The first format, also the first option under "Options -> JSON output format" is displayed as an example below:

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

## USING THE API

The app can be run using an API without triggering a user interface to get different outputs and file conversions.

The input file must be either .xlsx, .xls, .xlsm, .csv or .tsv.

Please note that if any of the parameters include spaces then they may need to be surrounded by double quotes e.g. "my xlsx sheet name" depending on how you choose to start the API.

It must be run with the following paramters with a space in-between each parameter, ignore the <> symbols:

#### Required parameters:

1. API Action, one of the following:
    - flatten
    - unflatten-top-base
    - unflatten-base-top
2. Input filepath, must be the full filepath
3. Output filepath, including new filename
4. All the parent column indexes, 0 being the lowest number

#### Required **only** for `flatten` action:

5. ID column index, **required** for flatten action, e.g.:
    - -id-0
    - -id-A
6. Parent column index, **required** for flatten action, e.g.:
    - -parent-2
    - -parent-C

#### Optional (but important) parameters:

7. Input sheet name, if not provided defaults to first sheet of the input file if it's an xlsx file, e.g.:
    - -input-sheet-Sheet1
8. Output sheet name, if not provided uses the input sheet name or Sheet1, e.g.:
    - "-output-sheet-New Sheet"

7. Delimiter, a delimiter character for the output file if it's a csv or tsv, defaults to comma, examples below:
    - -delim-tab
    - -delim-,
    - -delim-|
8. Flags (can be used one after the other):
    - e.g. -odjr

| Flag    | Used for                    |
|---------|-------------                |
| -o      | Overwrite new file          |
| -d      | Include detail columns      |
| -j      | Justify output cells left   |
| -r      | Reverse order (top-base)    |
| -i      | Add an index column         |

Some examples of the flatten command line below which would flatten the hierarchy at column index 2, column C (ignore <> characters):

```
flatten <"input filepath here.xlsx"> <"output filepath here.xlsx"> 0,2,4,6 -id-0 -parent-2 
```