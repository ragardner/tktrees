# tktrees

TkTrees is a hierarchy management app written in Python.

## Download

Download a release from:

https://github.com/ragardner/tktrees/releases

Unzip the folder somewhere you can find it again.

## How to run

Requirements:

- Python 3.9 or newer.
- tkinter is required (it usually comes with Python on Windows).

On Windows you can double-click the file named `TKTREES.pyw`.

Otherwise open a terminal in the unzipped folder and run:

```
python3 TKTREES.pyw
```

On Windows that command is often `python TKTREES.pyw` instead.

## Getting started

After the app opens, use the File menu or right-click the empty space. File -> Open loads an existing file.

Help is under Help -> View Help in the app, or in `DOCUMENTATION.md` in the same folder.

## Bundled Dependencies

For convenience in environments where pip is not available, this project includes copies of the following third-party libraries. These are unmodified and provided under their original licenses. Users are responsible for complying with these licenses.

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

## License

TkTrees is licensed under AGPL-3.0 and is the copyright of R. A. Gardner.

- Author: github.com/ragardner
- Source: github.com/ragardner/tktrees
- Email: github@ragardner.simplelogin.com