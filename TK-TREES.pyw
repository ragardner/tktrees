# ruff: noqa: F401
# SPDX-License-Identifier: AGPL-3.0-only
# Copyright © R. A. Gardner

"""
tk-Trees - Hierarchy Management Tool

This program is for management of hierarchy data in table format

Requires Python >= 3.9
"""

if __name__ == "__main__":
    from sys import argv

    from src import app

    app.run_app(argv)
