# SPDX-License-Identifier: AGPL-3.0-only
# Copyright (c) R. A. Gardner

"""Tables used by Session / CLI tests. DOCUMENTATION.md Animals example."""

from __future__ import annotations

# Header + rows, ID / Parent adjacency (format 0).
ANIMALS_TABLE = [
    ["ID", "Parent", "Name"],
    ["Animals", "", "All animals"],
    ["Cats", "Animals", "Cat family"],
    ["Lion", "Cats", "Lion"],
]

ANIMALS_ID_COL = 0
ANIMALS_PARENT_COLS = [1]

# Deeper forest for sibling order, numeric sort_key, and topnodes_order.
# Mix of branched tops, leaf tops, and numbered IDs (Item2 vs Item10, Ch2 vs Ch10).
FOREST_TABLE = [
    ["ID", "Parent", "Name"],
    ["Life", "", "Life"],
    ["Animals", "Life", "Animals"],
    ["Plants", "Life", "Plants"],
    ["Moss", "", "Moss"],
    ["Item10", "", "Item10"],
    ["Item2", "", "Item2"],
    ["Cats", "Animals", "Cats"],
    ["Dogs", "Animals", "Dogs"],
    ["Aardvark", "Animals", "Aardvark"],
    ["Zebra", "Animals", "Zebra"],
    ["Lion", "Cats", "Lion"],
    ["Tiger", "Cats", "Tiger"],
    ["Wolf", "Dogs", "Wolf"],
    ["Oak", "Plants", "Oak"],
    ["Ch10", "Oak", "Ch10"],
    ["Ch2", "Oak", "Ch2"],
    ["Ch3", "Oak", "Ch3"],
]

# Same IDs, two parent columns with different membership.
TWO_HIER_TABLE = [
    ["ID", "H1", "H2", "Name"],
    ["Root", "", "", "Root"],
    ["A", "Root", "", "A"],
    ["B", "A", "Root", "B"],
    ["C", "", "Root", "C"],
    ["D", "C", "", "D"],
]
