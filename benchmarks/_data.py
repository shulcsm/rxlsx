from decimal import Decimal
"""
ipg data 2021
cols to AX
27378 rows
"""
ROWS = 10_000
#  COLUMNS = 50
ROW = [
    1,
    2,
    3,
    4,
    5,
    6,
    7,
    8,
    9,
    10,
    True,
    False,
    0.1,
    Decimal("0.2"),
    "foo",
    "bar",
    "baz",
    "qux",
    "fred",
    "thud",
]

data = [ROW] * ROWS
