#!/usr/bin/env python

from _data import data
from pyexcelerate import Workbook
import os


if __name__ == "__main__":
    wb = Workbook()
    ws = wb.new_sheet("Test 1", data=data)
    wb.save(os.devnull)
