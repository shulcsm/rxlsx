#!/usr/bin/env python

from _data import data
import openpyxl
import os


if __name__ == "__main__":
    wb = openpyxl.workbook.Workbook(write_only=True)
    ws = wb.create_sheet()
    for row in data:
        ws.append(row)
    wb.save(os.devnull)
