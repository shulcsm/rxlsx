#!/usr/bin/env python

from _data import data
import rxlsx
import os


if __name__ == "__main__":
    wb = rxlsx.Workbook()
    ws = wb.create_sheet()
    for row in data:
        ws.append(row)
    wb.save(os.devnull)
