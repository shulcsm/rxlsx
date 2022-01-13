#!/usr/bin/env python

from _data import data
import xlsxwriter.workbook
import os


if __name__ == "__main__":
    wb = xlsxwriter.workbook.Workbook(os.devnull, {"constant_memory": True})
    ws = wb.add_worksheet()
    for rix, row in enumerate(data):
        for cix, value in enumerate(row):
            ws.write(rix, cix, value)

    wb.close()
