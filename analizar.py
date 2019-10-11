#!/usr/bin/python

from tkinter import *
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook


root = Tk()
root.withdraw()
root.update()


def get_file():
    work_book = None
    path_string = askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if path_string != "":
        work_book = load_workbook(path_string)
    return work_book


def get_sheet(wb, index):
    sheet_names = wb.sheetnames
    sheet = wb[sheet_names[index]]
    return sheet


def get_sr_record(sheet, record_col, value_col):
    values = []
    fields = []
    current_record = {}
    last_recorded_value = None
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=value_col):
        if last_recorded_value != row[record_col-1].value:
            current_recorded_id = row[record_col-1].value
            current_total_value = sum(values)
            current_record.update({current_recorded_id: (fields, current_total_value)})
            values = []
            fields = []
        else:
            values.append(row[value_col-1].value)
            fields.append(row[record_col-1].coordinate)
        last_recorded_value = row[record_col-1].value
    return current_record


def get_vs_record(sheet, record_col, value_col):
    current_record = {}
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=value_col):
        current_record.update({row[record_col-1].value: (row[record_col-1].coordinate, row[value_col-1].value)})
    return current_record


if __name__ == "__main__":
    current_wb = get_file()
    ag_sheet = get_sheet(current_wb, 0)
    cs_sheet = get_sheet(current_wb, 1)
    ag_record = get_vs_record(ag_sheet, 3, 10)
    cs_record = get_sr_record(cs_sheet, 6, 15)
    print(cs_record)
    for key in ag_record:
        if key in cs_record:
            print(cs_record[key])
            print(ag_record[key])
    root.destroy()