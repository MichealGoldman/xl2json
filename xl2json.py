""" 
        Title: xl2json
        Author: Harold Goldman
        Email: hgold90@entergy.com
        Date: 9/12/2017
        Description: Python program for copying xlsx/xls to json.
"""

import argparse
import json
import os.path
from collections import OrderedDict
import xlrd


def get_column_names(sheet):
    """
    Description: get the Names from the columns
    """
    try:
        column_names = []
        for value in sheet.row_values(0, 0, sheet.row_len(0)):
            column_names.append(value)
        return column_names
    except IndexError:
        return
    except Exception as exception:
        print "Exception in get_column_names {}".format(exception)


def get_row(row, column_names):
    """
    Description: get the data from the row
    """
    try:
        row_data = OrderedDict()
        for cell in row:
            row_data[column_names[row.index(cell)]] = cell.value
        return row_data
    except Exception as exception:
        print "Exception in get_row {}".format(exception)


def get_sheet(sheet, column_names):
    """
    Description: get data from sheet
    """
    try:
        sheet_data = []
        for row in range(1, sheet.nrows):
            sheet_data.append(get_row(sheet.row(row), column_names))
        return sheet_data
    except Exception as exception:
        print "Exception in get_sheet {}".format(exception)


def get_workbook(workbook):
    """
    Description: get data from xl workbook
    """
    try:
        workbook_data = OrderedDict()
        for sheet in range(0, workbook.nsheets):
            workbook_data[workbook.sheet_by_index(sheet).name] = get_sheet(
                workbook.sheet_by_index(sheet), 
                get_column_names(workbook.sheet_by_index(sheet)))
        return workbook_data
    except Exception as exception:
        print "Exception in get_workbook {}".format(exception)


def write_json(xls, workbookdata):
    """
    Description: open excel file file
    """
    try:
        with open((xls.replace("xlsx", "json")).replace("xls", "json"), "wb") as outfile:
            outfile.write(json.dumps(workbookdata, indent=4, separators=(',', ": ")))
        print "JSON written to {}".format(outfile.name)
    except Exception as exception:
        print "Exception in write_json {}".format(exception)


def get_args():
    """
    Description: get args
    """
    parser = argparse.ArgumentParser()
    parser.add_argument("f", help="the excel to convert to json", type=str)
    return parser.parse_args()


def run_main(infile):
    """
    Description: gather data and create json
    """
    try:
        if os.path.isfile(infile):
            write_json(infile, get_workbook(xlrd.open_workbook(infile)))
        else:
            print "Invalid filename provided."
    except Exception as exception:
        print "Exception in main {}".format(exception)



if __name__ == "__main__":
    args = get_args()
    run_main(args.f)
    exit()
