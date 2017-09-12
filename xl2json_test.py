""" 
        Title: xl2json_test
        Author: Harold Goldman
        Email: hgold90@entergy.com
        Date: 9/12/2017
        Description: Tests for testing xl2json using pytest.
"""

import filecmp
import json
import pytest
import xlrd
import xl2json

@pytest.fixture(scope = "module")
def test_data():
    """
     read in test data
    """
    with open('tests\\test_data.json', "r") as data:
        return json.load(data)

def test_main():
    """
    test the main function
    """
    xl2json.run_main("tests\\test.xlsx")
    assert filecmp.cmp('tests\\test.json', 'tests\\test.json.test')

def test_get_args():
    """
    test get_args
    """
    assert "xl2json_test.py" == xl2json.get_args().f

def test_get_workbook(test_data):
    """
    test get_workbook
    """
    assert(test_data["workbook_data"]  == str(xl2json.get_workbook(xlrd.open_workbook("tests\\test.xlsx"))))

def test_write_json(test_data):
    """
    test write json
    """
    xl2json.write_json("tests\test.xls", test_data["workbook_data"])
    assert filecmp.cmp('tests\\test.json', 'tests\\test.json.test')

def test_get_sheet(test_data):
    """
    test get_sheet
    """
   assert test_data["sheet_data"] == str(xl2json.get_sheet(
       xlrd.open_workbook("tests\\test.xlsx").sheet_by_index(1), 
       test_data["column_names"].split(",")))

def test_get_row(test_data):
    """
    test_get_row
    """
    assert str(test_data["row_data"]) == str(xl2json.get_row(
        xlrd.open_workbook("tests\\test.xlsx").sheet_by_index(1).row(1), 
        test_data["column_names"].split(",")))

def test_get_column_names(test_data):
    """
    test get_column_names
    """
    assert  test_data["column_names_text"] == str(xl2json.get_column_names(xlrd.open_workbook("tests\\test.xlsx").sheet_by_index(1)))
