from pandas import *
import glob
# import os

import numpy as np
import xlwings as xw
# from xlwings import Workbook, Range


def get_file_list():
    """get the list of files in the folder"""
    wb = xw.Workbook.caller()
    path_input = xw.Range('Macro', 'FilePath').value
    l_file_path = glob.glob(path_input + '[!~]*.*')
    l_file_name = [l.split('/')[-1] for l in l_file_path]
    xw.Range('Macro', 'FileField').clear_contents()
    xw.Range('Macro', 'C_FilePath').options(transpose=True).value = l_file_path
    xw.Range('Macro', 'C_FileName').options(transpose=True).value = l_file_name
    xw.Sheet('Macro').activate()
    wb.macro('ShowMsg')("Choose DataType for all the listed files")


def import_data():
    wb = xw.Workbook.caller()
    l_input = [l for l in xw.Range('Macro', 'FileField').value if l[0] != None]
    df_input = DataFrame(l_input, columns=['FilePath', 'FileName', 'DataType'])
    # test datatype input
    if any(df_input.duplicated('DataType')) or any(df_input.DataType.isnull()):
        wb.macro('ShowMsg')("DataType column has duplicates and/or empty cell" +
                            "\n" + "Please fix it and rerun the macro")
        return
    # import_data
    # exception for inbound freight
    dict_df = {}
    for index, row in df_input.iterrows():
        if row['FilePath'].split('.')[-1] == 'csv':
            dict_df[row['DataType']] = read_csv(
                row['FilePath'], na_filter=False)
        elif row['FilePath'].split('.')[-1] == 'xlsx':
            dict_df[row['DataType']] = read_excel(row['FilePath'], sheetname="All Accounts" if row[
                                                  'DataType'] == 'Inbound Freight' else 0, keep_default_na=False)
        else:
            dict_df[row['DataType']] = read_html(
                row['FilePath'], header=0)[0].fillna("")
