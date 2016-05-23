from pandas import *
import glob
#import os

import numpy as np
import xlwings as xw
#from xlwings import Workbook, Range

def get_file_list():
    """get the list of files in the folder"""
    wb=xw.Workbook.caller()
    path_input = xw.Range('Macro','C2').value
    l_file_path = glob.glob(path_input + '*.*')
    l_file_name = [l.split('/')[-1] for l in l_file_path]
    xw.Range('Macro','B10:D40').clear_contents()
    xw.Range('Macro','B10').options(transpose=True).value = l_file_path
    xw.Range('Macro','C10').options(transpose=True).value = l_file_name
    xw.Sheet('Macro').activate()
    wb.macro('ShowMsg')("Choose DataType for all the listed files")

def check_datatype():
    wb=xw.Workbook.caller()
    l_input = [l for l in xw.Range('Macro',"B10:D40").value if l[0]!=None]
    df_input = DataFrame(l_input,columns=['FilePath','FileName','DataType'])
    test_dup = any(df_input.duplicated('DataType'))
    test_empty = any(df_input.DataType.isnull())
    if test_dup or test_empty:
        wb.macro('ShowMsg')("DataType column has duplicates and/or empty cell" + "\n" + "Please fix it and rerun the macro")
    else:
        import_data()

def import_data():
    wb=xw.Workbook.caller()
    wb.macro('ShowMsg')("Data Import Completed")

