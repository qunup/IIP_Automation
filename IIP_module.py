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
        return 0
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
    return dict_df


def master_shop():
    global dict_df
    global wb
    dict_df = import_data()
    wb = xw.Workbook.caller()
    if dict_df == 0:
        # wb.macro('ShowMsg')("Data Import Failed")
        return 0
    l_output = [l for l in xw.Range(
        'Macro', 'OutputField').value if l[0] != None]
    df_output = DataFrame(
        l_output, columns=['FilePath', 'FileName', 'DataType', 'Status'])
    dict_ms = {"BOM": ms_BOM,
               "COOP": ms_COOP,
               "Future Sample Receipts": ms_Sample,
               "Gross Cost Margin": ms_Margin,
               "Inbound Freight": ms_Freight,
               "Receipts": ms_Receipts,
               "Sales, Discounts, Points": ms_Sales}
    for index, row in df_output.iterrows():
        try:
            row['Status'] = dict_ms[row['DataType']](row)
        except:
            row['Status'] = "Fail"
    wb.active()
    xw.Range('Macro', 'C_Status').options(transpose=True).value = df_output['Status'].tolist()
    wb.macro('ShowMsg')("Done!")


# separate routines to process each master file

def ms_BOM(row):
    return "Fail"

def ms_COOP(row):
    return "Fail"

def ms_Sample(row):
    return "Fail"

def ms_Margin(row):
    return "Fail"

def ms_Freight(row):
    return "Fail"

def ms_Receipts(row):
    return "Fail"

def ms_Sales(row):
    wb_output = xw.Workbook(row['FilePath'])
    
    # tab PIM Vendor
    xw.Range('PIM Vendor','A1').table.clear_contents()
    xw.Range('PIM Vendor','A1').value = dict_df['PIM Vendors'].columns.tolist()
    xw.Range('PIM Vendor','A2').value = dict_df['PIM Vendors'].values
    
    # tab Looker Pull
    xw.Range('Looker Pull','D2:L2').vertical.clear_contents()
    xw.Range('Looker Pull','D2').value = dict_df['Sales, Discounts, Points'].values
    n_row = xw.Range('Looker Pull','D2').vertical.last_cell.row
    xw.Range('Looker Pull','D' + str(n_row + 1)).vertical.clear_contents()
    xw.Range('Looker Pull','A3:C3').vertical.clear_contents()
    f_copy_formula = wb.macro('Copy_Formula')
    f_copy_formula('Looker Pull','A2:C2','A2:C'+ str(n_row))
    xw.Range('Looker Pull','M3:O3').vertical.clear_contents()
    f_copy_formula('Looker Pull','M2:O2','M2:O'+ str(n_row))
    xw.Range('Looker Pull',str(n_row + 1)+":"+str(n_row + 1)).clear_contents()
    
    # save
    wb_output.save()
    wb_output.close()
    return "Success"