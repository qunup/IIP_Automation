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
    xw.Range('File','A2:B100').clear_contents()
    xw.Range('File','A2').options(transpose=True).value = l_file_path
    xw.Range('File','B2').options(transpose=True).value = l_file_name
    xw.Sheet('File').activate()
