# -*- coding: utf-8 -*-
"""
Created on Wed Jul 22 13:04:50 2020

@author: Davit
"""

import openpyxl as opx
import os 
import pandas as pd
import numpy as np

os.chdir(r"D:\Bank of Russia\finist_new")

wb = opx.load_workbook("finist_new.xlsm")


def Delete_Engines():
    
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    ActiveSheet.Select
    
    # цикл для расчета влкадок с расчетными модулями
    for sheet in wb.sheetnames:
        if sheet.startswith('B_') and not sheet == 'B_ALL':
            wb.remove(sheet)
            
    # цикл для мертвых поименованных диапазонов 
    