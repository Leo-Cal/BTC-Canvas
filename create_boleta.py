import xlwings as xw
import numpy  as np
import pandas as pd

def create_boleta(workbook,dataframe):

    wb = xw.Book(workbook)
    sht = wb.sheets['New SQL']

    sht.range('AC1').value = dataframe

    return 0