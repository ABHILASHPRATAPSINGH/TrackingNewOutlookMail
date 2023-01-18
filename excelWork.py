import os

import win32com.client
import pythoncom
import re
import win32com.client as win32
import json
import requests
import pandas

from utils import getSrcFolderPath


class StockExlVBA:
    def __init__(self,filepath):
        self.filepath=filepath

    def calling_Exl_Macro(self):
        xl=win32.gencache.EnsureDispatch('Excel.Application')
        wb=xl.Workbooks.Open(self.filepath)
        xl.Application.Run("Marketwatch.xlsm!FinancialExpress.TestMacro")
        wb.Close(True)

    def update5DaysData(self):
        stock_data=self.getDataframeFromExl()
        outputFile=os.path.join(getSrcFolderPath(),'Get5DaysData.xlsx')
        datatoexl=pandas.ExcelWriter(outputFile)
        stock_data.to_excel(datatoexl,sheet_name='All',index=False)
        datatoexl.save()

    def getDataframeFromExl(self):
        df=pandas.read_excel(self.filepath,sheet_name='All',skiprows=5,usecols='A:AR',header=1)
        df=df[['Stokes',
               'p(t)','p(t+1)','p(t+2)','p(t+3)','p(t+4)',
               '%(t)','%(t+1)','%(t+3)','%(t+3)','%(t+4)',
               'v(t)','v(t+1)','v(t+4)','v(t+3)','v(t+4)',
               'del(t)','del(t+1)','del(t+2)','del(t+3)','del(t+4)',
               'Sector','Market Cap (Rs. in Cr.)','URL']]
        return df

    def getListOfDictFromDataframe(self,df):
        listOfDict_allStock=[]
        for row_stock in range(len(df)):
            data_dict={}
            data_dict['sname']=str(df.loc[row_stock,'Stocks'])
            data_dict['sprice_1'] = str(df.loc[row_stock, 'p(t)'])
            data_dict['sprice_2'] = str(df.loc[row_stock, 'p(t+1)'])
            data_dict['sprice_3'] = str(df.loc[row_stock, 'p(t+2)'])
            data_dict['sprice_4'] = str(df.loc[row_stock, 'p(t+3)'])
            data_dict['sprice_5'] = str(df.loc[row_stock, 'p(t+4)'])
            data_dict['srate_1'] = str(df.loc[row_stock, '%(t)'])
            data_dict['srate_2'] = str(df.loc[row_stock, '%(t+1)'])
            data_dict['srate_3'] = str(df.loc[row_stock, '%(t+2)'])
            data_dict['srate_4'] = str(df.loc[row_stock, '%(t+3)'])
            data_dict['srate_5'] = str(df.loc[row_stock, '%(t+4)'])
            data_dict['svol_1'] = str(df.loc[row_stock, 'v(t)'])
            data_dict['svol_2'] = str(df.loc[row_stock, 'v(t+1)'])
            data_dict['svol_3'] = str(df.loc[row_stock, 'v(t+2)'])
            data_dict['svol_4'] = str(df.loc[row_stock, 'v(t+3)'])
            data_dict['svol_5'] = str(df.loc[row_stock, 'v(t+4)'])
            data_dict['sdel_1'] = str(df.loc[row_stock, 'del(t)'])
            data_dict['sdel_2'] = str(df.loc[row_stock, 'del(t+1)'])
            data_dict['sdel_3'] = str(df.loc[row_stock, 'del(t+2)'])
            data_dict['sdel_4'] = str(df.loc[row_stock, 'del(t+3)'])
            data_dict['sdel_5'] = str(df.loc[row_stock, 'del(t+4)'])
            listOfDict_allStock.append(data_dict)
        return listOfDict_allStock