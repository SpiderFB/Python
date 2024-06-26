import os
import win32com.client

excel_app = win32com.client.Dispatch("Excel.Application")

MDG = 'C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx'
ATLAS = 'C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx'

def fun_cp(system_name):
    source_wb = excel_app.Workbooks.Open(system_name)

    # print(f"swf:{source_wb}")
    source_sheet = source_wb.Sheets('Sheet1')
    dest_wb = excel_app.Workbooks.Open('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx')
    # source_sheet.Copy(Before = dest_wb.Sheets(1))
    source_sheet.Copy(dest_wb.Sheets(1))
    dest_wb.Sheets(1).Name=os.path.basename(system_name)
    dest_wb.Save()
    dest_wb.Close()
    source_wb.Close()
    excel_app.Quit()
fun_cp(MDG)
fun_cp(ATLAS)