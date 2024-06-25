import pandas as pd

MBEW_MDG = "C:\\Users\2095421\Downloads\Migration\MBEW_M1Q_Ext.XLSX"
MBEW_ATLAS = "C:\\Users\2095421\Downloads\Migration\MBEW_APB_Ext.XLSX"

MBEW_DataValidation = "C:\\Users\2095421\Downloads\Migration\MBEW_DataValidation.xlsx"

#change xxx with the sheet name that includes the data
data = pd.read_excel(MBEW_MDG, sheet_name="Sheet1")

#save it to the 'new_tab' in destfile
data.to_excel(destfile, sheet_name='MDG_MBEW')

#change xxx with the sheet name that includes the data
data = pd.read_excel(MBEW_ATLAS, sheet_name="Sheet1")

#save it to the 'new_tab' in destfile
data.to_excel(destfile, sheet_name='MDG_ATLAS')