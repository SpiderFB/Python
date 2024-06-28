import pandas as pd

MBEW_MDG = "C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx"
MBEW_ATLAS = "C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx"

MBEW_DataValidation = "C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx"

df = pd.read_excel(MBEW_MDG, sheet_name = 'Sheet1')
ATLAS_df = pd.read_excel(MBEW_ATLAS, sheet_name = 'Sheet1')

with pd.ExcelWriter(MBEW_DataValidation, engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='MBEW_MDG', index=False)
    ATLAS_df.to_excel(writer, sheet_name='MBEW_ATLAS', index=False)
print("Copied")