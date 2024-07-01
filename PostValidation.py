import os
import win32com.client
import pandas as pd

TABLE = input(" TABLE NAME : ").upper()

MDG = 'C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx'
ATLAS = 'C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx'

df = pd.DataFrame()

DataValidationPath = 'C:/Users/2095421/Downloads/Migration/' + TABLE + '_DataVALIDATION.xlsx'

df.to_excel(DataValidationPath)
scope_df = pd.read_excel('C:/Users/2095421/Downloads/Migration/Field_Scope_Data_Migration.xlsx', sheet_name=TABLE)
scope_df.to_excel(DataValidationPath, sheet_name='Sheet1', index=False)

print(TABLE + " Data Validation File created at - " + DataValidationPath )
print("\n")



def fun_cp(system_name):
    excel_app = win32com.client.Dispatch("Excel.Application")
    source_wb = excel_app.Workbooks.Open(system_name)
    source_sheet = source_wb.Sheets('Sheet1')

    dest_wb = excel_app.Workbooks.Open('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx')
    source_sheet.Copy(dest_wb.Sheets(1))    
    print(os.path.basename(system_name), " File copied to the DataValidation file Successfully!" )
    scope_sheet = dest_wb.Sheets('Sheet1')
    dest_wb.Sheets(1).Name=os.path.basename(system_name)
    dest_sheet = dest_wb.Sheets(os.path.basename(system_name))
    dest_sheet.Columns(1).Insert()
    dest_sheet.Cells(1, 1).Value = 'KEY_' + os.path.basename(system_name)
    dest_sheet.Range("A1").Interior.Color = 65535
    dest_sheet.Range("A1").Font.Bold = True
    print("Key creation done Successfully!")
    
    # For now MBEW Table only
    row = 2
    while dest_sheet.Cells(row, 2).Value is not None and dest_sheet.Cells(row, 3).Value is not None:
        dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value)
        row += 1

    print('Count of ', os.path.basename(system_name) + ": ", row-1)
    for row in range(2, scope_sheet.usedRange.Rows.Count+1):
        if scope_sheet.Cells(row, 2).Value == 'Scope':
            technical_name = scope_sheet.Cells(row, 1).Value
            for col in range(1, dest_sheet.UsedRange.Columns.Count+1):
                if dest_sheet.Cells(1, col).Value == technical_name:
                    dest_sheet.Cells(1, col).Interior.ColorIndex = 4
                    break

    print("Scope fields marked with Green Successfully!")
    
    dest_wb.Save()
    dest_wb.Close()
    source_wb.Close()
    excel_app.Quit()
    print("\n")

fun_cp(MDG)
fun_cp(ATLAS)

df1 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_M1Q_Ext.xlsx')
df2 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_APB_Ext.xlsx')

df1.set_index('KEY_MBEW_M1Q_Ext.xlsx', inplace=True)
df2.set_index('KEY_MBEW_APB_Ext.xlsx', inplace=True)

df3 = pd.read_excel('C:/Users/2095421/Downloads/Migration/Field_Scope_Data_Migration.xlsx', sheet_name=TABLE)
df_scope = df3[df3['Comment'] == 'Scope']
scope_field = df_scope['Technical Name']

for col_name in scope_field:
    print(col_name)
    if col_name != 'MATNR' and col_name != 'BWKEY' :
        result_df = pd.DataFrame(columns=['Key','Value-Matched?','Material','Plant', 'ATLAS', 'MDG'])
        for index in df2.index:
            if index in df1.index:
                if df2.loc[index, col_name] == df1.loc[index, col_name] or (pd.isna(df2.loc[index, col_name]) and pd.isna(df1.loc[index, col_name])):
                    # print(f'Key {index}: Yes')
                    pass
                else:
                    # print(f'Key {index}: No')
                    result_df.loc[len(result_df)] = [index, 'No',df2.loc[index, 'MATNR'],df2.loc[index, 'BWKEY'], df1.loc[index, col_name], df1.loc[index, col_name]]
            else:
                # print(f'Key {index}: Key not found')
                result_df.loc[len(result_df)] = [index, 'Key Not Found',df2.loc[index, 'MATNR'],df2.loc[index, 'BWKEY'],'','']

        if not result_df.empty:
            writer = pd.ExcelWriter('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', engine='openpyxl', mode='a')
            result_df.to_excel(writer, sheet_name=col_name)
            writer.close()