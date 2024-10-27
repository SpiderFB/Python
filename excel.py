import pandas as pd

dummy_file = input("Please give MDG file path ").strip('\"')
data_frame = pd.read_excel(dummy_file, sheet_name='Sheet1')
data_frame = data_frame[data_frame['WERKS'] != 'JP13']
data_frame.to_excel(dummy_file, sheet_name='Sheet1', index=False)

# excel_data1 = pd.read_excel(file, sheet_name='Sheet1', index_col = 1) #To make a particular Comun as 1st/Index column
# excel_data2 = pd.read_excel(file, sheet_name='Sheet2')
# excel_data = pd.read_excel(file, sheet_name='Sheet1', usecols=['ABC']) #To read a particular column of a Excel
# excel_data = pd.read_excel(file, sheet_name='Sheet1', header=None) #Read 1st row as Non-header
# print(excel_data1)
# print(excel_data1.columns.ravel()) #Print only header names
# print(excel_data1['ABC']) #To print only a particular column
# print(excel_data1['ABC'].tolist()) #To print the coumun data in a list
# print('Excel Sheet to JSON:', excel_data1.to_json(orient='records')) #Convert Excel record into JSON data
# print('Excel Sheet to Dictionary:', excel_data1.to_dict(orient='records')) #Convert Excel record into Dictionary data
# print('Excel Sheet to Csv:', excel_data1.to_csv(orient='records')) #Convert Excel record into CSV data

# #to Add/Concat multiple sheets
# newData = pd.concat([excel_data1, excel_data2])
# print(newData)

# #To print Top/Bottom
# print(excel_data1.head(1)) #by default head() gives first 5 as output
# print(excel_data1.tail(2)) #by default tail() gives last 5 as output

# print(excel_data1.shape) #To print the number of Rows & Columns