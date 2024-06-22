import pandas as pd
excel_data = pd.read_excel('abc.xlsx', sheet_name='Sheet1')
# excel_data = pd.read_excel('abc.xlsx', sheet_name='Sheet1', usecols=['ABC']) #To read a particular column of a Excel
# excel_data = pd.read_excel('abc.xlsx', sheet_name='Sheet1', header=None) #Read 1st row as Non-header
print(excel_data)
# print(excel_data.columns.ravel()) #Print only header names
# print(excel_data['ABC']) #To print only a particular column
# print(excel_data['ABC'].tolist()) #To print the coumun data in a list
# print('Excel Sheet to JSON:', excel_data.to_json(orient='records')) #Convert Excel record into JSON data
# print('Excel Sheet to Dictionary:', excel_data.to_dict(orient='records')) #Convert Excel record into Dictionary data
# print('Excel Sheet to Csv:', excel_data.to_csv(orient='records')) #Convert Excel record into CSV data