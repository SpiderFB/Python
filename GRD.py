import pandas as pd
import time

# Path to the excel file

DataValidationPath = 'C:/Users/2095421/Downloads/Migration/GRD/VALIDATION.xlsx'

start_time = time.time()

# Read the excel file into two dataframes
df1 = pd.read_excel(DataValidationPath, sheet_name='MDG')
# df1 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG))
df2 = pd.read_excel(DataValidationPath, sheet_name='ATLAS')
# df2 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS))

# Set 'MATNR' as the index for both dataframes
df1.set_index('MATNR', inplace=True)
df2.set_index('MATNR', inplace=True)

# Find the common indices between df1 and df2
common_indices = df1.index.intersection(df2.index)

df3 = pd.read_excel(DataValidationPath, sheet_name="Sheet1")
df_scope = df3[df3['Comment'] == 'Scope']
scope_field = df_scope['Technical Name']


# Define the scope fields
# scope_field = ['/VSO/R_BOT_IND']

# Create a dictionary to hold the mismatched dataframes
mismatched_dataframes = {}

# Iterate over each field in the scope
for col_name in scope_field:
    if col_name not in df1.columns:
        print(col_name, "--------------Not Found")
    else:
        print(col_name)
        # Compare the FIELD values of df1 and df2 for the common indices
        # comparison = df1.loc[common_indices, col_name] == df2.loc[common_indices, col_name]
        comparison = df1.loc[common_indices, col_name].eq(df2.loc[common_indices, col_name]) | (df1.loc[common_indices, col_name].isna() & df2.loc[common_indices, col_name].isna())



        # If they don't match, create a new dataframe with the mismatched indices
        if not comparison.all():
            mismatched_indices = common_indices[~comparison]
            df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Value in df2': df2.loc[mismatched_indices, col_name].squeeze(),
                'Value in df1': df1.loc[mismatched_indices, col_name].squeeze()
            })
            # Add the mismatched dataframe to the dictionary
            mismatched_dataframes[col_name] = df_mismatch



# Save all the mismatched dataframes to new sheets in the excel file
with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a') as writer:
    for sheet_name, df_mismatch in mismatched_dataframes.items():
        df_mismatch.to_excel(writer, sheet_name=sheet_name.replace("/",""))

# Print a success message
print("All mismatched dataframes have been written to the excel file.")

end_time = time.time()
execution_time = end_time - start_time
print(f"Done in -- {execution_time} seconds")



# # Find the indices that are in df1 but not in df2, and vice versa
# df1_not_df2 = df1.index.difference(df2.index)
# df2_not_df1 = df2.index.difference(df1.index)

# # Create new dataframes with these indices
# df1_not_df2_df = pd.DataFrame({
#     'Index': df1_not_df2,
#     'In df1, not in df2': 'yes'
# })

# df2_not_df1_df = pd.DataFrame({
#     'Index': df2_not_df1,
#     'In df2, not in df1': 'yes'
# })

# # Save these dataframes to new sheets in the excel file
# with pd.ExcelWriter('C:/Users/2095421/Downloads/Migration/GRD/VALIDATION.xlsx', engine='openpyxl', mode='a') as writer:
#     df1_not_df2_df.to_excel(writer, sheet_name='In df1 not df2')
#     df2_not_df1_df.to_excel(writer, sheet_name='In df2 not df1')