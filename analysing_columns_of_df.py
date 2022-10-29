# this scripts analyses mupliple columns of a given dataframe of an excel workbook
# and reports the results in an excel workbook where each sheet corresponds to a certain column
# input: the excel workbook, the sheetname of the workbook to be analysed and the desired columns to be analyzed
import pandas as pd
import time
timestr = time.strftime("%Y%m%d%H%M")
#Inputs
sheet_dict = pd.read_excel(input("Enter tha excel workbook:"), sheet_name=None)
sheetname = input("Enter the desired sheetname:")
columns = list(map(str, input("Enter column names separated by comma:").split(','))) #creates a list of the columns that the user gave
data = sheet_dict.get(sheetname) 
df_list = [] # create an empty list to put our new dfs

# for each column of the original dataframe count the number of unique values and store the results in a new dataframe and that in a list of dfs 
for column in columns:
    df = data[column].value_counts().to_frame().reset_index()
    df_list.append(df)
    
# create a dictionary from with the list of columns as keys and the list of dataframes as values 
frames_dict = dict(zip(columns, df_list))

# write an excel book with the keys (columns) of the dictionary being the sheetname and the values (dataframes) being the data
with pd.ExcelWriter(sheetname +'_column_analysis'+ timestr +'_.xlsx') as writer:
    for sheet, frame in frames_dict.items():
        frame.to_excel(writer, sheet_name=sheet, index=False)

