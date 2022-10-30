# this scripts analyzes mupliple columns of a given dataframe of an excel workbook and reports the results in an excel workbook 
# where each sheet corresponds to concerns a certain column
# input: the excel workbook, the sheetname of the df to be analysed and the desired columns to be analyzed
import pandas as pd
import time
import os
path = os.getcwd() #get the absolute path to the current working directory
files = os.listdir(path) #list of files in the cwd
excel_files = [f for f in files if f[-4:] == 'xlsx'] #only excel files in the cwd
timestr = time.strftime("%Y%m%d%H%M") #time string for the output file

#input the excel workbook and catch errors if the user gave wrong filename
while True:
    try:
        workbook = input("Enter the excel workbook:")
        if workbook in excel_files:
            break
        else:
            print(workbook,'is not a valid name.\nValid files are:', excel_files)
    except Exception:
        print(workbook,'is not a valid name.\nValid files are:', excel_files)

#create a dictionary of the imported workbook with keys the sheetnames and values the dataframes in each sheet
sheet_dict = pd.read_excel(workbook, sheet_name=None)
#input the sheetname for the desired data to be analysed and catch for errors if the user gave wrong sheetname
while True:
    sheetname = input("Enter sheetname:")
    if sheetname not in sheet_dict.keys():
        print("Sheetname not valid.\nValid sheetnames are:", sheet_dict.keys())
    else:
        break

#assign the desired df of the workbook based on the given sheetname to a data variable
data = sheet_dict.get(sheetname)
# Clean up data
data.dropna(how='all', inplace=True) # drop the rows where in all columns of the df the values are NaN, so keep rows that NaN are only in some columns
data.fillna(value=0, inplace=True)
# input the columns to be analysed and catch for errors if the user gave columns not present in the dataframe
while True:
    columns = list(map(str, input("Enter column names separated by comma:").split(','))) #creates a list of the columns that the user gave
    column_problem = False
    for i in columns:
        if i not in data.columns:
            print(i, 'is not a valid column name.\nValid column names are:', data.columns)
            column_problem = True
            break
    if not column_problem:
        break

df_list = [] # create an empty list to put our new dfs

# for each column of the original dataframe count the number of unique values and store the results in a new dataframe and that in a list of dfs
for column in columns:
    df = data[column].value_counts().to_frame().reset_index()
    df_list.append(df)

# create a dictionary with the list of columns as keys and the list of dataframes as values
frames_dict = dict(zip(columns, df_list))

# write an excel book with the keys (columns) of the dictionary being the sheetname and the values (dataframes) being the data
with pd.ExcelWriter(sheetname +'_column_analysis'+ timestr +'_.xlsx') as writer:
    for sheet, frame in frames_dict.items():
        frame.to_excel(writer, sheet_name=sheet, index=False)
  
#give feedback to the user that all went right
print('All done!')

