# compare 2datasheets of the same excel workbook on a single column
# input in command line: the excel workbook
# input during the running of the script: sheetnames to be compared and the column name that they will be compared on
# output: excel workbook in the cwd with 3 sheets: the mutual values, unique in df1, unique in df2
import pandas as pd
import os
import time
path = os.getcwd() #get the absolute path to the current working directory
files = os.listdir(path) #list of files in the cwd
excel_files = [f for f in files if f[-4:] == 'xlsx'] #only excel files in the cwd
timestr = time.strftime("%Y%m%d%H%M") #time string for the output file

#request data
while True:
    try:
        workbook = input("Enter the excel workbook:")
        if workbook in excel_files:
            break
        else:
            print(workbook,'is not a valid name.\nValid files are:', excel_files)
    except Exception:
        print(workbook,'is not a valid name.\nValid files are:', excel_files)

sheet_dict = pd.read_excel(workbook, sheet_name=None)
while True:
    sheetname1 = input("Enter 1st sheetname containing data for comparison:")
    if sheetname1 not in sheet_dict.keys():
        print("Sheetname not valid.\nValid sheetnames are:", sheet_dict.keys())
    else:
        break
while True:
    sheetname2 = input("Enter 2nd sheetname containing data for comparison:")
    if sheetname2 not in sheet_dict.keys():
        print("Sheetname not valid.\nValid sheetnames are:", sheet_dict.keys())
    else:
        break

#assign the values of the dict to the corresponding dfs based on the sheetname
df1 = sheet_dict.get(sheetname1)
df2 = sheet_dict.get(sheetname2)
while True:
    column_to_compare = input("Enter mutual column name to compare the dfs:")
    if column_to_compare not in df1.columns and column_to_compare not in df2.columns:
        print("Sheetname not valid. Try again using the mutual column name of the dfs")
    else:
        break

# Clean up data
df1.dropna(how='all', inplace=True) # drop the rows where in all columns of the df the values are NaN, so keep rows that NaN are only in some columns
df2.dropna(how='all', inplace=True)
# fill NaN values with desired value (possibly with a dict to fill different columns with different values)
df1.fillna(value=0, inplace=True)
df2.fillna(value=0, inplace=True)

initial_df1 = df1[column_to_compare].size
initial_df2 = df2[column_to_compare].size
#compare the dfs
mutual = df1[df1[column_to_compare].isin(df2[column_to_compare])] # mutual patients
df1only = df1[~df1[column_to_compare].isin(df2[column_to_compare])] # patients unique in dataframe1
df2only = df2[~df2[column_to_compare].isin(df1[column_to_compare])] # patients unique in dataframe2

mutual_patients = mutual[column_to_compare].size
rest_df1 = initial_df1 - mutual_patients
rest_df2 = initial_df2 - mutual_patients
# print some valuable info
print("Initial patients",sheetname1,":",initial_df1,"\nInitial patients",sheetname2,":",initial_df2,"\nMutual patients:",mutual_patients,"\n Rest df1:", rest_df1, "\n Rest df2:", rest_df2)
# store the dfs into a dictionary
frames = {'Mutual_patients': mutual, sheetname1: df1only, sheetname2: df2only}

#loop through the dict of dfs and put each on a specific sheet of an excel workbook
#in the filename we include the date the file was created
with pd.ExcelWriter('BL_samples_compared_'+ timestr +'_.xlsx') as writer:
    for sheet, frame in frames.items():
        frame.to_excel(writer, sheet_name=sheet, index=False)

print("All done!")
