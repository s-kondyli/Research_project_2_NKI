# compare 2datasheets of the same excel workbook on a single column
# input in command line: the excel workbook
# input during the running of the script: sheetnames to be compared and the column name that they will be compared on
# output: excel workbook in the cwd with 3 sheets: the mutual values, unique in df1, unique in df2


import pandas as pd
import sys
import time
timestr = time.strftime("%Y%m%d%H%M") #optional how many details you include on the datetime: you could only include month_year ("%Y%m")

#request data 
sheetname1 = input('First sheetname for the comparison:')
sheetname2 = input('Second sheetname for the comparison:')
column_to_compare = input('column name to compare the files:')

#import data
sheet_dict = pd.read_excel(sys.argv[1], sheet_name=None)
#assign the values of the dict to the corresponding dfs based on the sheetname
df1 = sheet_dict.get(sheetname1)
df2 = sheet_dict.get(sheetname2)
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
