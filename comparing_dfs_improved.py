import pandas as pd
import sys
import time

# this script compares 2 dfs that can be imported from excel workbooks with multiple sheets 
# the sheetnames of the desired dfs to be compared are provided

timestr = time.strftime("%Y%m%d%H%M") #optional how many details you include on the datetime: you could only include month_year ("%Y%m")
#import data 
sheet_dict1 = pd.read_excel(sys.argv[1], sheet_name=None)
sheet_dict2 = pd.read_excel(sys.argv[3], sheet_name=None)
sheetname1 = sys.argv[2]
sheetname2 = sys.argv[4]
df1 = sheet_dict1.get(sheetname1)
df2 = sheet_dict2.get(sheetname2)
# Clean up data 
df1.dropna(how='all', inplace=True) # drop the rows where in all columns of the df the values are NaN, so keep rows that NaN are only in some columns
df2.dropna(how='all', inplace=True)
# fill NaN values with desired value (possibly with a dict to fill different columns with different values)
df1.fillna(value=0, inplace=True)
df2.fillna(value=0, inplace=True)

initial_df1 = df1['HiX#'].size
initial_df2 = df2['HiX#'].size
#compare the dfs
mutual = df1[df1['HiX#'].isin(df2['HiX#'])] # mutual patients
df1only = df1[~df1['HiX#'].isin(df2['HiX#'])] # patients unique in dataframe1
df2only = df2[~df2['HiX#'].isin(df1['HiX#'])] # patients unique in dataframe2

mutual_patients = mutual['HiX#'].size
rest_df1 = initial_df1 - mutual_patients
rest_df2 = initial_df2 - mutual_patients
# print some valuable info
print("Initial patients df1:",initial_df1,"\nInitial patients df2:",initial_df2,"\nMutual patients:",mutual_patients,"\n Rest df1:", rest_df1, "\n Rest df2:", rest_df2)
# store the dfs into a dictionary
frames = {'Mutual_patients': mutual, 'bs': df1only, '6m_prior_bs_def_maybe': df2only}

#loop through the dict of dfs and put each on a specific sheet of an excel workbook
#in the filename we include the date the file was created
with pd.ExcelWriter('BL_samples_compared_'+ timestr +'_.xlsx') as writer:
    for sheet, frame in frames.items():  # .use .items for python 3.X
        frame.to_excel(writer, sheet_name=sheet, index=False)
