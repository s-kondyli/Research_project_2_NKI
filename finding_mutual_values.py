import pandas as pd
import sys
# select values from a dataframe based on a list 
data = pd.read_excel(sys.argv[1])
data.fillna(value=0, inplace=True)
initial_patients = data['HiX#'].size
df = pd.read_excel(sys.argv[2], names=['patient'])

check = df['patient'].tolist()
data = data[data['HiX#'].isin(check)] #select only the values of the dataframe that are in the list 
mutual_patients = data['HiX#'].size
difference = initial_patients - mutual_patients
print("Initial patients:",initial_patients,"\nMutual patients:",mutual_patients,"\nDifference:",difference )
data.to_excel(excel_writer='mutual_patients.xlsx', index=False, sheet_name='Mutual_Data')


