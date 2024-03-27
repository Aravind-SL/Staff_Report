import os
import pandas as pd
import random

def generate_random_marks(row):
    return random.randint(0, 40)


folder_path1 = './Generated_Files/B.E_(CSE)_III YEAR-'
folder_path2 = '.xlsx'
output_path = './Generated_Files'

coursecode = ['CSPC601', 'CSPC602', 'CSPC603', 'CSPC604', 'CSPC605', 'CSPC606']
if not os.path.exists(output_path):
    os.makedirs(output_path)

    
for i in range(1, 7):
    df = pd.read_excel('./files/B.E_(CSE)_III YEAR- CSPC602-TEST1-MARKS.xls', skiprows=10, header=0)
    df['Marks'] = df.apply(lambda row: generate_random_marks(row), axis=1)
    df = df.drop(['Unnamed: 4'], axis=1)
    final_path = folder_path1+coursecode[i-1]+folder_path2
    df.to_excel(final_path, index=False)