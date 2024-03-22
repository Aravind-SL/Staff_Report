import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


folder_path = './files'
output_path = './output'
output_file = '_output.xlsx'
file_list = os.listdir(folder_path)
excel_files = [file for file in file_list if file.endswith('.xlsx') or file.endswith('.xls')]

for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path, skiprows=10, header=0)

    total_students = len(df)
    total_students_appeared = len(df[df['Marks'] != 'Absent'])
    total_absent = total_students - total_students_appeared
    avg_marks = df[df['Marks'] != 'Absent']['Marks'].astype(float).mean()
    less_than_15 = len(df[df['Marks'].astype(float) < 15])
    between_15_and_30 = len(df[(df['Marks'].astype(float) > 15) & (df['Marks'].astype(float) <= 30)])
    more_than_30 = len(df[df['Marks'].astype(float) > 30])

    results = [{
        'Total Students': total_students,
        'Total Students Appeared': total_students_appeared,
        'Total Absent': total_absent,
        'Average Marks': avg_marks,
        'Students Less than 15': less_than_15,
        'Students Between 15 and 30': between_15_and_30,
        'Students More than 30': more_than_30
    }]
    results_df = pd.DataFrame(results)


    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'Sheet1'


    for r in dataframe_to_rows(results_df, index=False, header=True):
        ws1.append(r)


    mark_less_than = float(input(f"Enter a mark to filter students who scored less than that mark for file {file}: "))

    filtered_less_than_df = df[df['Marks'].astype(float) < mark_less_than]

    ws2 = wb.create_sheet('Sheet2')
    for r in dataframe_to_rows(filtered_less_than_df, index=False, header=True):
        ws2.append(r)

    mark_more_than = float(input(f"Enter a mark to filter students who scored more than that mark for file {file}: "))

    filtered_more_than_df = df[df['Marks'].astype(float) > mark_more_than]

    ws3 = wb.create_sheet('Sheet3')
    for r in dataframe_to_rows(filtered_more_than_df, index=False, header=True):
        ws3.append(r)


    output_file_path = os.path.join(output_path, file.replace('.xlsx', '') + output_file)
    wb.save(output_file_path)

    print("Results saved successfully to:", output_file_path)
