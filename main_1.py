import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import xlsxwriter

# folder_path = './Generated_Files'
folder_path = './files'
output_path = './output'
output_file = '_output.xlsx'

# course_code =['CSPC601', 'CSPC602', 'CSPC603', 'CSPC604', 'CSPC605', 'CSPC606']

course_code = []
marks = []

file_list = os.listdir(folder_path)
excel_files = [file for file in file_list if file.endswith('.xlsx') or file.endswith('.xls')]
if not os.path.exists(output_path):
    os.makedirs(output_path)
    
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path)
    course_code.append(df.iloc[7, 3])
    marks.append(df.iloc[8, 3])
print(excel_files)
print(marks)

# Assuming 'excel_files' is a list of file names and 'folder_path' is the path to the folder containing the files
for i, file in enumerate(excel_files):
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path, skiprows=10) 
    if 'Unnamed: 4' in df:
        df = df.drop(['Unnamed: 4'], axis=1)
    df = df.dropna()      
    
    course_detail = pd.read_excel(file_path, nrows=9, header=None)
    course_detail = course_detail.iloc[:8, 1:].fillna('').values.tolist()
    
    mark_40 = marks[i]*0.4
    mark_75 = marks[i]*0.75
    # mark_40 = 16
    # mark_75 = 30
    total_students = len(df)
    total_students_appeared = len(df[df['Marks'] != 'Absent'])
    total_absent = total_students - total_students_appeared
    avg_marks = df[df['Marks'] != 'Absent']['Marks'].astype(float).mean()
    less_than_16 = len(df[df['Marks'].astype(float) < mark_40])
    between_16_and_30 = len(df[(df['Marks'].astype(float) >= mark_40) & (df['Marks'].astype(float) < mark_75)])
    more_than_30 = len(df[df['Marks'].astype(float) >= mark_75])

    # Create a new workbook for each input file
    output_file_name = f'{os.path.splitext(file)[0]}_output.xlsx'
    wb = xlsxwriter.Workbook(os.path.join(output_path, output_file_name))
    ws1 = wb.add_worksheet('Student Summary')

    title = "Staff Report"
    title_range = 'A1:F2'
    title_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 20, 'bg_color': '#FFFF00', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
    header_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 12, 'bg_color': '#B0E0E6', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
    cell_format = wb.add_format({'border': 2, 'align': 'center', 'valign': 'vcenter', 'bold': True})

    ws1.merge_range(title_range, title, title_format)

    for i, detail in enumerate(course_detail):
        ws1.write_row(i + 2, 0, detail, cell_format)

    headers = ["Attribute", "Value"]
    ws1.write_row('A13', headers, header_format)
    ws1.set_column(0, 2, 25)
    ws1.set_row(0, 35)
    ws1.set_row(3, 35)

    results = [
        ["Total Students", total_students],
        ["Total Students Appeared", total_students_appeared],
        ["Total Absent", total_absent],
        ["Average Marks & %", avg_marks],
        ["Less Than 15", less_than_16],
        ["Between 15-30", between_16_and_30],
        ["More than 30", more_than_30]
    ]
    for i, (attribute, value) in enumerate(results, start=13):
        ws1.write(i, 0, attribute, cell_format)
        ws1.write(i, 1, value, cell_format)

    mark_less_than = float(input(f"Enter a mark to filter students who scored less than that mark for file {file}: "))

    filtered_less_than_df = df[df['Marks'].astype(float) < mark_less_than]
    filtered_less_than_df = filtered_less_than_df.sort_values(by='Marks', ascending=True)
    if 'Unnamed: 4' in filtered_less_than_df:
        filtered_less_than_df = filtered_less_than_df.drop(['Unnamed: 4'], axis=1)
    ws2 = wb.add_worksheet('Slow Learners')
    header_row = filtered_less_than_df.columns.tolist()
    for col, header in enumerate(header_row):
        ws2.write(0, col, header, header_format)

    # Write data rows to the sheet
    for row_idx, row_data in enumerate(filtered_less_than_df.values, start=1):
        for col_idx, cell_data in enumerate(row_data):
            ws2.write(row_idx, col_idx, cell_data)

    mark_more_than = float(input(f"Enter a mark to filter students who scored more than that mark for file {file}: "))

    filtered_more_than_df = df[df['Marks'].astype(float) > mark_more_than]
    filtered_more_than_df = filtered_more_than_df.sort_values(by='Marks', ascending=False)
    if 'Unnamed: 4' in filtered_more_than_df:
        filtered_more_than_df = filtered_more_than_df.drop(['Unnamed: 4'], axis=1)
    ws3 = wb.add_worksheet('Fast Learners')
    header_row = filtered_more_than_df.columns.tolist()
    for col, header in enumerate(header_row):
        ws3.write(0, col, header, header_format)

    # Write data rows to the sheet
    for row_idx, row_data in enumerate(filtered_more_than_df.values, start=1):
        for col_idx, cell_data in enumerate(row_data):
            ws3.write(row_idx, col_idx, cell_data)
    wb.close()
    print(f"Results saved successfully to: ./output/{os.path.splitext(file)[0]}_output.xlsx")
    
output_folder_path = './output'
output_files = [file for file in os.listdir(output_folder_path) if file.endswith('.xlsx')]

data_sheet1 = []
for idx, file in enumerate(output_files):
    file_path = os.path.join(output_folder_path, file)
    df = pd.read_excel(file_path, sheet_name='Student Summary',skiprows=12, header=0)
    df = df.drop(columns='Unnamed: 2', axis=1) if 'Unnamed: 2' in df else df
    # print(df['Value'].values.tolist())
    data_sheet1.append([course_code[idx]] + df['Value'].values.tolist())
# print(data_sheet1)

columns_sheet1 = ['Course Code', 'Total Students', 'Total Students Appeared', 'Total Absent', 
                  'Average Marks', 'Students Less than 16', 
                  'Students Between 16 and 30', 'Students More than 30']

df_sheet1_transposed = pd.DataFrame(data_sheet1, columns=columns_sheet1).T.reset_index()
df_sheet1_transposed.columns = df_sheet1_transposed.iloc[0]
df_sheet1_transposed = df_sheet1_transposed.drop(0)
print(df_sheet1_transposed)

sheet2_roll_counts = {}
sheet3_roll_counts = {}

df_combined_sheet2 = pd.DataFrame()
df_combined_sheet3 = pd.DataFrame()

for file in output_files:
    file_path = os.path.join(output_folder_path, file)

    df_sheet2 = pd.read_excel(file_path, sheet_name='Slow Learners', header=0)
    df_sheet3 = pd.read_excel(file_path, sheet_name='Fast Learners', header=0)

    df_combined_sheet2 = pd.concat([df_combined_sheet2, df_sheet2], ignore_index=True)
    df_combined_sheet3 = pd.concat([df_combined_sheet3, df_sheet3], ignore_index=True)

for idx, row in df_combined_sheet2.iterrows():
    sheet2_roll_counts[row['Roll No.']] = sheet2_roll_counts.get(row['Roll No.'], {'Count': 0, 'Name': 'Unknown'})
    sheet2_roll_counts[row['Roll No.']]['Count'] += 1
    sheet2_roll_counts[row['Roll No.']]['Name'] = row['Student Name']

for idx, row in df_combined_sheet3.iterrows():
    sheet3_roll_counts[row['Roll No.']] = sheet3_roll_counts.get(row['Roll No.'], {'Count': 0, 'Name': 'Unknown'})
    sheet3_roll_counts[row['Roll No.']]['Count'] += 1
    sheet3_roll_counts[row['Roll No.']]['Name'] = row['Student Name']

consolidated_sheet2 = sorted(sheet2_roll_counts.items(), key=lambda x: x[1]['Count'], reverse=True)
consolidated_sheet3 = sorted(sheet3_roll_counts.items(), key=lambda x: x[1]['Count'], reverse=True)

df_consolidated_sheet2 = pd.DataFrame([(roll_no, data['Name'], data['Count']) for roll_no, data in consolidated_sheet2], columns=['Roll No.', 'Student Name', 'Count'])
df_consolidated_sheet3 = pd.DataFrame([(roll_no, data['Name'], data['Count']) for roll_no, data in consolidated_sheet3], columns=['Roll No.', 'Student Name', 'Count'])

top_n = int(input("Enter the count of students: "))

df_consolidated_sheet2_count = df_consolidated_sheet2[df_consolidated_sheet2['Count'] >= top_n]
df_consolidated_sheet3_count = df_consolidated_sheet3[df_consolidated_sheet3['Count'] >= top_n]

output_excel_path = './combined_output_new.xlsx'
with pd.ExcelWriter(output_excel_path) as writer:
    df_sheet1_transposed.to_excel(writer, sheet_name='Student Summary', index=False)
    df_consolidated_sheet2_count.to_excel(writer, sheet_name='Slow Learners', index=False)
    df_consolidated_sheet3_count.to_excel(writer, sheet_name='Fast Learners', index=False)

print("Data saved successfully to", output_excel_path)