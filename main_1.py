import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT


folder_path = './files'
output_path = './output'
output_file = '_output.xlsx'
course_code = []
file_list = os.listdir(folder_path)
excel_files = [file for file in file_list if file.endswith('.xlsx') or file.endswith('.xls')]
if not os.path.exists(output_path):
    os.makedirs(output_path)
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path)
    course_code.append(df.iloc[7, 3])
    
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    df = pd.read_excel(file_path, skiprows=10, header=0)

    total_students = len(df)
    total_students_appeared = len(df[df['Marks'] != 'Absent'])
    total_absent = total_students - total_students_appeared
    avg_marks = round(df[df['Marks'] != 'Absent']['Marks'].astype(float).mean(), 2)
    less_than_16 = len(df[df['Marks'].astype(float) < 16])
    between_16_and_30 = len(df[(df['Marks'].astype(float) >= 16) & (df['Marks'].astype(float) <= 30)])
    more_than_30 = len(df[df['Marks'].astype(float) > 30])

    results = [{
        'Total Students': total_students,
        'Total Students Appeared': total_students_appeared,
        'Total Absent': total_absent,
        'Average Marks': avg_marks,
        'Students Less than 15': less_than_16,
        'Students Between 15 and 30': between_16_and_30,
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
    filtered_less_than_df = filtered_less_than_df.sort_values(by='Marks', ascending=True)
    if 'Unnamed: 4' in filtered_less_than_df:
        filtered_less_than_df= filtered_less_than_df.drop(['Unnamed: 4'], axis=1)
    ws2 = wb.create_sheet('Sheet2')
    for r in dataframe_to_rows(filtered_less_than_df, index=False, header=True):
        ws2.append(r)

    mark_more_than = float(input(f"Enter a mark to filter students who scored more than that mark for file {file}: "))

    filtered_more_than_df = df[df['Marks'].astype(float) > mark_more_than]
    filtered_more_than_df = filtered_more_than_df.sort_values(by='Marks', ascending=False)
    if 'Unnamed: 4' in filtered_more_than_df:
        filtered_more_than_df= filtered_more_than_df.drop(['Unnamed: 4'], axis=1)
    ws3 = wb.create_sheet('Sheet3')
    for r in dataframe_to_rows(filtered_more_than_df, index=False, header=True):
        ws3.append(r)


    output_file_path = os.path.join(output_path, file.replace('.xlsx', '') + output_file)
    wb.save(output_file_path)

    print("Results saved successfully to:", output_file_path)
    
# Task 2    
output_folder_path = './output'
output_files = [file for file in os.listdir(output_folder_path) if file.endswith('.xlsx')]

data_sheet1 = []
for idx, file in enumerate(output_files):
    file_path = os.path.join(output_folder_path, file)
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=0)
    data_sheet1.append([course_code[idx]] + df.iloc[0].values.tolist())

columns_sheet1 = ['Course Code', 'Total Students', 'Total Students Appeared', 'Total Absent', 
                  'Average Marks', 'Students Less than 15', 
                  'Students Between 15 and 30', 'Students More than 30']

# Transpose df_sheet1
df_sheet1_transposed = pd.DataFrame(data_sheet1, columns=columns_sheet1).T.reset_index()
df_sheet1_transposed.columns = df_sheet1_transposed.iloc[0]
df_sheet1_transposed = df_sheet1_transposed.drop(0)
sheet2_roll_counts = {}
sheet3_roll_counts = {}

df_combined_sheet2 = pd.DataFrame()
df_combined_sheet3 = pd.DataFrame()

for file in output_files:
    file_path = os.path.join(output_folder_path, file)

    df_sheet2 = pd.read_excel(file_path, sheet_name='Sheet2', header=0)
    df_sheet3 = pd.read_excel(file_path, sheet_name='Sheet3', header=0)

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

top5_sheet2 = sorted(sheet2_roll_counts.items(), key=lambda x: x[1]['Count'], reverse=True)
top5_sheet3 = sorted(sheet3_roll_counts.items(), key=lambda x: x[1]['Count'], reverse=True)

df_top5_sheet2 = pd.DataFrame([(roll_no, data['Name'], data['Count']) for roll_no, data in top5_sheet2], columns=['Roll No.', 'Student Name', 'Count'])
df_top5_sheet3 = pd.DataFrame([(roll_no, data['Name'], data['Count']) for roll_no, data in top5_sheet3], columns=['Roll No.', 'Student Name', 'Count'])
# Create a new Word document
doc = Document()

# Add a title to the document
doc.add_heading('Data Summary', level=1)

# Function to add a dataframe as a table to the document
def add_dataframe_table(document, dataframe, table_title):
    # Add table title
    doc.add_heading(table_title, level=2)

    # Add a table to the document
    table = doc.add_table(rows=1, cols=len(dataframe.columns))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Add column headers to the table
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(dataframe.columns):
        hdr_cells[i].text = column

    # Add data to the table
    for _, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        for i, cell_value in enumerate(row):
            row_cells[i].text = str(cell_value)

# Add df_sheet1_transposed as a table to the document
add_dataframe_table(doc, df_sheet1_transposed, 'Course Summary')

# Add df_top5_sheet2 as a table to the document
add_dataframe_table(doc, df_top5_sheet2, 'Slow Learners')

# Add df_top5_sheet3 as a table to the document
add_dataframe_table(doc, df_top5_sheet3, 'Fast Learners')

# Save the document
doc.save('data_summary.docx')

