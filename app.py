import os
import pandas as pd
import streamlit as st
from openpyxl import Workbook
import xlsxwriter
from io import BytesIO
from zipfile import ZipFile

st.title("Excel File Processor")

# Step 1: File Upload
uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    course_code = []
    marks = []
    output_path = './output'

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    output_files = []
    zip_buffer = BytesIO()

    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)
        course_code.append(df.iloc[7, 3])
        marks.append(df.iloc[8, 3])

    st.write("Processing files:", [file.name for file in uploaded_files])

    for i, uploaded_file in enumerate(uploaded_files):
        df = pd.read_excel(uploaded_file, skiprows=10)
        if 'Unnamed: 4' in df:
            df = df.drop(['Unnamed: 4'], axis=1)
        df = df.dropna()

        course_detail = pd.read_excel(uploaded_file, nrows=10, header=None)
        course_detail = course_detail.iloc[:10, 1:].fillna('').values.tolist()

        mark_40 = marks[i] * 0.4
        mark_75 = marks[i] * 0.75

        total_students = len(df)
        total_students_appeared = len(df[df['Marks'] != 'A'])
        total_absent = total_students - total_students_appeared
        avg_marks = df[df['Marks'] != 'A']['Marks'].astype(float).mean().round(2)
        df = df.replace('A', 0)
        less_than_16 = len(df[df['Marks'].astype(float) < mark_40]) - total_absent
        between_16_and_30 = len(df[(df['Marks'].astype(float) >= mark_40) & (df['Marks'].astype(float) < mark_75)]) - total_absent
        more_than_30 = len(df[df['Marks'].astype(float) >= mark_75]) - total_absent

        # Creating output Excel file
        output_file_name = f'{os.path.splitext(uploaded_file.name)[0]}_output.xlsx'
        wb = xlsxwriter.Workbook(os.path.join(output_path, output_file_name))
        ws1 = wb.add_worksheet('Student Summary')

        # Format and write data
        title_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 20, 'bg_color': '#FFFF00', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
        header_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 12, 'bg_color': '#B0E0E6', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
        cell_format = wb.add_format({'border': 2, 'align': 'center', 'valign': 'vcenter', 'bold': True})

        ws1.merge_range('A1:F2', "Staff Report", title_format)
        for i, detail in enumerate(course_detail):
            ws1.write_row(i + 2, 0, detail, cell_format)

        headers = ["Attribute", "Value"]
        ws1.write_row('A13', headers, header_format)
        ws1.set_column(0, 2, 25)

        results = [
            ["Total Students", total_students],
            ["Total Students Appeared", total_students_appeared],
            ["Total Absent", total_absent],
            ["Average Marks", avg_marks],
            ["Less Than 40%", less_than_16],
            ["Between 40% - 75%", between_16_and_30],
            ["More than 75%", more_than_30]
        ]
        for i, (attribute, value) in enumerate(results, start=13):
            ws1.write(i, 0, attribute, cell_format)
            ws1.write(i, 1, value, cell_format)

        mark_less_than = st.number_input(
            f"Enter a mark to filter students who scored less than that mark for file {uploaded_file.name}:",
            min_value=0.0,  # Changed to float
            value=float(mark_40)  # Ensure value is a float
        )
        filtered_less_than_df = df[df['Marks'].astype(float) < mark_less_than]
        filtered_less_than_df = filtered_less_than_df.sort_values(by='Marks', ascending=True)
        ws2 = wb.add_worksheet('Slow Learners')
        for col, header in enumerate(filtered_less_than_df.columns.tolist()):
            ws2.write(0, col, header, header_format)
        for row_idx, row_data in enumerate(filtered_less_than_df.values, start=1):
            for col_idx, cell_data in enumerate(row_data):
                ws2.write(row_idx, col_idx, cell_data)

        mark_more_than = st.number_input(
            f"Enter a mark to filter students who scored more than that mark for file {uploaded_file.name}:",
            min_value=0.0,  # Changed to float
            value=float(mark_75)  # Ensure value is a float
        )
        filtered_more_than_df = df[df['Marks'].astype(float) > mark_more_than]
        filtered_more_than_df = filtered_more_than_df.sort_values(by='Marks', ascending=False)
        ws3 = wb.add_worksheet('Fast Learners')
        for col, header in enumerate(filtered_more_than_df.columns.tolist()):
            ws3.write(0, col, header, header_format)
        for row_idx, row_data in enumerate(filtered_more_than_df.values, start=1):
            for col_idx, cell_data in enumerate(row_data):
                ws3.write(row_idx, col_idx, cell_data)

        wb.close()
        st.write(f"Results saved successfully to: {output_file_name}")
        output_files.append(output_file_name)

        # Add individual output file to zip
        with ZipFile(zip_buffer, 'a') as zip_file:
            zip_file.write(os.path.join(output_path, output_file_name), output_file_name)

    st.success("Processing Complete!")

    # Process the combined output
    data_sheet1 = []
    for idx, file in enumerate(output_files):
        file_path = os.path.join(output_path, file)
        df = pd.read_excel(file_path, sheet_name='Student Summary', skiprows=12, header=0)
        df = df.drop(columns='Unnamed: 2', axis=1) if 'Unnamed: 2' in df else df
        data_sheet1.append([course_code[idx]] + df['Value'].values.tolist())

    columns_sheet1 = ['Course Code', 'Total Students', 'Total Students Appeared', 'Total Absent', 
                      'Average Marks', 'Students Less than 40%', 
                      'Students Between 40% and 75%', 'Students More than 75%']

    df_sheet1_transposed = pd.DataFrame(data_sheet1, columns=columns_sheet1).T.reset_index()
    df_sheet1_transposed.columns = df_sheet1_transposed.iloc[0]
    df_sheet1_transposed = df_sheet1_transposed.drop(0)

    sheet2_roll_counts = {}
    sheet3_roll_counts = {}

    df_combined_sheet2 = pd.DataFrame()
    df_combined_sheet3 = pd.DataFrame()

    for file in output_files:
        file_path = os.path.join(output_path, file)

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

    top_n = st.number_input("Enter the count of students:", min_value=1, value=1, step=1)

    df_consolidated_sheet2_count = df_consolidated_sheet2[df_consolidated_sheet2['Count'] >= top_n]
    df_consolidated_sheet3_count = df_consolidated_sheet3[df_consolidated_sheet3['Count'] >= top_n]

    output_excel_path = './combined_output_new.xlsx'
    with pd.ExcelWriter(output_excel_path) as writer:
        df_sheet1_transposed.to_excel(writer, sheet_name='Student Summary', index=False)
        df_consolidated_sheet2_count.to_excel(writer, sheet_name='Slow Learners', index=False)
        df_consolidated_sheet3_count.to_excel(writer, sheet_name='Fast Learners', index=False)

    # Add combined output to zip
    with ZipFile(zip_buffer, 'a') as zip_file:
        zip_file.write(output_excel_path, os.path.basename(output_excel_path))

    # Provide download link for zip file
    zip_buffer.seek(0)
    st.download_button(
        label="Download ZIP",
        data=zip_buffer,
        file_name="processed_files.zip",
        mime="application/zip"
    )

    st.success("Processing Complete!")
