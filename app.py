import os
import pandas as pd
import streamlit as st
from io import BytesIO
from zipfile import ZipFile
import xlsxwriter
import math

# Function to process each uploaded file
def process_file(uploaded_file, marks):
    custom_headers = ['S.No.', 'Roll No.', 'Student Name', 'Marks']
    df = pd.read_excel(uploaded_file, skiprows=10, names=custom_headers, usecols="A:D") 
    if 'Unnamed: 4' in df:
        df = df.drop(['Unnamed: 4'], axis=1)
    df = df.dropna()

    course_detail = pd.read_excel(uploaded_file, nrows=10, header=None).iloc[:10, 1:].fillna('').values.tolist()
    mark_40 = marks * 0.4
    mark_75 = marks * 0.75
    total_students = len(df)
    total_students_appeared = len(df[df['Marks'] != 'A'])
    total_absent = total_students - total_students_appeared
    df=df[df['Marks']!='A']
    avg_marks = df['Marks'].astype(float).mean()
    avg_marks = round(avg_marks, 2)
    less_than_16 = len(df[df['Marks'].astype(float) < mark_40]) 
    between_16_and_30 = len(df[(df['Marks'].astype(float) >= mark_40) & (df['Marks'].astype(float) < mark_75)]) 
    more_than_30 = len(df[df['Marks'].astype(float) >= mark_75]) 

    return df, course_detail, total_students, total_students_appeared, total_absent, avg_marks, less_than_16, between_16_and_30, more_than_30, mark_40, mark_75

# Function to create the summary sheet
def create_summary_sheet(wb, course_detail, summary_data):
    ws = wb.add_worksheet('Student Summary')

    title_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 20, 'bg_color': '#FFFF00', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
    header_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 12, 'bg_color': '#B0E0E6', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
    cell_format = wb.add_format({'border': 2, 'align': 'center', 'valign': 'vcenter', 'bold': True})

    ws.merge_range('A1:F2', "Staff Report", title_format)
    for i, detail in enumerate(course_detail):
        ws.write_row(i + 2, 0, detail, cell_format)

    headers = ["Attribute", "Value"]
    ws.write_row('A13', headers, header_format)
    ws.set_column(0, 2, 25)

    results = [
        ["Total Students", summary_data['total_students']],
        ["Total Students Appeared", summary_data['total_students_appeared']],
        ["Total Absent", summary_data['total_absent']],
        ["Average Marks", summary_data['avg_marks']],
        ["Less Than 40%", summary_data['less_than_16']],
        ["Between 40% - 75%", summary_data['between_16_and_30']],
        ["More than 75%", summary_data['more_than_30']]
    ]
    
    # Loop through the results and check for NaN or Infinity
    for i, (attribute, value) in enumerate(results, start=13):
        ws.write(i, 0, attribute, cell_format)
        
        # Check if value is NaN or Infinity, replace with 'N/A' or 0
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            value = 'N/A'  # or use 0 if you prefer a numeric value
        
        ws.write(i, 1, value, cell_format)
    
# Function to create learner sheets
def create_learner_sheet(wb, sheet_name, df, mark, comparison, header_format):
    filtered_df = df[df['Marks'].astype(float).apply(comparison, args=(mark,))]
    filtered_df = filtered_df.sort_values(by='Marks', ascending=not comparison(mark, mark))
    
    ws = wb.add_worksheet(sheet_name)
    for col, header in enumerate(filtered_df.columns.tolist()):
        ws.write(0, col, header, header_format)
    for row_idx, row_data in enumerate(filtered_df.values, start=1):
        for col_idx, cell_data in enumerate(row_data):
            ws.write(row_idx, col_idx, cell_data)

# Function to save the processed file and return the file path
def save_processed_file(output_path, uploaded_file, df, course_detail, summary_data, mark_40, mark_75):
    # Ensure the output directory exists
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    output_file_name = f'{os.path.splitext(uploaded_file.name)[0]}_output.xlsx'
    output_file_path = os.path.join(output_path, output_file_name)

    wb = xlsxwriter.Workbook(output_file_path)
    header_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 12, 'bg_color': '#B0E0E6', 'border': 2, 'align': 'center', 'valign': 'vcenter'})

    create_summary_sheet(wb, course_detail, summary_data)

    st.text(uploaded_file.name[:-5])

    col1, col2 = st.columns(2)
    with col1:
        mark_less_than = st.number_input(
            f"Enter a mark to filter out slow learners: ",
            min_value=0.0,  
            value=float(mark_40),
            key=f"mark_less_than_{uploaded_file.name}"  
        )
        create_learner_sheet(wb, 'Slow Learners', df, mark_less_than, lambda x, y: x < y, header_format)
    with col2:
        mark_more_than = st.number_input(
            f"Enter a mark to filter out fast learners: ",
            min_value=0.0, 
            value=float(mark_75),
            key=f"mark_more_than_{uploaded_file.name}"  
        )
        create_learner_sheet(wb, 'Fast Learners', df, mark_more_than, lambda x, y: x > y, header_format)

    wb.close()
    
    if os.path.exists(output_file_path):
        return output_file_path
    else:
        raise FileNotFoundError(f"Failed to save the processed file: {output_file_path}")

# Function to generate the consolidated Excel file
def generate_consolidated_file(output_files, course_code, output_path, top_n):
    data_sheet1 = []
    sheet2_roll_counts = {}
    sheet3_roll_counts = {}

    df_combined_sheet2 = pd.DataFrame()
    df_combined_sheet3 = pd.DataFrame()

    for idx, file in enumerate(output_files):
        file_path = file  
        df_sheet1 = pd.read_excel(file_path, sheet_name='Student Summary', skiprows=12, header=0).drop(columns='Unnamed: 2', axis=1, errors='ignore')
        data_sheet1.append([course_code[idx]] + df_sheet1['Value'].values.tolist())

        df_combined_sheet2 = pd.concat([df_combined_sheet2, pd.read_excel(file_path, sheet_name='Slow Learners', header=0)], ignore_index=True)
        df_combined_sheet3 = pd.concat([df_combined_sheet3, pd.read_excel(file_path, sheet_name='Fast Learners', header=0)], ignore_index=True)

    df_sheet1_transposed = pd.DataFrame(data_sheet1, columns=['Course Code', 'Total Students', 'Total Students Appeared', 'Total Absent', 
                                                              'Average Marks', 'Students Less than 40%', 
                                                              'Students Between 40% and 75%', 'Students More than 75%']).T.reset_index()
    df_sheet1_transposed.columns = df_sheet1_transposed.iloc[0]
    df_sheet1_transposed = df_sheet1_transposed.drop(0)

    for _, row in df_combined_sheet2.iterrows():
        sheet2_roll_counts[row['Roll No.']] = sheet2_roll_counts.get(row['Roll No.'], {'Count': 0, 'Name': 'Unknown'})
        sheet2_roll_counts[row['Roll No.']]['Count'] += 1
        sheet2_roll_counts[row['Roll No.']]['Name'] = row['Student Name']

    for _, row in df_combined_sheet3.iterrows():
        sheet3_roll_counts[row['Roll No.']] = sheet3_roll_counts.get(row['Roll No.'], {'Count': 0, 'Name': 'Unknown'})
        sheet3_roll_counts[row['Roll No.']]['Count'] += 1
        sheet3_roll_counts[row['Roll No.']]['Name'] = row['Student Name']

    df_consolidated_sheet2 = pd.DataFrame([(roll_no, data['Name'], data['Count']) for roll_no, data in sorted(sheet2_roll_counts.items(), key=lambda x: x[1]['Count'], reverse=True)],
                                          columns=['Roll No.', 'Student Name', 'Count']).head(top_n)
    df_consolidated_sheet3 = pd.DataFrame([(roll_no, data['Name'], data['Count']) for roll_no, data in sorted(sheet3_roll_counts.items(), key=lambda x: x[1]['Count'], reverse=True)],
                                          columns=['Roll No.', 'Student Name', 'Count']).head(top_n)

    consolidated_file_path = os.path.join(output_path, "consolidated_output.xlsx")
    wb = xlsxwriter.Workbook(consolidated_file_path)
    header_format = wb.add_format({'bold': True, 'font': 'Times New Roman', 'font_size': 12, 'bg_color': '#B0E0E6', 'border': 2, 'align': 'center', 'valign': 'vcenter'})

    ws1 = wb.add_worksheet("Sheet1")
    for i, col in enumerate(df_sheet1_transposed.columns.tolist()):
        ws1.write(0, i, col, header_format)
    for row_idx, row_data in enumerate(df_sheet1_transposed.values, start=1):
        for col_idx, cell_data in enumerate(row_data):
            ws1.write(row_idx, col_idx, cell_data)

    ws2 = wb.add_worksheet("Slow Learners")
    for i, col in enumerate(df_consolidated_sheet2.columns.tolist()):
        ws2.write(0, i, col, header_format)
    for row_idx, row_data in enumerate(df_consolidated_sheet2.values, start=1):
        for col_idx, cell_data in enumerate(row_data):
            ws2.write(row_idx, col_idx, cell_data)

    ws3 = wb.add_worksheet("Fast Learners")
    for i, col in enumerate(df_consolidated_sheet3.columns.tolist()):
        ws3.write(0, i, col, header_format)
    for row_idx, row_data in enumerate(df_consolidated_sheet3.values, start=1):
        for col_idx, cell_data in enumerate(row_data):
            ws3.write(row_idx, col_idx, cell_data)

    wb.close()

    return consolidated_file_path

# Function to process all uploaded files and prepare the combined Excel file
def process_uploaded_files(uploaded_files, output_path):
    output_files = []
    course_code = []
    marks = []

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    for uploaded_file in uploaded_files:
        df = pd.read_excel(uploaded_file)
        course_code.append(df.iloc[7, 3])
        marks.append(df.iloc[8, 3])

    for i, uploaded_file in enumerate(uploaded_files):
        df, course_detail, *summary_data = process_file(uploaded_file, marks[i])
        output_file_path = save_processed_file(output_path, uploaded_file, df, course_detail, 
                                               dict(zip(['total_students', 'total_students_appeared', 'total_absent', 'avg_marks', 'less_than_16', 'between_16_and_30', 'more_than_30'], summary_data)),
                                               summary_data[-2], summary_data[-1])
        output_files.append(output_file_path)

    return output_files, course_code

# Streamlit UI
st.title("Excel File Processing")
uploaded_files = st.file_uploader("Upload Excel files", accept_multiple_files=True)
top_n = st.number_input("Enter the number of top roll numbers to display:", min_value=1, value=10)
zip_file_name = st.text_input("Enter the name for the zip file: ")
output_path = "processed_files"

if uploaded_files:
    # Process the uploaded files
    output_files, course_code = process_uploaded_files(uploaded_files, output_path)
    consolidated_file_path = generate_consolidated_file(output_files, course_code, output_path, top_n)

    # Create an in-memory zip buffer to store the files before saving to disk
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "a") as zip_file:
        for output_file in output_files + [consolidated_file_path]:
            zip_file.write(output_file, arcname=os.path.basename(output_file))

    # Move the buffer pointer to the start
    zip_buffer.seek(0)

    if st.button("Save Zip File"):
        # Define the folder where the zip file will be saved
        zip_folder = "/home/aravind/Projects/Staff_Report/zip_files"
        
        # Create the directory if it doesn't exist
        if not os.path.exists(zip_folder):
            os.makedirs(zip_folder)

        # Save the zip file to the specified location
        zip_file_path = os.path.join(zip_folder, f"{zip_file_name}.zip")
        with open(zip_file_path, "wb") as f:
            f.write(zip_buffer.getvalue())

        st.success(f"Zip file saved at: {zip_file_path}")

    # Delete temporary files after zipping
    for file_path in output_files + [consolidated_file_path]:
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            st.error(f"Error deleting file {file_path}: {e}")