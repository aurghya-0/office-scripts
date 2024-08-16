import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment

# Load the CSV file into a DataFrame
df = pd.read_csv('CSE.csv')
df = df.fillna(method='ffill')

# Rename columns to match your headers
df.rename(columns={
    'Student': 'student_name',
    'Registration Id': 'registration_id',
    'Batch Year': 'batch_year',
    'Programme Section': 'programme_section',
    'Course': 'course_name',
    'Present': 'attendance_percentage',
    'Overall Present': 'overall_attendance'
}, inplace=True)


# Exclude courses with "PRACTICAL" or "LECTURE" tags
df = df[~df['course_name'].str.contains('PRACTICAL|LECTURE', case=False, na=False)]

# Apply the eligibility logic
def determine_eligibility(row):
    if row['overall_attendance'] >= 75:
        return 'Allowed for all courses'
    elif row['course_name'] == 'Intermediate Internship [CSE(INT)501]  [PROJECT]':
        return 'Allowed for all courses due to Project'
    elif row['course_name'] == 'Basic Internship [CSE(INT)301]  [PROJECT]':
        return 'Allowed for all courses due to Project'
    elif row['course_name'] == 'Advanced Internship [CSE(INT)701]  [PROJECT]':
        return 'Allowed for all courses due to Project'
    elif row['attendance_percentage'] >= 75:
        return 'Allowed for this course'
    else:
        return 'Not allowed'

# Apply the function to each row
df['eligibility'] = df.apply(determine_eligibility, axis=1)

# Determine overall eligibility for each student
overall_eligibility = df.groupby('student_name')['eligibility'].apply(
    lambda x: 'Allowed for all courses' if 'Allowed for all courses' in x.values else 'Allowed for this course' if 'Allowed for this course' in x.values else 'Not allowed'
)

# Apply the function to each row
df['eligibility'] = df.apply(determine_eligibility, axis=1)

# Sort the DataFrame by student_name and course_name
df = df[['student_name', 'registration_id', 'batch_year', 'programme_section', 'course_name', 'attendance_percentage', 'overall_attendance', 'eligibility']].sort_values(by=['student_name', 'course_name'])

# Create an Excel workbook and add a worksheet
wb = Workbook()
ws = wb.active
ws.title = "Student Exam Eligibility"

# Add the DataFrame data to the worksheet
for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)

# Merge cells for student_name, overall_attendance, and eligibility where applicable
current_student = None
start_row = 2  # Excel rows start from 1, and the first row is the header
for row in range(2, ws.max_row + 1):
    student_name = ws.cell(row=row, column=1).value
    eligibility = ws.cell(row=row, column=8).value
    overall_attendance = ws.cell(row=row, column=7).value
    if student_name != current_student:
        if current_student is not None:
            # Merge student_name cells
            ws.merge_cells(start_row=start_row, start_column=1, end_row=row-1, end_column=1)
            # Merge eligibility cells for "Allowed for all courses" or "Not allowed"
            if ws.cell(row=start_row, column=8).value == 'Allowed for all courses' or \
               ws.cell(row=start_row, column=8).value == 'Allowed for all courses due to Project' or \
               ws.cell(row=start_row, column=8).value == 'Not allowed':
                ws.merge_cells(start_row=start_row, start_column=8, end_row=row-1, end_column=8)
            # Merge overall_attendance cells
            ws.merge_cells(start_row=start_row, start_column=7, end_row=row-1, end_column=7)
        current_student = student_name
        start_row = row

# Handle the last student group
if current_student is not None:
    ws.merge_cells(start_row=start_row, start_column=1, end_row=ws.max_row, end_column=1)
    if ws.cell(row=start_row, column=8).value == 'Allowed for all courses' or \
       ws.cell(row=start_row, column=8).value == 'Allowed for all courses due to Project' or \
       ws.cell(row=start_row, column=8).value == 'Not allowed':
        ws.merge_cells(start_row=start_row, start_column=8, end_row=ws.max_row, end_column=8)
    # Merge overall_attendance cells
    ws.merge_cells(start_row=start_row, start_column=7, end_row=ws.max_row, end_column=7)

# Center align the merged cells
for col in ['A', 'G', 'H']:
    for cell in ws[col]:
        cell.alignment = Alignment(horizontal='center', vertical='center')

# Save the workbook
wb.save('student_exam_eligibility.xlsx')