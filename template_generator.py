import pandas as pd
from openpyxl import Workbook
import io
from datetime import date


def create_excel_template():
    """Create a template Excel file with all required sheets and formats"""
    # Create a Pandas Excel writer
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # Sheet 1: CourseData
    course_data = pd.DataFrame({
        'Course Name': ['Course A', 'Course B', 'Course C', 'Course D', 'Course E'],
        'Delivery Type': ['F2F', 'F2F', 'F2F', 'F2F', 'F2F'],
        'Language': ['EN', 'EN', 'EN', 'EN', 'EN'],
        'Runs': [3, 2, 4, 1, 2],
        'Duration': [5, 5, 5, 5, 5]  # All courses 5 days
    })
    course_data.to_excel(writer, sheet_name='CourseData', index=False)

    # Sheet 2: TrainerData
    trainer_data = pd.DataFrame({
        'Name': ['Trainer 1', 'Trainer 2', 'Trainer 3', 'Trainer 4', 'Trainer 5'],
        'Title': ['Champion', 'Consultant', 'Senior Consultant', 'Partner', 'Freelancer'],
        'Max_Days': [150, 120, 180, 200, 100]
    })
    trainer_data.to_excel(writer, sheet_name='TrainerData', index=False)

    # Sheet 3: PriorityData
    priority_data = pd.DataFrame({
        'Title': ['Champion', 'Consultant', 'Senior Consultant', 'Partner', 'DMD', 'GM', 'MD', 'Freelancer'],
        'Priority': [1, 2, 2, 3, 4, 5, 5, 6]  # Lower number = higher priority
    })
    priority_data.to_excel(writer, sheet_name='PriorityData', index=False)

    # Sheet 4: AnnualLeaves
    annual_leaves = pd.DataFrame({
        'Name': ['Trainer 1', 'Trainer 1', 'Trainer 2', 'Trainer 3', 'Trainer 4'],
        'Start_Date': [date(2025, 3, 10), date(2025, 8, 15), date(2025, 5, 1), date(2025, 7, 20), date(2025, 12, 20)],
        'End_Date': [date(2025, 3, 20), date(2025, 8, 25), date(2025, 5, 15), date(2025, 8, 5), date(2025, 12, 31)]
    })
    annual_leaves.to_excel(writer, sheet_name='AnnualLeaves', index=False)

    # Sheet 5: AffinityMatrix
    affinity_matrix = pd.DataFrame({
        'Course 1': ['Course A', 'Course B', 'Course A'],
        'Course 2': ['Course B', 'Course C', 'Course D'],
        'Gap Weeks': [3, 2, 4]  # Minimum weeks between courses
    })
    affinity_matrix.to_excel(writer, sheet_name='AffinityMatrix', index=False)

    # Sheet 6: PublicHolidays
    public_holidays = pd.DataFrame({
        'Start Date': [date(2025, 1, 1), date(2025, 5, 1), date(2025, 12, 24)],
        'End Date': [date(2025, 1, 2), date(2025, 5, 5), date(2025, 12, 26)]
    })
    public_holidays.to_excel(writer, sheet_name='PublicHolidays', index=False)

    # Sheet 7: Fleximatrix using the WIDE format (consultants as columns)
    # Define the courses and their categories
    courses = [
        {"CourseName": "Course A", "CategoryName": "Category 1", "Language": "EN"},
        {"CourseName": "Course B", "CategoryName": "Category 1", "Language": "EN"},
        {"CourseName": "Course C", "CategoryName": "Category 2", "Language": "EN"},
        {"CourseName": "Course D", "CategoryName": "Category 2", "Language": "EN"},
        {"CourseName": "Course E", "CategoryName": "Category 3", "Language": "EN"},
    ]

    # Get the trainer names for columns
    trainers = trainer_data['Name'].tolist()

    # Create the DataFrame structure for wide-format Fleximatrix
    rows = []
    for course in courses:
        row = {
            "CourseName": course["CourseName"],
            "CategoryName": course["CategoryName"],
            "Language": course["Language"],
            "Champion": ""  # Will be filled with the champion's name
        }

        # Add a column for each trainer (initially empty)
        for trainer in trainers:
            row[trainer] = ""

        rows.append(row)

    # Create the DataFrame
    fleximatrix_df = pd.DataFrame(rows)

    # Fill in the qualifications ("U") and champions based on the scenario
    # Course A: Trainer 1 is champion, Trainers 1, 2, 3 can teach it
    fleximatrix_df.at[0, "Champion"] = "Trainer 1"
    fleximatrix_df.at[0, "Trainer 1"] = "U"
    fleximatrix_df.at[0, "Trainer 2"] = "U"
    fleximatrix_df.at[0, "Trainer 3"] = "U"

    # Course B: Trainer 2 is champion, Trainers 2, 3 can teach it
    fleximatrix_df.at[1, "Champion"] = "Trainer 2"
    fleximatrix_df.at[1, "Trainer 2"] = "U"
    fleximatrix_df.at[1, "Trainer 3"] = "U"

    # Course C: Trainer 3 is champion, Trainers 3, 4, 5 can teach it
    fleximatrix_df.at[2, "Champion"] = "Trainer 3"
    fleximatrix_df.at[2, "Trainer 3"] = "U"
    fleximatrix_df.at[2, "Trainer 4"] = "U"
    fleximatrix_df.at[2, "Trainer 5"] = "U"

    # Course D: Trainer 4 is champion, Trainers 4, 5 can teach it
    fleximatrix_df.at[3, "Champion"] = "Trainer 4"
    fleximatrix_df.at[3, "Trainer 4"] = "U"
    fleximatrix_df.at[3, "Trainer 5"] = "U"

    # Course E: Trainer 1 is champion, Trainers 1, 3, 5 can teach it
    fleximatrix_df.at[4, "Champion"] = "Trainer 1"
    fleximatrix_df.at[4, "Trainer 1"] = "U"
    fleximatrix_df.at[4, "Trainer 3"] = "U"
    fleximatrix_df.at[4, "Trainer 5"] = "U"

    # Save to Excel
    fleximatrix_df.to_excel(writer, sheet_name='Fleximatrix', index=False)

    # Sheet 8: Week Restrictions
    week_restrictions = pd.DataFrame([
        # Course C can't run on the last week of any month (e.g., financial statement courses)
        {'Course': 'Course C', 'Week Type': 'Last', 'Restricted': True,
         'Notes': 'Financial statement courses should not run in last week of month'},
        # Course D can't run on the third week of any month (e.g., payroll courses)
        {'Course': 'Course D', 'Week Type': 'Third', 'Restricted': True,
         'Notes': 'Payroll courses should not run in third week of month'},
        # Course A can't run on the first week of any month
        {'Course': 'Course A', 'Week Type': 'First', 'Restricted': True, 'Notes': 'Do not run in first week of month'},
        # Course E can't run on the second week of any month
        {'Course': 'Course E', 'Week Type': 'Second', 'Restricted': True,
         'Notes': 'Do not run in second week of month'},
    ])
    week_restrictions.to_excel(writer, sheet_name='WeekRestrictions', index=False)

    # Sheet 9: MonthlyDemand
    monthly_demand = pd.DataFrame({
        'Month': list(range(1, 13)),  # 1-12 for Jan-Dec
        'Percentage': [8.3, 8.3, 16.7, 8.3, 8.3, 0, 8.3, 8.3, 16.7, 8.3, 8.3, 0]  # Percentages (adding up to ~100%)
    })
    monthly_demand.to_excel(writer, sheet_name='MonthlyDemand', index=False)

    # Save and return
    writer.close()
    output.seek(0)
    return output

# This function can be used to generate the template for download
# Example usage in a Streamlit app:
# template = create_excel_template()
# st.download_button("Download Excel Template", data=template, file_name="course_scheduler_template.xlsx",
#                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")