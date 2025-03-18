# Course Scheduler Application

A Streamlit application for optimizing course schedules and trainer assignments.

## Setup

1. Install requirements:pip install -r requirements.txt
2. Run the application: streamlit run app.py
## Data Format

Upload an Excel file with the following sheets:

### CourseData
Contains course information with columns:
- Course Name
- Methodology
- Language
- Runs
- Duration

### TrainerData
Contains trainer information with columns:
- Name (use initials or codes)
- Title
- Max_Days

### Fleximatrix
Defines which trainers can deliver which courses using a grid format:
- **Rows**: Each course+language combination
- **Columns**: 
- CourseName, CategoryName, Language, Champion + trainer initials
- **Cell values**: 
- "U" indicates a trainer is qualified to deliver the course
- Champion column contains the trainer code who is designated champion

Example:
| CourseName | CategoryName | Language | Champion | AAD | AAG | AEB |
|------------|--------------|----------|----------|-----|-----|-----|
| Course A   | Finance      | E        | AAD      | U   |     |     |
| Course B   | Finance      | E        | AEB      | U   |     | U   |

### WeekRestrictions (Optional)
Contains restrictions on when specific courses can run:
- Course: Course name
- Week Type: Position in month (First, Second, Third, Fourth, Last)
- Restricted: Boolean (TRUE/FALSE)
- Notes: Optional explanation

### Other Required Sheets
- PriorityData: Defines priority levels for trainer titles
- AnnualLeaves: Trainer leave periods
- AffinityMatrix: Course timing relationships
- PublicHolidays: Holiday periods 
- MonthlyDemand: Monthly course distribution targets

See the application for details on the specific format for each sheet.