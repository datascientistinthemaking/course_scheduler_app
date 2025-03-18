import pandas as pd
import numpy as np
import datetime
import streamlit as st
import io
import matplotlib.pyplot as plt
import seaborn as sns
from ortools.sat.python import cp_model

st.set_page_config(page_title="Course Scheduler", layout="wide")

# Import template generator
from template_generator import create_excel_template

# Add template download button
st.markdown("### Download Excel Template")
st.write("Not sure about the format? Download our template file:")
if st.button("Generate Template"):
    template = create_excel_template()
    st.download_button(
        "Download Excel Template",
        data=template,
        file_name="course_scheduler_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

class CourseScheduler:
    def __init__(self):
        self.course_run_data = None
        self.consultant_data = None
        self.priority_data = None
        self.annual_leaves = None
        self.affinity_matrix_data = None
        self.public_holidays_data = None
        self.weekly_calendar = None
        self.week_to_month_map = {}
        self.weekly_working_days = {}
        self.fleximatrix = {}  # Dictionary mapping (course, language) to list of qualified trainers
        self.course_champions = {}  # Dictionary to store champions for each course

    def load_data_from_excel(self, uploaded_file):
        """Load all required data from an Excel file with multiple sheets"""
        try:
            # Read the Excel file with multiple sheets
            excel_file = pd.ExcelFile(uploaded_file)

            # Load each required sheet
            self.course_run_data = pd.read_excel(excel_file, sheet_name='CourseData')
            self.consultant_data = pd.read_excel(excel_file, sheet_name='TrainerData')
            self.priority_data = pd.read_excel(excel_file, sheet_name='PriorityData')
            self.annual_leaves = pd.read_excel(excel_file, sheet_name='AnnualLeaves')
            self.affinity_matrix_data = pd.read_excel(excel_file, sheet_name='AffinityMatrix')
            self.public_holidays_data = pd.read_excel(excel_file, sheet_name='PublicHolidays')

            # Parse date columns
            date_columns = ['Start_Date', 'End_Date']
            for col in date_columns:
                if col in self.annual_leaves.columns:
                    self.annual_leaves[col] = pd.to_datetime(self.annual_leaves[col]).dt.date

            if 'Start Date' in self.public_holidays_data.columns:
                self.public_holidays_data['Start Date'] = pd.to_datetime(
                    self.public_holidays_data['Start Date']).dt.date

            if 'End Date' in self.public_holidays_data.columns:
                self.public_holidays_data['End Date'] = pd.to_datetime(self.public_holidays_data['End Date']).dt.date

            # Load fleximatrix - which trainers can teach which courses (wide format)
            flexi_sheet = pd.read_excel(excel_file, sheet_name='Fleximatrix')

            # Process the wide-format Fleximatrix
            self.fleximatrix, self.course_champions = self.process_wide_fleximatrix(flexi_sheet)

            # Load adjusted monthly demand data
            self.monthly_demand = pd.read_excel(excel_file, sheet_name='MonthlyDemand')

            # Load course week restrictions (if exists)
            self.week_restrictions = {}
            try:
                # Check if the sheet exists
                restrictions_sheet = pd.read_excel(excel_file, sheet_name='WeekRestrictions')

                # Process each restriction
                for _, row in restrictions_sheet.iterrows():
                    course = row['Course']
                    week_type = row['Week Type']  # 'First', 'Second', 'Third', 'Fourth', 'Last'
                    restricted = row['Restricted']  # Boolean: True = cannot run, False = can run

                    if course not in self.week_restrictions:
                        self.week_restrictions[course] = {}

                    self.week_restrictions[course][week_type] = restricted

                print(f"Loaded {len(restrictions_sheet)} week restrictions for courses")
            except Exception as e:
                print(f"No week restrictions found or error loading them: {e}")
                self.week_restrictions = {}

            return True

        except Exception as e:
            st.error(f"Error loading data: {e}")
            return False

    def process_wide_fleximatrix(self, flexi_sheet):
        """Process a wide-format Fleximatrix where consultants are columns"""
        fleximatrix = {}
        course_champions = {}

        # Get list of consultants (all columns except CourseName, CategoryName, Language, and Champion)
        consultants = [col for col in flexi_sheet.columns
                       if col not in ['CourseName', 'CategoryName', 'Language', 'Champion']]

        # Process each row (course+language combination)
        for _, row in flexi_sheet.iterrows():
            course = row['CourseName']
            language = row['Language']
            champion = row['Champion']

            # Initialize empty list for this course+language
            fleximatrix[(course, language)] = []

            # Add the champion
            if champion and str(champion).strip():
                course_champions[(course, language)] = champion

            # Add all qualified consultants (marked with "U")
            for consultant in consultants:
                if row[consultant] == "U":
                    fleximatrix[(course, language)].append(consultant)

        return fleximatrix, course_champions


    def initialize_calendar(self, year, weekend_selection):
        """Initialize the weekly calendar based on year and weekend selection"""

        # Function to Generate Work Week Start Dates
        def get_week_start_dates(year, weekend_selection):
            first_day = datetime.date(year, 1, 1)
            if weekend_selection == "FS":  # Friday-Saturday
                week_start = 6  # Sunday
            else:  # "Saturday-Sunday"
                week_start = 0  # Monday

            weeks = []
            current_date = first_day

            # Find the first day that matches our week start
            while current_date.weekday() != week_start:
                current_date += datetime.timedelta(days=1)

            # Generate all week start dates for the year
            while current_date.year == year:
                weeks.append(current_date)
                current_date += datetime.timedelta(days=7)  # Move to next week

            return weeks

        # Generate Work Week Start Dates
        self.weekly_calendar = get_week_start_dates(year, weekend_selection)

        # Ensure correct mapping of weeks to months
        self.week_to_month_map = {}
        self.week_position_in_month = {}  # Track position of week within its month

        # First, map weeks to months
        for i, week_start in enumerate(self.weekly_calendar):
            month = week_start.month  # Extract the month from the start date
            self.week_to_month_map[i + 1] = month  # Map week number to month

        # Then, calculate position of each week within its month
        for week_num, month in self.week_to_month_map.items():
            # Get all weeks in this month
            month_weeks = [w for w, m in self.week_to_month_map.items() if m == month]
            month_weeks.sort()

            # Find position (1-indexed)
            position = month_weeks.index(week_num) + 1

            # Also mark if it's the last week of the month
            is_last = (position == len(month_weeks))

            # Store as a dictionary with position info
            self.week_position_in_month[week_num] = {
                'position': position,
                'total_weeks': len(month_weeks),
                'is_first': (position == 1),
                'is_second': (position == 2),
                'is_third': (position == 3),
                'is_fourth': (position == 4),
                'is_last': is_last
            }

        # Calculate working days for each week accounting for public holidays
        self.calculate_working_days()

        return True


    def calculate_working_days(self):
        """Calculate working days for each week accounting for public holidays"""
        self.weekly_working_days = {}

        for i, week_start in enumerate(self.weekly_calendar):
            week_num = i + 1
            working_days = 5  # Assume a full workweek

            # Check if this week is completely inside a long holiday
            fully_inside_long_holiday = False

            for _, row in self.public_holidays_data.iterrows():
                holiday_length = (row["End Date"] - row["Start Date"]).days + 1

                # If it's a long holiday (25+ days) and the week is fully inside it
                if holiday_length >= 25:
                    if week_start >= row["Start Date"] and (week_start + datetime.timedelta(days=4)) <= row["End Date"]:
                        fully_inside_long_holiday = True
                        break

            # If fully inside a long holiday, block the week completely
            if fully_inside_long_holiday:
                working_days = 0
            else:
                # Adjust only affected days for regular holidays
                for _, row in self.public_holidays_data.iterrows():
                    holiday_days_in_week = sum(
                        1 for d in range(5)  # Check only workdays
                        if row["Start Date"] <= (week_start + datetime.timedelta(days=d)) <= row["End Date"]
                        and (week_start + datetime.timedelta(days=d)).weekday() not in [5, 6]  # Exclude weekends
                    )
                    working_days -= holiday_days_in_week
                    working_days = max(0, working_days)  # Ensure no negative values

            self.weekly_working_days[week_num] = working_days


    def is_trainer_available(self, trainer_name, week_num):
        """Check if trainer is available during the given week (not on leave)"""
        week_start = self.weekly_calendar[week_num - 1]
        week_end = week_start + datetime.timedelta(days=4)  # Assuming 5-day courses

        # Check against all leave periods
        for _, leave in self.annual_leaves[self.annual_leaves["Name"] == trainer_name].iterrows():
            # If leave period overlaps with course week, trainer is unavailable
            if (leave["Start_Date"] <= week_end and leave["End_Date"] >= week_start):
                return False
        return True


    def get_trainer_priority(self, trainer_name):
        """Get priority level for a trainer based on their title"""
        title = self.consultant_data.loc[self.consultant_data["Name"] == trainer_name, "Title"].iloc[0]
        priority = self.priority_data.loc[self.priority_data["Title"] == title, "Priority"].iloc[0]
        return priority


    def is_freelancer(self, trainer_name):
        """Check if a trainer is a freelancer"""
        title = self.consultant_data.loc[self.consultant_data["Name"] == trainer_name, "Title"].iloc[0]
        return title == "Freelancer"


    def run_optimization(self):
        """Run the optimization model to create the schedule"""
        print("Starting optimization with trainer assignments...")

        # Get total F2F runs
        total_f2f_runs = sum(self.course_run_data["Runs"])

        # Create adjusted monthly demand dictionary from the demand dataframe
        # Convert percentages to actual course counts
        adjusted_f2f_demand = {}
        total_percentage = self.monthly_demand['Percentage'].sum()

        for _, row in self.monthly_demand.iterrows():
            month = row['Month']
            percentage = row['Percentage'] / total_percentage  # Normalize to ensure percentages sum to 1
            # Calculate how many courses should be in this month (round to nearest integer)
            demand = round(percentage * total_f2f_runs)
            adjusted_f2f_demand[month] = demand

        # Adjust rounding errors to ensure total matches
        total_allocated = sum(adjusted_f2f_demand.values())
        if total_allocated != total_f2f_runs:
            # Find the month with the highest demand and adjust
            max_month = max(adjusted_f2f_demand.items(), key=lambda x: x[1])[0]
            adjusted_f2f_demand[max_month] += (total_f2f_runs - total_allocated)

        print("Monthly demand (converted from percentages):")
        for month, demand in sorted(adjusted_f2f_demand.items()):
            print(f"Month {month}: {demand} courses")
        # Initialize model and variables
        model = cp_model.CpModel()
        schedule = {}
        trainer_assignments = {}
        max_weeks = len(self.weekly_calendar)

        # All penalty lists
        month_deviation_penalties = []
        affinity_penalties = []
        trainer_utilization_penalties = []
        champion_assignment_penalties = []

        # Create the schedule variables for each course and run
        for _, row in self.course_run_data.iterrows():
            course, methodology, language, runs, duration = row["Course Name"], row["Methodology"], row["Language"], row[
                "Runs"], row["Duration"]

            for i in range(runs):
                # Create a variable for the start week
                start_week = model.NewIntVar(1, max_weeks, f"start_week_{course}_{i}")
                schedule[(course, methodology, language, i)] = start_week

                # Create trainer assignment variable
                qualified_trainers = self.fleximatrix.get((course, language), [])
                if not qualified_trainers:
                    print(f"Warning: No qualified trainers for {course} ({language})")
                    continue

                trainer_var = model.NewIntVar(0, len(qualified_trainers) - 1, f"trainer_{course}_{i}")
                trainer_assignments[(course, methodology, language, i)] = trainer_var

                # Only schedule in weeks with enough working days AND available trainers
                valid_weeks = []
                for w, days in self.weekly_working_days.items():
                    if days >= duration:
                        # WEEK RESTRICTION CHECK: Skip if this course can't run in this week position
                        if course in self.week_restrictions:
                            # Get week position info
                            week_info = self.week_position_in_month.get(w, {})

                            # Check each restriction type
                            skip_week = False

                            if week_info.get('is_first') and self.week_restrictions[course].get('First', False):
                                skip_week = True
                            elif week_info.get('is_second') and self.week_restrictions[course].get('Second', False):
                                skip_week = True
                            elif week_info.get('is_third') and self.week_restrictions[course].get('Third', False):
                                skip_week = True
                            elif week_info.get('is_fourth') and self.week_restrictions[course].get('Fourth', False):
                                skip_week = True
                            elif week_info.get('is_last') and self.week_restrictions[course].get('Last', False):
                                skip_week = True

                            if skip_week:
                                # Skip this week for this course due to restriction
                                continue

                        # Check if at least one qualified trainer is available this week
                        for trainer in qualified_trainers:
                            if self.is_trainer_available(trainer, w):
                                valid_weeks.append(w)
                                break

                if valid_weeks:
                    model.AddAllowedAssignments([start_week], [[w] for w in valid_weeks])
                else:
                    print(f"Warning: No valid weeks for {course} run {i + 1}")

        # CONSTRAINT 1: Track which month each course is assigned to and enforce monthly distribution
        print("Adding fixed monthly distribution constraints")

        # For each month, explicitly track which courses are scheduled in it
        for month in range(1, 13):
            target_demand = adjusted_f2f_demand.get(month, 0)

            # Create a list of all Boolean variables indicating if a course is scheduled in this month
            courses_in_month = []

            # Get all weeks that belong to this month
            month_weeks = [week for week, m in self.week_to_month_map.items() if m == month]

            if not month_weeks:  # Skip months with no weeks
                continue

            # For each course run, create a Boolean variable indicating if it's in this month
            for (course, methodology, language, i), week_var in schedule.items():
                is_in_month = model.NewBoolVar(f"{course}_{i}_in_month_{month}")

                # Add constraints: is_in_month is True if and only if week_var is one of month's weeks
                week_choices = []
                for week in month_weeks:
                    is_in_this_week = model.NewBoolVar(f"{course}_{i}_in_week_{week}")
                    model.Add(week_var == week).OnlyEnforceIf(is_in_this_week)
                    model.Add(week_var != week).OnlyEnforceIf(is_in_this_week.Not())
                    week_choices.append(is_in_this_week)

                # is_in_month is true if any week_choice is true
                model.AddBoolOr(week_choices).OnlyEnforceIf(is_in_month)
                model.AddBoolAnd([choice.Not() for choice in week_choices]).OnlyEnforceIf(is_in_month.Not())

                courses_in_month.append(is_in_month)

            # CORE CONSTRAINT: Force the sum of all courses in this month to equal the target demand
            if courses_in_month and target_demand > 0:  # Only if we have variables for this month
                model.Add(sum(courses_in_month) == target_demand)
                print(f"  Month {month}: Enforcing exactly {target_demand} courses (hard constraint)")

        # CONSTRAINT 2: Minimum spacing between runs of same course
        print("Adding spacing between runs of the same course")

        for course_name in set(self.course_run_data["Course Name"]):
            course_runs = []
            for (course, methodology, language, i), var in schedule.items():
                if course == course_name:
                    course_runs.append((i, var))

            course_runs.sort()
            for i in range(len(course_runs) - 1):
                run_num1, var1 = course_runs[i]
                run_num2, var2 = course_runs[i + 1]

                # Hard constraint for minimum spacing
                min_gap = 2  # Weeks between runs of same course
                model.Add(var2 >= var1 + min_gap)

        # CONSTRAINT 3: Add constraints for course affinities
        print(f"Adding affinity constraints for course pairs")
        for _, row in self.affinity_matrix_data.iterrows():
            c1, c2, gap_weeks = row["Course 1"], row["Course 2"], row["Gap Weeks"]

            c1_runs = []
            c2_runs = []

            for (course, _, _, i), var in schedule.items():
                if course == c1:
                    c1_runs.append((i, var))
                elif course == c2:
                    c2_runs.append((i, var))

            # Skip if either course not found in schedule (might have 0 runs)
            if not c1_runs or not c2_runs:
                continue

            # Sort by run number
            c1_runs.sort()
            c2_runs.sort()

            # Only check first run of each course to reduce constraints
            run1, var1 = c1_runs[0]
            run2, var2 = c2_runs[0]

            # Soft affinity constraint
            too_close = model.NewBoolVar(f"affinity_too_close_{c1}_{c2}_{run1}_{run2}")

            # Two options: either var2 >= var1 + gap_weeks OR var2 <= var1 - gap_weeks
            far_enough_after = model.NewBoolVar(f"far_after_{c1}_{c2}_{run1}_{run2}")
            far_enough_before = model.NewBoolVar(f"far_before_{c1}_{c2}_{run1}_{run2}")

            model.Add(var2 >= var1 + gap_weeks).OnlyEnforceIf(far_enough_after)
            model.Add(var2 <= var1 - gap_weeks).OnlyEnforceIf(far_enough_before)

            model.AddBoolOr([far_enough_after, far_enough_before]).OnlyEnforceIf(too_close.Not())
            model.Add(var2 < var1 + gap_weeks).OnlyEnforceIf(too_close)
            model.Add(var2 > var1 - gap_weeks).OnlyEnforceIf(too_close)

            affinity_penalties.append(too_close)
            print(f"  Added affinity constraint: {c1} and {c2} should be {gap_weeks} weeks apart")

        # CONSTRAINT 4: Trainer-specific constraints
        print("Adding trainer assignment constraints")

        # 4.1: Trainer availability - only assign trainers who are available during the scheduled week
        for (course, methodology, language, i), week_var in schedule.items():
            trainer_var = trainer_assignments.get((course, methodology, language, i))
            if trainer_var is None:
                continue

            qualified_trainers = self.fleximatrix.get((course, language), [])

            # For each week, set which trainers are available
            for week in range(1, max_weeks + 1):
                is_this_week = model.NewBoolVar(f"{course}_{i}_in_week_{week}")
                model.Add(week_var == week).OnlyEnforceIf(is_this_week)
                model.Add(week_var != week).OnlyEnforceIf(is_this_week.Not())

                # For each trainer, check availability
                for t_idx, trainer in enumerate(qualified_trainers):
                    if not self.is_trainer_available(trainer, week):
                        # If trainer is not available this week, cannot select this trainer
                        is_this_trainer = model.NewBoolVar(f"{course}_{i}_trainer_{t_idx}_week_{week}")
                        model.Add(trainer_var == t_idx).OnlyEnforceIf(is_this_trainer)
                        model.Add(trainer_var != t_idx).OnlyEnforceIf(is_this_trainer.Not())

                        # Cannot have both is_this_week and is_this_trainer be true
                        model.AddBoolOr([is_this_week.Not(), is_this_trainer.Not()])

        # 4.2: Workload limits - track and limit total days per trainer
        trainer_workload = {name: [] for name in self.consultant_data["Name"]}

        # Track course assignments for each trainer
        for (course, methodology, language, i), trainer_var in trainer_assignments.items():
            duration = self.course_run_data.loc[self.course_run_data["Course Name"] == course, "Duration"].iloc[0]
            qualified_trainers = self.fleximatrix.get((course, language), [])

            for t_idx, trainer in enumerate(qualified_trainers):
                is_assigned = model.NewBoolVar(f"{course}_{i}_assigned_to_{trainer}")
                model.Add(trainer_var == t_idx).OnlyEnforceIf(is_assigned)
                model.Add(trainer_var != t_idx).OnlyEnforceIf(is_assigned.Not())

                # Accumulate workload
                trainer_workload[trainer].append((is_assigned, duration))

                # 4.3: Champion priority - add penalty for not using the champion for their courses
                champion = self.course_champions.get((course, language))
                if champion == trainer:
                    not_using_champion = model.NewBoolVar(f"not_using_champion_{course}_{i}")
                    model.Add(trainer_var != t_idx).OnlyEnforceIf(not_using_champion)
                    model.Add(trainer_var == t_idx).OnlyEnforceIf(not_using_champion.Not())
                    champion_assignment_penalties.append(not_using_champion)

        # Add constraints for maximum workload
        for trainer, workload_items in trainer_workload.items():
            if not workload_items:
                continue

            max_days = self.consultant_data.loc[self.consultant_data["Name"] == trainer, "Max_Days"].iloc[0]

            # Calculate total workload
            total_workload = model.NewIntVar(0, max_days, f"total_workload_{trainer}")
            weighted_sum = []
            for is_assigned, duration in workload_items:
                weighted_sum.append((is_assigned, duration))

            if weighted_sum:
                weighted_terms = []
                for is_assigned, duration in weighted_sum:
                    term = model.NewIntVar(0, duration, f"term_{id(is_assigned)}_{duration}")
                    model.Add(term == duration).OnlyEnforceIf(is_assigned)
                    model.Add(term == 0).OnlyEnforceIf(is_assigned.Not())
                    weighted_terms.append(term)
                model.Add(total_workload == sum(weighted_terms))
                # Enforce maximum workload
                model.Add(total_workload <= max_days)

                # 4.4: For non-freelancers, encourage higher utilization (soft constraint)
                if not self.is_freelancer(trainer):
                    # Calculate target utilization as a percentage of max days
                    target_days = int(max_days * 0.7)  # Target 70% utilization

                    # Add penalty for underutilization
                    under_target = model.NewBoolVar(f"{trainer}_under_target")
                    model.Add(total_workload < target_days).OnlyEnforceIf(under_target)
                    model.Add(total_workload >= target_days).OnlyEnforceIf(under_target.Not())

                    # Higher priority trainers get stronger penalty for underutilization
                    priority = self.get_trainer_priority(trainer)
                    penalty_weight = 7 - priority  # Invert priority (1 becomes 6, 6 becomes 1)

                    for _ in range(penalty_weight):
                        trainer_utilization_penalties.append(under_target)

        # Combined objective function with balanced weights
        model.Minimize(
            5 * sum(month_deviation_penalties) +  # Monthly distribution (higher priority)
            2 * sum(affinity_penalties) +  # Affinity constraints (medium priority)
            4 * sum(champion_assignment_penalties) +  # Champion assignment (high priority)
            3 * sum(trainer_utilization_penalties)  # Trainer utilization (medium-high priority)
        )

        # Initialize solver with increased time limit
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 300  # 5 minutes time limit

        # Solve the model
        status = solver.Solve(model)

        # Print status information
        print(f"Solver status: {solver.StatusName(status)}")
        print(f"Objective value: {solver.ObjectiveValue()}")
        print(f"Wall time: {solver.WallTime():.2f} seconds")

        print("Final Weekly Course Schedule with Trainers:")
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            schedule_results = []

            for (course, methodology, language, i), week_var in schedule.items():
                assigned_week = solver.Value(week_var)
                start_date = self.weekly_calendar[assigned_week - 1].strftime("%Y-%m-%d")

                # Get trainer assignment
                if (course, methodology, language, i) in trainer_assignments:
                    trainer_var = trainer_assignments[(course, methodology, language, i)]
                    trainer_idx = solver.Value(trainer_var)

                    if 0 <= trainer_idx < len(self.fleximatrix[(course, language)]):
                        trainer = self.fleximatrix[(course, language)][trainer_idx]

                        # Check if this is a champion course
                        is_champion = "✓" if self.course_champions.get((course, language)) == trainer else " "

                        schedule_results.append({
                            "Week": assigned_week,
                            "Start Date": start_date,
                            "Course": course,
                            "Methodology": methodology,
                            "Language": language,
                            "Run": i + 1,
                            "Trainer": trainer,
                            "Champion": is_champion
                        })

            # Convert to DataFrame
            schedule_df = pd.DataFrame(schedule_results)

            # Sort by week first, then by course name
            schedule_df = schedule_df.sort_values(by=["Week", "Course"])

            return status, schedule_df, solver, schedule, trainer_assignments
        else:
            return status, None, solver, schedule, trainer_assignments


    def plot_weekly_course_bar_chart(self, schedule, solver):
        """Creates a bar chart showing number of courses per week."""
        max_weeks = len(self.weekly_calendar)

        # Count courses per week
        weekly_counts = {week: 0 for week in range(1, max_weeks + 1)}
        for (_, _, _, _), week_var in schedule.items():
            assigned_week = solver.Value(week_var)
            weekly_counts[assigned_week] += 1

        # Convert to dataframe for plotting
        df = pd.DataFrame({
            'Week': list(weekly_counts.keys()),
            'Courses': list(weekly_counts.values())
        })

        # Create a bar chart
        plt.figure(figsize=(12, 6))
        bars = plt.bar(df['Week'], df['Courses'], width=0.8)

        # Add week dates as second x-axis
        ax1 = plt.gca()
        ax2 = ax1.twiny()

        # Select a subset of week numbers for readability
        step = max(1, max_weeks // 12)  # Show about 12 date labels
        selected_weeks = list(range(1, max_weeks + 1, step))

        # Get date strings for selected weeks
        date_labels = [self.weekly_calendar[w - 1].strftime("%b %d") for w in selected_weeks]

        # Set positions and labels
        ax2.set_xticks([w for w in selected_weeks])
        ax2.set_xticklabels(date_labels, rotation=45)

        # Formatting
        plt.title("Number of Courses per Week", fontsize=16)
        ax1.set_xlabel("Week Number", fontsize=14)
        ax1.set_ylabel("Number of Courses", fontsize=14)
        ax2.set_xlabel("Week Start Date", fontsize=14)

        # Set y limit to make small differences more visible
        max_courses = max(weekly_counts.values())
        ax1.set_ylim(0, max(4, max_courses + 1))

        # Highlight weeks with no courses
        for week, count in weekly_counts.items():
            if count == 0:
                plt.axvspan(week - 0.4, week + 0.4, color='lightgray', alpha=0.3)

        # Add grid for readability
        plt.grid(axis='y', linestyle='--', alpha=0.7)

        plt.tight_layout()

        # Return the figure
        return plt.gcf()


    def plot_trainer_workload_chart(self, schedule, trainer_assignments, solver):
        """Creates a bar chart showing trainer workload and utilization"""
        # Calculate days assigned to each trainer
        trainer_days = {name: 0 for name in self.consultant_data["Name"]}

        for (course, methodology, language, i), week_var in schedule.items():
            duration = self.course_run_data.loc[self.course_run_data["Course Name"] == course, "Duration"].iloc[0]

            # Get trainer assignment
            trainer_var = trainer_assignments[(course, methodology, language, i)]
            trainer_idx = solver.Value(trainer_var)

            if 0 <= trainer_idx < len(self.fleximatrix[(course, language)]):
                trainer = self.fleximatrix[(course, language)][trainer_idx]
                trainer_days[trainer] += duration

        # Prepare data for plotting
        trainers = []
        assigned_days = []
        max_days_list = []
        colors = []

        for name, days in sorted(trainer_days.items(), key=lambda x: x[1], reverse=True):
            if days > 0:  # Only include trainers with assignments
                trainers.append(name)
                assigned_days.append(days)
                max_days = self.consultant_data.loc[self.consultant_data["Name"] == name, "Max_Days"].iloc[0]
                max_days_list.append(max_days)

                # Different color for freelancers
                title = self.consultant_data.loc[self.consultant_data["Name"] == name, "Title"].iloc[0]
                colors.append("tab:orange" if title == "Freelancer" else "tab:blue")

        # Create figure and axis
        fig, ax = plt.subplots(figsize=(12, 6))

        # Plot assigned days
        bars = ax.bar(trainers, assigned_days, label="Assigned Days", color=colors)

        # Plot max days as a line
        ax.plot(trainers, max_days_list, 'rx-', label="Maximum Days", linewidth=2)

        # Add percentage labels
        for i, (bar, max_days) in enumerate(zip(bars, max_days_list)):
            height = bar.get_height()
            percentage = (height / max_days * 100) if max_days > 0 else 0
            ax.text(bar.get_x() + bar.get_width() / 2., height + 5,
                    f"{percentage:.1f}%", ha='center', va='bottom', fontsize=9)

        # Improve readability
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        # Add labels and title
        plt.ylabel("Training Days")
        plt.title("Trainer Workload and Utilization")
        plt.legend()
        plt.grid(axis='y', linestyle='--', alpha=0.7)

        # Return the figure
        return plt.gcf()

    def generate_monthly_validation(self, schedule, solver):
        """Generates a monthly validation report for the schedule"""
        # Get adjusted F2F demand from the monthly demand dataframe
        adjusted_f2f_demand = {}

        # Create the dictionary using the same logic as in run_optimization
        total_f2f_runs = sum(self.course_run_data["Runs"])
        total_percentage = self.monthly_demand['Percentage'].sum()

        for _, row in self.monthly_demand.iterrows():
            month = row['Month']
            percentage = row['Percentage'] / total_percentage
            demand = round(percentage * total_f2f_runs)
            adjusted_f2f_demand[month] = demand

        # Count courses by month
        monthly_counts = {month: 0 for month in range(1, 13)}

        for (course, methodology, language, i), week_var in schedule.items():
            assigned_week = solver.Value(week_var)
            month = self.week_to_month_map[assigned_week]
            monthly_counts[month] += 1

        # Create validation dataframe
        validation_data = []

        total_target = 0
        total_actual = 0
        all_match = True

        for month in range(1, 13):
            target = adjusted_f2f_demand.get(month, 0)
            actual = monthly_counts[month]
            diff = actual - target

            status = "MATCH" if diff == 0 else "OVER" if diff > 0 else "UNDER"
            if diff != 0:
                all_match = False

            validation_data.append({
                "Month": month,
                "Target": target,
                "Actual": actual,
                "Difference": diff,
                "Status": status
            })

            total_target += target
            total_actual += actual

        # Add total row
        validation_data.append({
            "Month": "TOTAL",
            "Target": total_target,
            "Actual": total_actual,
            "Difference": total_actual - total_target,
            "Status": "MATCH" if all_match else "MISMATCH"
        })

        return pd.DataFrame(validation_data)

    def generate_trainer_utilization_report(self, schedule, trainer_assignments, solver):
        """Generates a report on trainer utilization"""
        # Calculate days assigned to each trainer
        trainer_days = {name: 0 for name in self.consultant_data["Name"]}
        trainer_courses = {name: 0 for name in self.consultant_data["Name"]}
        champion_courses = {name: 0 for name in self.consultant_data["Name"]}

        for (course, methodology, language, i), week_var in schedule.items():
            duration = self.course_run_data.loc[self.course_run_data["Course Name"] == course, "Duration"].iloc[0]

            # Get trainer assignment
            trainer_var = trainer_assignments[(course, methodology, language, i)]
            trainer_idx = solver.Value(trainer_var)

            if 0 <= trainer_idx < len(self.fleximatrix[(course, language)]):
                trainer = self.fleximatrix[(course, language)][trainer_idx]
                trainer_days[trainer] += duration
                trainer_courses[trainer] += 1

                # Check if this is a champion course
                if self.course_champions.get((course, language)) == trainer:
                    champion_courses[trainer] += 1

        # Create trainer utilization report dataframe
        utilization_data = []

        for _, row in self.consultant_data.iterrows():
            name = row["Name"]
            title = row["Title"]
            max_days = row["Max_Days"]
            assigned_days = trainer_days[name]
            utilization = (assigned_days / max_days * 100) if max_days > 0 else 0
            courses = trainer_courses[name]
            champion = champion_courses[name]

            utilization_data.append({
                "Name": name,
                "Title": title,
                "Max Days": max_days,
                "Assigned Days": assigned_days,
                "Utilization %": round(utilization, 1),
                "Courses": courses,
                "Champion Courses": champion
            })

        # Add total row
        total_days = sum(trainer_days.values())
        max_days_sum = self.consultant_data["Max_Days"].sum()
        overall_utilization = (total_days / max_days_sum * 100) if max_days_sum > 0 else 0

        utilization_data.append({
            "Name": "TOTAL",
            "Title": "",
            "Max Days": max_days_sum,
            "Assigned Days": total_days,
            "Utilization %": round(overall_utilization, 1),
            "Courses": sum(trainer_courses.values()),
            "Champion Courses": sum(champion_courses.values())
        })

        return pd.DataFrame(utilization_data)

    def generate_excel_report(self, schedule_df, monthly_validation_df, trainer_utilization_df):
        """Generate an Excel report with multiple sheets containing all results"""
        import io  # Add this line
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            schedule_df.to_excel(writer, sheet_name='Schedule', index=False)
            monthly_validation_df.to_excel(writer, sheet_name='Monthly Validation', index=False)
            trainer_utilization_df.to_excel(writer, sheet_name='Trainer Utilization', index=False)

        output.seek(0)
        return output


# Create Streamlit application
def main():


    st.title("Course Scheduler App")
    st.write("Upload your Excel file with course and trainer data to generate an optimized schedule.")

    # Initialize session state for storing the scheduler object
    if 'scheduler' not in st.session_state:
        st.session_state.scheduler = CourseScheduler()

    if 'schedule_df' not in st.session_state:
        st.session_state.schedule_df = None

    if 'validation_df' not in st.session_state:
        st.session_state.validation_df = None

    if 'utilization_df' not in st.session_state:
        st.session_state.utilization_df = None

    # Step 1: Upload Excel File
    st.header("1. Upload Data")

    col1, col2 = st.columns(2)

    with col1:
        uploaded_file = st.file_uploader("Upload Excel file with course and trainer data", type=["xlsx"])

        if uploaded_file is not None:
            if st.button("Load Data"):
                with st.spinner("Loading data from Excel..."):
                    success = st.session_state.scheduler.load_data_from_excel(uploaded_file)
                    if success:
                        st.success("Data loaded successfully!")
                    else:
                        st.error("Failed to load data. Please check your Excel file format.")

    with col2:
        st.subheader("Required Excel Sheets")
        st.markdown("""
        Your Excel file should contain the following sheets:

        1. **CourseData**: Course information
           - Columns: Course Name, Methodology, Language, Runs, Duration

        2. **TrainerData**: Trainer information
           - Columns: Name, Title, Max_Days

        3. **PriorityData**: Trainer priority levels
           - Columns: Title, Priority

        4. **AnnualLeaves**: Trainer leave periods
           - Columns: Name, Start_Date, End_Date

        5. **AffinityMatrix**: Course affinity relationships
           - Columns: Course 1, Course 2, Gap Weeks

        6. **PublicHolidays**: Holiday periods
           - Columns: Start Date, End Date

        7. **Fleximatrix**: Trainer-course qualifications (wide format)
           - Columns: CourseName, CategoryName, Language, Champion, [Trainer1], [Trainer2], etc.
           - "U" in cells indicates trainer qualification
           - Champion column contains the trainer code who is the champion

        8. **WeekRestrictions**: Course week position constraints (optional)
           - Columns: Course, Week Type, Restricted, Notes

        9. **MonthlyDemand**: Monthly course distribution targets
           - Columns: Month, Demand
        """)

    if st.session_state.scheduler.course_run_data is not None:
        # Step 2: Configure Calendar
        st.header("2. Configure Calendar")

        col1, col2 = st.columns(2)

        with col1:
            year = st.number_input("Select scheduling year", min_value=2020, max_value=2030, value=2025)
            weekend_options = {"FS": "Friday-Saturday", "SS": "Saturday-Sunday"}
            weekend_selection = st.selectbox("Select weekend configuration",
                                             options=list(weekend_options.keys()),
                                             format_func=lambda x: weekend_options[x])

            if st.button("Initialize Calendar"):
                with st.spinner("Generating calendar..."):
                    success = st.session_state.scheduler.initialize_calendar(year, weekend_selection)
                    if success:
                        st.success(f"Calendar created with {len(st.session_state.scheduler.weekly_calendar)} weeks")
                    else:
                        st.error("Failed to create calendar")

        # Step 3: Run Optimization
        if st.session_state.scheduler.weekly_calendar is not None:
            st.header("3. Run Optimization")

            if st.button("Optimize Schedule"):
                with st.spinner("Running optimization (this may take a few minutes)..."):
                    status, schedule_df, solver, schedule, trainer_assignments = st.session_state.scheduler.run_optimization()

                    if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                        st.session_state.schedule_df = schedule_df

                        # Generate validation reports
                        st.session_state.validation_df = st.session_state.scheduler.generate_monthly_validation(
                            schedule, solver)
                        st.session_state.utilization_df = st.session_state.scheduler.generate_trainer_utilization_report(
                            schedule, trainer_assignments, solver)

                        st.success(f"Optimization completed successfully! Status: {solver.StatusName(status)}")
                    else:
                        st.error(f"Optimization failed. Status: {solver.StatusName(status)}")

            # Step 4: View Results
            if st.session_state.schedule_df is not None:
                st.header("4. Results")

                # Display tabs for different views
                tab1, tab2, tab3, tab4 = st.tabs(
                    ["Schedule", "Monthly Validation", "Trainer Utilization", "Visualizations"])

                with tab1:
                    st.subheader("Course Schedule")
                    st.dataframe(st.session_state.schedule_df)

                with tab2:
                    st.subheader("Monthly Validation")
                    st.dataframe(st.session_state.validation_df)

                with tab3:
                    st.subheader("Trainer Utilization")
                    st.dataframe(st.session_state.utilization_df)

                with tab4:
                    st.subheader("Visualizations")

                    # We need to rerun the plotting functions
                    if st.button("Generate Visualizations"):
                        with st.spinner("Generating visualizations..."):
                            # Retrieve solver and schedule from the last run
                            _, _, solver, schedule, trainer_assignments = st.session_state.scheduler.run_optimization()

                            # Generate and display charts
                            col1, col2 = st.columns(2)

                            with col1:
                                st.subheader("Weekly Course Distribution")
                                fig1 = st.session_state.scheduler.plot_weekly_course_bar_chart(schedule, solver)
                                st.pyplot(fig1)

                            with col2:
                                st.subheader("Trainer Workload")
                                fig2 = st.session_state.scheduler.plot_trainer_workload_chart(schedule,
                                                                                              trainer_assignments,
                                                                                              solver)
                                st.pyplot(fig2)

                # Step 5: Export Results
                st.header("5. Export Results")

                if st.button("Generate Excel Report"):
                    with st.spinner("Generating Excel report..."):
                        output = st.session_state.scheduler.generate_excel_report(
                            st.session_state.schedule_df,
                            st.session_state.validation_df,
                            st.session_state.utilization_df
                        )

                        st.download_button(
                            label="Download Excel Report",
                            data=output,
                            file_name="course_schedule_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )


if __name__ == "__main__":
    main()