import unittest
import pandas as pd
from datetime import datetime
from app import CourseScheduler

class TestCourseScheduler(unittest.TestCase):
    def setUp(self):
        self.scheduler = CourseScheduler()
        
    def create_test_data(self, scenario):
        """Helper method to create test data for different scenarios"""
        if scenario == "q4_affinity":
            # Test data for Q4 affinity relaxation
            course_data = {
                "Course Name": ["Course A", "Course B", "Course A", "Course B"],
                "Delivery Type": ["F2F"] * 4,
                "Language": ["English"] * 4,
                "Runs": [1] * 4,
                "Duration": [5] * 4
            }
            
            monthly_demand = {
                "Month": list(range(1, 13)),
                "Percentage": [5] * 9 + [25, 25, 20]  # Higher demand in Q4
            }
            
            affinity_data = {
                "Course 1": ["Course A", "Course A"],
                "Course 2": ["Course B", "Course B"],
                "Gap Weeks": [4, 4]  # Should be reduced in Q4
            }
            
            return pd.DataFrame(course_data), pd.DataFrame(monthly_demand), pd.DataFrame(affinity_data)
            
        elif scenario == "trainer_availability":
            # Test data for trainer availability constraints
            course_data = {
                "Course Name": ["Course C", "Course D"],
                "Delivery Type": ["F2F"] * 2,
                "Language": ["English"] * 2,
                "Runs": [2] * 2,
                "Duration": [5] * 2
            }
            
            consultant_data = {
                "Name": ["Trainer 1", "Trainer 2"],
                "Type": ["Internal"] * 2
            }
            
            return pd.DataFrame(course_data), pd.DataFrame(consultant_data)
            
        elif scenario == "high_demand":
            # Test data for high demand periods
            course_data = {
                "Course Name": [f"Course {i}" for i in range(10)],
                "Delivery Type": ["F2F"] * 10,
                "Language": ["English"] * 10,
                "Runs": [3] * 10,  # 30 total runs
                "Duration": [5] * 10
            }
            return pd.DataFrame(course_data)

    def test_q4_affinity_relaxation(self):
        """Test that affinity constraints are properly relaxed in Q4"""
        course_data, monthly_demand, affinity_data = self.create_test_data("q4_affinity")
        
        # Set up the scheduler with test data
        self.scheduler.course_run_data = course_data
        self.scheduler.monthly_demand = monthly_demand
        self.scheduler.affinity_matrix_data = affinity_data
        
        # Run optimization
        status, schedule_df, solver, _, _, _ = self.scheduler.run_optimization(
            affinity_weight=5,  # High weight to make affinity important
            solver_time_minutes=1
        )
        
        # Verify optimization succeeded
        self.assertIn(status, [1, 2])  # OPTIMAL or FEASIBLE
        
        if schedule_df is not None:
            # Check Q4 scheduling
            q4_courses = schedule_df[schedule_df['Week'].apply(
                lambda w: self.scheduler.week_to_month_map[w] in [10, 11, 12]
            )]
            
            # Verify that some courses are scheduled closer together in Q4
            if len(q4_courses) >= 2:
                q4_gaps = []
                for course in ['Course A', 'Course B']:
                    course_weeks = q4_courses[q4_courses['Course'] == course]['Week'].sort_values()
                    if len(course_weeks) >= 2:
                        gap = course_weeks.iloc[1] - course_weeks.iloc[0]
                        q4_gaps.append(gap)
                
                # Check if any gaps are less than the original 4 weeks
                self.assertTrue(any(gap < 4 for gap in q4_gaps), 
                              "No reduced gaps found in Q4 scheduling")

    def test_trainer_availability(self):
        """Test that trainer availability constraints are respected"""
        course_data, consultant_data = self.create_test_data("trainer_availability")
        
        # Set up the scheduler with test data
        self.scheduler.course_run_data = course_data
        self.scheduler.consultant_data = consultant_data
        
        # Create a simple fleximatrix where trainers are qualified for all courses
        self.scheduler.fleximatrix = {
            ('Course C', 'English'): ['Trainer 1', 'Trainer 2'],
            ('Course D', 'English'): ['Trainer 1', 'Trainer 2']
        }
        
        # Make some weeks unavailable for trainers
        self.scheduler.trainer_availability = {
            'Trainer 1': {week: False for week in range(1, 5)},  # Unavailable weeks 1-4
            'Trainer 2': {week: False for week in range(5, 9)}   # Unavailable weeks 5-8
        }
        
        # Run optimization
        status, schedule_df, solver, _, _, _ = self.scheduler.run_optimization(
            solver_time_minutes=1
        )
        
        # Verify optimization succeeded
        self.assertIn(status, [1, 2])  # OPTIMAL or FEASIBLE
        
        if schedule_df is not None:
            # Check that no courses are scheduled when trainers are unavailable
            for _, row in schedule_df.iterrows():
                week = row['Week']
                trainer = row['Trainer']
                self.assertFalse(
                    trainer in self.scheduler.trainer_availability and 
                    week in self.scheduler.trainer_availability[trainer] and 
                    not self.scheduler.trainer_availability[trainer][week],
                    f"Course scheduled in week {week} when trainer {trainer} is unavailable"
                )

    def test_high_demand_handling(self):
        """Test scheduler's behavior under high demand conditions"""
        course_data = self.create_test_data("high_demand")
        
        # Set up the scheduler with test data
        self.scheduler.course_run_data = course_data
        
        # Run optimization
        status, schedule_df, solver, _, _, unscheduled = self.scheduler.run_optimization(
            solver_time_minutes=2,
            prioritize_all_courses=True
        )
        
        # Verify optimization attempted
        self.assertIn(status, [1, 2, 3])  # OPTIMAL, FEASIBLE, or INFEASIBLE
        
        if status in [1, 2]:  # If solution found
            # Check distribution of courses
            if schedule_df is not None:
                scheduled_courses = len(schedule_df)
                total_courses = sum(course_data['Runs'])
                
                # Calculate scheduling success rate
                success_rate = scheduled_courses / total_courses
                
                # Log the results
                print(f"\nHigh Demand Test Results:")
                print(f"Total courses: {total_courses}")
                print(f"Scheduled courses: {scheduled_courses}")
                print(f"Success rate: {success_rate:.2%}")
                print(f"Unscheduled courses: {len(unscheduled)}")
                
                # Verify reasonable success rate (adjust threshold as needed)
                self.assertGreaterEqual(success_rate, 0.7, 
                                      "Less than 70% of courses were successfully scheduled")

if __name__ == '__main__':
    unittest.main() 