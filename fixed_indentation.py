# FIX 1: Indentation fix for line 658
# Replace the original code with this properly indented version:

            # Add penalty for being too close
            for _ in range(affinity_weight):
                affinity_penalties.append(too_close)

            print(f"  Added affinity constraint: {c1} and {c2} should be {gap_weeks} weeks apart (reduced to {reduced_gap} in Q4)")


# FIX 2: Indentation fix for the try-except block at line 1028
# Replace the original code with this properly indented version:

            for (course, delivery_type, language, i), week_var in schedule.items():
                try:
                    total_courses += 1
                    assigned_week = solver.Value(week_var)
                    if assigned_week > 0:  # Course was scheduled
                        scheduled_courses += 1
                        start_date = self.weekly_calendar[assigned_week - 1].strftime("%Y-%m-%d")

                        # Get trainer assignment if it exists
                        if (course, delivery_type, language, i) in trainer_assignments:
                            trainer_var = trainer_assignments[(course, delivery_type, language, i)]
                            trainer_idx = solver.Value(trainer_var)

                            # Check if this course has qualified trainers in the fleximatrix
                            if (course, language) in self.fleximatrix and len(self.fleximatrix[(course, language)]) > 0:
                                if 0 <= trainer_idx < len(self.fleximatrix[(course, language)]):
                                    trainer = self.fleximatrix[(course, language)][trainer_idx]
                                    is_champion = "✓" if self.course_champions.get((course, language)) == trainer else " "
                                else:
                                    # Handle out of range index
                                    trainer = "Unknown (Index Error)"
                                    is_champion = " "
                            else:
                                # Handle missing course in fleximatrix
                                trainer = "No Qualified Trainers"
                                is_champion = " "

                            schedule_results.append({
                                "Week": assigned_week,
                                "Start Date": start_date,
                                "Course": course,
                                "Delivery Type": delivery_type,
                                "Language": language,
                                "Run": i + 1,
                                "Trainer": trainer,
                                "Champion": is_champion
                            })

                    else:  # Course wasn't scheduled
                        unscheduled_courses.append({
                            "Course": course,
                            "Delivery Type": delivery_type,
                            "Language": language,
                            "Run": i + 1
                        })
                except Exception as e:
                    print(f"Error processing result for {course} (run {i+1}): {e}")
                    # Add to unscheduled due to error
                    unscheduled_courses.append({
                        "Course": course,
                        "Delivery Type": delivery_type,
                        "Language": language,
                        "Run": i + 1,
                        "Error": str(e)
                    })


# INSTRUCTIONS:
# 1. Open app.py in your code editor
# 2. Find line 658 (the affinity_penalties.append line) and fix its indentation as shown in Fix 1
# 3. Find line 1028 (the try-except block) and fix the indentation as shown in Fix 2
# 4. Save the file and run your application again 