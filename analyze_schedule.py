import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def analyze_scheduling_capacity():
    # Given data
    total_courses = 750
    total_trainers = 31
    course_duration = 5  # days
    
    # Monthly distribution percentages from the image
    monthly_distribution = {
        1: 0.0986,  # 9.86%
        2: 0.0914,  # 9.14%
        3: 0.0743,  # 7.43% (March - with holidays)
        4: 0.0886,  # 8.86%
        5: 0.0857,  # 8.57%
        6: 0.0800,  # 8.00%
        7: 0.0771,  # 7.71%
        8: 0.0829,  # 8.29%
        9: 0.0914,  # 9.14%
        10: 0.1114, # 11.14%
        11: 0.0886, # 8.86%
        12: 0.0300  # 3.00%
    }
    
    # Calculate courses per month
    courses_per_month = {month: int(round(pct * total_courses)) 
                        for month, pct in monthly_distribution.items()}
    
    # Assuming 20 working days per month (except March with holidays)
    working_days = {month: 20 for month in range(1, 13)}
    working_days[3] = 15  # March has public holidays
    
    # Calculate theoretical capacity
    max_courses_per_trainer = {month: working_days[month] // course_duration 
                             for month in range(1, 13)}
    
    theoretical_capacity = {month: max_courses_per_trainer[month] * total_trainers 
                          for month in range(1, 13)}
    
    # Analysis results
    results = pd.DataFrame({
        'Month': range(1, 13),
        'Required Courses': [courses_per_month[m] for m in range(1, 13)],
        'Working Days': [working_days[m] for m in range(1, 13)],
        'Max Courses/Trainer': [max_courses_per_trainer[m] for m in range(1, 13)],
        'Theoretical Capacity': [theoretical_capacity[m] for m in range(1, 13)]
    })
    
    results['Utilization %'] = (results['Required Courses'] / results['Theoretical Capacity'] * 100).round(1)
    
    # Plotting
    plt.figure(figsize=(15, 8))
    plt.bar(results['Month'], results['Required Courses'], alpha=0.5, label='Required Courses')
    plt.plot(results['Month'], results['Theoretical Capacity'], 'r-', label='Theoretical Capacity')
    plt.xlabel('Month')
    plt.ylabel('Number of Courses')
    plt.title('Course Requirements vs. Theoretical Capacity')
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.savefig('scheduling_analysis.png')
    plt.close()
    
    return results

if __name__ == "__main__":
    results = analyze_scheduling_capacity()
    print("\nScheduling Capacity Analysis:")
    print("=" * 80)
    print(results.to_string(index=False))
    
    # Identify potential bottlenecks
    bottlenecks = results[results['Utilization %'] > 80]
    if not bottlenecks.empty:
        print("\nPotential Bottleneck Months (>80% utilization):")
        print("-" * 80)
        print(bottlenecks.to_string(index=False)) 