import pandas as pd
import re

def time_analysis(path):
    df = pd.read_excel(path)
    
    time_classes = [
        (0, 2, '0-2 hrs'),
        (2, 4, '2-4 hrs'),
        (4, 8, '4-8 hrs'),
        (8, 12, '8-12 hrs'),
        (12, 24, '12-24 hrs'),
        (24, 48, '24-48 hrs'),
        (48, float('inf'), '>48 hrs')
    ]
    
    def convert_time_to_hours(time_str):
        if isinstance(time_str, str):
            if 'Pending' in time_str:
                return 'Screen Pending'
            if 'Less Than Minute' in time_str:
                return 1 / 60
            if 'Minute' in time_str:
                minutes = int(re.search(r'(\d+) Minute', time_str).group(1))
                return minutes / 60
            if 'Hour' in time_str:
                hours = int(re.search(r'(\d+) Hour', time_str).group(1))
                return hours
            if 'Day' in time_str:
                days = int(re.search(r'(\d+) Day', time_str).group(1))
                return days * 24
        return None
    
    df['Time in Hours'] = df['Screening Completed In'].apply(convert_time_to_hours)
    
    def classify_time(hours):
        if hours == 'Screen Pending':
            return hours
        for lower, upper, label in time_classes:
            if lower <= hours < upper:
                return label
        return '>48 hrs'
    
    df['Time Class'] = df['Time in Hours'].apply(classify_time)
    df['Under 5 Minutes'] = df['Time in Hours'].apply(lambda x: x < (5 / 60) if isinstance(x, (int, float)) else 0)
    
    time_class_counts = df.groupby(['Job Title', 'Time Class']).size().unstack(fill_value=0)
    total_counts = time_class_counts.sum(axis=1)
    time_class_percentage = time_class_counts.div(total_counts, axis=0) * 100
    
    under_5_minutes_percentage = df.groupby('Job Title')['Under 5 Minutes'].mean() * 100
    under_5_minutes_df = under_5_minutes_percentage.reset_index()
    under_5_minutes_df.columns = ['Job Title', '<5mins']
    
    result_df = under_5_minutes_df.reset_index().merge(time_class_percentage, on='Job Title')
    output_file_path = "Screening_Time_Analysis.xlsx"
    result_df.to_excel(output_file_path)


def answer_analysis(path):
    df = pd.read_excel(path)
    result = df.groupby(['Job Title', 'Question', 'Correct Answer']).size().unstack(fill_value=0)
    result['% Correct'] = result['Correct Answer'] / (result['Correct Answer'] + result['Wrong Answer'] + result['Not Applicable']) * 100
    result['% Wrong'] = result['Wrong Answer'] / (result['Correct Answer'] + result['Wrong Answer'] + result['Not Applicable']) * 100
    result['% NA'] = result['Not Applicable'] / (result['Correct Answer'] + result['Wrong Answer'] + result['Not Applicable']) * 100
    final_result = result[["% Correct", "% Wrong", "% NA"]].reset_index()
    final_result.to_excel('Answer_Analysis.xlsx', index=False)
    # print(final_result)


# answer_analysis('Tests/test.xlsx')
time_analysis("Tests/Time Analysis Test.xlsx")