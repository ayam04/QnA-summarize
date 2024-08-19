import os
import pymongo
import pandas as pd
from datetime import datetime
from bson.objectid import ObjectId
from dotenv import load_dotenv

load_dotenv()

client = pymongo.MongoClient(os.getenv("conn_string"))
db = client[os.getenv("database")]
collection = db[os.getenv("collection")]

documents = list(collection.find({
        "screeningCompletedAt": {"$exists": True, "$ne": None},
        "screeningTriggeredOn": {"$exists": True, "$ne": None},
        "jobId": {"$exists": True, "$ne": None},
    }))

print(documents[0])

def remove_object_ids(data):
    if isinstance(data, dict):
        for key, value in data.items():
            if isinstance(value, ObjectId):
                data[key] = str(value)
            elif isinstance(value, dict):
                data[key] = remove_object_ids(value)
            elif isinstance(value, list):
                data[key] = [remove_object_ids(item) for item in value]
    elif isinstance(data, list):
        data = [remove_object_ids(item) for item in data]
    return data

def time_analysis():    
    df = pd.DataFrame(documents)
    
    time_classes = [
        (0, 2, '0-2 hrs'),
        (2, 4, '2-4 hrs'),
        (4, 8, '4-8 hrs'),
        (8, 12, '8-12 hrs'),
        (12, 24, '12-24 hrs'),
        (24, 48, '24-48 hrs'),
        (48, float('inf'), '>48 hrs')
    ]
    
    # Convert ObjectId to string and inspect column names
    df = remove_object_ids(df)
    print(df.columns)  # Print the columns to check the names
    
    df['Time in Hours'] = (df['screeningCompletedAt'] - df['screeningTriggeredOn']).dt.total_seconds() / 3600
    
    def classify_time(hours):
        for lower, upper, label in time_classes:
            if lower <= hours < upper:
                return label
        return '>48 hrs'
    
    df['Time Class'] = df['Time in Hours'].apply(classify_time)
    df['Under 5 Minutes'] = df['Time in Hours'].apply(lambda x: x < (5 / 60))
    
    time_class_counts = df.groupby(['jobId', 'Time Class']).size().unstack(fill_value=0)
    total_counts = time_class_counts.sum(axis=1)
    time_class_percentage = time_class_counts.div(total_counts, axis=0) * 100
    
    under_5_minutes_percentage = df.groupby('jobId')['Under 5 Minutes'].mean() * 100
    under_5_minutes_df = under_5_minutes_percentage.reset_index()
    under_5_minutes_df.columns = ['Job ID', '<5mins']
    
    result_df = under_5_minutes_df.merge(time_class_percentage, left_on='Job ID', right_on='jobId')
    result_df.to_excel("Screening_Time_Analysis.xlsx", index=False)

def average_time_analysis():
    df = pd.DataFrame(documents)
    
    time_classes = [
        (0, 2, '0-2 hrs'),
        (2, 4, '2-4 hrs'),
        (4, 8, '4-8 hrs'),
        (8, 12, '8-12 hrs'),
        (12, 24, '12-24 hrs'),
        (24, 48, '24-48 hrs'),
        (48, float('inf'), '>48 hrs')
    ]
    
    df['Time in Hours'] = (df['screeningCompletedAt'] - df['screeningTriggeredOn']).dt.total_seconds() / 3600
    df = df[df['Time in Hours'].notnull()]
    
    def classify_time(hours):
        for lower, upper, label in time_classes:
            if lower <= hours < upper:
                return label
        return '>48 hrs'
    
    df['Time Class'] = df['Time in Hours'].apply(classify_time)
    df['Under 5 Minutes'] = df['Time in Hours'].apply(lambda x: x < (5 / 60))
    
    time_class_counts = df['Time Class'].value_counts()
    total_counts = len(df)
    time_class_percentage = (time_class_counts / total_counts) * 100
    under_5_minutes_percentage = df['Under 5 Minutes'].mean() * 100
    
    time_class_df = time_class_percentage.reset_index()
    time_class_df.columns = ['Time Class', 'Percentage']
    
    under_5_minutes_df = pd.DataFrame({
        'Time Class': ['Under 5 Minutes'],
        'Percentage': [under_5_minutes_percentage]
    })
    
    combined_df = pd.concat([time_class_df, under_5_minutes_df], ignore_index=True)
    combined_df.to_excel("Avg_Screening_Time_Analysis.xlsx", sheet_name='Time Analysis', index=False)

def answer_analysis():
    df = pd.DataFrame(documents)
    
    result = df.groupby(['jobId', 'qualifyingQuestions', 'candidateStatus']).size().unstack(fill_value=0)
    
    if 'Correct Answer' in result.columns:
        result['% Correct'] = result['Correct Answer'] / (result['Correct Answer'] + result['Wrong Answer'] + result.get('Not Applicable', 0)) * 100
        result['% Wrong'] = result['Wrong Answer'] / (result['Correct Answer'] + result['Wrong Answer'] + result.get('Not Applicable', 0)) * 100
        result['% NA'] = result.get('Not Applicable', 0) / (result['Correct Answer'] + result['Wrong Answer'] + result.get('Not Applicable', 0)) * 100
    else:
        result['% Correct'] = result['Correct'] / (result['Correct'] + result['Wrong'] + result.get('NA', 0)) * 100
        result['% Wrong'] = result['Wrong'] / (result['Correct'] + result['Wrong'] + result.get('NA', 0)) * 100
        result['% NA'] = result.get('NA', 0) / (result['Correct'] + result['Wrong'] + result.get('NA', 0)) * 100
    
    final_result = result[["% Correct", "% Wrong", "% NA"]].reset_index()
    final_result.to_excel('Answer_Analysis.xlsx', index=False)

time_analysis()
# average_time_analysis()
# answer_analysis()