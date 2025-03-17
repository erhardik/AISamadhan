import pandas as pd

# Create sample data
data = {
    'Parameter': ['Attendance Bonus', 'Subject 1'],
    'Subject Name': ['', 'Mathematics'],
    'Theory Marks': [7, 32],
    'Practical Marks': [None, 33],
    'Theory Credits': [None, 3],
    'Practical Credits': [None, 1],
    'Value': [7, None]
}

# Create DataFrame and save to Excel
df = pd.DataFrame(data)
df.to_excel('input.xlsx', index=False)
print('Sample input Excel file created successfully!')