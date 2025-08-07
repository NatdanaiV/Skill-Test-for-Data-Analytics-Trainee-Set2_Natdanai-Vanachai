import pandas as pd
import glob
import os

new_employee_df = pd.read_excel('New_Employee_202502.xlsx')
result = []
daily_files = glob.glob('Daily_report_*.xls')

for file in daily_files:
    base = os.path.basename(file)
    team_member = base.split('_')[2] + ' ' + base.split('_')[3].replace('.xls', '')
    df = pd.read_excel(file)
    passed = df[(df['Interview'] == 'Yes') & (df['Status'] == 'Pass')]   
    for _, row in passed.iterrows():
        name = row['Candidate Name']
        matched = new_employee_df[new_employee_df['Employee Name'] == name]       
        if not matched.empty:
            employee = matched.iloc[0]
            result.append({
                'Employee Name': employee['Employee Name'],
                'Join Date': employee['Join Date'],
                'Role': employee['Role'],
                'Team Member': team_member
            })
result_df = pd.DataFrame(result)
print(result_df)
result_df.to_excel('Summary_New_Employees.xlsx', index=False)