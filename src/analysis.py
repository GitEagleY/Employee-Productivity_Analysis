import pandas as pd
import random
from datetime import datetime, timedelta
from openpyxl import Workbook

# Load transformed data
transformed_data = pd.read_excel("../Data/employee_tasks_transformed.xlsx")

# Analyze productivity
employee_summary = transformed_data.groupby('EmployeeName').agg(
    TotalTasks=('TaskID', 'count'),
    CompletedTasks=('CompletionPercentage', lambda x: (x > 0).sum()),
    AvgDelay=('DelayDays', 'mean'),
).reset_index()

# Save analysis
employee_summary.to_excel("../Data/employee_productivity_analysis.xlsx", index=False)
print("Productivity analysis saved: employee_productivity_analysis.xlsx")
