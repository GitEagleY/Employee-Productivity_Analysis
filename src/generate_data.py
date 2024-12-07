import pandas as pd
import random
from datetime import datetime, timedelta
from openpyxl import Workbook

# Generate random data
random.seed(42)
employees = ['Alice', 'Bob', 'Charlie', 'Diana', 'Edward']
task_status = ['Completed', 'Pending', 'Overdue']
start_date = datetime(2024, 1, 1)

data = {
    'TaskID': [f"T{str(i).zfill(4)}" for i in range(1, 101)],
    'EmployeeName': [random.choice(employees) for _ in range(100)],
    'TaskDescription': [f"Task {i}" for i in range(1, 101)],
    'AssignedDate': [start_date + timedelta(days=random.randint(0, 90)) for _ in range(100)],
    'Deadline': [start_date + timedelta(days=random.randint(5, 100)) for _ in range(100)],
    'CompletionDate': [
        None if random.random() < 0.2 else start_date + timedelta(days=random.randint(6, 120))
        for _ in range(100)
    ],
    'Status': [random.choice(task_status) for _ in range(100)],
}

# Create DataFrame
raw_data = pd.DataFrame(data)
raw_data.to_excel("../Data/employee_tasks_raw.xlsx", index=False)
print("Dummy raw data generated: employee_tasks_raw.xlsx")
