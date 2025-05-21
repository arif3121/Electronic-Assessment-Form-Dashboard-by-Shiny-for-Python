import pandas as pd
import os

filename = "student_records.xlsx"
print("Testing update at:", os.path.abspath(filename))

student_id = input("Enter an existing Student_ID from your Excel file: ").strip()

marks = 88
comment = "Automated test comment"

# Read file and print before update
df = pd.read_excel(filename)
print("Before update:", df[df["Student_ID"].astype(str) == student_id])

# Update marks and comment for this student
df.loc[df["Student_ID"].astype(str) == student_id, "Marks"] = marks
df.loc[df["Student_ID"].astype(str) == student_id, "Comment"] = comment

# Save back to Excel
df.to_excel(filename, index=False)
print("After update:", pd.read_excel(filename)[pd.read_excel(filename)["Student_ID"].astype(str) == student_id])
