import pandas as pd
from collections import defaultdict

# Load your dataset
file_path = "datass.xlsx"  # Update the file path as needed
df = pd.read_excel(file_path)

# Extract relevant columns (including all information)
students = df[['Informazioni cronologiche', 'Vor- und Zuname:', 'Unikennung:',
                'Matrikelnummer:', 'Uni E-Mail Adresse',
                'Sollten Sie einen Kurs zu einer bestimmten Uhrzeit benötigen (aufgrund einer Erkrankung, Kinderbetreuung oder Angehörigenpflege), so können Sie uns hier den Grund und Ihre präferierte Uhrzeit mitteilen. Ansonsten lassen Sie dieses Feld bitte frei.',
                ' [Wunsch 1]', ' [Wunsch 2]', ' [Wunsch 3]', ' [Wunsch 4]',
                ' [Wunsch 5]', ' [Wunsch 6]', ' [Wunsch 7]', ' [Wunsch 8]',
                ' [Wunsch 9]', ' [Wunsch 10]', ' [Wunsch 11]', ' [Wunsch 12]',
                ' [Wunsch 13]', ' [Wunsch 14]', ' [Wunsch 15]']]

# Initialize a dictionary to store students assigned to each group
group_assignments = defaultdict(list)

# Set the maximum number of students per group
num_groups = 15
max_students_per_group = len(students) // num_groups

# Function to assign students to groups based on preferences
def assign_students_to_groups(students):
    for _, student in students.iterrows():
        # Store all student information in a dictionary
        student_info = student.to_dict()

        # Go through the student's preferences
        for i in range(1, num_groups + 1):  # 15 preferences
            group = student[f' [Wunsch {i}]']
            if pd.notna(group) and len(group_assignments[group]) < max_students_per_group:
                # Append the student info dictionary to the appropriate group
                group_assignments[group].append(student_info)
                break  # Stop once the student is assigned to a group

# Assign students to groups
assign_students_to_groups(students)

# Prepare the output data to include all information
output_data = []
for group, assigned_students in group_assignments.items():
    for student in assigned_students:
        student_with_group = {**student, 'Assigned Group': group}  # Include assigned group
        output_data.append(student_with_group)

# Check for unassigned students
assigned_student_count = len(output_data)
remaining_students_count = len(students) - assigned_student_count

if remaining_students_count > 0:
    # Identify unassigned students
    assigned_ids = {student['Unikennung:'] for student in output_data}
    remaining_students = students[~students['Unikennung:'].isin(assigned_ids)]

    # Assign remaining students to the first few groups (Group 1, Group 2, Group 3)
    for i, student in enumerate(remaining_students.iterrows()):
        student_info = student[1].to_dict()
        group_to_assign = f'Group {i % 3 + 1}'  # Assign to Group 1, 2, or 3
        group_assignments[group_to_assign].append(student_info)

        # Add to output data
        student_with_group = {**student_info, 'Assigned Group': group_to_assign}
        output_data.append(student_with_group)

# Convert the output data into a DataFrame
output_df = pd.DataFrame(output_data)

# Save the group distribution to a new Excel file
output_file = "group_distribution_with_info.xlsx"  # Change as needed
output_df.to_excel(output_file, index=False)

print(f"Group distribution with student information saved to {output_file}")
