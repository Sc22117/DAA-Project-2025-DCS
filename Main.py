# Start coding here, and add your contribution with date in comments
# Contributed by : Shraddha Chauhan on 17-04-2025
import pandas as pd
import json

def clean_df(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
    df.fillna("", inplace=True)
    return df

def parse_excel(file_name, semester):
    data = pd.read_excel(file_name, sheet_name=None, engine='openpyxl')

    faculty_df = clean_df(data.get('Faculty', pd.DataFrame()))
    subject_df = clean_df(data.get('Subjects', pd.DataFrame()))
    section_df = clean_df(data.get('Sections', pd.DataFrame()))
    venue_df = clean_df(data.get('Venue', pd.DataFrame()))

    subjects = []

    for _, row in faculty_df.iterrows():
        subject_name = row.get('subject', '').strip()
        subject_type = row.get('type', '').strip()

        # Match subject from Subjects sheet
        matched_subject = subject_df[
            (subject_df['subject'].str.strip() == subject_name) &
            (subject_df['type'].str.strip() == subject_type)
        ]

        if matched_subject.empty:
            continue

        subj_row = matched_subject.iloc[0]
        subject_code = subj_row.get('subject_code', '')
        credits = subj_row.get('credit', '')

        # Match Section
        section = section_df.iloc[0].get('section', '') if not section_df.empty else ""
        strength = section_df.iloc[0].get('strength', '') if not section_df.empty else ""

        # Match Venue
        venue = venue_df.iloc[0].get('venue', '') if not venue_df.empty else ""
        capacity = venue_df.iloc[0].get('capacity', '') if not venue_df.empty else ""

        subject = {
            "id": str(row.get('id', '')),
            "faculty_name": str(row.get('name', '')),
            "subject_name": subject_name,
            "subject_code": str(subject_code),
            "type": subject_type,
            "credits": str(credits),
            "section": str(section),
            "strength": str(strength),
            "venue": str(venue),
            "capacity": str(capacity),
            "semester": semester
        }

        # Optional: only add if subject_code is not blank
        if subject["subject_code"]:
            subjects.append(subject)

    return subjects

# Parse each file
sem4 = parse_excel(r"C:\Users\shrad\OneDrive\Documents\Work\dynamic_schedular\4th Sem Details.xlsx", "4th_sem")
sem6 = parse_excel(r"C:\Users\shrad\OneDrive\Documents\Work\dynamic_schedular\6th sem details.xlsx", "6th_sem")


# Combine results
all_data = {
    "4th_sem": sem4,
    "6th_sem": sem6
}

# Save to JSON
with open("parsed_subjects.json", "w") as f:
    json.dump(all_data, f, indent=4)

print(" Parsing complete. Data saved to 'parsed_subjects.json'")
#Contributed By: Nitya Bhardwaj on 19-04-2025
import pandas as pd
import pickle

# Define file paths for both semester details
file_path_4th_sem = "C:\\Users\\NITYA\\Downloads\\4th Sem Details.xlsx"
file_path_6th_sem = "C:\\Users\\NITYA\\Downloads\\6th Sem Details.xlsx"

# Function to process and validate data for a given semester
def process_semester(file_path, semester_name):
    print(f"\nProcessing {semester_name}...")

    # Load the Excel file
    excel_file = pd.ExcelFile(file_path)

    # Parse sheets into dataframes
    faculty_df = excel_file.parse("Faculty")
    subjects_df = excel_file.parse("Subjects")
    sections_df = excel_file.parse("Sections")
    venue_df = excel_file.parse("Venue")

    # Print previews
    print(f"{semester_name} Faculty:")
    print(faculty_df.head())

    print(f"\n{semester_name} Subjects:")
    print(subjects_df.head())

    print(f"\n{semester_name} Sections:")
    print(sections_df.head())

    print(f"\n{semester_name} Venue:")
    print(venue_df.head())

    # Validate data
    def validate_df(df, name):
        print(f"\n📄 Validating '{name}' sheet for {semester_name}:")
        print("➤ Missing values:\n", df.isnull().sum())
        print("➤ Duplicates:", df.duplicated().sum())
        print("➤ Total rows:", len(df))
        print("-" * 40)

    validate_df(faculty_df, "Faculty")
    validate_df(subjects_df, "Subjects")
    validate_df(sections_df, "Sections")
    validate_df(venue_df, "Venue")

    # Mapping faculty to subjects
    faculty_subject_map = {}
    for _, row in faculty_df.iterrows():
        faculty_name = row['Name']
        subject = row['Subject']
        if faculty_name in faculty_subject_map:
            faculty_subject_map[faculty_name].append(subject)
        else:
            faculty_subject_map[faculty_name] = [subject]

    # Print faculty-to-subject mapping
    print(f"\nFaculty to Subject Mapping for {semester_name}:")
    for faculty, subjects in faculty_subject_map.items():
        print(f"{faculty}: {', '.join(subjects)}")

    # Map section to strength and subjects to credits
    subjects_df.columns = subjects_df.columns.str.strip()  # Clean column names
    section_strength_map = dict(zip(sections_df['Section'], sections_df['Strength']))
    subject_credit_map = dict(zip(subjects_df['Subject'], subjects_df['Credit']))

    print(f"\nSection → Strength Mapping for {semester_name}:")
    for section, strength in section_strength_map.items():
        print(f"{section}: {strength} students")

    # Assign rooms to sections
    available_rooms = venue_df['Venue'].tolist()
    room_assignment = {}
    for _, section_row in sections_df.iterrows():
        section = section_row['Section']
        strength = section_row['Strength']
        for room in available_rooms:
            room_capacity = venue_df.loc[venue_df['Venue'] == room, 'Capacity'].values[0]
            if room_capacity >= strength:
                room_assignment[section] = room
                available_rooms.remove(room)
                break

    print(f"\nRoom Assignments for {semester_name}:")
    for section, room in room_assignment.items():
        print(f"{section}: {room}")

    # Tag subjects as Lab or Theory
    subjects_df['Type'] = subjects_df['Subject'].apply(lambda x: 'Lab' if 'Lab' in x else 'Theory')
    print(f"\nSubjects with Tags for {semester_name}:")
    print(subjects_df[['Subject', 'Type']])

    # Faculty and Room availability matrices
    availability_matrix = {faculty: [1] * 5 for faculty in faculty_df['Name']}
    room_availability_matrix = {room: [1] * 5 for room in venue_df['Venue']}

    print(f"\nFaculty Availability Matrix for {semester_name}:")
    for faculty, availability in availability_matrix.items():
        print(f"{faculty}: {availability}")

    print(f"\nRoom Availability Matrix for {semester_name}:")
    for room, availability in room_availability_matrix.items():
        print(f"{room}: {availability}")

    # Save processed data for the semester
    pickle.dump(room_assignment, open(f'{semester_name}_room_assignment.pkl', 'wb'))
    pickle.dump(availability_matrix, open(f'{semester_name}_availability_matrix.pkl', 'wb'))
    pickle.dump(room_availability_matrix, open(f'{semester_name}_room_availability_matrix.pkl', 'wb'))

    print(f"\n{semester_name} data processing complete!")
    return {
        "faculty_df": faculty_df,
        "subjects_df": subjects_df,
        "sections_df": sections_df,
        "venue_df": venue_df,
        "faculty_subject_map": faculty_subject_map,
        "room_assignment": room_assignment,
        "availability_matrix": availability_matrix,
        "room_availability_matrix": room_availability_matrix
    }

# Process both semesters
data_4th_sem = process_semester(file_path_4th_sem, "4th Semester")
data_6th_sem = process_semester(file_path_6th_sem, "6th Semester")

