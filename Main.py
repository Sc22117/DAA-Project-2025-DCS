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
