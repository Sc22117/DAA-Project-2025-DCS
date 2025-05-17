# Contributed by Varima Dudeja on May 6, 2025

import firebase_admin
from firebase_admin import credentials, db
import pandas as pd
import random
from datetime import datetime, timedelta
from collections import defaultdict

# Initialize Firebase
cred = credentials.Certificate("serviceAccountKey.json")
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://dynamic-classroom-scheduler-default-rtdb.asia-southeast1.firebasedatabase.app/'
})

def allocate_faculties():
    """Randomly allocate faculties to sections based on subjects"""
    faculty = db.reference('faculty').get()
    subjects = db.reference('subjects').get()
    sections = db.reference('sections').get()
    
    # Create faculty pool by subject
    faculty_pool = {}
    for fac_id, fac_info in faculty.items():
        subject = fac_info['subject']
        if subject not in faculty_pool:
            faculty_pool[subject] = []
        faculty_pool[subject].append(fac_id)
    
    updates = {}
    for section_id, section_info in sections.items():
        # Allocate faculty for each subject
        for subj_code, subj_info in subjects.items():
            if subj_info['subject'] in faculty_pool:
                # Randomly select faculty for this subject
                selected_faculty = random.choice(faculty_pool[subj_info['subject']])
                updates[f'sections/{section_id}/allocations/{subj_code}'] = {
                    'faculty_id': selected_faculty,
                    'faculty_name': faculty[selected_faculty]['name'],
                    'subject_name': subj_info['subject']
                }
    
    db.reference().update(updates)

def allocate_venues():
    """Allocate venues to sections based on capacity and type"""
    sections = db.reference('sections').get()
    venues = db.reference('venues').get()
    
    # Sort venues by capacity
    sorted_venues = sorted(venues.items(), key=lambda x: x[1]['capacity'])
    
    updates = {}
    for section_id, section_info in sections.items():
        strength = section_info['strength']
        
        # Find suitable venue
        for venue_id, venue_info in sorted_venues:
            if venue_info['capacity'] >= strength:
                updates[f'sections/{section_id}/venue'] = venue_id
                break
    
    db.reference().update(updates)

#Contributed by Nitya Bhardwaj on May 8, 2025

from openpyxl import Workbook
def check_clashes(new_slot, existing_timetables):
    """Check if new slot clashes with existing schedules"""
    if not isinstance(new_slot, dict):
        return False

    for section_id, timetable in existing_timetables.items():
        if not isinstance(timetable, dict):
            continue

        for day, day_slots in timetable.items():
            if not isinstance(day_slots, dict):
                continue

            for slot_key, slot in day_slots.items():
                if not isinstance(slot, dict):
                    continue

                try:
                    # Faculty can't be in two places at once
                    faculty_clash = new_slot.get('faculty_id') == slot.get('faculty_id')
                    # Venue can't be double booked
                    venue_clash = new_slot.get('venue') == slot.get('venue')
                    # Time overlap check
                    time_clash = ((new_slot.get('start') <= slot.get('start') < new_slot.get('end')) or
                                (slot.get('start') <= new_slot.get('start') < slot.get('end')))

                    if faculty_clash and time_clash:
                        return True
                    if venue_clash and time_clash:
                        return True
                except (TypeError, ValueError):
                    continue

    return False

#Contributed by Varima Dudeja on May 9, 2025

def allocate_electives(sections, elective_subjects, venues):
    """Allocate elective subjects to sections with proper venue assignment"""
    section_ids = list(sections.keys())

    for subj_code, subj_info in elective_subjects.items():
        # Get suitable venues for this elective (classroom type)
        suitable_venues = [
            v_id for v_id, v_info in venues.items()
            if v_info.get('type') == 'Classroom'  # Electives typically in classrooms
        ]

        for section_id in section_ids:
            if 'allocations' not in sections[section_id]:
                sections[section_id]['allocations'] = {}

            # Assign random faculty and venue for this elective
            faculty_id = random.choice(subj_info['faculty_ids'])
            venue = random.choice(suitable_venues) if suitable_venues else ''

            sections[section_id]['allocations'][subj_code] = {
                'subject_name': f"{subj_info['name']} (Elective)",
                'faculty_id': faculty_id,
                'is_elective': True,
                'venue': venue,  # Store venue at allocation time
                'original_name': subj_info['name']  # Store original name for reference
            }

    return sections


#Contributed by Shraddha Chauhan on May 10, 2025

def generate_time_slots(start_time, end_time, period_duration, break_duration, has_lunch=True):
    """Generate time slots with proper breaks"""
    start = datetime.strptime(start_time, "%H:%M")
    end = datetime.strptime(end_time, "%H:%M")
    period_delta = timedelta(minutes=period_duration)

    time_slots = []
    current_time = start

    # Morning slots
    while current_time + period_delta <= end:
        end_time = current_time + period_delta
        time_slots.append({
            'start': current_time.strftime("%H:%M"),
            'end': end_time.strftime("%H:%M"),
            'is_lunch': False
        })
        current_time = end_time

        # Add short break between periods (except before lunch)
        if not (has_lunch and current_time.hour >= 12 and current_time.hour < 13):
            current_time += timedelta(minutes=5)

    # Insert lunch break if needed
    if has_lunch:
        lunch_start = datetime.strptime("12:00", "%H:%M")
        lunch_end = lunch_start + timedelta(minutes=break_duration)

        # Find where to insert lunch
        for i, slot in enumerate(time_slots):
            if datetime.strptime(slot['start'], "%H:%M") <= lunch_start < datetime.strptime(slot['end'], "%H:%M"):
                # Split the slot at lunch time
                pre_lunch = {
                    'start': slot['start'],
                    'end': lunch_start.strftime("%H:%M"),
                    'is_lunch': False
                }
                lunch_slot = {
                    'start': lunch_start.strftime("%H:%M"),
                    'end': lunch_end.strftime("%H:%M"),
                    'is_lunch': True
                }
                post_lunch = {
                    'start': lunch_end.strftime("%H:%M"),
                    'end': slot['end'],
                    'is_lunch': False
                }

                # Replace current slot with split slots
                time_slots[i:i+1] = [pre_lunch, lunch_slot, post_lunch]
                break

    return time_slots

#Contributed by Yashita Jain on May 11, 2025

def generate_timetable(section_id, working_days, period_duration, start_time, end_time, lunch_duration):
    """Generate timetable for a section"""
    # Get all data from Firebase
    all_sections = db.reference('sections').get() or {}
    subjects = db.reference('subjects').get() or {}
    venues = db.reference('venues').get() or {}
    faculty = db.reference('faculty').get() or {}

    # Get current section data
    section = all_sections.get(section_id)
    if not section:
        print(f"Section {section_id} not found")
        return

    # Generate time slots
    time_slots = generate_time_slots(start_time, end_time, period_duration, lunch_duration)

    # Prepare days (exactly the number of working days requested)
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][:working_days]

    # Initialize timetable structure
    timetable = {day: {} for day in days}

    # Get all subjects for this section
    section_subjects = section.get('allocations', {})

    # Calculate total periods needed
    total_periods = sum(
        subjects.get(subj_code, {}).get('credits', 0)
        for subj_code in section_subjects
    )

    # Prepare subject allocations with proper weighting
    subject_allocations = []
    for subj_code, alloc in section_subjects.items():
        credits = subjects.get(subj_code, {}).get('credits', 0)
        subject_allocations.extend([(subj_code, alloc)] * credits)

    # Special handling for electives (longer slots)
    for subj_code, alloc in section_subjects.items():
        if alloc.get('is_elective', False):
            # Add extra slots for electives (2-hour blocks)
            subject_allocations.extend([(subj_code, alloc)] * 2)

    random.shuffle(subject_allocations)

    # Assign subjects to time slots
    for day in days:
        day_slots = [s for s in time_slots if not s['is_lunch']]

        for slot in day_slots:
            if not subject_allocations:
                break

            subj_code, alloc = subject_allocations.pop()
            subject = subjects.get(subj_code, {})

            # For electives, check group and assign appropriate time
            if alloc.get('is_elective', False):
                if alloc.get('group', 1) == 1 and day not in ["Friday", "Saturday"]:
                    continue
                if alloc.get('group', 1) == 2 and day in ["Friday", "Saturday"]:
                    continue

            # Select appropriate venue
            venue_type = 'Lab' if 'Lab' in subject.get('type', '') else 'Classroom'
            suitable_venues = [
                v_id for v_id, v_info in venues.items()
                if v_info.get('type') == venue_type and
                   v_info.get('capacity', 0) >= section.get('strength', 0)
            ]
            venue = random.choice(suitable_venues) if suitable_venues else ''

            new_slot = {
                'day': day,
                'start': slot['start'],
                'end': slot['end'],
                'subject_code': subj_code,
                'subject_name': alloc.get('subject_name', ''),
                'faculty_id': alloc.get('faculty_id', ''),
                'faculty_name': faculty.get(alloc.get('faculty_id', ''), {}).get('name', ''),
                'venue': venue,
                'is_elective': alloc.get('is_elective', False)
            }

            # For electives, make longer slots (2 periods)
            if alloc.get('is_elective', False):
                # Find next consecutive slot
                next_slot_index = day_slots.index(slot) + 1
                if next_slot_index < len(day_slots):
                    next_slot = day_slots[next_slot_index]
                    new_slot['end'] = next_slot['end']
                    # Remove the next slot from allocation pool
                    day_slots.pop(next_slot_index)

            timetable[day][slot['start']] = new_slot

    # Save to Firebase
    db.reference(f'sections/{section_id}/timetable').set(timetable)

    # Generate Excel
    generate_excel(section_id, timetable)


#Contributed by Varima Dudeja on May 13, 2025

def generate_excel(section_id, timetable):
    """Generate Excel timetable with proper elective display and separate electives sheet"""
    wb = Workbook()
    ws_main = wb.active
    ws_main.title = f"Section {section_id}"

    # Headers for main sheet
    headers = ["Day", "Time", "Subject", "Faculty", "Venue", "Type"]
    ws_main.append(headers)

    # Prepare data for electives details sheet
    electives_data = []

    # Add data to main sheet and collect electives info
    for day, slots in timetable.items():
        for slot in sorted(slots.values(), key=lambda x: x['start']):
            # For electives, show just "Elective" in main sheet
            if slot.get('is_elective', False):
                subject_display = "Elective"
                # Collect details for electives sheet
                electives_data.append([
                    slot.get('subject_name', '').replace(' (Elective)', ''),
                    slot.get('faculty_name', ''),
                    slot.get('venue', ''),
                    day,
                    f"{slot['start']}-{slot['end']}"
                ])
            else:
                subject_display = slot.get('subject_name', '')

            ws_main.append([
                day,
                f"{slot['start']}-{slot['end']}",
                subject_display,
                slot['faculty_name'],
                slot['venue'],
                "Elective" if slot.get('is_elective', False) else "Regular"
            ])

    # Create electives details sheet if there are any electives
    if electives_data:
        ws_electives = wb.create_sheet("Electives Details")
        ws_electives.append(["Elective Subject", "Faculty", "Venue", "Day", "Time"])
        for row in electives_data:
            ws_electives.append(row)

    filename = f"Timetable_{section_id}.xlsx"
    wb.save(filename)
    print(f"Timetable saved to {filename}")

def main():
    # Initialize data if needed
    if not db.reference('faculty').get():
        'faculty': pd.read_excel(file_path, sheet_name='Faculty'),
        'subjects': pd.read_excel(file_path, sheet_name='Subjects'),
        'sections': pd.read_excel(file_path, sheet_name='Sections'),
        'venues': pd.read_excel(file_path, sheet_name='Venues')

        # Process faculty
        faculty_data = {
        row['ID']: {
            'name': row['Name'],
            'subject': row['Subject'],
            'type': row['Type']
        }
        for _, row in data['faculty'].iterrows()
        }
        db.reference('faculty').set(faculty_data)
    
        # Process subjects
        subjects_data = {
        row['Subject Code']: {
            'subject': row['Subject'],
            'type': row['Type'],
            'credits': int(row['Credits'])
        }
        for _, row in data['subjects'].iterrows()
        }
        db.reference('subjects').set(subjects_data)
    
        # Process sections
        sections_data = {
        row['Section']: {
            'strength': int(row['Strength']),
            'allocations': {},
            'timetable': {}
        }
        for _, row in data['sections'].iterrows()
    }
    db.reference('sections').set(sections_data)
    
    # Process venues
    venues_data = {
        row['Classroom/Lab']: {
            'capacity': int(row['Capacity']),
            'type': 'Lab' if 'Lab' in row['Classroom/Lab'] else 'Classroom'
        }
        for _, row in data['venues'].iterrows()
    }
    db.reference('venues').set(venues_data)

    # Get user inputs
    section_id = input("Enter section ID (e.g., DS1): ").strip().upper()
    working_days = int(input("Enter number of working days (5 or 6): "))
    period_duration = int(input("Enter period duration in minutes: "))
    start_time = input("Enter start time (HH:MM): ")
    end_time = input("Enter end time (HH:MM): ")
    lunch_duration = int(input("Enter lunch duration in minutes: "))

    generate_timetable(
        section_id=section_id,
        working_days=working_days,
        period_duration=period_duration,
        start_time=start_time,
        end_time=end_time,
        lunch_duration=lunch_duration
    )

if __name__ == "__main__":
    main()
