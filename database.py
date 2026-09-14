import sqlite3
import pandas as pd
import os

DB_NAME = "timesheet.db"

def get_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the database schema and inserts sample data if empty."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Instructors Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS instructors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        email TEXT,
        phone TEXT,
        department TEXT,
        employment_status TEXT CHECK(employment_status IN ('SC-Based', 'Part-Time', 'Pro-Rated'))
    )
    """)
    
    # 2. Class Loads Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS class_loads (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        instructor_id INTEGER NOT NULL,
        subject_name TEXT NOT NULL,
        section_code TEXT NOT NULL,
        room TEXT NOT NULL,
        day_of_week TEXT NOT NULL CHECK(day_of_week IN ('M', 'T', 'W', 'TH', 'F', 'S', 'SU')),
        start_time TEXT NOT NULL, -- HH:MM format
        end_time TEXT NOT NULL,   -- HH:MM format
        type TEXT NOT NULL CHECK(type IN ('LEC', 'LAB')),
        units REAL NOT NULL,
        hours REAL NOT NULL,
        load_category TEXT NOT NULL CHECK(load_category IN ('Regular Load', 'Excess/Overload', 'Consultation')),
        enrollment INTEGER DEFAULT 0,
        FOREIGN KEY(instructor_id) REFERENCES instructors(id) ON DELETE CASCADE
    )
    """)
    
    # 3. Timesheet Overrides Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS timesheet_overrides (
        instructor_id INTEGER NOT NULL,
        year INTEGER NOT NULL,
        month INTEGER NOT NULL,
        cutoff_type TEXT NOT NULL CHECK(cutoff_type IN ('11-25', '26-10')),
        row_name TEXT NOT NULL,
        date_str TEXT NOT NULL, -- YYYY-MM-DD format
        hours REAL NOT NULL,
        PRIMARY KEY(instructor_id, year, month, cutoff_type, row_name, date_str),
        FOREIGN KEY(instructor_id) REFERENCES instructors(id) ON DELETE CASCADE
    )
    """)
    
    conn.commit()
    
    # Prepopulate with sample data if instructors table is empty
    cursor.execute("SELECT COUNT(*) as count FROM instructors")
    if cursor.fetchone()['count'] == 0:
        # Prepopulate Jane Doe (SC-Based)
        cursor.execute("""
        INSERT INTO instructors (name, email, phone, department, employment_status)
        VALUES ('Jane Doe', 'jane.doe@university.edu', '555-0192', 'Computer Science', 'SC-Based')
        """)
        jane_id = cursor.lastrowid
        
        # Prepopulate class loads for Jane Doe
        # CP5 Lecture Wed 12:30PM-2:00PM (1.5h)
        # CP5 Lab Fri 11:00AM-2:00PM (3h)
        # Data Structures Lecture Tue 9:00AM-12:00PM (3h)
        # Consultation Mon 10:00AM-12:00PM (2h)
        classes = [
            (jane_id, 'Computer Programming 5', 'ICTE201P', 'Room 301', 'W', '12:30', '14:00', 'LEC', 2.0, 1.5, 'Regular Load', 35),
            (jane_id, 'Computer Programming 5', 'ICTE201P', 'Lab 3', 'F', '11:00', '14:00', 'LAB', 1.0, 3.0, 'Regular Load', 35),
            (jane_id, 'Data Structures', 'ICTE202P', 'Room 402', 'T', '09:00', '12:00', 'LEC', 3.0, 3.0, 'Regular Load', 30),
            (jane_id, 'Consultation Hours', 'N/A', 'Faculty Room', 'M', '10:00', '12:00', 'LEC', 0.0, 2.0, 'Consultation', 0)
        ]
        cursor.executemany("""
        INSERT INTO class_loads (instructor_id, subject_name, section_code, room, day_of_week, start_time, end_time, type, units, hours, load_category, enrollment)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, classes)
        
        # Prepopulate John Smith (Part-Time)
        cursor.execute("""
        INSERT INTO instructors (name, email, phone, department, employment_status)
        VALUES ('John Smith', 'john.smith@university.edu', '555-0143', 'Information Technology', 'Part-Time')
        """)
        john_id = cursor.lastrowid
        
        # John Smith class loads
        john_classes = [
            (john_id, 'Web Development', 'ICTE302', 'Lab 5', 'TH', '13:00', '16:00', 'LAB', 2.0, 3.0, 'Excess/Overload', 28),
            (john_id, 'Database Systems', 'ICTE305', 'Room 204', 'M', '14:00', '17:00', 'LEC', 3.0, 3.0, 'Regular Load', 32),
            (john_id, 'Database Systems', 'ICTE305', 'Lab 1', 'W', '09:00', '12:00', 'LAB', 1.0, 3.0, 'Regular Load', 32)
        ]
        cursor.executemany("""
        INSERT INTO class_loads (instructor_id, subject_name, section_code, room, day_of_week, start_time, end_time, type, units, hours, load_category, enrollment)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, john_classes)
        
        conn.commit()
        
    conn.close()

# Instructor functions
def get_instructors():
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM instructors", conn)
    conn.close()
    return df

def add_instructor(name, email, phone, department, employment_status):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO instructors (name, email, phone, department, employment_status)
    VALUES (?, ?, ?, ?, ?)
    """, (name, email, phone, department, employment_status))
    conn.commit()
    instructor_id = cursor.lastrowid
    conn.close()
    return instructor_id

def update_instructor(instructor_id, name, email, phone, department, employment_status):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE instructors
    SET name=?, email=?, phone=?, department=?, employment_status=?
    WHERE id=?
    """, (name, email, phone, department, employment_status, instructor_id))
    conn.commit()
    conn.close()

def delete_instructor(instructor_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM instructors WHERE id=?", (instructor_id,))
    conn.commit()
    conn.close()

# Class Load functions
def get_class_loads(instructor_id):
    conn = get_connection()
    df = pd.read_sql_query("SELECT * FROM class_loads WHERE instructor_id=?", conn, params=(int(instructor_id),))
    conn.close()
    return df

def add_class_load(instructor_id, subject_name, section_code, room, day_of_week, start_time, end_time, type_, units, hours, load_category, enrollment):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO class_loads (instructor_id, subject_name, section_code, room, day_of_week, start_time, end_time, type, units, hours, load_category, enrollment)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (int(instructor_id), subject_name, section_code, room, day_of_week, start_time, end_time, type_, float(units), float(hours), load_category, int(enrollment)))
    conn.commit()
    conn.close()

def delete_class_load(class_load_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM class_loads WHERE id=?", (int(class_load_id),))
    conn.commit()
    conn.close()

# Timesheet Overrides functions
def get_timesheet_overrides(instructor_id, year, month, cutoff_type):
    """
    Returns a dictionary of overrides for a specific timesheet:
    Key: (row_name, date_str)
    Value: hours (float)
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    SELECT row_name, date_str, hours FROM timesheet_overrides
    WHERE instructor_id=? AND year=? AND month=? AND cutoff_type=?
    """, (int(instructor_id), int(year), int(month), cutoff_type))
    rows = cursor.fetchall()
    conn.close()
    
    overrides = {}
    for r in rows:
        overrides[(r['row_name'], r['date_str'])] = r['hours']
    return overrides

def save_timesheet_overrides(instructor_id, year, month, cutoff_type, overrides_list):
    """
    Saves the list of overrides.
    overrides_list is a list of dicts: [{'row_name': ..., 'date_str': ..., 'hours': ...}]
    """
    conn = get_connection()
    cursor = conn.cursor()
    
    # First, delete existing overrides for this combination
    cursor.execute("""
    DELETE FROM timesheet_overrides
    WHERE instructor_id=? AND year=? AND month=? AND cutoff_type=?
    """, (int(instructor_id), int(year), int(month), cutoff_type))
    
    # Insert new overrides
    to_insert = [
        (int(instructor_id), int(year), int(month), cutoff_type, x['row_name'], x['date_str'], float(x['hours']))
        for x in overrides_list
    ]
    cursor.executemany("""
    INSERT OR REPLACE INTO timesheet_overrides (instructor_id, year, month, cutoff_type, row_name, date_str, hours)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, to_insert)
    
    conn.commit()
    conn.close()

def clear_timesheet_overrides(instructor_id, year, month, cutoff_type):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    DELETE FROM timesheet_overrides
    WHERE instructor_id=? AND year=? AND month=? AND cutoff_type=?
    """, (int(instructor_id), int(year), int(month), cutoff_type))
    conn.commit()
    conn.close()
