import json
from datetime import datetime
from models import Base, Student, Session as SessionModel, Enrollment, Weekday
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import os

# Load your scraped JSON data
with open("./PortalExtension/supabase_setup/report_data.json", "r") as f:
    data = json.load(f)

# Connect to your Supabase PostgreSQL DB
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
db = SessionLocal()

def parse_time_range(time_range):
    try:
        start_str, end_str = time_range.split(" - ")
        return datetime.strptime(start_str.strip(), "%H:%M:%S").time(), datetime.strptime(end_str.strip(), "%H:%M:%S").time()
    except Exception:
        return None, None

def get_or_create_student(db, student_id, full_name):
    student = db.query(Student).filter_by(student_id=student_id).first()
    if not student:
        parts = full_name.split(" ")
        first, last = parts[0], " ".join(parts[1:]) if len(parts) > 1 else ""
        student = Student(student_id=int(student_id), first_name=first, last_name=last)
        db.add(student)
        db.commit()
    return student

def get_or_create_session(db, weekday_str, start, end):
    weekday = Weekday[weekday_str.lower()]
    session = db.query(SessionModel).filter_by(weekday=weekday, start_time=start, end_time=end).first()
    if not session:
        session = SessionModel(weekday=weekday, start_time=start, end_time=end)
        db.add(session)
        db.commit()
    return session

# Process each record
for record in data:
    if not all(k in record for k in ("Day", "Times", "Student ID", "Student Name", "Course")):
        continue

    day = record["Day"]
    start, end = parse_time_range(record["Times"])
    if not (start and end):
        continue

    student = get_or_create_student(db, record["Student ID"], record["Student Name"])
    session = get_or_create_session(db, day, start, end)

    enrollment = db.query(Enrollment).filter_by(student_id=student.student_id, session_id=session.id).first()
    if not enrollment:
        enrollment = Enrollment(student_id=student.student_id, session_id=session.id, course=record["Course"])
        db.add(enrollment)

db.commit()
db.close()