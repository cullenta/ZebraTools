import json
from datetime import datetime
from supabase_setup.models import Student, Session as SessionModel, Enrollment, Weekday, Course, Trial, MakeupSession
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import os


def insert_data_from_json():

    # Load data
    with open("./supabase_setup/report_data.json", "r") as f:
        data = json.load(f)

    # DB setup
    DATABASE_URL = os.getenv("DATABASE_URL")
    engine = create_engine(DATABASE_URL)
    SessionLocal = sessionmaker(bind=engine)
    db = SessionLocal()

    def parse_time_range(time_range):
        try:
            start_str, end_str = time_range.split(" - ")
            return datetime.strptime(start_str.strip(), "%H:%M:%S").time(), datetime.strptime(end_str.strip(), "%H:%M:%S").time()
        except:
            return None, None

    def get_or_create_student(student_id, full_name):
        student = db.query(Student).filter_by(student_id=student_id).first()
        if not student:
            first, *rest = full_name.strip().split()
            last = " ".join(rest)
            student = Student(student_id=int(student_id), first_name=first, last_name=last)
            db.add(student)
            db.commit()
        return student

    def get_or_create_session(day_str, start, end):
        weekday = Weekday[day_str.lower()]
        session = db.query(SessionModel).filter_by(weekday=weekday, start_time=start, end_time=end).first()
        if not session:
            session = SessionModel(weekday=weekday, start_time=start, end_time=end)
            db.add(session)
            db.commit()
        return session

    def get_or_create_course(code):
        if not code or code == "N/A":
            return None
        course = db.query(Course).filter_by(course_code=code).first()
        if not course:
            course = Course(course_code=code, display_name="")
            db.add(course)
            db.commit()
        return course

    seen_enrollments = set()
    seen_trials = set()
    seen_makeups = set()

    for row in data:
        if not all(k in row for k in ("Day", "Times", "Student ID", "Student Name")):
            continue

        start, end = parse_time_range(row["Times"])
        if not start or not end:
            continue

        student = get_or_create_student(row["Student ID"], row["Student Name"])
        session = get_or_create_session(row["Day"], start, end)
        course = get_or_create_course(row.get("Stream"))
        course_code = course.course_code if course else None

        trial_date = row.get("Trial Date", "").strip()
        makeup_date = row.get("Make Up Date", "").strip()

        if trial_date:
            trial_date_obj = datetime.strptime(trial_date, "%b %d, %Y").date()
            seen_trials.add((session.id, student.first_name, student.last_name, course_code, trial_date_obj))
            exists = db.query(Trial).filter_by(session_id=session.id, first_name=student.first_name, last_name=student.last_name, date=trial_date_obj).first()
            if not exists:
                db.add(Trial(session_id=session.id, first_name=student.first_name, last_name=student.last_name, course=course_code, date=trial_date_obj))
        elif makeup_date:
            makeup_date_obj = datetime.strptime(makeup_date, "%b %d, %Y").date()
            seen_makeups.add((student.student_id, session.id, makeup_date_obj))
            exists = db.query(MakeupSession).filter_by(student_id=student.student_id, session_id=session.id, date=makeup_date_obj).first()
            if not exists:
                db.add(MakeupSession(session_id=session.id, student_id=student.student_id, course=course_code, date=makeup_date_obj))
        else:
            seen_enrollments.add((student.student_id, session.id))
            enrollment = db.query(Enrollment).filter_by(student_id=student.student_id, session_id=session.id).first()
            if enrollment:
                if enrollment.course != course_code:
                    enrollment.course = course_code
            else:
                db.add(Enrollment(student_id=student.student_id, session_id=session.id, course=course_code))

    # Delete removed enrollments
    for enr in db.query(Enrollment).all():
        if (enr.student_id, enr.session_id) not in seen_enrollments:
            db.delete(enr)

    # Delete removed trial entries
    for trial in db.query(Trial).all():
        if (trial.session_id, trial.first_name, trial.last_name, trial.course, trial.date) not in seen_trials:
            db.delete(trial)

    # Delete removed makeup sessions
    for mu in db.query(MakeupSession).all():
        if (mu.student_id, mu.session_id, mu.date) not in seen_makeups:
            db.delete(mu)

    db.commit()
    db.close()


insert_data_from_json()