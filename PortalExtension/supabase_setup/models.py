from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Time, Enum, Date, Boolean
from sqlalchemy.orm import relationship
from supabase_setup.db import Base
import enum

class Weekday(enum.Enum):
    monday = "Monday"
    tuesday = "Tuesday"
    wednesday = "Wednesday"
    thursday = "Thursday"
    friday = "Friday"
    saturday = "Saturday"
    sunday = "Sunday"


class Student(Base):
    __tablename__ = "students"

    id = Column(Integer, primary_key=True, index=True)
    student_id = Column(Integer, unique=True)
    first_name = Column(String)
    last_name = Column(String)
    lms_password = Column(String)

    notes = relationship("Note", back_populates="student")
    scratchlogin = relationship("ScratchLogin", back_populates="student")
    courses = relationship("Enrollment", back_populates="student")
    makeup_sessions = relationship("MakeupSession", back_populates="student")


class Note(Base):
    __tablename__ = "notes"

    id = Column(Integer, primary_key=True, index=True)
    note = Column(String)
    date = Column(DateTime)
    creator = Column(String)
    student_id = Column(Integer, ForeignKey("students.student_id"))



    student = relationship("Student", back_populates="notes")
    

class ScratchLogin(Base):
    __tablename__ = "scratchlogins"

    id=Column(Integer, primary_key=True, index=True)
    login=Column(String)
    password = Column(String, default="zebra123")
    student_id = Column(Integer, ForeignKey("students.student_id"), unique=True)


    student = relationship("Student", back_populates="scratchlogin")


class Session(Base):
    __tablename__ = "sessions"

    id = Column(Integer, primary_key=True, index=True)

    weekday = Column(Enum(Weekday), nullable=False)
    start_time = Column(Time, nullable=False)
    end_time = Column(Time, nullable=False)

    students = relationship("Enrollment", back_populates="session")
    makeup_students = relationship("MakeupSession", back_populates="session")
    


class Enrollment(Base):
    __tablename__ = "enrollments"

    id = Column(Integer, primary_key=True, index=True)

    session_id = Column(Integer, ForeignKey("sessions.id"))
    student_id = Column(Integer, ForeignKey("students.student_id"))
    course = Column(String, ForeignKey("courses.course_code"))

    student = relationship("Student", back_populates="courses")
    session = relationship("Session", back_populates="students")
    course_obj = relationship("Course")

    
class Attendance(Base):
    __tablename__ = "attendance"
    
    id = Column(Integer, primary_key=True, index=True)

    student_id = Column(Integer, ForeignKey("students.student_id"))
    session_id = Column(Integer, ForeignKey("sessions.id"))
    date = Column(Date)
    attend = Column(Boolean)

    note_id = Column(Integer, ForeignKey("notes.id"))
    note = relationship("Note")


class Trial(Base):
    
    __tablename__ = "trial_classes"

    id = Column(Integer, primary_key=True, index=True)
    session_id = Column(Integer, ForeignKey("sessions.id"))
    first_name = Column(String)
    last_name = Column(String)
    date = Column(Date, nullable=False)


    course = Column(String, ForeignKey("courses.course_code"))

    course_obj = relationship("Course")
    session_obj = relationship("Session")


class MakeupSession(Base):
    __tablename__ = "makeup_sessions"

    id = Column(Integer, primary_key=True, index=True)

    student_id = Column(Integer, ForeignKey("students.student_id"))
    session_id = Column(Integer, ForeignKey("sessions.id"))
    course = Column(String, ForeignKey("courses.course_code"))  # Optional override if it differs from the session's default
    date = Column(Date, nullable=False)

    student = relationship("Student", back_populates="makeup_sessions")
    session = relationship("Session", back_populates="makeup_students")


class Course(Base):
    __tablename__ = "courses"

    course_code=Column(String, primary_key=True, unique=True)
    display_name=Column(String)

