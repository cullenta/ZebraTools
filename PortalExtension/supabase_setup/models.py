from sqlalchemy import Column, Integer, String, DateTime, ForeignKey, Time, Enum, Date, Boolean
from sqlalchemy.orm import relationship
from db import Base
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
    scratchlogin = relationship("ScratchLogin", back_populates="student", uselist=False)
    courses = relationship("Enrollment", back_populates="student")


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

    login = Column(String)
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
    


class Enrollment(Base):
    __tablename__ = "enrollments"

    id = Column(Integer, primary_key=True, index=True)

    session_id = Column(Integer, ForeignKey("sessions.id"))
    student_id = Column(Integer, ForeignKey("students.student_id"))
    course = Column(String)

    student = relationship("Student", back_populates="courses")
    session = relationship("Session", back_populates="students")



class Attendance(Base):
    __tablename__ = "attendance"
    
    id = Column(Integer, primary_key=True, index=True)

    student_id = Column(Integer, ForeignKey("students.student_id"))
    session_id = Column(Integer, ForeignKey("sessions.id"))
    date = Column(Date)
    attend = Column(Boolean)

    note_id = Column(Integer, ForeignKey("notes.id"))
    note = relationship("Note")

