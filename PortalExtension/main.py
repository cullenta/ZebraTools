# main.py
from fastapi import FastAPI, Depends
from sqlalchemy.orm import Session
from supabase_setup.db import SessionLocal
from supabase_setup.models import Student
from supabase_setup.schemas import AttendanceOut
from fastapi.middleware.cors import CORSMiddleware


# data_refresh()

app = FastAPI()

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


# all the students enrolled in regular classes or registered for trial classes

@app.get("/attendance/", response_model=list[AttendanceOut])
def read_attendance(skip: int = 0, limit: int = 100, db: Session = Depends(get_db)):
    return db.query(Student).offset(skip).limit(limit).all()

# attendance of particular student by ID

# notes for particular student by ID

# list of students who attend during given class time (as student objects with ID or TRIAL, notes, passwords, and their attendance record, whether it's a makeup)


# scratch login for a student by ID

