# main.py
from fastapi import FastAPI, Depends, BackgroundTasks
from sqlalchemy.orm import Session, joinedload
from supabase_setup.db import SessionLocal
from supabase_setup.models import *
from supabase_setup.schemas import StudentData, TrialData
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




# @app.post("/force-scrape/")
# def force_scrape():
#     result = subprocess.run(["python", "scrape_regclass.py"], capture_output=True, text=True)
    
#     insert_data_from_json()

#     if result.returncode != 0:
#         return {"status": "error", "details": result.stderr}
#     return {"status": "success", "output": result.stdout}



# @app.get("/scrape-status/")
# def get_scrape_status():
#     return scrape_status.get_status()


# all enrolled in regular classes

@app.get("/students/", response_model=list[StudentData])
def read_students(skip: int = 0, limit: int = 100, db: Session = Depends(get_db)):
    enrollments = (
        db.query(Enrollment)
        .options(
            joinedload(Enrollment.student).joinedload(Student.scratchlogin)
        )
        .offset(skip)
        .limit(limit)
        .all()
    )

    return [
        StudentData(
            course=str(enrollment.course_obj.display_name),
            name=f"{enrollment.student.first_name} {enrollment.student.last_name}".strip(),
            lmsusername=str(enrollment.student.student_id),
            lmspassword=str(enrollment.student.lms_password),
            scratchlogin=enrollment.student.scratchlogin[0].login if enrollment.student.scratchlogin else None,
            scratchpass=enrollment.student.scratchlogin[0].password if enrollment.student.scratchlogin else None,
            laptop=None  # You can add this field if you're tracking it elsewhere
        )
        for enrollment in enrollments
    ]

# all trial students

@app.get("/trials/", response_model=list[TrialData])
def read_trials(skip: int = 0, limit: int = 100, db: Session = Depends(get_db)):
    trials = (
        db.query(Trial)
        .offset(skip)
        .limit(limit)
        .all()
    )

    return [
        TrialData(
            course=str(trial.course_obj.display_name),
            name=f"{trial.first_name} {trial.last_name}".strip(),
            date = trial.date,
            time = trial.session_obj.start_time
        )
        for trial in trials
    ]

# attendance of particular student by ID

# notes for particular student by ID

# list of students who attend during given class time (as student objects with ID, notes, and their attendance record, whether it's a makeup)

# scratch login for a student by ID

