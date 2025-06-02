import pandas as pd
from fuzzywuzzy import process
from sqlalchemy.orm import Session
from db import SessionLocal
from models import Student, ScratchLogin

# Load the TSV file
df = pd.read_csv("PortalExtension/supabase_setup/scratch.csv", header=None, names=["login", "password", "name", "notes", "empty"])
df = df.dropna(subset=["login"])  # Remove rows without a name

print(df)

db: Session = SessionLocal()

# Create a name-to-student mapping
students = db.query(Student).all()
student_name_map = {f"{s.first_name} {s.last_name}".lower(): s for s in students}

for _, row in df.iterrows():
    if not row["name"] or pd.isna(row["name"]):
        scratch_login = ScratchLogin(
                    login=row["login"].strip(),
                    password=(row["password"].strip() if pd.notna(row["password"]) else "zebra123"),
                    student_id=None,
                )
        db.add(scratch_login)

    else:
        name_input = row["name"].strip().lower()
        match_name, score = process.extractOne(name_input, student_name_map.keys())

        print(name_input)

        if score >= 90:  # Only accept strong matches
            student = student_name_map[match_name]

            if student.scratchlogin:  # Skip if login already exists
                continue

            existing_login = db.query(ScratchLogin).filter_by(student_id=student.student_id).first()

            if not existing_login:
                scratch_login = ScratchLogin(
                    login=row["login"].strip(),
                    password=(row["password"].strip() if pd.notna(row["password"]) else "zebra123"),
                    student_id=student.student_id,
                )
                db.add(scratch_login)
            else:
                print(f"Skipped: {student} (student_id={student.student_id}) already has a login.")
        else:
            scratch_login = ScratchLogin(
                    login=row["login"].strip(),
                    password=(row["password"].strip() if pd.notna(row["password"]) else "zebra123"),
                    student_id=None,
                )
            db.add(scratch_login)

db.commit()
db.close()
