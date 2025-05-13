# supabase_setup/schemas.py
from pydantic import BaseModel

class AttendanceOut(BaseModel):
    day: str
    times: str
    stream: str
    course: str
    student_id: int
    student_name: str
    instructor_name: str
    makeup_date: str | None
    trial_date: str | None

    class Config:
        orm_mode = True
