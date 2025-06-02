# supabase_setup/schemas.py
from pydantic import BaseModel
from datetime import date, time

class StudentData(BaseModel):
    course: str
    name: str
    lmsusername: str
    lmspassword: str
    scratchlogin: str | None
    scratchpass:str | None
    laptop: int | None
    
    class Config:
        orm_mode = True

class TrialData(BaseModel):
    course: str
    name: str
    date: date
    time: time
