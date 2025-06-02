import json
from datetime import datetime
from models import Base, Student, Session as SessionModel, Enrollment, Weekday, Course, Trial, MakeupSession
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import os


# DB setup
DATABASE_URL = os.getenv("DATABASE_URL")
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
db = SessionLocal()


