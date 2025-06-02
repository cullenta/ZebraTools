from db import Base, engine
from models import Student, Note, ScratchLogin, Session, Enrollment, Attendance

# Drop all existing tables
# Base.metadata.drop_all(bind=engine)

Base.metadata.create_all(bind=engine)

