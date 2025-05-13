import psycopg2
from dotenv import load_dotenv
import os
import supabase

# Load env vars
load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")
print(f"Connecting to: {DATABASE_URL}")  # Debug print

try:
    connection = psycopg2.connect(DATABASE_URL)
    print("Connection successful!")

    cursor = connection.cursor()
    cursor.execute("SELECT * FROM attendance;")
    result = cursor.fetchone()
    print("Current Time:", result)

    

    cursor.close()
    connection.close()

except Exception as e:
    print(f"Failed to connect: {e}")