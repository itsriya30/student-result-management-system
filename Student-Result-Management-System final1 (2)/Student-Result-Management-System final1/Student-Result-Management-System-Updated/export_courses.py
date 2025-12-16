import sqlite3
import pandas as pd

def export_courses_to_excel():
    conn = sqlite3.connect("ResultManagementSystem.db")
    query = "SELECT * FROM course"
    df = pd.read_sql_query(query, conn)
    conn.close()

    df.to_excel("Exported_Courses.xlsx", index=False)  # Completely replaces old file
    print("Courses exported successfully, old data removed.")

if __name__ == "__main__":
    export_courses_to_excel()
