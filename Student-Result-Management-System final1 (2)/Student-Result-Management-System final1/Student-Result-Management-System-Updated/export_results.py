import sqlite3
import pandas as pd

def export_results_to_excel():
    conn = sqlite3.connect("ResultManagementSystem.db")
    query = "SELECT * FROM result"
    df = pd.read_sql_query(query, conn)
    conn.close()

    df.to_excel("Exported_Results.xlsx", index=False)  # Completely replaces old file
    print("Results exported successfully, old data removed.")

if __name__ == "__main__":
    export_results_to_excel()
