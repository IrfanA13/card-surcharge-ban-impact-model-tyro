# run_sql.py
# Runs every query in sql/week1_queries.sql on my database and prints the results.
# Each query in the sql file starts with a comment like "-- Q1" and ends with a semicolon.

from pathlib import Path
import sqlite3
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
con = sqlite3.connect(ROOT / "data" / "tyro_project.db")
sql_text = (ROOT / "sql" / "week1_queries.sql").read_text()
pd.set_option("display.width", 200)

for block in sql_text.split(";"):
    lines = block.strip().splitlines()
    title = next((l for l in lines if l.startswith("-- Q")), None)
    if title is None:
        continue
    # drop the comment lines and keep only the actual sql
    query = "\n".join(l for l in lines if not l.strip().startswith("--"))
    print("\n" + title)
    print(pd.read_sql_query(query, con).to_string(index=False))
