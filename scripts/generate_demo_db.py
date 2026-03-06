"""
Generate the TPC-H demo DuckDB database.

Run once from the project root:
    uv run python scripts/generate_demo_db.py

Scale factor (sf):
  0.01 → ~1 MB,  fast
  0.1  → ~12 MB, default (good for demos)
  1.0  → ~120 MB, more realistic

The generated file is saved to data/tpch.duckdb.
"""
import pathlib
import duckdb

SF = 0.1
DB_PATH = pathlib.Path("data/tpch.duckdb")

DB_PATH.parent.mkdir(parents=True, exist_ok=True)

print(f"Generating TPC-H database at {DB_PATH} (scale factor={SF})...")
conn = duckdb.connect(str(DB_PATH))
conn.execute("INSTALL tpch;")
conn.execute("LOAD tpch;")
conn.execute(f"CALL dbgen(sf={SF});")

tables = conn.execute("SHOW TABLES;").fetchall()
print("Tables created:")
for (tbl,) in tables:
    count = conn.execute(f"SELECT count(*) FROM {tbl}").fetchone()[0]
    print(f"  {tbl}: {count:,} rows")

conn.close()
print(f"\nDone. Database saved to {DB_PATH.resolve()}")
