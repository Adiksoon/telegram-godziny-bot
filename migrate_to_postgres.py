import os
import sqlite3
import psycopg2
from pathlib import Path
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "work_hours.sqlite3"

load_dotenv(ROOT / ".env")

DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

if not DATABASE_URL:
    print("Blad: Brak DATABASE_URL w pliku .env!")
    exit(1)

if not DB_PATH.exists():
    print(f"Blad: Nie znaleziono lokalnej bazy danych SQLite pod adresem {DB_PATH}")
    exit(1)

print("Inicjalizacja nowej bazy PostgreSQL...")
import bot
bot.IS_POSTGRES = True
bot.DATABASE_URL = DATABASE_URL
bot.init_db()

sqlite_conn = sqlite3.connect(DB_PATH)
sqlite_conn.row_factory = sqlite3.Row

pg_conn = psycopg2.connect(DATABASE_URL)
pg_cur = pg_conn.cursor()

try:
    # 1. Migrate user_settings
    print("Migracja tabeli user_settings...")
    sqlite_settings = sqlite_conn.execute("SELECT * FROM user_settings").fetchall()
    pg_cur.execute("TRUNCATE TABLE user_settings CASCADE")
    for row in sqlite_settings:
        pg_cur.execute(
            "INSERT INTO user_settings (user_id, hourly_rate, updated_at) VALUES (%s, %s, %s)",
            (row["user_id"], row["hourly_rate"], row["updated_at"])
        )
    print(f"Skopiowano {len(sqlite_settings)} wierszy.")

    # 2. Migrate shifts
    print("Migracja tabeli shifts...")
    sqlite_shifts = sqlite_conn.execute("SELECT * FROM shifts ORDER BY id").fetchall()
    pg_cur.execute("TRUNCATE TABLE shifts CASCADE")
    for row in sqlite_shifts:
        pg_cur.execute(
            "INSERT INTO shifts (id, user_id, start_at, end_at, hourly_rate, auto_closed, paid_out_at, created_at) "
            "VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
            (row["id"], row["user_id"], row["start_at"], row["end_at"], row["hourly_rate"], row["auto_closed"], row["paid_out_at"], row["created_at"])
        )
    print(f"Skopiowano {len(sqlite_shifts)} wierszy.")

    # 3. Migrate breaks
    print("Migracja tabeli breaks...")
    sqlite_breaks = sqlite_conn.execute("SELECT * FROM breaks ORDER BY id").fetchall()
    pg_cur.execute("TRUNCATE TABLE breaks CASCADE")
    for row in sqlite_breaks:
        pg_cur.execute(
            "INSERT INTO breaks (id, shift_id, start_at, end_at) VALUES (%s, %s, %s, %s)",
            (row["id"], row["shift_id"], row["start_at"], row["end_at"])
        )
    print(f"Skopiowano {len(sqlite_breaks)} wierszy.")

    # Update serial sequences in PostgreSQL after manual ID insertion
    print("Aktualizacja licznikow sekwencji ID...")
    pg_cur.execute("SELECT setval(pg_get_serial_sequence('shifts', 'id'), COALESCE(MAX(id), 1)) FROM shifts")
    pg_cur.execute("SELECT setval(pg_get_serial_sequence('breaks', 'id'), COALESCE(MAX(id), 1)) FROM breaks")

    pg_conn.commit()
    print("SUKCES: Migracja zakonczona pomyslnie!")

except Exception as e:
    pg_conn.rollback()
    print(f"BLAD podczas migracji: {e}")

finally:
    sqlite_conn.close()
    pg_cur.close()
    pg_conn.close()
