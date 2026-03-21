"""Lightweight SQLite migrations for columns added after initial deploy."""

from sqlalchemy import inspect, text

from models import db


def ensure_schema() -> None:
    engine = db.engine

    def add_column(table: str, column_sql: str) -> None:
        insp = inspect(engine)
        cols = {c["name"] for c in insp.get_columns(table)}
        col_name = column_sql.split()[0]
        if col_name in cols:
            return
        with engine.connect() as conn:
            conn.execute(text(f"ALTER TABLE {table} ADD COLUMN {column_sql}"))
            conn.commit()

    add_column("users", "google_sub VARCHAR(255)")
    add_column("inspections", "image_paths_json TEXT")
    add_column("inspections", "video_path TEXT")
