from flask import Flask, render_template, request, jsonify, redirect, url_for, session, abort, send_file
import os
import secrets
import json
import re
import io
from datetime import datetime, date, timedelta
from functools import wraps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DATABASE_URL = os.environ.get("DATABASE_URL", "")

if DATABASE_URL:
    import psycopg2
    import psycopg2.extras
else:
    import sqlite3

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=365)
DB = os.path.join(os.path.dirname(__file__), "workout.db")
PASSWORD = os.environ.get("APP_PASSWORD", "")
API_KEY = os.environ.get("API_KEY", "")


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not PASSWORD:
            return f(*args, **kwargs)
        if not session.get("authed"):
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return decorated


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == PASSWORD:
            session.permanent = True
            session["authed"] = True
            return redirect(request.args.get("next") or "/")
        error = "Wrong password."
    return render_template("login.html", error=error)


def get_db():
    if DATABASE_URL:
        conn = psycopg2.connect(DATABASE_URL)
        conn.autocommit = False
        return conn
    else:
        conn = sqlite3.connect(DB)
        # Plain dicts (like Postgres RealDictCursor) so row.get() works in both modes
        conn.row_factory = lambda cur, row: {col[0]: val for col, val in zip(cur.description, row)}
        return conn


class _PgCursor:
    """Wraps a psycopg2 connection to mimic the sqlite3 context-manager + execute API."""
    def __init__(self, conn):
        self._conn = conn
        self._cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)

    def execute(self, sql, params=()):
        sql = sql.replace("?", "%s")
        self._cur.execute(sql, params)
        return self

    def executescript(self, sql):
        for stmt in sql.split(";"):
            stmt = stmt.strip()
            if stmt:
                self._cur.execute(stmt)

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    @property
    def lastrowid(self):
        self._cur.execute("SELECT lastval()")
        return self._cur.fetchone()["lastval"]

    def __enter__(self):
        return self

    def __exit__(self, exc, *_):
        if exc:
            self._conn.rollback()
        else:
            self._conn.commit()
        self._cur.close()
        self._conn.close()


def week_start_of(d):
    """Sunday-based week start (US convention) for a given date."""
    return d - timedelta(days=(d.weekday() + 1) % 7)


def db():
    """Return a unified context-manager cursor for either SQLite or Postgres."""
    conn = get_db()
    if DATABASE_URL:
        return _PgCursor(conn)
    return conn


_PG_SCHEMA = """
    CREATE TABLE IF NOT EXISTS body_weight (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL,
        weight REAL NOT NULL,
        unit TEXT NOT NULL DEFAULT 'lbs'
    );
    CREATE TABLE IF NOT EXISTS templates (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL,
        description TEXT
    );
    CREATE TABLE IF NOT EXISTS template_exercises (
        id SERIAL PRIMARY KEY,
        template_id INTEGER NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
        exercise_name TEXT NOT NULL,
        sets INTEGER,
        reps INTEGER,
        sort_order INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS workouts (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL,
        name TEXT,
        notes TEXT,
        template_id INTEGER REFERENCES templates(id),
        duration_min INTEGER,
        rating INTEGER,
        calories INTEGER,
        avg_heart_rate INTEGER
    );
    CREATE TABLE IF NOT EXISTS sets (
        id SERIAL PRIMARY KEY,
        workout_id INTEGER NOT NULL REFERENCES workouts(id) ON DELETE CASCADE,
        exercise_name TEXT NOT NULL,
        set_number INTEGER NOT NULL,
        reps INTEGER,
        weight REAL,
        unit TEXT DEFAULT 'lbs',
        rpe REAL,
        completed INTEGER DEFAULT 1,
        superset_group INTEGER,
        is_warmup INTEGER DEFAULT 0,
        exercise_note TEXT
    );
    CREATE TABLE IF NOT EXISTS cardio_sessions (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL,
        activity TEXT NOT NULL,
        duration_min REAL,
        distance REAL,
        distance_unit TEXT DEFAULT 'mi',
        calories INTEGER,
        notes TEXT,
        elevation_gain REAL,
        avg_heart_rate INTEGER,
        pool_length REAL,
        pool_length_unit TEXT,
        laps INTEGER,
        elapsed_time_min REAL
    );
    CREATE TABLE IF NOT EXISTS cardio_splits (
        id SERIAL PRIMARY KEY,
        session_id INTEGER NOT NULL REFERENCES cardio_sessions(id) ON DELETE CASCADE,
        mile_number INTEGER NOT NULL,
        pace_seconds INTEGER,
        peak_hr_zone TEXT,
        stroke_type TEXT,
        distance_per_stroke REAL,
        split_distance REAL DEFAULT 1.0
    );
    CREATE TABLE IF NOT EXISTS exercise_bank (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL UNIQUE,
        muscle_group TEXT
    );
    CREATE TABLE IF NOT EXISTS cardio_activities (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL UNIQUE,
        category TEXT
    );
    CREATE TABLE IF NOT EXISTS sleep_logs (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL,
        bedtime TEXT,
        wake_time TEXT,
        duration_min INTEGER,
        quality INTEGER,
        notes TEXT
    );
    CREATE TABLE IF NOT EXISTS weekly_notes (
        id SERIAL PRIMARY KEY,
        week_start TEXT NOT NULL UNIQUE,
        title TEXT,
        body TEXT,
        is_rest_week INTEGER DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS rest_days (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL UNIQUE,
        reason TEXT
    );
    CREATE TABLE IF NOT EXISTS sweet_days (
        id SERIAL PRIMARY KEY,
        date TEXT NOT NULL UNIQUE
    )
"""

_SQLITE_SCHEMA = """
    CREATE TABLE IF NOT EXISTS body_weight (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        weight REAL NOT NULL,
        unit TEXT NOT NULL DEFAULT 'lbs'
    );
    CREATE TABLE IF NOT EXISTS templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        description TEXT
    );
    CREATE TABLE IF NOT EXISTS template_exercises (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER NOT NULL REFERENCES templates(id) ON DELETE CASCADE,
        exercise_name TEXT NOT NULL,
        sets INTEGER,
        reps INTEGER,
        sort_order INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS workouts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        name TEXT,
        notes TEXT,
        template_id INTEGER REFERENCES templates(id),
        duration_min INTEGER,
        rating INTEGER,
        calories INTEGER,
        avg_heart_rate INTEGER
    );
    CREATE TABLE IF NOT EXISTS sets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workout_id INTEGER NOT NULL REFERENCES workouts(id) ON DELETE CASCADE,
        exercise_name TEXT NOT NULL,
        set_number INTEGER NOT NULL,
        reps INTEGER,
        weight REAL,
        unit TEXT DEFAULT 'lbs',
        rpe REAL,
        completed INTEGER DEFAULT 1,
        superset_group INTEGER,
        is_warmup INTEGER DEFAULT 0,
        exercise_note TEXT
    );
    CREATE TABLE IF NOT EXISTS cardio_sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        activity TEXT NOT NULL,
        duration_min REAL,
        distance REAL,
        distance_unit TEXT DEFAULT 'mi',
        calories INTEGER,
        notes TEXT
    );
    CREATE TABLE IF NOT EXISTS cardio_splits (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id INTEGER NOT NULL REFERENCES cardio_sessions(id) ON DELETE CASCADE,
        mile_number INTEGER NOT NULL,
        pace_seconds INTEGER,
        peak_hr_zone TEXT,
        stroke_type TEXT,
        distance_per_stroke REAL,
        split_distance REAL DEFAULT 1.0
    );
    CREATE TABLE IF NOT EXISTS exercise_bank (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        muscle_group TEXT
    );
    CREATE TABLE IF NOT EXISTS cardio_activities (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        category TEXT
    );
    CREATE TABLE IF NOT EXISTS sleep_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        bedtime TEXT,
        wake_time TEXT,
        duration_min INTEGER,
        quality INTEGER,
        notes TEXT
    );
    CREATE TABLE IF NOT EXISTS weekly_notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        week_start TEXT NOT NULL UNIQUE,
        title TEXT,
        body TEXT,
        is_rest_week INTEGER DEFAULT 0,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS rest_days (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL UNIQUE,
        reason TEXT
    );
    CREATE TABLE IF NOT EXISTS sweet_days (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL UNIQUE
    )
"""


MUSCLE_GROUPS = [
    "Chest", "Back", "Front Delts", "Side Delts", "Rear Delts", "Triceps", "Biceps", "Forearms",
    "Quads", "Hamstrings", "Glutes", "Calves", "Abs", "Other",
]

# Starter exercise bank, seeded on first run (grouped by muscle)
_MUSCLE_GROUP_MAP = {
    # Chest
    "BB Bench": "Chest",
    "BB Incline Bench": "Chest",
    "DB Bench Press": "Chest",
    "DB Incline Bench": "Chest",
    "Incline Press Machine": "Chest",
    "Smith Machine Bench": "Chest",
    "Smith Machine Incline Bench": "Chest",
    "Chest Press Machine": "Chest",
    "Pec Dec": "Chest",
    "Push Ups": "Chest",
    "Decline Bench": "Chest",
    "DB Decline Bench": "Chest",
    "Floor Press": "Chest",
    "DB Fly": "Chest",
    "Incline DB Fly": "Chest",
    "Low-to-High Cable Fly": "Chest",
    "High-to-Low Cable Fly": "Chest",
    "Cable Fly": "Chest",
    "Chest Dip": "Chest",
    "Landmine Press": "Chest",
    "Hammer Strength Chest Press": "Chest",
    "Deficit Push Ups": "Chest",
    "Svend Press": "Chest",
    # Back
    "Wide Grip Lat Pulldown": "Back",
    "CG Lat Pulldown": "Back",
    "Front Pulldown": "Back",
    "Single Arm Lat Pulldown": "Back",
    "Lat Pullover": "Back",
    "Pull-Up": "Back",
    "Chin Up": "Back",
    "BB Row": "Back",
    "DB Row": "Back",
    "T-Bar Row": "Back",
    "Hammer Row": "Back",
    "Cable Row": "Back",
    "Seated Row": "Back",
    "Single Arm Cable Row": "Back",
    "DB Chest Supported Row": "Back",
    "Chest supported row machine": "Back",
    "Back Extension": "Back",
    "Shrugs": "Back",
    "Deadlift": "Back",
    "Rack Pull": "Back",
    "Pendlay Row": "Back",
    "Seal Row": "Back",
    "Meadows Row": "Back",
    "Kroc Row": "Back",
    "Inverted Row": "Back",
    "Straight Arm Pulldown": "Back",
    "Neutral Grip Pull-Up": "Back",
    "Weighted Pull-Up": "Back",
    "Assisted Pull-Up": "Back",
    "Iso-Lateral Lat Pulldown": "Back",
    "DB Shrugs": "Back",
    # Front Delts
    "DB Shoulder Press": "Front Delts",
    "Shoulder Press Machine": "Front Delts",
    "Overhead Press": "Front Delts",
    "Front Raises": "Front Delts",
    "Arnold Press": "Front Delts",
    "Push Press": "Front Delts",
    "Z Press": "Front Delts",
    "Landmine Shoulder Press": "Front Delts",
    "Smith Machine Shoulder Press": "Front Delts",
    "Cable Front Raise": "Front Delts",
    "Plate Front Raise": "Front Delts",
    # Side Delts
    "DB Side Raises": "Side Delts",
    "Cable Side Raises": "Side Delts",
    "Side Raise Machine": "Side Delts",
    "Upright Row": "Side Delts",
    "Lean-Away Cable Side Raise": "Side Delts",
    "Seated DB Side Raise": "Side Delts",
    "Behind-the-Back Cable Side Raise": "Side Delts",
    "Cable Y-Raise": "Side Delts",
    "Lu Raise": "Side Delts",
    # Rear Delts
    "Rear Delt Fly": "Rear Delts",
    "Cable Face Pull": "Rear Delts",
    "Single Arm Cable Rear Delt": "Rear Delts",
    "Reverse Pec Dec": "Rear Delts",
    "Bent Over DB Rear Delt Fly": "Rear Delts",
    "Chest Supported Rear Delt Raise": "Rear Delts",
    "Cable Reverse Fly": "Rear Delts",
    "Band Pull-Apart": "Rear Delts",
    # Triceps
    "Tricep Pushdown": "Triceps",
    "Rope tricep pushdown": "Triceps",
    "EZ Bar Skull Crushers": "Triceps",
    "DB Overhead Tricep Extension": "Triceps",
    "Tricep Extension Machine": "Triceps",
    "Dip": "Triceps",
    "Seated Dip Machine": "Triceps",
    "Close Grip Bench": "Triceps",
    "JM Press": "Triceps",
    "BB Skull Crushers": "Triceps",
    "Tate Press": "Triceps",
    "Overhead Cable Tricep Extension": "Triceps",
    "Single Arm Cable Pushdown": "Triceps",
    "DB Kickback": "Triceps",
    "Cable Kickback": "Triceps",
    "Diamond Push Ups": "Triceps",
    "Bench Dips": "Triceps",
    # Biceps
    "DB Curl": "Biceps",
    "Single Arm DB Curl": "Biceps",
    "Hammer Curl": "Biceps",
    "DB Hammer Curl": "Biceps",
    "EZ Bar Curl": "Biceps",
    "Straight Bar Curl": "Biceps",
    "Cable Curl": "Biceps",
    "Incline Curl": "Biceps",
    "Preacher Curl": "Biceps",
    "DB Preacher Curl": "Biceps",
    "Zottman Curl": "Biceps",
    "Spider Curl": "Biceps",
    "Bayesian Cable Curl": "Biceps",
    "Concentration Curl": "Biceps",
    "Drag Curl": "Biceps",
    "Cross-Body Hammer Curl": "Biceps",
    "Rope Hammer Curl": "Biceps",
    "Machine Curl": "Biceps",
    "Cable Preacher Curl": "Biceps",
    "21s": "Biceps",
    # Forearms
    "Reverse Curl": "Forearms",
    "BTB Curls": "Forearms",
    "Wrist Curl": "Forearms",
    "Reverse Wrist Curl": "Forearms",
    "Wrist Roller": "Forearms",
    "Plate Pinch": "Forearms",
    "Farmer's Carry": "Forearms",
    "Dead Hang": "Forearms",
    # Quads
    "Back Squat": "Quads",
    "Front Squat": "Quads",
    "KB Goblet Squat": "Quads",
    "Split Squat": "Quads",
    "Leg Press": "Quads",
    "Leg Extension": "Quads",
    "SL Leg Extension": "Quads",
    "Hack Squat": "Quads",
    "Pendulum Squat": "Quads",
    "Belt Squat": "Quads",
    "Smith Machine Squat": "Quads",
    "Bulgarian Split Squat": "Quads",
    "Walking Lunge": "Quads",
    "Reverse Lunge": "Quads",
    "Step-Ups": "Quads",
    "Sissy Squat": "Quads",
    # Hamstrings
    "BB RDL": "Hamstrings",
    "Hamstring Curl": "Hamstrings",
    "SL Hamstring Curl": "Hamstrings",
    "Sumo Deadlift": "Hamstrings",
    "DB RDL": "Hamstrings",
    "SL RDL": "Hamstrings",
    "Good Morning": "Hamstrings",
    "Nordic Curl": "Hamstrings",
    "Glute Ham Raise": "Hamstrings",
    "Seated Hamstring Curl": "Hamstrings",
    # Glutes
    "Hip Thrust": "Glutes",
    "Glute Bridge": "Glutes",
    "Hip Abduction Machine": "Glutes",
    "Hip Adduction Machine": "Glutes",
    "Cable Glute Kickback": "Glutes",
    # Calves
    "Calf Raise": "Calves",
    "Seated Calf Raise": "Calves",
    "Standing Calf Raise": "Calves",
    "Leg Press Calf Raise": "Calves",
    "Tibialis Raise": "Calves",
    # Abs
    "Leg raises(abs)": "Abs",
    "Hanging Leg Raise": "Abs",
    "Hanging Knee Raise": "Abs",
    "Cable Crunch": "Abs",
    "Decline Sit-Up": "Abs",
    "Ab Wheel Rollout": "Abs",
    "Plank": "Abs",
    "Side Plank": "Abs",
    "Dead Bug": "Abs",
    "Russian Twist": "Abs",
    "Pallof Press": "Abs",
    "Cable Woodchopper": "Abs",
    # Other
    "Power Clean": "Other",
    "KB Swing": "Other",
    "Thrusters": "Other",
    "Sled Push": "Other",
    "Turkish Get-Up": "Other",
    "Burpees": "Other",
}

CARDIO_CATEGORIES = ["Run & Walk", "Machines", "Sports", "Other"]

# Starter cardio activities, seeded on first run. The first 14 have their own
# fields on the cardio log (splits, pool length, ski runs...); the rest use the standard fields.
_CARDIO_ACTIVITY_MAP = {
    # Run & Walk
    "Run": "Run & Walk", "Walk": "Run & Walk", "Incline Walk": "Run & Walk",
    "Hike": "Run & Walk", "Trail Run": "Run & Walk", "Sprints": "Run & Walk",
    # Machines
    "Bike": "Machines", "Row": "Machines", "Swim": "Machines", "Elliptical": "Machines",
    "Stairmaster": "Machines", "Jump Rope": "Machines",
    "Assault Bike": "Machines", "Spin Class": "Machines", "Ski Erg": "Machines",
    "Arc Trainer": "Machines", "VersaClimber": "Machines",
    # Sports
    "Basketball": "Sports", "Tennis": "Sports", "Pickleball": "Sports", "Skiing": "Sports",
    "Golf (Walking)": "Sports",
    "Soccer": "Sports", "Volleyball": "Sports", "Hockey": "Sports", "Racquetball": "Sports",
    "Squash": "Sports", "Snowboarding": "Sports",
    # Other
    "Boxing": "Other", "Kickboxing": "Other", "HIIT Class": "Other", "Circuit Training": "Other",
    "Rock Climbing": "Other", "Kayaking": "Other", "Dance": "Other",
}


def add_cardio_activity(conn, name, category="Other"):
    """Remember a cardio activity (no-op if it already exists)."""
    name = (name or "").strip()
    if not name or name == "Other":
        return
    if category not in CARDIO_CATEGORIES:
        category = "Other"
    conn.execute("INSERT INTO cardio_activities (name, category) VALUES (?, ?) ON CONFLICT (name) DO NOTHING", (name, category))


def get_cardio_activities():
    """{category: [names]} in CARDIO_CATEGORIES order."""
    with db() as conn:
        rows = conn.execute("SELECT name, category FROM cardio_activities ORDER BY name").fetchall()
    groups = {c: [] for c in CARDIO_CATEGORIES}
    for r in rows:
        groups.setdefault(r["category"] if r["category"] in groups else "Other", []).append(r["name"])
    return {c: names for c, names in groups.items() if names}


def init_db():
    with db() as conn:
        conn.executescript(_PG_SCHEMA if DATABASE_URL else _SQLITE_SCHEMA)
    # Seed cardio activities (plus any already logged) if empty
    with db() as conn:
        if conn.execute("SELECT COUNT(*) as c FROM cardio_activities").fetchone()["c"] == 0:
            for name, category in _CARDIO_ACTIVITY_MAP.items():
                add_cardio_activity(conn, name, category)
            for r in conn.execute("SELECT DISTINCT activity FROM cardio_sessions").fetchall():
                add_cardio_activity(conn, r["activity"])
    # Migrate exercise_bank from sets if empty
    with db() as conn:
        count = conn.execute("SELECT COUNT(*) as c FROM exercise_bank").fetchone()["c"]
        if count == 0:
            exercises = conn.execute(
                "SELECT DISTINCT exercise_name FROM sets ORDER BY exercise_name"
            ).fetchall()
            names = [ex["exercise_name"] for ex in exercises] or list(_MUSCLE_GROUP_MAP)
            for name in names:
                muscle = _MUSCLE_GROUP_MAP.get(name, "Other")
                try:
                    conn.execute(
                        "INSERT INTO exercise_bank (name, muscle_group) VALUES (?, ?) ON CONFLICT (name) DO NOTHING"
                        if not DATABASE_URL else
                        "INSERT INTO exercise_bank (name, muscle_group) VALUES (%s, %s) ON CONFLICT (name) DO NOTHING",
                        (name, muscle)
                    )
                except Exception:
                    pass
    # user_settings table
    with db() as conn:
        if DATABASE_URL:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS user_settings (
                    key TEXT PRIMARY KEY, value TEXT
                )
            """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS user_settings (
                    key TEXT PRIMARY KEY, value TEXT
                )
            """)

    # sleep_logs table (migration for existing DBs)
    with db() as conn:
        if DATABASE_URL:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS sleep_logs (
                    id SERIAL PRIMARY KEY,
                    date TEXT NOT NULL,
                    bedtime TEXT,
                    wake_time TEXT,
                    duration_min INTEGER,
                    quality INTEGER,
                    notes TEXT
                )
            """)
        else:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS sleep_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    bedtime TEXT,
                    wake_time TEXT,
                    duration_min INTEGER,
                    quality INTEGER,
                    notes TEXT
                )
            """)

    # Add new columns if they don't exist yet
    if DATABASE_URL:
        import psycopg2
        pg = psycopg2.connect(DATABASE_URL)
        cur = pg.cursor()
        for col, typ in [("incline", "REAL"), ("speed", "REAL")]:
            cur.execute(f"""
                ALTER TABLE cardio_sessions ADD COLUMN IF NOT EXISTS {col} {typ}
            """)
        cur.execute("ALTER TABLE workouts ADD COLUMN IF NOT EXISTS duration_min INTEGER")
        cur.execute("ALTER TABLE workouts ADD COLUMN IF NOT EXISTS rating INTEGER")
        cur.execute("ALTER TABLE workouts ADD COLUMN IF NOT EXISTS calories INTEGER")
        cur.execute("ALTER TABLE workouts ADD COLUMN IF NOT EXISTS avg_heart_rate INTEGER")
        cur.execute("ALTER TABLE sets ADD COLUMN IF NOT EXISTS superset_group INTEGER")
        cur.execute("ALTER TABLE sets ADD COLUMN IF NOT EXISTS is_warmup INTEGER DEFAULT 0")
        cur.execute("ALTER TABLE sets ADD COLUMN IF NOT EXISTS exercise_note TEXT")
        for col, typ in [("elevation_gain", "REAL"), ("avg_heart_rate", "INTEGER"), ("pool_length", "REAL"), ("pool_length_unit", "TEXT"), ("laps", "INTEGER"), ("elapsed_time_min", "REAL")]:
            cur.execute(f"ALTER TABLE cardio_sessions ADD COLUMN IF NOT EXISTS {col} {typ}")
        for col, typ in [("stroke_type", "TEXT"), ("distance_per_stroke", "REAL")]:
            cur.execute(f"ALTER TABLE cardio_splits ADD COLUMN IF NOT EXISTS {col} {typ}")
        cur.execute("""
            CREATE TABLE IF NOT EXISTS cardio_splits (
                id SERIAL PRIMARY KEY,
                session_id INTEGER NOT NULL REFERENCES cardio_sessions(id) ON DELETE CASCADE,
                mile_number INTEGER NOT NULL,
                pace_seconds INTEGER,
                peak_hr_zone TEXT
            )
        """)
        cur.execute("ALTER TABLE cardio_splits ADD COLUMN IF NOT EXISTS split_distance REAL DEFAULT 1.0")
        cur.execute("ALTER TABLE weekly_notes ADD COLUMN IF NOT EXISTS is_rest_week INTEGER DEFAULT 0")
        cur.execute("CREATE TABLE IF NOT EXISTS rest_days (id SERIAL PRIMARY KEY, date TEXT NOT NULL UNIQUE, reason TEXT)")
        cur.execute("ALTER TABLE rest_days ADD COLUMN IF NOT EXISTS reason TEXT")
        cur.execute("CREATE TABLE IF NOT EXISTS sweet_days (id SERIAL PRIMARY KEY, date TEXT NOT NULL UNIQUE)")
        # Migrate last_sweet_date setting → sweet_days table
        cur.execute("INSERT INTO sweet_days (date) SELECT value FROM user_settings WHERE key='last_sweet_date' AND value IS NOT NULL AND value != '' ON CONFLICT (date) DO NOTHING")
        pg.commit(); cur.close(); pg.close()
    else:
        sq = sqlite3.connect(DB)
        for col in ["incline", "speed"]:
            try:
                sq.execute(f"ALTER TABLE cardio_sessions ADD COLUMN {col} REAL")
                sq.commit()
            except Exception:
                pass
        for col in ["duration_min INTEGER", "rating INTEGER", "calories INTEGER", "avg_heart_rate INTEGER"]:
            try:
                sq.execute(f"ALTER TABLE workouts ADD COLUMN {col}")
                sq.commit()
            except Exception:
                pass
        for col, default in [("superset_group", None), ("is_warmup", 0), ("exercise_note", None)]:
            try:
                dflt = f" DEFAULT {default}" if default is not None else ""
                sq.execute(f"ALTER TABLE sets ADD COLUMN {col} INTEGER{dflt}")
                sq.commit()
            except Exception:
                pass
        for col in ["elevation_gain REAL", "avg_heart_rate INTEGER", "pool_length REAL", "pool_length_unit TEXT", "laps INTEGER", "elapsed_time_min REAL"]:
            try:
                sq.execute(f"ALTER TABLE cardio_sessions ADD COLUMN {col}")
                sq.commit()
            except Exception:
                pass
        for col in ["stroke_type TEXT", "distance_per_stroke REAL", "split_distance REAL DEFAULT 1.0"]:
            try:
                sq.execute(f"ALTER TABLE cardio_splits ADD COLUMN {col}")
                sq.commit()
            except Exception:
                pass
        try:
            sq.execute("""
                CREATE TABLE IF NOT EXISTS cardio_splits (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    session_id INTEGER NOT NULL REFERENCES cardio_sessions(id) ON DELETE CASCADE,
                    mile_number INTEGER NOT NULL,
                    pace_seconds INTEGER,
                    peak_hr_zone TEXT
                )
            """)
            sq.commit()
        except Exception:
            pass
        try:
            sq.execute("ALTER TABLE weekly_notes ADD COLUMN is_rest_week INTEGER DEFAULT 0")
            sq.commit()
        except Exception:
            pass
        try:
            sq.execute("CREATE TABLE IF NOT EXISTS rest_days (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL UNIQUE, reason TEXT)")
            sq.commit()
        except Exception:
            pass
        try:
            sq.execute("ALTER TABLE rest_days ADD COLUMN reason TEXT")
            sq.commit()
        except Exception:
            pass
        try:
            sq.execute("CREATE TABLE IF NOT EXISTS sweet_days (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL UNIQUE)")
            sq.commit()
        except Exception:
            pass
        # Migrate last_sweet_date setting → sweet_days table
        try:
            row = sq.execute("SELECT value FROM user_settings WHERE key='last_sweet_date'").fetchone()
            if row and row[0]:
                sq.execute("INSERT OR IGNORE INTO sweet_days (date) VALUES (?)", (row[0],))
                sq.commit()
        except Exception:
            pass
        sq.close()
    # One-time: sort existing exercises into the finer muscle groups (Quads, Abs...)
    # and add the expanded starter exercises. Only touches exercises filed under Legs/Other.
    with db() as conn:
        done = conn.execute("SELECT value FROM user_settings WHERE key='muscle_groups_v2'").fetchone()
        if not done:
            for name, group in _MUSCLE_GROUP_MAP.items():
                conn.execute(
                    "UPDATE exercise_bank SET muscle_group=? WHERE name=? AND (muscle_group IS NULL OR muscle_group IN ('Legs', 'Other'))",
                    (group, name),
                )
                conn.execute("INSERT INTO exercise_bank (name, muscle_group) VALUES (?, ?) ON CONFLICT (name) DO NOTHING", (name, group))
            conn.execute("INSERT INTO user_settings (key, value) VALUES ('muscle_groups_v2', 'done')")


init_db()


@app.context_processor
def inject_globals():
    today = date.today().strftime("%Y-%m-%d")
    now = datetime.now().strftime("%A, %B %-d, %Y")
    settings = get_settings()
    weekday = date.today().weekday()
    day = get_schedule(settings)[weekday]
    day_info = (WEEKDAYS[weekday], day["label"], day["color"]) if day["label"] else None
    return {"today": today, "now": now, "day_workout": day_info, "settings": settings}


@app.template_filter("format_date")
def format_date(value):
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%b %-d, %Y")
    except Exception:
        return value


@app.template_filter("format_date_short")
def format_date_short(value):
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%-m/%-d/%y")
    except Exception:
        return value




_TYPE_MAP = {
    "push": ("type-push", "Push"),
    "pull": ("type-pull", "Pull"),
    "legs": ("type-legs", "Legs"),
    "chest": ("type-chest", "Chest & Back"),
    "back": ("type-chest", "Chest & Back"),
    "shoulder": ("type-shoulders", "Shoulders & Arms"),
    "arm": ("type-arms", "Shoulders & Arms"),
}

def _type_key(name):
    n = name.lower()
    for key in _TYPE_MAP:
        if key in n:
            return key
    return None

@app.template_filter("workout_type_class")
def workout_type_class(name):
    key = _type_key(name)
    return _TYPE_MAP[key][0] if key else "type-other"

@app.template_filter("workout_type_label")
def workout_type_label(name):
    key = _type_key(name)
    return _TYPE_MAP[key][1] if key else "Other"


# ---------- helpers ----------

def int_in_range(value, low, high):
    """Optional whole number from a form/JSON value; None if blank or out of range."""
    try:
        n = int(float(value))
    except (TypeError, ValueError):
        return None
    return n if low <= n <= high else None


WEEKDAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
_SCHEDULE_COLORS = ["#3b82f6", "#10b981", "#f59e0b", "#8b5cf6", "#ec4899", "#06b6d4", "#ef4444"]

# Everything adjustable on the Settings page. Stored as strings in user_settings.
SETTINGS_DEFAULTS = {
    "schedule": "",              # JSON list of 7 {label, color}, Monday first; blank label = no badge
    "cardio_goal_min": "150",    # weekly cardio goal (minutes)
    "sleep_goal_hours": "8",     # nightly sleep target
    "streak_goal_days": "7",     # dashboard streak goal
    "streak_skip_weekends": "1", # weekends don't break streaks
    "goal_weight": "",
    "height_inches": "",
    "age": "",
    "sex": "",
}


def get_settings():
    with db() as conn:
        rows = conn.execute("SELECT key, value FROM user_settings").fetchall()
    settings = dict(SETTINGS_DEFAULTS)
    settings.update({r["key"]: r["value"] for r in rows if r["value"] is not None})
    return settings


def setting_num(settings, key):
    """Numeric setting, falling back to the default if the stored value is bad."""
    try:
        return float(settings.get(key) or SETTINGS_DEFAULTS[key])
    except ValueError:
        return float(SETTINGS_DEFAULTS[key])


def get_schedule(settings):
    try:
        days = json.loads(settings.get("schedule") or "[]")
    except ValueError:
        days = []
    return [
        {
            "label": (days[i].get("label") or "").strip() if i < len(days) else "",
            "color": (days[i].get("color") if i < len(days) else None) or _SCHEDULE_COLORS[i],
        }
        for i in range(7)
    ]


def get_setting(key, default=None):
    with db() as conn:
        row = conn.execute("SELECT value FROM user_settings WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default

def set_setting(key, value):
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        if DATABASE_URL:
            conn.execute(
                "INSERT INTO user_settings (key, value) VALUES (%s, %s) ON CONFLICT (key) DO UPDATE SET value=EXCLUDED.value",
                (key, value)
            )
        else:
            conn.execute("INSERT OR REPLACE INTO user_settings (key, value) VALUES (?, ?)", (key, value))


def get_prs():
    """Return dict of exercise_name -> {weight, reps, date, set_id}."""
    with db() as conn:
        rows = conn.execute("""
            SELECT s.exercise_name, s.weight, s.reps, s.unit, w.date, s.id
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            WHERE s.weight IS NOT NULL AND s.reps IS NOT NULL AND s.completed = 1
            ORDER BY s.exercise_name, s.weight DESC, s.reps DESC
        """).fetchall()
    prs = {}
    for r in rows:
        name = r["exercise_name"]
        if name not in prs:
            prs[name] = dict(r)
    return prs


def get_pr_timeline():
    """Return dict of exercise_name -> chronological list of PR events
    (new estimated-1RM highs), each {date, weight, reps, unit, e1rm, id}."""
    with db() as conn:
        rows = conn.execute("""
            SELECT s.id, s.exercise_name, s.weight, s.reps, s.unit, w.date
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            WHERE s.weight IS NOT NULL AND s.reps IS NOT NULL AND s.completed = 1
            ORDER BY s.exercise_name, w.date, s.id
        """).fetchall()
    timeline = {}
    running_max = {}
    for r in rows:
        name = r["exercise_name"]
        weight = r["weight"]
        reps = r["reps"]
        e1rm = round(weight * (1 + reps / 30), 1) if reps > 1 else weight
        if name not in running_max or e1rm > running_max[name]:
            running_max[name] = e1rm
            timeline.setdefault(name, []).append({
                "id": r["id"], "date": r["date"], "weight": weight,
                "reps": reps, "unit": r["unit"], "e1rm": e1rm,
            })
    return timeline


def get_estimated_1rm():
    """Return dict of exercise_name -> {estimated_1rm, weight, reps, unit, date} using Epley formula."""
    with db() as conn:
        rows = conn.execute("""
            SELECT s.id, s.exercise_name, s.weight, s.reps, s.unit, w.date
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            WHERE s.weight IS NOT NULL AND s.reps IS NOT NULL
              AND s.reps >= 1 AND s.reps <= 12 AND s.completed = 1
            ORDER BY s.exercise_name
        """).fetchall()
    estimates = {}
    for r in rows:
        name = r["exercise_name"]
        weight = r["weight"]
        reps = r["reps"]
        # Epley formula: weight * (1 + reps/30)
        e1rm = round(weight * (1 + reps / 30), 1) if reps > 1 else weight
        if name not in estimates or e1rm > estimates[name]["estimated_1rm"]:
            estimates[name] = {
                "id": r["id"],
                "estimated_1rm": e1rm,
                "estimated_3rm": round(e1rm / (1 + 3/30), 1),
                "estimated_5rm": round(e1rm / (1 + 5/30), 1),
                "weight": weight,
                "reps": reps,
                "unit": r["unit"],
                "date": r["date"],
            }
    return estimates


def get_prs_by_reps():
    """Return dict of reps -> {exercise_name -> {weight, reps, date, unit}}."""
    with db() as conn:
        rows = conn.execute("""
            SELECT s.exercise_name, s.weight, s.reps, s.unit, w.date
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            WHERE s.weight IS NOT NULL AND s.reps IS NOT NULL AND s.completed = 1
            ORDER BY s.exercise_name, s.reps, s.weight DESC
        """).fetchall()
    # For each exercise+rep combo, keep the heaviest weight
    by_reps = {}
    for r in rows:
        rep = r["reps"]
        name = r["exercise_name"]
        if rep not in by_reps:
            by_reps[rep] = {}
        if name not in by_reps[rep]:
            by_reps[rep][name] = dict(r)
    return by_reps


def get_cardio_prs():
    """Return dict of cardio PRs: longest_run, fastest_pace (Run, per mile), most_calories."""
    with db() as conn:
        rows = conn.execute("""
            SELECT date, activity, duration_min, distance, distance_unit, calories
            FROM cardio_sessions
        """).fetchall()
    prs = {}
    for r in rows:
        # Longest run (by distance, converted to miles)
        if r["activity"] == "Run" and r["distance"]:
            dist_mi = r["distance"]
            if r["distance_unit"] == "km":
                dist_mi = r["distance"] * 0.621371
            elif r["distance_unit"] == "m":
                dist_mi = r["distance"] * 0.000621371
            if "longest_run" not in prs or dist_mi > prs["longest_run"]["distance_mi"]:
                prs["longest_run"] = {"distance_mi": dist_mi, "distance": r["distance"], "unit": r["distance_unit"], "date": r["date"]}

            # Fastest pace (min/mile)
            if r["duration_min"] and dist_mi > 0:
                pace = r["duration_min"] / dist_mi
                if "fastest_pace" not in prs or pace < prs["fastest_pace"]["pace"]:
                    prs["fastest_pace"] = {"pace": pace, "distance": r["distance"], "unit": r["distance_unit"], "date": r["date"]}

        # Most calories burned in a session (any activity)
        if r["calories"] and ("most_calories" not in prs or r["calories"] > prs["most_calories"]["calories"]):
            prs["most_calories"] = {"calories": r["calories"], "activity": r["activity"], "date": r["date"]}

    with db() as conn:
        split_rows = conn.execute("""
            SELECT cs.pace_seconds, cs.mile_number, sess.date
            FROM cardio_splits cs JOIN cardio_sessions sess ON cs.session_id = sess.id
            WHERE cs.pace_seconds IS NOT NULL
        """).fetchall()
    for r in split_rows:
        if "fastest_split" not in prs or r["pace_seconds"] < prs["fastest_split"]["pace_seconds"]:
            prs["fastest_split"] = {"pace_seconds": r["pace_seconds"], "mile": r["mile_number"], "date": r["date"]}

    return prs


# ---------- routes ----------

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    with db() as conn:
        recent_workouts = conn.execute("""
            SELECT w.*, COUNT(DISTINCT s.exercise_name) AS exercise_count
            FROM workouts w
            LEFT JOIN sets s ON s.workout_id = w.id
            GROUP BY w.id
            ORDER BY w.date DESC
            LIMIT 6
        """).fetchall()
        recent_workout_ids = [w["id"] for w in recent_workouts]
        recent_exercises = {}
        if recent_workout_ids:
            placeholders = ",".join("?" if not DATABASE_URL else "%s" for _ in recent_workout_ids)
            rows = conn.execute(
                f"""SELECT workout_id, exercise_name, COUNT(*) as set_count,
                    CAST(ROUND(AVG(reps)) AS INTEGER) as avg_reps
                    FROM sets WHERE workout_id IN ({placeholders})
                    GROUP BY workout_id, exercise_name
                    ORDER BY workout_id, MIN(id)""",
                recent_workout_ids
            ).fetchall()
            for r in rows:
                recent_exercises.setdefault(r["workout_id"], []).append(r)
        latest_weight = conn.execute(
            "SELECT * FROM body_weight ORDER BY date DESC LIMIT 1"
        ).fetchone()
        recent_weights = conn.execute(
            "SELECT weight FROM body_weight ORDER BY date DESC LIMIT 7"
        ).fetchall()
        workout_count = conn.execute("SELECT COUNT(*) as c FROM workouts").fetchone()["c"]
        pr_count = len(get_prs())
        def best_1rm(exercise):
            ph = "%s" if DATABASE_URL else "?"
            row = conn.execute(
                f"SELECT weight as w, unit FROM sets WHERE exercise_name={ph} AND reps=1 ORDER BY weight DESC LIMIT 1",
                (exercise,)
            ).fetchone()
            if row:
                return {"w": row["w"], "unit": row["unit"], "estimated": False}
            # Epley formula: e1rm = weight * (1 + reps/30), best set within 1-12 reps
            row = conn.execute(
                f"SELECT weight, reps, unit FROM sets WHERE exercise_name={ph} AND reps BETWEEN 2 AND 12 AND weight > 0 ORDER BY weight*(1+reps/30.0) DESC LIMIT 1",
                (exercise,)
            ).fetchone()
            if row:
                e1rm = round(row["weight"] * (1 + row["reps"] / 30.0), 1)
                return {"w": e1rm, "unit": row["unit"], "estimated": True}
            return None
        bench_1rm = best_1rm("BB Bench")
        squat_1rm = best_1rm("Back Squat")
        latest_sleep = conn.execute(
            "SELECT * FROM sleep_logs ORDER BY date DESC LIMIT 1"
        ).fetchone()
        sleep_is_recent = bool(
            latest_sleep and latest_sleep["date"] >= (datetime.today().date() - timedelta(days=1)).strftime("%Y-%m-%d")
        )
        _sleep_cutoff = (datetime.today().date() - timedelta(days=6)).strftime("%Y-%m-%d")
        avg_sleep_row = conn.execute(
            "SELECT AVG(duration_min) as avg_duration, AVG(quality) as avg_quality FROM sleep_logs WHERE date >= ?",
            (_sleep_cutoff,)
        ).fetchone()
        avg_sleep = dict(avg_sleep_row) if avg_sleep_row else None
        if avg_sleep and avg_sleep["avg_duration"] is not None:
            avg_sleep["avg_duration"] = round(avg_sleep["avg_duration"])
        # Streaks: consecutive days up to today
        def calc_streak(dates_set, skip_weekends=False):
            def counts(d):
                if skip_weekends and d.weekday() >= 5:  # Sat=5, Sun=6
                    return True
                return d.strftime("%Y-%m-%d") in dates_set
            streak = 0
            d = date.today()
            while counts(d):
                streak += 1
                d -= timedelta(days=1)
            # also count if yesterday was the last day (streak not broken yet today)
            if streak == 0:
                d = date.today() - timedelta(days=1)
                while counts(d):
                    streak += 1
                    d -= timedelta(days=1)
            return streak
        workout_dates = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM workouts").fetchall()}
        cardio_dates = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM cardio_sessions").fetchall()}
        sleep_dates = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM sleep_logs").fetchall()}
        weight_dates = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM body_weight").fetchall()}
        rest_day_rows = conn.execute("SELECT date, reason FROM rest_days").fetchall()
        rest_day_dates = {r["date"] for r in rest_day_rows}
        rest_day_map = {r["date"]: r["reason"] for r in rest_day_rows}
        def calc_streak(dates_set, skip_weekends=False):
            def counts(d):
                ds = d.strftime("%Y-%m-%d")
                if ds in rest_day_dates:
                    return True
                if skip_weekends and d.weekday() >= 5:
                    return True
                return ds in dates_set
            streak = 0
            d = date.today()
            while counts(d):
                streak += 1
                d -= timedelta(days=1)
            if streak == 0:
                d = date.today() - timedelta(days=1)
                while counts(d):
                    streak += 1
                    d -= timedelta(days=1)
            return streak
        skip_weekends = get_settings()["streak_skip_weekends"] == "1"
        workout_streak = calc_streak(workout_dates, skip_weekends=skip_weekends)
        cardio_streak = calc_streak(cardio_dates, skip_weekends=skip_weekends)
        sleep_streak = calc_streak(sleep_dates, skip_weekends=skip_weekends)
        weight_streak = calc_streak(weight_dates, skip_weekends=skip_weekends)
        last_cardio_row = conn.execute("SELECT date FROM cardio_sessions ORDER BY date DESC LIMIT 1").fetchone()
        # This-week day dots (Sun=0 .. Sat=6)
        _today = date.today()
        _wk_start_d = week_start_of(_today)
        _wk_start = _wk_start_d.strftime("%Y-%m-%d")
        _wk_end   = (_wk_start_d + timedelta(days=6)).strftime("%Y-%m-%d")
        _week_workout_dates  = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM workouts WHERE date >= ? AND date <= ?", (_wk_start, _wk_end)).fetchall()}
        _week_cardio_dates   = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM cardio_sessions WHERE date >= ? AND date <= ?", (_wk_start, _wk_end)).fetchall()}
        _week_sleep_dates    = {r["date"] for r in conn.execute("SELECT DISTINCT date FROM sleep_logs WHERE date >= ? AND date <= ?", (_wk_start, _wk_end)).fetchall()}
        _week_rest_dates     = {r["date"] for r in conn.execute("SELECT date FROM rest_days WHERE date >= ? AND date <= ?", (_wk_start, _wk_end)).fetchall()}
    with db() as conn:
        _sweet_rows = conn.execute("SELECT date FROM sweet_days ORDER BY date DESC").fetchall()
    sweet_dates = [r["date"] for r in _sweet_rows]
    last_sweet = sweet_dates[0] if sweet_dates else None
    if last_sweet:
        days_since_sweet = (datetime.today().date() - datetime.strptime(last_sweet, "%Y-%m-%d").date()).days
    else:
        days_since_sweet = None

    # Build week day summary for dashboard
    _today = date.today()
    week_days = []
    for i in range(7):
        d = week_start_of(_today) + timedelta(days=i)
        ds = d.strftime("%Y-%m-%d")
        week_days.append({
            "label": d.strftime("%a"),
            "date": ds,
            "is_today": ds == _today.strftime("%Y-%m-%d"),
            "future": d > _today,
            "lift": ds in _week_workout_dates,
            "cardio": ds in _week_cardio_dates,
            "sleep": ds in _week_sleep_dates,
            "rest": ds in _week_rest_dates,
        })

    last_lift_date = recent_workouts[0]["date"] if recent_workouts else None
    days_since_lift = (datetime.today().date() - datetime.strptime(last_lift_date, "%Y-%m-%d").date()).days if last_lift_date else None
    last_cardio_date = last_cardio_row["date"] if last_cardio_row else None
    days_since_cardio = (datetime.today().date() - datetime.strptime(last_cardio_date, "%Y-%m-%d").date()).days if last_cardio_date else None

    return render_template(
        "dashboard.html",
        recent_workouts=recent_workouts,
        recent_exercises=recent_exercises,
        latest_weight=latest_weight,
        weight_trend=[r["weight"] for r in reversed(recent_weights)] if recent_weights else [],
        workout_count=workout_count,
        pr_count=pr_count,
        bench_1rm=bench_1rm,
        squat_1rm=squat_1rm,
        latest_sleep=latest_sleep,
        sleep_is_recent=sleep_is_recent,
        avg_sleep=avg_sleep,
        workout_streak=workout_streak,
        cardio_streak=cardio_streak,
        sleep_streak=sleep_streak,
        weight_streak=weight_streak,
        days_since_sweet=days_since_sweet,
        last_sweet=last_sweet,
        sweet_dates=sweet_dates,
        days_since_lift=days_since_lift,
        days_since_cardio=days_since_cardio,
        week_days=week_days,
        today_str=date.today().isoformat(),
        rest_day_dates=list(rest_day_dates),
        rest_day_map=rest_day_map,
    )


@app.route("/api/rest-days/delete", methods=["POST"])
@login_required
def api_rest_days_delete():
    data = request.json or {}
    start = data.get("start") or data.get("date")
    end = data.get("end") or start
    if not start:
        return jsonify({"error": "date required"}), 400
    from datetime import date as _d
    cur = _d.fromisoformat(start)
    last = _d.fromisoformat(end)
    deleted = []
    with db() as conn:
        while cur <= last:
            ds = cur.isoformat()
            conn.execute(
                "DELETE FROM rest_days WHERE date = ?" if not DATABASE_URL else "DELETE FROM rest_days WHERE date = %s",
                (ds,)
            )
            deleted.append(ds)
            cur += timedelta(days=1)
    return jsonify({"deleted": deleted})


@app.route("/api/rest-days", methods=["POST"])
@login_required
def api_rest_days():
    data = request.json or {}
    start = data.get("start") or data.get("date")
    end = data.get("end") or start
    reason = data.get("reason") or None
    if not start:
        return jsonify({"error": "date or start required"}), 400
    from datetime import date as _d
    cur = _d.fromisoformat(start)
    last = _d.fromisoformat(end)
    inserted = []
    with db() as conn:
        while cur <= last:
            ds = cur.isoformat()
            conn.execute(
                "INSERT INTO rest_days (date, reason) VALUES (?, ?) ON CONFLICT (date) DO UPDATE SET reason=excluded.reason"
                if not DATABASE_URL else
                "INSERT INTO rest_days (date, reason) VALUES (%s, %s) ON CONFLICT (date) DO UPDATE SET reason=excluded.reason",
                (ds, reason)
            )
            inserted.append(ds)
            cur += timedelta(days=1)
    return jsonify({"inserted": inserted})


@app.route("/sweet/log", methods=["POST"])
@login_required
def log_sweet():
    data = request.json or {}
    day = data.get("date") or datetime.today().strftime("%Y-%m-%d")
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        conn.execute(
            f"INSERT INTO sweet_days (date) VALUES ({ph}) ON CONFLICT (date) DO NOTHING",
            (day,)
        )
    return jsonify({"ok": True, "date": day})


@app.route("/sweet/remove", methods=["POST"])
@login_required
def remove_sweet():
    data = request.json or {}
    day = data.get("date")
    if not day:
        return jsonify({"error": "date required"}), 400
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        conn.execute(f"DELETE FROM sweet_days WHERE date = {ph}", (day,))
    return jsonify({"ok": True, "date": day})


@app.route("/workouts")
@login_required
def workouts():
    with db() as conn:
        all_workouts = conn.execute("""
            SELECT w.*, COUNT(DISTINCT s.exercise_name) AS exercise_count,
                   COUNT(s.id) AS set_count,
                   COALESCE(SUM(CASE WHEN s.weight > 0 AND s.reps IS NOT NULL THEN s.weight * s.reps ELSE 0 END), 0) AS volume
            FROM workouts w
            LEFT JOIN sets s ON s.workout_id = w.id
            GROUP BY w.id
            ORDER BY w.date DESC
        """).fetchall()
        templates = conn.execute("SELECT * FROM templates ORDER BY name").fetchall()
        workout_ids = [w["id"] for w in all_workouts]
        workout_exercises = {}
        if workout_ids:
            placeholders = ",".join("?" if not DATABASE_URL else "%s" for _ in workout_ids)
            rows = conn.execute(
                f"""SELECT workout_id, exercise_name, COUNT(*) as set_count,
                    CAST(ROUND(AVG(reps)) AS INTEGER) as avg_reps
                    FROM sets WHERE workout_id IN ({placeholders})
                    GROUP BY workout_id, exercise_name
                    ORDER BY workout_id, MIN(id)""",
                workout_ids
            ).fetchall()
            for r in rows:
                workout_exercises.setdefault(r["workout_id"], []).append(r)

        # Workouts containing an all-time PR set (max weight or e1RM)
        pr_set_ids = set()
        for pr in get_prs().values():
            pr_set_ids.add(pr["id"])
        for pr in get_estimated_1rm().values():
            pr_set_ids.add(pr["id"])
        pr_workout_ids = set()
        if pr_set_ids and workout_ids:
            placeholders = ",".join("?" if not DATABASE_URL else "%s" for _ in pr_set_ids)
            pr_rows = conn.execute(
                f"SELECT DISTINCT workout_id FROM sets WHERE id IN ({placeholders})",
                list(pr_set_ids)
            ).fetchall()
            pr_workout_ids = {r["workout_id"] for r in pr_rows}

    today = datetime.today().date()
    week_start = week_start_of(today).strftime("%Y-%m-%d")
    month_start = today.replace(day=1).strftime("%Y-%m-%d")

    week_stats = {"workouts": 0, "volume": 0, "sets": 0}
    month_stats = {"workouts": 0, "volume": 0, "sets": 0}
    for w in all_workouts:
        if w["date"] >= month_start:
            month_stats["workouts"] += 1
            month_stats["volume"] += w["volume"] or 0
            month_stats["sets"] += w["set_count"] or 0
            if w["date"] >= week_start:
                week_stats["workouts"] += 1
                week_stats["volume"] += w["volume"] or 0
                week_stats["sets"] += w["set_count"] or 0

    days_since_last = None
    if all_workouts:
        last_date = datetime.strptime(all_workouts[0]["date"], "%Y-%m-%d").date()
        days_since_last = (today - last_date).days

    # Days since last performed, per workout type (Push/Pull/Legs/etc.)
    last_performed_raw = {}
    for w in all_workouts:
        label = workout_type_label(w["name"] or "")
        if label not in last_performed_raw:
            w_date = datetime.strptime(w["date"], "%Y-%m-%d").date()
            last_performed_raw[label] = (today - w_date).days
    type_order = ["Push", "Pull", "Legs", "Chest & Back", "Shoulders & Arms"]
    last_performed = {label: last_performed_raw[label] for label in type_order if label in last_performed_raw}
    for label, days in last_performed_raw.items():
        if label not in last_performed:
            last_performed[label] = days

    # Weekly volume trend (last 8 weeks, oldest first)
    weeks = []
    for i in range(7, -1, -1):
        wk_start = week_start_of(today) - timedelta(days=i * 7)
        wk_end = wk_start + timedelta(days=6)
        weeks.append({"start": wk_start.strftime("%Y-%m-%d"), "end": wk_end.strftime("%Y-%m-%d"), "volume": 0})
    for w in all_workouts:
        for wk in weeks:
            if wk["start"] <= w["date"] <= wk["end"]:
                wk["volume"] += w["volume"] or 0
                break
    chart_labels = [
        f"{datetime.strptime(wk['start'], '%Y-%m-%d').strftime('%b %-d')}–{datetime.strptime(wk['end'], '%Y-%m-%d').strftime('%-d')}"
        for wk in weeks
    ]
    chart_volume = [round(wk["volume"]) for wk in weeks]

    # Volume by workout type
    type_volume = {}
    for w in all_workouts:
        label = workout_type_label(w["name"] or "")
        type_volume[label] = type_volume.get(label, 0) + (w["volume"] or 0)
    total_type_volume = sum(type_volume.values())
    type_volume = sorted(type_volume.items(), key=lambda x: -x[1])

    return render_template(
        "workouts.html", workouts=all_workouts, templates=templates, workout_exercises=workout_exercises,
        week_stats=week_stats, month_stats=month_stats, days_since_last=days_since_last,
        last_performed=last_performed, chart_labels=chart_labels, chart_volume=chart_volume,
        pr_workout_ids=pr_workout_ids, type_volume=type_volume, total_type_volume=total_type_volume,
    )


@app.route("/workouts/new", methods=["GET", "POST"])
@login_required
def new_workout():
    if request.method == "POST":
        data = request.json
        date = data.get("date", datetime.today().strftime("%Y-%m-%d"))
        name = data.get("name", "")
        notes = data.get("notes", "")
        template_id = data.get("template_id") or None
        duration_min = data.get("duration_min") or None
        exercises = data.get("exercises", [])

        prior_e1rm = {k: v["estimated_1rm"] for k, v in get_estimated_1rm().items()}

        new_prs = []
        with db() as conn:
            rating = data.get("rating") or None
            cur = conn.execute(
                "INSERT INTO workouts (date, name, notes, template_id, duration_min, rating, calories, avg_heart_rate) VALUES (?,?,?,?,?,?,?,?)",
                (date, name, notes, template_id, duration_min, rating,
                 int_in_range(data.get("calories"), 1, 10000), int_in_range(data.get("avg_heart_rate"), 30, 250)),
            )
            workout_id = cur.lastrowid
            for ex in exercises:
                ss_group = ex.get("superset_group")
                ex_note = ex.get("note", "").strip() or None
                for i, s in enumerate(ex.get("sets", []), 1):
                    is_warmup = 1 if s.get("is_warmup") else 0
                    conn.execute(
                        "INSERT INTO sets (workout_id, exercise_name, set_number, reps, weight, unit, rpe, completed, superset_group, is_warmup, exercise_note) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        (
                            workout_id,
                            ex["name"],
                            i,
                            s.get("reps"),
                            s.get("weight"),
                            s.get("unit", "lbs"),
                            s.get("rpe"),
                            1 if s.get("completed", True) else 0,
                            ss_group,
                            is_warmup,
                            ex_note,
                        ),
                    )
                    weight = s.get("weight")
                    reps = s.get("reps")
                    if is_warmup:
                        continue
                    if weight and reps and reps >= 1 and reps <= 12 and s.get("completed", True):
                        e1rm = round(weight * (1 + reps / 30), 1) if reps > 1 else weight
                        prior = prior_e1rm.get(ex["name"])
                        if prior is None or e1rm > prior:
                            prior_e1rm[ex["name"]] = e1rm
                            new_prs.append({
                                "name": ex["name"], "weight": weight, "reps": reps,
                                "unit": s.get("unit", "lbs"), "e1rm": e1rm,
                                "prev_e1rm": prior,
                            })
        return jsonify({"id": workout_id, "new_prs": new_prs})

    with db() as conn:
        templates = conn.execute("SELECT * FROM templates ORDER BY name").fetchall()
        template_exercises = conn.execute(
            "SELECT * FROM template_exercises ORDER BY template_id, sort_order"
        ).fetchall()
        exercise_bank = conn.execute(
            "SELECT name, muscle_group FROM exercise_bank ORDER BY muscle_group, name"
        ).fetchall()
        last_sets_rows = conn.execute("""
            SELECT s.exercise_name, s.set_number, s.reps, s.weight, s.unit, s.rpe, w.date
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            JOIN (
                SELECT s2.exercise_name, MAX(w2.date) as max_date
                FROM sets s2 JOIN workouts w2 ON w2.id = s2.workout_id
                GROUP BY s2.exercise_name
            ) latest ON latest.exercise_name = s.exercise_name AND latest.max_date = w.date
            ORDER BY s.exercise_name, s.set_number
        """).fetchall()
        last_by_reps_rows = conn.execute("""
            SELECT s.exercise_name, s.reps, s.weight, s.unit, w.date
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            WHERE s.weight IS NOT NULL AND s.reps IS NOT NULL
            ORDER BY s.exercise_name, s.reps, w.date, s.id
        """).fetchall()
    grouped = {}
    for ex in exercise_bank:
        mg = ex["muscle_group"] or "Other"
        grouped.setdefault(mg, []).append(ex["name"])
    last_sets = {}
    for r in last_sets_rows:
        last_sets.setdefault(r["exercise_name"], []).append(
            {"reps": r["reps"], "weight": r["weight"], "unit": r["unit"], "rpe": r["rpe"], "date": r["date"]}
        )
    # last_by_reps[exercise][reps] = most recent {weight, unit, date} logged at that rep count
    last_by_reps = {}
    for r in last_by_reps_rows:
        last_by_reps.setdefault(r["exercise_name"], {})[r["reps"]] = {
            "weight": r["weight"], "unit": r["unit"], "date": r["date"]
        }
    import json as _json
    return render_template(
        "new_workout.html",
        templates=templates,
        template_exercises=template_exercises,
        exercise_bank=exercise_bank,
        exercise_bank_json=_json.dumps(grouped),
        last_sets_json=_json.dumps(last_sets),
        last_by_reps_json=_json.dumps(last_by_reps),
    )


@app.route("/workouts/<int:workout_id>")
@login_required
def workout_detail(workout_id):
    with db() as conn:
        workout = conn.execute("SELECT * FROM workouts WHERE id=?", (workout_id,)).fetchone()
        if not workout:
            return redirect(url_for("workouts"))
        sets = conn.execute(
            "SELECT * FROM sets WHERE workout_id=? ORDER BY id",
            (workout_id,),
        ).fetchall()

    exercises = {}
    for s in sets:
        exercises.setdefault(s["exercise_name"], []).append(s)

    superset_groups = {}
    for s in sets:
        sg = s.get("superset_group")
        if sg:
            superset_groups[s["exercise_name"]] = sg

    total_volume = sum(
        (s["weight"] or 0) * (s["reps"] or 0)
        for s in sets if s["weight"] and s["reps"]
    )
    total_sets = len(sets)

    prs = get_prs()
    return render_template("workout_detail.html", workout=workout, exercises=exercises, prs=prs,
                           total_volume=total_volume, total_sets=total_sets,
                           superset_groups=superset_groups)


@app.route("/api/export")
def api_export():
    key = request.headers.get("X-API-Key") or request.args.get("api_key")
    if not API_KEY or key != API_KEY:
        return jsonify({"error": "unauthorized"}), 401

    days = request.args.get("days", type=int, default=7)
    cutoff = (datetime.today() - timedelta(days=days)).strftime("%Y-%m-%d")
    ph = "%s" if DATABASE_URL else "?"

    with db() as conn:
        workouts = conn.execute(
            f"SELECT * FROM workouts WHERE date >= {ph} ORDER BY date, id", (cutoff,)
        ).fetchall()
        workout_ids = [w["id"] for w in workouts]
        sets_by_workout = {}
        if workout_ids:
            placeholders = ",".join(ph for _ in workout_ids)
            rows = conn.execute(
                f"""SELECT * FROM sets WHERE workout_id IN ({placeholders})
                    ORDER BY workout_id, id""",
                workout_ids
            ).fetchall()
            for r in rows:
                sets_by_workout.setdefault(r["workout_id"], []).append({
                    "exercise_name": r["exercise_name"],
                    "set_number": r["set_number"],
                    "reps": r["reps"],
                    "weight": r["weight"],
                    "unit": r["unit"],
                    "rpe": r["rpe"],
                    "is_warmup": bool(r["is_warmup"]),
                    "superset_group": r["superset_group"],
                    "exercise_note": r["exercise_note"],
                })

    result = []
    for w in workouts:
        result.append({
            "id": w["id"],
            "date": w["date"],
            "name": w["name"],
            "notes": w["notes"],
            "duration_min": w["duration_min"],
            "rating": w["rating"],
            "calories": w["calories"],
            "avg_heart_rate": w["avg_heart_rate"],
            "sets": sets_by_workout.get(w["id"], []),
        })
    return jsonify(result)


def _api_auth():
    key = request.headers.get("X-API-Key") or request.args.get("api_key")
    if not API_KEY or key != API_KEY:
        return False
    return True

def _api_cutoff():
    days = request.args.get("days", type=int, default=30)
    return (datetime.today() - timedelta(days=days)).strftime("%Y-%m-%d")


@app.route("/api/export/cardio")
def api_export_cardio():
    if not _api_auth():
        return jsonify({"error": "unauthorized"}), 401
    cutoff = _api_cutoff()
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        rows = conn.execute(
            f"SELECT * FROM cardio_sessions WHERE date >= {ph} ORDER BY date DESC", (cutoff,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/export/cardio-splits")
def api_export_cardio_splits():
    if not _api_auth():
        return jsonify({"error": "unauthorized"}), 401
    cutoff = _api_cutoff()
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        rows = conn.execute(
            f"""SELECT sp.*, cs.date, cs.activity
                FROM cardio_splits sp
                JOIN cardio_sessions cs ON sp.session_id = cs.id
                WHERE cs.date >= {ph}
                ORDER BY cs.date DESC, sp.mile_number""",
            (cutoff,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/export/bodyweight")
def api_export_bodyweight():
    if not _api_auth():
        return jsonify({"error": "unauthorized"}), 401
    cutoff = _api_cutoff()
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        rows = conn.execute(
            f"SELECT * FROM body_weight WHERE date >= {ph} ORDER BY date DESC", (cutoff,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/export/sleep")
def api_export_sleep():
    if not _api_auth():
        return jsonify({"error": "unauthorized"}), 401
    cutoff = _api_cutoff()
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        rows = conn.execute(
            f"SELECT * FROM sleep_logs WHERE date >= {ph} ORDER BY date DESC", (cutoff,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/export/rest-days")
def api_export_rest_days():
    if not _api_auth():
        return jsonify({"error": "unauthorized"}), 401
    cutoff = _api_cutoff()
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        rows = conn.execute(
            f"SELECT * FROM rest_days WHERE date >= {ph} ORDER BY date DESC", (cutoff,)
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/api/export/prs")
def api_export_prs():
    if not _api_auth():
        return jsonify({"error": "unauthorized"}), 401
    prs = get_prs()
    e1rm = get_estimated_1rm()
    result = {}
    for name, pr in prs.items():
        result[name] = dict(pr)
    for name, pr in e1rm.items():
        result.setdefault(name, {})["estimated_1rm"] = pr
    return jsonify(result)


@app.route("/workouts/<int:workout_id>/repeat", methods=["POST"])
@login_required
def repeat_workout(workout_id):
    with db() as conn:
        original = conn.execute("SELECT * FROM workouts WHERE id=?", (workout_id,)).fetchone()
        if not original:
            return redirect(url_for("workouts"))
        sets = conn.execute(
            "SELECT * FROM sets WHERE workout_id=? ORDER BY id",
            (workout_id,)
        ).fetchall()
        today_str = datetime.today().strftime("%Y-%m-%d")
        cur = conn.execute(
            "INSERT INTO workouts (date, name, notes, template_id) VALUES (?,?,?,?)",
            (today_str, original["name"], "", original.get("template_id"))
        )
        new_id = cur.lastrowid
        for s in sets:
            conn.execute(
                "INSERT INTO sets (workout_id, exercise_name, set_number, reps, weight, unit, rpe, completed, superset_group, is_warmup, exercise_note) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (new_id, s["exercise_name"], s["set_number"], s["reps"], s["weight"], s["unit"], s["rpe"], s["completed"], s.get("superset_group"), s.get("is_warmup", 0), s.get("exercise_note"))
            )
    return redirect(url_for("edit_workout", workout_id=new_id))


@app.route("/workouts/<int:workout_id>/save-template", methods=["POST"])
@login_required
def save_template_from_workout(workout_id):
    data = request.json or {}
    name = data.get("name", "").strip()
    with db() as conn:
        workout = conn.execute("SELECT * FROM workouts WHERE id=?", (workout_id,)).fetchone()
        if not workout:
            return jsonify({"error": "not found"}), 404
        if not name:
            name = workout["name"] or "My Template"
        ex_rows = conn.execute(
            "SELECT exercise_name, COUNT(*) as set_count, CAST(ROUND(AVG(reps)) AS INTEGER) as avg_reps FROM sets WHERE workout_id=? GROUP BY exercise_name ORDER BY MIN(id)",
            (workout_id,)
        ).fetchall()
        cur = conn.execute("INSERT INTO templates (name) VALUES (?)", (name,))
        template_id = cur.lastrowid
        for i, ex in enumerate(ex_rows):
            conn.execute(
                "INSERT INTO template_exercises (template_id, exercise_name, sets, reps, sort_order) VALUES (?,?,?,?,?)",
                (template_id, ex["exercise_name"], ex["set_count"], ex["avg_reps"], i)
            )
    return jsonify({"ok": True, "id": template_id})


@app.route("/workouts/<int:workout_id>/notes", methods=["POST"])
@login_required
def update_workout_notes(workout_id):
    data = request.json or {}
    with db() as conn:
        conn.execute("UPDATE workouts SET notes=? WHERE id=?", (data.get("notes", "") or "", workout_id))
    return jsonify({"ok": True})


@app.route("/weekly-notes")
@login_required
def weekly_notes_page():
    with db() as conn:
        notes = conn.execute(
            "SELECT * FROM weekly_notes ORDER BY week_start DESC"
        ).fetchall()
    notes_with_end = []
    for n in notes:
        n = dict(n)
        start = datetime.strptime(n["week_start"], "%Y-%m-%d").date()
        n["week_end"] = (start + timedelta(days=6)).strftime("%Y-%m-%d")
        notes_with_end.append(n)
    return render_template("weekly_notes.html", notes=notes_with_end)


@app.route("/api/weekly-notes", methods=["GET", "POST"])
def api_weekly_notes():
    if request.method == "POST":
        data = request.json or {}
        key = (request.headers.get("X-API-Key")
               or request.args.get("api_key")
               or data.get("api_key"))
        if not API_KEY or key != API_KEY:
            return jsonify({"error": "unauthorized"}), 401

        week_start = data.get("week_start")
        title = data.get("title", "")
        body = data.get("body", "")
        is_rest_week = 1 if data.get("is_rest_week") else 0
        if not week_start:
            return jsonify({"error": "week_start is required"}), 400

        now = datetime.utcnow().isoformat()
        with db() as conn:
            conn.execute(
                """INSERT INTO weekly_notes (week_start, title, body, is_rest_week, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT (week_start) DO UPDATE SET
                       title = excluded.title,
                       body = excluded.body,
                       is_rest_week = excluded.is_rest_week,
                       updated_at = excluded.updated_at""",
                (week_start, title, body, is_rest_week, now, now),
            )
            note = conn.execute(
                "SELECT * FROM weekly_notes WHERE week_start = ?", (week_start,)
            ).fetchone()
        return jsonify(dict(note))

    key = request.headers.get("X-API-Key") or request.args.get("api_key")
    if not API_KEY or key != API_KEY:
        return jsonify({"error": "unauthorized"}), 401

    week_start = request.args.get("week_start")
    with db() as conn:
        if week_start:
            notes = conn.execute(
                "SELECT * FROM weekly_notes WHERE week_start = ? ORDER BY week_start DESC",
                (week_start,),
            ).fetchall()
        else:
            notes = conn.execute(
                "SELECT * FROM weekly_notes ORDER BY week_start DESC"
            ).fetchall()
    return jsonify([dict(n) for n in notes])


@app.route("/api/notes/weekly-review", methods=["POST"])
def api_notes_weekly_review():
    key = request.headers.get("X-API-Key") or request.args.get("api_key")
    if not API_KEY or key != API_KEY:
        return jsonify({"error": "unauthorized"}), 401

    data = request.json or {}
    raw_date = data.get("date")
    content = data.get("content", "")
    title = data.get("title", "Weekly Review")
    if not raw_date:
        return jsonify({"error": "date is required"}), 400

    from datetime import date as _d
    try:
        d = _d.fromisoformat(raw_date)
    except ValueError:
        return jsonify({"error": "invalid date format, expected YYYY-MM-DD"}), 400

    week_start = week_start_of(d).isoformat()
    now = datetime.utcnow().isoformat()
    with db() as conn:
        conn.execute(
            """INSERT INTO weekly_notes (week_start, title, body, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT (week_start) DO UPDATE SET
                   title = excluded.title,
                   body = excluded.body,
                   updated_at = excluded.updated_at""",
            (week_start, title, content, now, now),
        )
        note = conn.execute(
            "SELECT * FROM weekly_notes WHERE week_start = ?", (week_start,)
        ).fetchone()
    return jsonify(dict(note)), 201


@app.route("/workouts/<int:workout_id>/edit", methods=["GET", "POST"])
@login_required
def edit_workout(workout_id):
    if request.method == "POST":
        data = request.json
        date = data.get("date")
        name = data.get("name", "")
        notes = data.get("notes", "")
        duration_min = data.get("duration_min") or None
        exercises = data.get("exercises", [])
        with db() as conn:
            conn.execute(
                "UPDATE workouts SET date=?, name=?, notes=?, duration_min=?, rating=?, calories=?, avg_heart_rate=? WHERE id=?",
                (date, name, notes, duration_min, data.get("rating") or None,
                 int_in_range(data.get("calories"), 1, 10000), int_in_range(data.get("avg_heart_rate"), 30, 250), workout_id),
            )
            conn.execute("DELETE FROM sets WHERE workout_id=?", (workout_id,))
            for ex in exercises:
                ss_group = ex.get("superset_group")
                ex_note = ex.get("note", "").strip() or None
                for i, s in enumerate(ex.get("sets", []), 1):
                    conn.execute(
                        "INSERT INTO sets (workout_id, exercise_name, set_number, reps, weight, unit, rpe, completed, superset_group, is_warmup, exercise_note) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                        (
                            workout_id,
                            ex["name"],
                            i,
                            s.get("reps"),
                            s.get("weight"),
                            s.get("unit", "lbs"),
                            s.get("rpe"),
                            1 if s.get("completed", True) else 0,
                            ss_group,
                            1 if s.get("is_warmup") else 0,
                            ex_note,
                        ),
                    )
        return jsonify({"id": workout_id})

    with db() as conn:
        workout = conn.execute("SELECT * FROM workouts WHERE id=?", (workout_id,)).fetchone()
        if not workout:
            return redirect(url_for("workouts"))
        sets = conn.execute(
            "SELECT * FROM sets WHERE workout_id=? ORDER BY id",
            (workout_id,),
        ).fetchall()
        exercise_bank = conn.execute(
            "SELECT name, muscle_group FROM exercise_bank ORDER BY muscle_group, name"
        ).fetchall()

    exercises = {}
    for s in sets:
        exercises.setdefault(s["exercise_name"], []).append(s)

    exercises_list = [
        {
            "name": ex_name,
            "superset_group": s_list[0].get("superset_group"),
            "note": s_list[0].get("exercise_note") or "",
            "sets": [
                {"reps": s["reps"], "weight": s["weight"], "unit": s["unit"], "rpe": s["rpe"], "is_warmup": bool(s.get("is_warmup"))}
                for s in s_list
            ],
        }
        for ex_name, s_list in exercises.items()
    ]

    grouped = {}
    for ex in exercise_bank:
        mg = ex["muscle_group"] or "Other"
        grouped.setdefault(mg, []).append(ex["name"])

    import json as _json
    return render_template(
        "edit_workout.html",
        workout=workout,
        exercise_bank=exercise_bank,
        exercise_bank_json=_json.dumps(grouped),
        exercises_json=_json.dumps(exercises_list),
    )


@app.route("/workouts/<int:workout_id>/delete", methods=["POST"])
@login_required
def delete_workout(workout_id):
    with db() as conn:
        conn.execute("DELETE FROM workouts WHERE id=?", (workout_id,))
    return jsonify({"ok": True})


@app.route("/weight")
@login_required
def weight():
    with db() as conn:
        entries = conn.execute(
            "SELECT * FROM body_weight ORDER BY date DESC"
        ).fetchall()
    peak = max((e["weight"] for e in entries), default=0) if entries else 0
    heaviest = max(entries, key=lambda e: e["weight"], default=None) if entries else None
    lightest = min(entries, key=lambda e: e["weight"], default=None) if entries else None
    # 7-day and 30-day change
    from datetime import datetime, timedelta
    def weight_change(days):
        if not entries: return None
        cutoff = (datetime.today() - timedelta(days=days)).strftime("%Y-%m-%d")
        past = [e for e in entries if e["date"] <= cutoff]
        if not past: return None
        return round(entries[0]["weight"] - past[0]["weight"], 1)
    week_change = weight_change(7)
    month_change = weight_change(30)
    year_change = weight_change(365)
    last_change = round(entries[0]["weight"] - entries[1]["weight"], 1) if len(entries) >= 2 else None
    goal_weight   = get_setting("goal_weight")
    height_inches = get_setting("height_inches")
    age           = get_setting("age")
    sex           = get_setting("sex")  # "male" or "female"
    goal_weight   = float(goal_weight) if goal_weight else None
    height_inches = float(height_inches) if height_inches else None
    age           = int(age) if age else None

    def calc_bmi(w):
        return round(w * 703 / (height_inches ** 2), 1) if height_inches else None

    def calc_bf(w):
        # Deurenberg formula: BFP = (1.20 * BMI) + (0.23 * age) - (10.8 * sex_val) - 5.4
        # sex_val: 1=male, 0=female
        if not height_inches or not age or not sex:
            return None
        bmi_val = w * 703 / (height_inches ** 2)
        sex_val = 1 if sex == "male" else 0
        return round((1.20 * bmi_val) + (0.23 * age) - (10.8 * sex_val) - 5.4, 1)

    current_weight = entries[0]["weight"] if entries else None
    bmi    = calc_bmi(current_weight) if current_weight else None
    bf_pct = calc_bf(current_weight) if current_weight else None
    to_goal = round(current_weight - goal_weight, 1) if current_weight and goal_weight else None

    # Per-entry computed stats for the table
    entry_stats = []
    for e in entries:
        entry_stats.append({
            "bmi": calc_bmi(e["weight"]),
            "bf":  calc_bf(e["weight"]),
        })

    goal_eta = None
    if goal_weight and len(entries) >= 2 and current_weight and current_weight > goal_weight:
        cutoff_30 = (datetime.today() - timedelta(days=30)).strftime("%Y-%m-%d")
        recent = [e for e in entries if e["date"] >= cutoff_30]
        if len(recent) >= 2:
            oldest = recent[-1]
            newest = recent[0]
            days_span = (datetime.strptime(newest["date"], "%Y-%m-%d") - datetime.strptime(oldest["date"], "%Y-%m-%d")).days
            if days_span > 0:
                rate_per_day = (newest["weight"] - oldest["weight"]) / days_span
                if rate_per_day < 0:
                    days_left = (current_weight - goal_weight) / abs(rate_per_day)
                    eta_date = (datetime.today() + timedelta(days=days_left)).strftime("%Y-%m-%d")
                    goal_eta = {"date": eta_date, "weeks": round(days_left / 7, 1)}

    return render_template("weight.html",
        entries=entries, peak_weight=peak,
        week_change=week_change, month_change=month_change,
        year_change=year_change, last_change=last_change,
        goal_weight=goal_weight, height_inches=height_inches,
        bmi=bmi, bf_pct=bf_pct, to_goal=to_goal,
        age=age, sex=sex, entry_stats=entry_stats,
        heaviest=heaviest, lightest=lightest,
        goal_eta=goal_eta,
    )


@app.route("/weight/settings", methods=["POST"])
@login_required
def save_weight_settings():
    data = request.get_json()
    if "goal_weight" in data and data["goal_weight"]:
        set_setting("goal_weight", str(data["goal_weight"]))
    if "height_inches" in data and data["height_inches"]:
        set_setting("height_inches", str(data["height_inches"]))
    if "age" in data and data["age"]:
        set_setting("age", str(data["age"]))
    if "sex" in data and data["sex"]:
        set_setting("sex", str(data["sex"]))
    return jsonify({"ok": True})


@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings_page():
    if request.method == "POST":
        f = request.form
        schedule = [
            {"label": f.get(f"day{i}_label", "").strip()[:40],
             "color": f.get(f"day{i}_color") if re.fullmatch(r"#[0-9a-fA-F]{6}", f.get(f"day{i}_color", "")) else _SCHEDULE_COLORS[i]}
            for i in range(7)
        ]
        set_setting("schedule", json.dumps(schedule))
        for key in ["cardio_goal_min", "sleep_goal_hours", "streak_goal_days"]:
            try:
                if float(f.get(key, "")) > 0:
                    set_setting(key, f[key])
            except ValueError:
                pass
        set_setting("streak_skip_weekends", "1" if f.get("streak_skip_weekends") else "0")
        # Body stats may be cleared
        ft, inch = f.get("height_ft", "").strip(), f.get("height_in", "").strip()
        try:
            height = float(ft or 0) * 12 + float(inch or 0)
        except ValueError:
            height = 0
        set_setting("height_inches", str(height) if height else "")
        for key in ["goal_weight", "age"]:
            try:
                set_setting(key, str(float(f[key]) if key == "goal_weight" else int(f[key])) if f.get(key, "").strip() else "")
            except ValueError:
                pass
        set_setting("sex", f.get("sex") if f.get("sex") in ("male", "female") else "")
        return redirect(url_for("settings_page", saved=1))

    settings = get_settings()
    height = float(settings["height_inches"]) if settings["height_inches"] else None
    return render_template(
        "settings.html",
        schedule=list(zip(WEEKDAYS, get_schedule(settings))),
        height_ft=int(height // 12) if height else "",
        height_in=f"{round(height % 12, 1):g}" if height else "",
        saved=request.args.get("saved"),
    )


@app.route("/weight/add", methods=["POST"])
@login_required
def add_weight():
    data = request.json
    with db() as conn:
        prior_lightest = conn.execute("SELECT MIN(weight) as w FROM body_weight").fetchone()["w"]
        conn.execute(
            "INSERT INTO body_weight (date, weight, unit) VALUES (?,?,?)",
            (data["date"], data["weight"], data.get("unit", "lbs")),
        )
        goal_weight = get_setting("goal_weight")
        goal_weight = float(goal_weight) if goal_weight else None
    new_low = prior_lightest is not None and data["weight"] < prior_lightest
    goal_reached = goal_weight is not None and data["weight"] <= goal_weight and not (prior_lightest is not None and prior_lightest <= goal_weight)
    return jsonify({"ok": True, "new_low": new_low, "goal_reached": goal_reached})


@app.route("/weight/edit/<int:entry_id>", methods=["POST"])
@login_required
def edit_weight(entry_id):
    data = request.json
    with db() as conn:
        conn.execute(
            "UPDATE body_weight SET date=?, weight=?, unit=? WHERE id=?",
            (data["date"], data["weight"], data.get("unit", "lbs"), entry_id),
        )
    return jsonify({"ok": True})


@app.route("/weight/delete/<int:entry_id>", methods=["POST"])
@login_required
def delete_weight(entry_id):
    with db() as conn:
        conn.execute("DELETE FROM body_weight WHERE id=?", (entry_id,))
    return jsonify({"ok": True})


@app.route("/progress")
@login_required
def progress():
    with db() as conn:
        exercises = conn.execute(
            "SELECT DISTINCT exercise_name FROM sets ORDER BY exercise_name"
        ).fetchall()
        weight_data = conn.execute(
            "SELECT date, weight, unit FROM body_weight ORDER BY date"
        ).fetchall()
    return render_template("progress.html", exercises=exercises, weight_data=weight_data)


@app.route("/api/progress/<exercise_name>")
@login_required
def exercise_progress(exercise_name):
    with db() as conn:
        rows = conn.execute("""
            SELECT w.date, MAX(s.weight) as max_weight,
                   MAX(s.unit) as unit,
                   MAX(CASE WHEN s.reps IS NOT NULL AND s.reps > 0
                       THEN CASE WHEN s.reps = 1 THEN s.weight ELSE s.weight * (1 + s.reps / 30.0) END
                       ELSE s.weight END) as e1rm
            FROM sets s
            JOIN workouts w ON w.id = s.workout_id
            WHERE s.exercise_name = ? AND s.weight IS NOT NULL AND s.completed = 1
            GROUP BY w.date
            ORDER BY w.date
        """, (exercise_name,)).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/templates")
@login_required
def templates():
    with db() as conn:
        tmps = conn.execute("SELECT * FROM templates ORDER BY name").fetchall()
        exercises = {
            t["id"]: conn.execute(
                "SELECT * FROM template_exercises WHERE template_id=? ORDER BY sort_order",
                (t["id"],),
            ).fetchall()
            for t in tmps
        }
        exercise_bank = conn.execute(
            "SELECT name FROM exercise_bank ORDER BY name"
        ).fetchall()
    return render_template("templates.html", templates=tmps, exercises=exercises, exercise_bank=exercise_bank)


@app.route("/templates/save", methods=["POST"])
@login_required
def save_template():
    data = request.json
    with db() as conn:
        if data.get("id"):
            conn.execute(
                "UPDATE templates SET name=?, description=? WHERE id=?",
                (data["name"], data.get("description", ""), data["id"]),
            )
            conn.execute("DELETE FROM template_exercises WHERE template_id=?", (data["id"],))
            template_id = data["id"]
        else:
            cur = conn.execute(
                "INSERT INTO templates (name, description) VALUES (?,?)",
                (data["name"], data.get("description", "")),
            )
            template_id = cur.lastrowid
        for i, ex in enumerate(data.get("exercises", [])):
            conn.execute(
                "INSERT INTO template_exercises (template_id, exercise_name, sets, reps, sort_order) VALUES (?,?,?,?,?)",
                (template_id, ex["name"], ex.get("sets"), ex.get("reps"), i),
            )
    return jsonify({"id": template_id})


@app.route("/templates/<int:template_id>/delete", methods=["POST"])
@login_required
def delete_template(template_id):
    with db() as conn:
        conn.execute("DELETE FROM templates WHERE id=?", (template_id,))
    return jsonify({"ok": True})


@app.route("/api/template/<int:template_id>")
@login_required
def get_template(template_id):
    with db() as conn:
        tmpl = conn.execute("SELECT * FROM templates WHERE id=?", (template_id,)).fetchone()
        exs = conn.execute(
            "SELECT * FROM template_exercises WHERE template_id=? ORDER BY sort_order",
            (template_id,),
        ).fetchall()
    if not tmpl:
        return jsonify({}), 404
    return jsonify({"template": dict(tmpl), "exercises": [dict(e) for e in exs]})


@app.route("/prs")
@login_required
def prs():
    all_prs = get_prs()
    prs_by_reps = get_prs_by_reps()
    estimated_1rm = get_estimated_1rm()
    with db() as conn:
        muscle_rows = conn.execute("SELECT name, muscle_group FROM exercise_bank").fetchall()
    _delt_merge = {"Front Delts", "Side Delts", "Rear Delts"}
    _arm_merge = {"Biceps", "Triceps", "Forearms"}
    _leg_merge = {"Quads", "Hamstrings", "Glutes", "Calves"}
    def _pr_group(raw):
        if raw in _delt_merge: return "Shoulders"
        if raw in _arm_merge: return "Arms"
        if raw in _leg_merge: return "Legs"
        return raw or "Other"
    muscle_map = {r["name"]: _pr_group(r["muscle_group"]) for r in muscle_rows}
    group_order = ["Chest", "Back", "Shoulders", "Arms", "Legs", "Abs", "Other"]
    def _build_grouped(items):
        grouped = {}
        for name, val in items:
            group = muscle_map.get(name, "Other")
            grouped.setdefault(group, {})[name] = val
        def _sort_group(d):
            return dict(sorted(d.items(), key=lambda x: x[1].get("weight") or 0, reverse=True))
        ordered = {g: _sort_group(grouped[g]) for g in group_order if g in grouped}
        for g, v in grouped.items():
            if g not in ordered:
                ordered[g] = _sort_group(v)
        return ordered
    prs_by_group = _build_grouped(all_prs.items())
    e1rm_by_group = _build_grouped(estimated_1rm.items())
    cardio_prs = get_cardio_prs()

    pr_timeline = get_pr_timeline()
    today = date.today()

    # Days since last PR + next-milestone progress, per exercise
    for name, pr in all_prs.items():
        pr_date = datetime.strptime(pr["date"], "%Y-%m-%d").date()
        pr["days_since"] = (today - pr_date).days
        if pr["weight"]:
            step = 25
            next_milestone = pr["weight"] + step - (pr["weight"] % step)
            if next_milestone <= pr["weight"]:
                next_milestone += step
            pr["next_milestone"] = next_milestone
            pr["milestone_pct"] = round((pr["weight"] / next_milestone) * 100, 1)

    # Add days_since to prs_by_reps for tab coloring
    for rep, exercises in prs_by_reps.items():
        for name, pr in exercises.items():
            pr_date = datetime.strptime(pr["date"], "%Y-%m-%d").date()
            pr["days_since"] = (today - pr_date).days

    # Recent PRs feed (flatten timeline, newest first)
    recent_prs = []
    for name, events in pr_timeline.items():
        for e in events:
            recent_prs.append({"name": name, **e})
    recent_prs.sort(key=lambda x: (x["date"], x["id"]), reverse=True)

    cutoff_30 = (today - timedelta(days=30)).strftime("%Y-%m-%d")
    cutoff_90 = (today - timedelta(days=90)).strftime("%Y-%m-%d")
    recent_prs_30 = [p for p in recent_prs if p["date"] >= cutoff_30]
    recent_prs_90 = [p for p in recent_prs if p["date"] >= cutoff_90]

    month_start = today.replace(day=1).strftime("%Y-%m-%d")
    year_start = today.replace(month=1, day=1).strftime("%Y-%m-%d")
    pr_count_month = sum(1 for p in recent_prs if p["date"] >= month_start)
    pr_count_year = sum(1 for p in recent_prs if p["date"] >= year_start)

    return render_template(
        "prs.html", prs=all_prs, prs_by_reps=prs_by_reps, estimated_1rm=estimated_1rm,
        prs_by_group=prs_by_group, e1rm_by_group=e1rm_by_group, cardio_prs=cardio_prs,
        pr_timeline=pr_timeline, recent_prs_30=recent_prs_30, recent_prs_90=recent_prs_90,
        pr_count_month=pr_count_month, pr_count_year=pr_count_year,
    )


@app.route("/cardio")
@login_required
def cardio():
    with db() as conn:
        sessions = conn.execute(
            "SELECT * FROM cardio_sessions ORDER BY date DESC"
        ).fetchall()

    today = datetime.today().date()
    week_start = week_start_of(today).strftime("%Y-%m-%d")
    month_start = today.replace(day=1).strftime("%Y-%m-%d")

    year_start = today.replace(month=1, day=1).strftime("%Y-%m-%d")

    week_stats = {"sessions": 0, "minutes": 0, "distance": 0}
    month_stats = {"sessions": 0, "minutes": 0, "distance": 0}
    year_stats = {"sessions": 0, "minutes": 0, "distance": 0}
    for s in sessions:
        if s["date"] >= year_start:
            year_stats["sessions"] += 1
            year_stats["minutes"] += s["duration_min"] or 0
            year_stats["distance"] += s["distance"] or 0
        if s["date"] >= month_start:
            month_stats["sessions"] += 1
            month_stats["minutes"] += s["duration_min"] or 0
            month_stats["distance"] += s["distance"] or 0
            if s["date"] >= week_start:
                week_stats["sessions"] += 1
                week_stats["minutes"] += s["duration_min"] or 0
                week_stats["distance"] += s["distance"] or 0

    cardio_prs = get_cardio_prs()

    # Weekly minutes trend (last 8 weeks, oldest first)
    weeks = []
    for i in range(7, -1, -1):
        wk_start = week_start_of(today) - timedelta(days=i * 7)
        wk_end = wk_start + timedelta(days=6)
        weeks.append({"start": wk_start.strftime("%Y-%m-%d"), "end": wk_end.strftime("%Y-%m-%d"), "minutes": 0})
    for s in sessions:
        for wk in weeks:
            if wk["start"] <= s["date"] <= wk["end"]:
                wk["minutes"] += s["duration_min"] or 0
                break
    chart_labels = [
        f"{datetime.strptime(wk['start'], '%Y-%m-%d').strftime('%b %-d')}–{datetime.strptime(wk['end'], '%Y-%m-%d').strftime('%-d')}"
        for wk in weeks
    ]
    chart_minutes = [round(wk["minutes"]) for wk in weeks]

    # Cardio goal streak — consecutive weeks hitting the goal, counting back from last completed week
    cardio_goal = setting_num(get_settings(), "cardio_goal_min")
    goal_streak = 0
    cur_wk_start = week_start_of(today)
    for i in range(1, 200):
        wk_s = cur_wk_start - timedelta(days=i * 7)
        wk_e = wk_s + timedelta(days=6)
        mins = sum((s["duration_min"] or 0) for s in sessions if wk_s.strftime("%Y-%m-%d") <= s["date"] <= wk_e.strftime("%Y-%m-%d"))
        if mins >= cardio_goal:
            goal_streak += 1
        else:
            break
    # Include current week if already hit
    if week_stats["minutes"] >= cardio_goal:
        goal_streak += 1

    # Per-session distance chart (last 30 sessions with distance, oldest first)
    dist_sessions = [s for s in reversed(sessions) if s["distance"]][-30:]
    chart_dist_labels = [s["date"] for s in dist_sessions]
    chart_dist_values = [float(s["distance"]) for s in dist_sessions]
    # Per-session duration chart (last 30 sessions with duration, oldest first)
    dur_sessions = [s for s in reversed(sessions) if s["duration_min"]][-30:]
    chart_dur_labels = [s["date"] for s in dur_sessions]
    chart_dur_values = [float(s["duration_min"]) for s in dur_sessions]

    # Load splits for all sessions
    splits_by_session = {}
    with db() as conn:
        all_splits = conn.execute("SELECT * FROM cardio_splits ORDER BY session_id, mile_number").fetchall()
    for sp in all_splits:
        splits_by_session.setdefault(sp["session_id"], []).append(dict(sp))

    return render_template(
        "cardio.html", sessions=sessions, week_stats=week_stats, month_stats=month_stats,
        year_stats=year_stats,
        cardio_prs=cardio_prs, chart_labels=chart_labels, chart_minutes=chart_minutes,
        chart_dist_labels=chart_dist_labels, chart_dist_values=chart_dist_values,
        chart_dur_labels=chart_dur_labels, chart_dur_values=chart_dur_values,
        goal_streak=goal_streak, cardio_goal=cardio_goal, splits_by_session=splits_by_session,
        cardio_activities=get_cardio_activities(),
    )


@app.route("/cardio/activities/add", methods=["POST"])
@login_required
def add_cardio_activity_route():
    data = request.json
    if not (data.get("name") or "").strip():
        return jsonify({"error": "Name required"}), 400
    with db() as conn:
        add_cardio_activity(conn, data["name"], data.get("category", "Other"))
    return jsonify({"ok": True})


@app.route("/cardio/add", methods=["POST"])
@login_required
def add_cardio():
    data = request.json
    prior_prs = get_cardio_prs()
    with db() as conn:
        add_cardio_activity(conn, data.get("activity"))
        cur = conn.execute(
            "INSERT INTO cardio_sessions (date, activity, duration_min, distance, distance_unit, calories, incline, speed, notes, elevation_gain, avg_heart_rate, pool_length, pool_length_unit, laps, elapsed_time_min) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (
                data["date"],
                data["activity"],
                data.get("duration_min") or None,
                data.get("distance") or None,
                data.get("distance_unit", "mi"),
                data.get("calories") or None,
                data.get("incline") or None,
                data.get("speed") or None,
                data.get("notes", ""),
                data.get("elevation_gain") or None,
                data.get("avg_heart_rate") or None,
                data.get("pool_length") or None,
                data.get("pool_length_unit", "yd"),
                data.get("laps") or None,
                data.get("elapsed_time_min") or None,
            ),
        )
        session_id = cur.lastrowid
        for split in data.get("splits", []):
            conn.execute(
                "INSERT INTO cardio_splits (session_id, mile_number, pace_seconds, peak_hr_zone, stroke_type, distance_per_stroke, split_distance) VALUES (?,?,?,?,?,?,?)",
                (session_id, split["mile"], split.get("pace_seconds"), split.get("peak_hr_zone"), split.get("stroke_type"), split.get("distance_per_stroke"), split.get("split_distance", 1.0)),
            )

    new_prs = []
    activity = data["activity"]
    distance = data.get("distance")
    distance_unit = data.get("distance_unit", "mi")
    duration_min = data.get("duration_min")
    calories = data.get("calories")

    if activity == "Run" and distance:
        dist_mi = distance
        if distance_unit == "km":
            dist_mi = distance * 0.621371
        elif distance_unit == "m":
            dist_mi = distance * 0.000621371

        prior_longest = prior_prs.get("longest_run", {}).get("distance_mi")
        if prior_longest is None or dist_mi > prior_longest:
            new_prs.append({"label": "Longest Run", "value": f"{distance} {distance_unit}"})

        if duration_min and dist_mi > 0:
            pace = duration_min / dist_mi
            prior_pace = prior_prs.get("fastest_pace", {}).get("pace")
            if prior_pace is None or pace < prior_pace:
                new_prs.append({"label": "Fastest Pace", "value": f"{int(pace)}:{round((pace % 1) * 60):02d} /mi"})

    if calories:
        prior_calories = prior_prs.get("most_calories", {}).get("calories")
        if prior_calories is None or calories > prior_calories:
            new_prs.append({"label": "Most Calories", "value": f"{calories} cal"})

    prior_split_pace = prior_prs.get("fastest_split", {}).get("pace_seconds")
    best_new_split = None
    for split in data.get("splits", []):
        ps = split.get("pace_seconds")
        if ps and (best_new_split is None or ps < best_new_split):
            best_new_split = ps
    if best_new_split is not None and (prior_split_pace is None or best_new_split < prior_split_pace):
        new_prs.append({"label": "Fastest Split", "value": f"{best_new_split // 60}:{best_new_split % 60:02d} /mi"})

    return jsonify({"id": cur.lastrowid, "new_prs": new_prs})


@app.route("/cardio/edit/<int:session_id>", methods=["POST"])
@login_required
def edit_cardio(session_id):
    data = request.json
    with db() as conn:
        add_cardio_activity(conn, data.get("activity"))
        conn.execute(
            "UPDATE cardio_sessions SET date=?, activity=?, duration_min=?, distance=?, distance_unit=?, calories=?, incline=?, speed=?, notes=?, elevation_gain=?, avg_heart_rate=?, pool_length=?, pool_length_unit=?, laps=?, elapsed_time_min=? WHERE id=?",
            (
                data["date"],
                data["activity"],
                data.get("duration_min") or None,
                data.get("distance") or None,
                data.get("distance_unit", "mi"),
                data.get("calories") or None,
                data.get("incline") or None,
                data.get("speed") or None,
                data.get("notes", ""),
                data.get("elevation_gain") or None,
                data.get("avg_heart_rate") or None,
                data.get("pool_length") or None,
                data.get("pool_length_unit", "yd"),
                data.get("laps") or None,
                data.get("elapsed_time_min") or None,
                session_id,
            ),
        )
        conn.execute("DELETE FROM cardio_splits WHERE session_id=?", (session_id,))
        for split in data.get("splits", []):
            conn.execute(
                "INSERT INTO cardio_splits (session_id, mile_number, pace_seconds, peak_hr_zone, stroke_type, distance_per_stroke, split_distance) VALUES (?,?,?,?,?,?,?)",
                (session_id, split["mile"], split.get("pace_seconds"), split.get("peak_hr_zone"), split.get("stroke_type"), split.get("distance_per_stroke"), split.get("split_distance", 1.0)),
            )
    return jsonify({"ok": True})


@app.route("/cardio/delete/<int:session_id>", methods=["POST"])
@login_required
def delete_cardio(session_id):
    with db() as conn:
        conn.execute("DELETE FROM cardio_sessions WHERE id=?", (session_id,))
    return jsonify({"ok": True})


@app.route("/exercises")
@login_required
def exercises():
    with db() as conn:
        bank = conn.execute(
            "SELECT * FROM exercise_bank ORDER BY muscle_group, name"
        ).fetchall()
        stats = conn.execute("""
            SELECT s.exercise_name, MAX(s.weight) as max_weight, s.unit,
                   COUNT(s.id) as total_sets, MAX(w.date) as last_date
            FROM sets s JOIN workouts w ON w.id = s.workout_id
            WHERE s.weight IS NOT NULL
            GROUP BY s.exercise_name, s.unit
        """).fetchall()
    stats_map = {r["exercise_name"]: dict(r) for r in stats}
    # Group by muscle group
    groups = {}
    for ex in bank:
        mg = ex["muscle_group"] or "Other"
        groups.setdefault(mg, []).append({"name": ex["name"], "id": ex["id"], "stats": stats_map.get(ex["name"])})
    ordered = {mg: groups[mg] for mg in MUSCLE_GROUPS if mg in groups}
    for mg in groups:
        if mg not in ordered:
            ordered[mg] = groups[mg]

    with db() as conn:
        cardio_stats = conn.execute("""
            SELECT activity, COUNT(*) as sessions, SUM(duration_min) as total_min, MAX(date) as last_date
            FROM cardio_sessions GROUP BY activity
        """).fetchall()
    cardio_stats_map = {r["activity"]: dict(r) for r in cardio_stats}
    cardio_groups = {
        category: [{"name": n, "stats": cardio_stats_map.get(n)} for n in names]
        for category, names in get_cardio_activities().items()
    }
    return render_template(
        "exercises.html", groups=ordered, muscle_groups=MUSCLE_GROUPS,
        cardio_groups=cardio_groups, cardio_categories=CARDIO_CATEGORIES,
    )


@app.route("/exercises/add", methods=["POST"])
@login_required
def add_exercise():
    data = request.json
    name = data.get("name", "").strip()
    muscle = data.get("muscle_group", "Other")
    if not name:
        return jsonify({"error": "Name required"}), 400
    with db() as conn:
        conn.execute(
            "INSERT INTO exercise_bank (name, muscle_group) VALUES (?, ?) ON CONFLICT (name) DO NOTHING"
            if not DATABASE_URL else
            "INSERT INTO exercise_bank (name, muscle_group) VALUES (%s, %s) ON CONFLICT (name) DO NOTHING",
            (name, muscle)
        )
    return jsonify({"ok": True})


@app.route("/api/exercise-bank")
@login_required
def api_exercise_bank():
    with db() as conn:
        rows = conn.execute(
            "SELECT name, muscle_group FROM exercise_bank ORDER BY muscle_group, name"
        ).fetchall()
    return jsonify([dict(r) for r in rows])


@app.route("/stats")
@login_required
def stats():
    ph = "%s" if DATABASE_URL else "?"
    with db() as conn:
        # --- Totals ---
        total_workouts = conn.execute("SELECT COUNT(*) as c FROM workouts").fetchone()["c"]
        total_sets     = conn.execute("SELECT COUNT(*) as c FROM sets").fetchone()["c"]
        total_reps     = conn.execute("SELECT COALESCE(SUM(reps),0) as c FROM sets WHERE reps IS NOT NULL").fetchone()["c"]
        total_volume   = conn.execute("SELECT COALESCE(SUM(weight*reps),0) as v FROM sets WHERE weight>0 AND reps IS NOT NULL").fetchone()["v"]
        total_cardio   = conn.execute("SELECT COUNT(*) as c FROM cardio_sessions").fetchone()["c"]
        cardio_mins    = conn.execute("SELECT COALESCE(SUM(duration_min),0) as m FROM cardio_sessions WHERE duration_min IS NOT NULL").fetchone()["m"]

        # --- Period stats (this week / month / year) ---
        if DATABASE_URL:
            _period_q = """
                SELECT
                    COALESCE(SUM(CASE WHEN w.date::date >= (date_trunc('week', CURRENT_DATE + INTERVAL '1 day') - INTERVAL '1 day') THEN s.weight*s.reps ELSE 0 END),0) as vol_week,
                    COALESCE(SUM(CASE WHEN w.date::date >= date_trunc('month', CURRENT_DATE) THEN s.weight*s.reps ELSE 0 END),0) as vol_month,
                    COALESCE(SUM(CASE WHEN w.date::date >= date_trunc('year', CURRENT_DATE) THEN s.weight*s.reps ELSE 0 END),0) as vol_year,
                    SUM(CASE WHEN w.date::date >= (date_trunc('week', CURRENT_DATE + INTERVAL '1 day') - INTERVAL '1 day') THEN 1 ELSE 0 END) as sets_week,
                    SUM(CASE WHEN w.date::date >= date_trunc('month', CURRENT_DATE) THEN 1 ELSE 0 END) as sets_month,
                    SUM(CASE WHEN w.date::date >= date_trunc('year', CURRENT_DATE) THEN 1 ELSE 0 END) as sets_year,
                    COALESCE(SUM(CASE WHEN w.date::date >= (date_trunc('week', CURRENT_DATE + INTERVAL '1 day') - INTERVAL '1 day') THEN s.reps ELSE 0 END),0) as reps_week,
                    COALESCE(SUM(CASE WHEN w.date::date >= date_trunc('month', CURRENT_DATE) THEN s.reps ELSE 0 END),0) as reps_month,
                    COALESCE(SUM(CASE WHEN w.date::date >= date_trunc('year', CURRENT_DATE) THEN s.reps ELSE 0 END),0) as reps_year
                FROM sets s JOIN workouts w ON s.workout_id = w.id
                WHERE s.weight > 0 AND s.reps IS NOT NULL
            """
        else:
            from datetime import date as _pd
            _today = _pd.today()
            _week_start  = week_start_of(_today).isoformat()
            _month_start = _today.replace(day=1).isoformat()
            _year_start  = _today.replace(month=1, day=1).isoformat()
            _period_q = f"""
                SELECT
                    COALESCE(SUM(CASE WHEN w.date >= '{_week_start}'  THEN s.weight*s.reps ELSE 0 END),0) as vol_week,
                    COALESCE(SUM(CASE WHEN w.date >= '{_month_start}' THEN s.weight*s.reps ELSE 0 END),0) as vol_month,
                    COALESCE(SUM(CASE WHEN w.date >= '{_year_start}'  THEN s.weight*s.reps ELSE 0 END),0) as vol_year,
                    SUM(CASE WHEN w.date >= '{_week_start}'  THEN 1 ELSE 0 END) as sets_week,
                    SUM(CASE WHEN w.date >= '{_month_start}' THEN 1 ELSE 0 END) as sets_month,
                    SUM(CASE WHEN w.date >= '{_year_start}'  THEN 1 ELSE 0 END) as sets_year,
                    COALESCE(SUM(CASE WHEN w.date >= '{_week_start}'  THEN s.reps ELSE 0 END),0) as reps_week,
                    COALESCE(SUM(CASE WHEN w.date >= '{_month_start}' THEN s.reps ELSE 0 END),0) as reps_month,
                    COALESCE(SUM(CASE WHEN w.date >= '{_year_start}'  THEN s.reps ELSE 0 END),0) as reps_year
                FROM sets s JOIN workouts w ON s.workout_id = w.id
                WHERE s.weight > 0 AND s.reps IS NOT NULL
            """
        pr = conn.execute(_period_q).fetchone()

        # --- Averages ---
        avg_sets_per_workout = conn.execute("""
            SELECT ROUND(AVG(sc),1) as a FROM (SELECT COUNT(*) as sc FROM sets GROUP BY workout_id)
        """).fetchone()["a"] or 0
        avg_reps_per_set = conn.execute("SELECT ROUND(AVG(reps),1) as a FROM sets WHERE reps IS NOT NULL").fetchone()["a"] or 0

        # --- Streak (consecutive weeks with ≥1 workout) ---
        week_rows = conn.execute(
            "SELECT DISTINCT to_char(date::date, 'IYYY-IW') as wk FROM workouts ORDER BY wk DESC"
            if DATABASE_URL else
            "SELECT DISTINCT strftime('%Y-%W', date) as wk FROM workouts ORDER BY wk DESC"
        ).fetchall()
        rest_week_starts_stats = {r["week_start"] for r in conn.execute("SELECT week_start FROM weekly_notes WHERE is_rest_week=1").fetchall()}
        streak_weeks = 0
        if week_rows:
            from datetime import date as _date
            fmt = "%G-%V" if DATABASE_URL else "%Y-%W"
            today_wk    = _date.today().strftime(fmt)
            prev_wk_str = (_date.today() - timedelta(days=7)).strftime(fmt)
            weeks = set(r["wk"] for r in week_rows)
            cur_week_d = _date.today()
            if cur_week_d.strftime(fmt) in weeks or (cur_week_d - timedelta(days=7)).strftime(fmt) in weeks:
                streak_weeks = 1
                i = 1
                while True:
                    check_d = _date.today() - timedelta(weeks=i)
                    check_ws = week_start_of(check_d).strftime("%Y-%m-%d")
                    if check_ws in rest_week_starts_stats:
                        streak_weeks += 1
                        i += 1
                        continue
                    if check_d.strftime(fmt) in weeks:
                        streak_weeks += 1
                        i += 1
                    else:
                        break

        # --- Days since last workout ---
        last_workout_date = conn.execute("SELECT MAX(date) as d FROM workouts").fetchone()["d"]
        days_since = None
        if last_workout_date:
            from datetime import date as _date2
            days_since = (_date2.today() - _date2.fromisoformat(last_workout_date)).days

        # --- Workout type breakdown ---
        type_rows = conn.execute("""
            SELECT name, COUNT(*) as c FROM workouts GROUP BY name ORDER BY c DESC
        """).fetchall()
        # Normalise to canonical types (exact match first, then substring)
        _exact_map = {
            "push": "Push", "pull": "Pull", "legs": "Legs",
            "chest and back": "Chest & Back", "chest & back": "Chest & Back",
            "shoulders & arms": "Shoulders & Arms", "arms and shoulders": "Shoulders & Arms",
            "shoulders and arms": "Shoulders & Arms",
        }
        _substr_rules = [
            ("push",               "Push"),
            ("pull",               "Pull"),
            ("chest",              "Chest & Back"),
            ("shoulder",           "Shoulders & Arms"),
            ("arms",               "Shoulders & Arms"),
            ("bis",                "Shoulders & Arms"),
            ("tris",               "Shoulders & Arms"),
            ("legs",               "Legs"),
        ]
        def _canonicalise(name):
            lo = name.lower()
            if lo in _exact_map:
                return _exact_map[lo]
            for substr, canon in _substr_rules:
                if substr in lo:
                    return canon
            return name
        type_counts = {}
        for r in type_rows:
            key = _canonicalise(r["name"])
            type_counts[key] = type_counts.get(key, 0) + r["c"]
        type_counts = sorted(type_counts.items(), key=lambda x: -x[1])

        # --- Top exercises by set count ---
        top_exercises = conn.execute("""
            SELECT exercise_name, COUNT(*) as c FROM sets GROUP BY exercise_name ORDER BY c DESC LIMIT 8
        """).fetchall()

        # --- Most volume by exercise (weight*reps) ---
        top_volume_exercises = conn.execute("""
            SELECT exercise_name, ROUND(SUM(weight*reps)) as v FROM sets
            WHERE weight>0 AND reps IS NOT NULL
            GROUP BY exercise_name ORDER BY v DESC LIMIT 8
        """).fetchall()

        # --- Workouts per month (last 12 months) ---
        monthly = conn.execute(
            "SELECT to_char(date::date, 'YYYY-MM') as mo, COUNT(*) as c FROM workouts WHERE date::date >= CURRENT_DATE - INTERVAL '12 months' GROUP BY mo ORDER BY mo"
            if DATABASE_URL else
            "SELECT strftime('%Y-%m', date) as mo, COUNT(*) as c FROM workouts WHERE date >= date('now', '-12 months') GROUP BY mo ORDER BY mo"
        ).fetchall()

        # --- Workouts per day of week ---
        dow_rows = conn.execute(
            "SELECT EXTRACT(DOW FROM date::date)::int as dow, COUNT(*) as c FROM workouts GROUP BY dow ORDER BY dow"
            if DATABASE_URL else
            "SELECT strftime('%w', date) as dow, COUNT(*) as c FROM workouts GROUP BY dow ORDER BY dow"
        ).fetchall()
        dow_labels = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]
        dow_counts = {str(int(r["dow"])): r["c"] for r in dow_rows}
        dow_data = [{"day": dow_labels[i], "count": dow_counts.get(str(i), 0)} for i in range(7)]

        # --- Longest gap between workouts ---
        workout_dates = conn.execute("SELECT date FROM workouts ORDER BY date ASC").fetchall()
        longest_gap = 0
        if len(workout_dates) > 1:
            from datetime import date as _d3
            for i in range(1, len(workout_dates)):
                gap = (_d3.fromisoformat(workout_dates[i]["date"]) - _d3.fromisoformat(workout_dates[i-1]["date"])).days
                if gap > longest_gap:
                    longest_gap = gap

        # --- Best estimated 1RMs (matches get_estimated_1rm() logic used by PRs page) ---
        e1rm_rows = conn.execute("""
            SELECT exercise_name,
                   MAX(CASE WHEN reps = 1 THEN weight ELSE weight*(1+reps/30.0) END) as e1rm,
                   MIN(unit) as unit
            FROM sets WHERE reps BETWEEN 1 AND 12 AND weight > 0 AND completed = 1
            GROUP BY exercise_name ORDER BY e1rm DESC LIMIT 8
        """).fetchall()

        # --- Muscle group volume split ---
        mg_rows = conn.execute("""
            SELECT eb.muscle_group, ROUND(SUM(s.weight*s.reps)) as v
            FROM sets s JOIN exercise_bank eb ON s.exercise_name = eb.name
            WHERE s.weight>0 AND s.reps IS NOT NULL
            GROUP BY eb.muscle_group ORDER BY v DESC
        """).fetchall()

        # --- Workout counts by period ---
        if DATABASE_URL:
            wc = conn.execute("""
                SELECT
                    SUM(CASE WHEN date::date >= (date_trunc('week', CURRENT_DATE + INTERVAL '1 day') - INTERVAL '1 day') THEN 1 ELSE 0 END) as week,
                    SUM(CASE WHEN date::date >= date_trunc('month', CURRENT_DATE) THEN 1 ELSE 0 END) as month,
                    SUM(CASE WHEN date::date >= date_trunc('year', CURRENT_DATE) THEN 1 ELSE 0 END) as year
                FROM workouts
            """).fetchone()
        else:
            from datetime import date as _wcd
            _tw = (_wcd.today() - __import__('datetime').timedelta(days=_wcd.today().weekday())).isoformat()
            _tm = _wcd.today().replace(day=1).isoformat()
            _ty = _wcd.today().replace(month=1, day=1).isoformat()
            wc = conn.execute(f"""
                SELECT
                    SUM(CASE WHEN date >= '{_tw}' THEN 1 ELSE 0 END) as week,
                    SUM(CASE WHEN date >= '{_tm}' THEN 1 ELSE 0 END) as month,
                    SUM(CASE WHEN date >= '{_ty}' THEN 1 ELSE 0 END) as year
                FROM workouts
            """).fetchone()

        # --- Biggest workout by volume ---
        biggest_workout = conn.execute("""
            SELECT w.date, w.name, ROUND(SUM(s.weight*s.reps)) as vol
            FROM sets s JOIN workouts w ON s.workout_id=w.id
            WHERE s.weight>0 AND s.reps IS NOT NULL
            GROUP BY w.id, w.date, w.name ORDER BY vol DESC LIMIT 1
        """).fetchone()

        # --- Most improved lift (first 30 days e1rm vs last 30 days e1rm) ---
        improved_lift = None
        lift_rows = conn.execute("""
            SELECT exercise_name FROM sets
            WHERE weight>0 AND reps BETWEEN 1 AND 12
            GROUP BY exercise_name HAVING COUNT(DISTINCT workout_id) >= 4
        """).fetchall()
        best_improvement = 0
        for row in lift_rows:
            ex = row["exercise_name"]
            if DATABASE_URL:
                first_e = conn.execute(
                    "SELECT MAX(weight*(1+reps/30.0)) as e FROM sets s JOIN workouts w ON s.workout_id=w.id WHERE s.exercise_name=%s AND s.weight>0 AND s.reps BETWEEN 1 AND 12 AND w.date::date <= (SELECT MIN(date::date) FROM workouts WHERE id IN (SELECT workout_id FROM sets WHERE exercise_name=%s)) + INTERVAL '45 days'",
                    (ex, ex)
                ).fetchone()
                last_e = conn.execute(
                    "SELECT MAX(weight*(1+reps/30.0)) as e FROM sets s JOIN workouts w ON s.workout_id=w.id WHERE s.exercise_name=%s AND s.weight>0 AND s.reps BETWEEN 1 AND 12 AND w.date::date >= CURRENT_DATE - INTERVAL '45 days'",
                    (ex,)
                ).fetchone()
            else:
                first_e = conn.execute(
                    "SELECT MAX(weight*(1+reps/30.0)) as e FROM sets s JOIN workouts w ON s.workout_id=w.id WHERE s.exercise_name=? AND s.weight>0 AND s.reps BETWEEN 1 AND 12 AND w.date <= date((SELECT MIN(date) FROM workouts WHERE id IN (SELECT workout_id FROM sets WHERE exercise_name=?)), '+45 days')",
                    (ex, ex)
                ).fetchone()
                last_e = conn.execute(
                    "SELECT MAX(weight*(1+reps/30.0)) as e FROM sets s JOIN workouts w ON s.workout_id=w.id WHERE s.exercise_name=? AND s.weight>0 AND s.reps BETWEEN 1 AND 12 AND w.date >= date('now', '-45 days')",
                    (ex,)
                ).fetchone()
            fe = first_e["e"] if first_e and first_e["e"] else None
            le = last_e["e"] if last_e and last_e["e"] else None
            if fe and le and fe > 0:
                pct = ((le - fe) / fe) * 100
                if pct > best_improvement:
                    best_improvement = pct
                    improved_lift = {"name": ex, "pct": round(pct, 1), "from": round(fe, 1), "to": round(le, 1)}

        # --- Avg weight trend per exercise (last 30d vs prev 30d) ---
        trend_rows = conn.execute("""
            SELECT exercise_name, COUNT(*) as c FROM sets WHERE weight>0 GROUP BY exercise_name HAVING COUNT(*) >= 10 ORDER BY COUNT(*) DESC LIMIT 12
        """).fetchall()
        trends = []
        for row in trend_rows:
            ex = row["exercise_name"]
            if DATABASE_URL:
                r = conn.execute(
                    "SELECT AVG(CASE WHEN w.date::date >= CURRENT_DATE-30 THEN s.weight END) as recent, AVG(CASE WHEN w.date::date BETWEEN CURRENT_DATE-60 AND CURRENT_DATE-31 THEN s.weight END) as prev FROM sets s JOIN workouts w ON s.workout_id=w.id WHERE s.exercise_name=%s AND s.weight>0",
                    (ex,)
                ).fetchone()
            else:
                r = conn.execute(
                    "SELECT AVG(CASE WHEN w.date >= date('now','-30 days') THEN s.weight END) as recent, AVG(CASE WHEN w.date BETWEEN date('now','-60 days') AND date('now','-31 days') THEN s.weight END) as prev FROM sets s JOIN workouts w ON s.workout_id=w.id WHERE s.exercise_name=? AND s.weight>0",
                    (ex,)
                ).fetchone()
            recent = r["recent"]
            prev   = r["prev"]
            if recent and prev and prev > 0:
                delta_pct = ((recent - prev) / prev) * 100
                trends.append({"name": ex, "recent": round(recent, 1), "prev": round(prev, 1), "delta": round(delta_pct, 1)})
        trends.sort(key=lambda x: -abs(x["delta"]))

        # --- PRs set this year / this month ---
        if DATABASE_URL:
            pr_counts = conn.execute("""
                SELECT
                    SUM(CASE WHEN w.date::date >= date_trunc('month', CURRENT_DATE) THEN 1 ELSE 0 END) as this_month,
                    SUM(CASE WHEN w.date::date >= date_trunc('year', CURRENT_DATE) THEN 1 ELSE 0 END) as this_year
                FROM (
                    SELECT s.exercise_name, s.weight, s.reps, w.date,
                           MAX(s.weight) OVER (PARTITION BY s.exercise_name ORDER BY w.date ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING) as prev_max
                    FROM sets s JOIN workouts w ON s.workout_id=w.id
                    WHERE s.weight>0 AND s.reps IS NOT NULL
                ) sub JOIN workouts w ON w.date = sub.date
                WHERE sub.weight > COALESCE(sub.prev_max, 0)
            """).fetchone()
        else:
            from datetime import date as _prc
            _prc_month = _prc.today().replace(day=1).isoformat()
            _prc_year  = _prc.today().replace(month=1, day=1).isoformat()
            pr_counts = conn.execute(f"""
                SELECT
                    SUM(CASE WHEN w.date >= '{_prc_month}' THEN 1 ELSE 0 END) as this_month,
                    SUM(CASE WHEN w.date >= '{_prc_year}'  THEN 1 ELSE 0 END) as this_year
                FROM (
                    SELECT s.exercise_name, s.weight, s.reps, w.date,
                           MAX(s.weight) OVER (PARTITION BY s.exercise_name ORDER BY w.date ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING) as prev_max
                    FROM sets s JOIN workouts w ON s.workout_id=w.id
                    WHERE s.weight>0 AND s.reps IS NOT NULL
                ) sub JOIN workouts w ON w.date = sub.date
                WHERE sub.weight > COALESCE(sub.prev_max, 0)
            """).fetchone()

        # --- Sets per muscle group this week ---
        if DATABASE_URL:
            mg_week_rows = conn.execute("""
                SELECT eb.muscle_group, COUNT(*) as c
                FROM sets s
                JOIN workouts w ON s.workout_id=w.id
                JOIN exercise_bank eb ON s.exercise_name=eb.name
                WHERE w.date::date >= (date_trunc('week', CURRENT_DATE + INTERVAL '1 day') - INTERVAL '1 day')
                GROUP BY eb.muscle_group ORDER BY c DESC
            """).fetchall()
        else:
            from datetime import date as _mgd
            _mgw = (_mgd.today() - __import__('datetime').timedelta(days=_mgd.today().weekday())).isoformat()
            mg_week_rows = conn.execute(f"""
                SELECT eb.muscle_group, COUNT(*) as c
                FROM sets s
                JOIN workouts w ON s.workout_id=w.id
                JOIN exercise_bank eb ON s.exercise_name=eb.name
                WHERE w.date >= '{_mgw}'
                GROUP BY eb.muscle_group ORDER BY c DESC
            """).fetchall()

        # --- Most neglected muscle group (fewest sets in last 30 days, among trained groups) ---
        if DATABASE_URL:
            neglected_rows = conn.execute("""
                SELECT eb.muscle_group, COUNT(*) as c
                FROM sets s JOIN workouts w ON s.workout_id=w.id
                JOIN exercise_bank eb ON s.exercise_name=eb.name
                WHERE w.date::date >= CURRENT_DATE - INTERVAL '30 days'
                GROUP BY eb.muscle_group ORDER BY c ASC LIMIT 3
            """).fetchall()
        else:
            neglected_rows = conn.execute("""
                SELECT eb.muscle_group, COUNT(*) as c
                FROM sets s JOIN workouts w ON s.workout_id=w.id
                JOIN exercise_bank eb ON s.exercise_name=eb.name
                WHERE w.date >= date('now', '-30 days')
                GROUP BY eb.muscle_group ORDER BY c ASC LIMIT 3
            """).fetchall()

        # --- Total weight lifted milestone label ---
        def volume_label(v):
            if v > 1_000_000: return f"{v/1_000_000:.1f}M lbs"
            if v > 1_000:     return f"{v/1_000:.0f}K lbs"
            return f"{v:.0f} lbs"

        # --- Weekly volume by muscle group (last 8 weeks) ---
        from datetime import date as _mgvd
        mg_vol_weeks = []
        for i in range(7, -1, -1):
            ws = _mgvd.today() - timedelta(days=_mgvd.today().weekday() + i * 7)
            we = ws + timedelta(days=6)
            mg_vol_weeks.append({"start": ws.isoformat(), "end": we.isoformat(), "label": ws.strftime("%b %-d")})
        mg_vol_data = {}
        if not DATABASE_URL:
            mg_vol_rows = conn.execute("""
                SELECT eb.muscle_group, w.date, ROUND(SUM(s.weight * s.reps)) as v
                FROM sets s
                JOIN workouts w ON s.workout_id = w.id
                JOIN exercise_bank eb ON s.exercise_name = eb.name
                WHERE s.weight > 0 AND s.reps IS NOT NULL AND w.date >= ?
                GROUP BY eb.muscle_group, w.date
            """, (mg_vol_weeks[0]["start"],)).fetchall()
        else:
            mg_vol_rows = conn.execute("""
                SELECT eb.muscle_group, w.date, ROUND(SUM(s.weight * s.reps)) as v
                FROM sets s
                JOIN workouts w ON s.workout_id = w.id
                JOIN exercise_bank eb ON s.exercise_name = eb.name
                WHERE s.weight > 0 AND s.reps IS NOT NULL AND w.date >= %s
                GROUP BY eb.muscle_group, w.date
            """, (mg_vol_weeks[0]["start"],)).fetchall()
        for r in mg_vol_rows:
            mg = r["muscle_group"] or "Other"
            if mg not in mg_vol_data:
                mg_vol_data[mg] = [0] * 8
            for wi, wk in enumerate(mg_vol_weeks):
                if wk["start"] <= r["date"] <= wk["end"]:
                    mg_vol_data[mg][wi] += r["v"] or 0
                    break

        # --- Weekly avg workouts ---
        if total_workouts and last_workout_date:
            from datetime import date as _d4
            first_date = conn.execute("SELECT MIN(date) as d FROM workouts").fetchone()["d"]
            span_weeks = max(1, (_d4.today() - _d4.fromisoformat(first_date)).days / 7)
            avg_workouts_per_week = round(total_workouts / span_weeks, 1)
        else:
            avg_workouts_per_week = 0

    return render_template("stats.html",
        total_workouts=total_workouts,
        total_sets=total_sets,
        total_reps=int(total_reps),
        total_volume=total_volume,
        volume_label=volume_label(total_volume),
        total_cardio=total_cardio,
        cardio_mins=int(cardio_mins),
        avg_sets_per_workout=avg_sets_per_workout,
        avg_reps_per_set=avg_reps_per_set,
        streak_weeks=streak_weeks,
        days_since=days_since,
        type_counts=type_counts,
        top_exercises=top_exercises,
        top_volume_exercises=top_volume_exercises,
        monthly=monthly,
        dow_data=dow_data,
        longest_gap=longest_gap,
        e1rm_rows=e1rm_rows,
        mg_rows=mg_rows,
        avg_workouts_per_week=avg_workouts_per_week,
        pr=pr,
        volume_label_fn=volume_label,
        wc=wc,
        biggest_workout=biggest_workout,
        improved_lift=improved_lift,
        trends=trends,
        pr_counts=pr_counts,
        mg_week_rows=mg_week_rows,
        neglected_rows=neglected_rows,
        mg_vol_weeks=mg_vol_weeks,
        mg_vol_data=mg_vol_data,
    )


@app.route("/export")
@login_required
def export():
    HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    HEADER_FILL = PatternFill("solid", start_color="1E3A5F")
    HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ROW_FONT = Font(name="Arial", size=10)
    ALT_FILL = PatternFill("solid", start_color="F0F4FA")
    thin = Side(style="thin", color="CCCCCC")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_header(ws, headers, col_widths):
        ws.append(headers)
        for i, width in enumerate(col_widths, 1):
            cell = ws.cell(1, i)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGN
            cell.border = BORDER
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.row_dimensions[1].height = 22

    def style_rows(ws, start_row=2):
        for row_idx, row in enumerate(ws.iter_rows(min_row=start_row), start=start_row):
            fill = ALT_FILL if row_idx % 2 == 0 else None
            for cell in row:
                cell.font = ROW_FONT
                cell.border = BORDER
                cell.alignment = Alignment(vertical="center")
                if fill:
                    cell.fill = fill

    WORKOUT_HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    WORKOUT_HEADER_FILL = PatternFill("solid", start_color="1E3A5F")
    EXERCISE_HEADER_FONT = Font(name="Arial", bold=True, color="1E3A5F", size=10)
    EXERCISE_HEADER_FILL = PatternFill("solid", start_color="D6E4F0")

    wb = Workbook()

    # --- Workouts tab (grouped) ---
    ws_w = wb.active
    ws_w.title = "Workouts"

    # Set column widths
    col_widths = [18, 22, 8, 10, 10, 8, 8, 10, 8, 28]
    for i, w in enumerate(col_widths, 1):
        ws_w.column_dimensions[get_column_letter(i)].width = w

    with db() as conn:
        all_workouts = conn.execute("""
            SELECT w.id, w.date, w.name, w.notes, w.duration_min, w.rating, w.calories, w.avg_heart_rate
            FROM workouts w ORDER BY w.date DESC, w.id DESC
        """).fetchall()
        all_sets = conn.execute("""
            SELECT s.workout_id, s.exercise_name, s.set_number, s.reps, s.weight, s.unit, s.rpe, s.is_warmup, s.exercise_note, s.superset_group
            FROM sets s ORDER BY s.workout_id, s.id
        """).fetchall()

    # Group sets by workout
    from collections import defaultdict
    sets_by_workout = defaultdict(list)
    for s in all_sets:
        sets_by_workout[s["workout_id"]].append(s)

    for workout in all_workouts:
        # Workout header row
        row = ws_w.max_row + 1
        ws_w.cell(row, 1, workout["date"]).font = WORKOUT_HEADER_FONT
        ws_w.cell(row, 1).fill = WORKOUT_HEADER_FILL
        ws_w.cell(row, 1).alignment = Alignment(horizontal="left", vertical="center")
        name_cell = ws_w.cell(row, 2, workout["name"] or "")
        name_cell.font = WORKOUT_HEADER_FONT
        name_cell.fill = WORKOUT_HEADER_FILL
        name_cell.alignment = Alignment(horizontal="left", vertical="center")
        dur = workout["duration_min"]
        rating = workout["rating"]
        meta_parts = []
        if dur:
            meta_parts.append(f"{dur} min")
        if rating:
            meta_parts.append(f"{rating}/5 stars")
        if workout["calories"]:
            meta_parts.append(f"{workout['calories']} cal")
        if workout["avg_heart_rate"]:
            meta_parts.append(f"{workout['avg_heart_rate']} bpm avg")
        meta_cell = ws_w.cell(row, 3, " · ".join(meta_parts))
        meta_cell.font = Font(name="Arial", color="D6E4F0", size=9, italic=True)
        meta_cell.fill = WORKOUT_HEADER_FILL
        meta_cell.alignment = Alignment(horizontal="left", vertical="center")
        if workout["notes"]:
            ws_w.cell(row, 4, workout["notes"]).font = Font(name="Arial", color="D6E4F0", size=9, italic=True)
            ws_w.cell(row, 4).fill = WORKOUT_HEADER_FILL
        for col in range(4, 11):
            ws_w.cell(row, col).fill = WORKOUT_HEADER_FILL
        ws_w.row_dimensions[row].height = 20

        # Exercise column headers
        row += 1
        for col, label in enumerate(["Exercise", "Set #", "Reps", "Weight", "Unit", "RPE", "Warmup", "Superset", "Note"], 2):
            c = ws_w.cell(row, col, label)
            c.font = EXERCISE_HEADER_FONT
            c.fill = EXERCISE_HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BORDER
        ws_w.row_dimensions[row].height = 18

        # Set rows
        sets = sets_by_workout.get(workout["id"], [])
        for i, s in enumerate(sets):
            row += 1
            fill = ALT_FILL if i % 2 == 0 else None
            for col in range(1, 11):
                c = ws_w.cell(row, col)
                c.font = ROW_FONT
                c.border = BORDER
                c.alignment = Alignment(horizontal="center", vertical="center")
                if fill:
                    c.fill = fill
            ws_w.cell(row, 2, s["exercise_name"]).alignment = Alignment(horizontal="left", vertical="center")
            ws_w.cell(row, 3, s["set_number"])
            ws_w.cell(row, 4, s["reps"])
            ws_w.cell(row, 5, s["weight"])
            ws_w.cell(row, 6, s["unit"])
            ws_w.cell(row, 7, s["rpe"])
            ws_w.cell(row, 8, "Yes" if s["is_warmup"] else "")
            ws_w.cell(row, 9, s["superset_group"] or "")
            ws_w.cell(row, 10, s["exercise_note"] or "").alignment = Alignment(horizontal="left", vertical="center")

        if not sets:
            row += 1
            ws_w.cell(row, 2, "No sets logged").font = Font(name="Arial", color="999999", italic=True, size=10)

        # Blank spacer row between workouts
        ws_w.append([])

    ws_w.freeze_panes = "A1"

    # --- Weight tab ---
    ws_bw = wb.create_sheet("Body Weight")
    with db() as conn:
        weight_rows = conn.execute(
            "SELECT date, weight, unit FROM body_weight ORDER BY date DESC"
        ).fetchall()
    style_header(ws_bw, ["Date", "Weight", "Unit"], [14, 12, 10])
    for r in weight_rows:
        ws_bw.append([r["date"], r["weight"], r["unit"]])
    style_rows(ws_bw)
    ws_bw.freeze_panes = "A2"

    # --- Cardio tab ---
    ws_c = wb.create_sheet("Cardio")
    with db() as conn:
        cardio_rows = conn.execute(
            "SELECT date, activity, duration_min, distance, distance_unit, calories, "
            "elevation_gain, avg_heart_rate, incline, speed, pool_length, pool_length_unit, "
            "laps, elapsed_time_min, notes FROM cardio_sessions ORDER BY date DESC"
        ).fetchall()
    style_header(ws_c,
        ["Date", "Activity", "Duration (min)", "Distance", "Unit", "Calories",
         "Elevation Gain (ft)", "Avg HR (bpm)", "Incline (%)", "Speed (mph)",
         "Pool Length", "Pool Unit", "Laps/Runs/Holes", "Elapsed Time (min)", "Notes"],
        [14, 16, 14, 12, 8, 10, 14, 12, 12, 12, 12, 10, 14, 16, 28])
    for r in cardio_rows:
        ws_c.append([r["date"], r["activity"], r["duration_min"],
                     r["distance"], r["distance_unit"], r["calories"],
                     r["elevation_gain"], r["avg_heart_rate"], r["incline"], r["speed"],
                     r["pool_length"], r["pool_length_unit"], r["laps"], r["elapsed_time_min"],
                     r["notes"]])
    style_rows(ws_c)
    ws_c.freeze_panes = "A2"

    # --- Sleep tab ---
    ws_sl = wb.create_sheet("Sleep")
    with db() as conn:
        sleep_rows = conn.execute(
            "SELECT date, bedtime, wake_time, duration_min, quality, notes FROM sleep_logs ORDER BY date DESC"
        ).fetchall()
    style_header(ws_sl, ["Date", "Bedtime", "Wake Time", "Duration (min)", "Score", "Notes"], [14, 12, 12, 16, 10, 28])
    for r in sleep_rows:
        ws_sl.append([r["date"], r["bedtime"], r["wake_time"], r["duration_min"], r["quality"], r["notes"]])
    style_rows(ws_sl)
    ws_sl.freeze_panes = "A2"

    # --- Cardio Splits tab ---
    ws_sp = wb.create_sheet("Cardio Splits")
    with db() as conn:
        split_rows = conn.execute(
            "SELECT cs.date, cs.activity, sp.mile_number, sp.split_distance, sp.pace_seconds, sp.peak_hr_zone, sp.stroke_type, sp.distance_per_stroke "
            "FROM cardio_splits sp JOIN cardio_sessions cs ON sp.session_id = cs.id ORDER BY cs.date DESC, sp.mile_number"
        ).fetchall()
    style_header(ws_sp, ["Date", "Activity", "Split #", "Distance (mi)", "Pace (sec)", "HR Zone", "Stroke Type", "Distance/Stroke"], [14, 16, 10, 14, 12, 12, 14, 16])
    for r in split_rows:
        ws_sp.append([r["date"], r["activity"], r["mile_number"], r["split_distance"], r["pace_seconds"], r["peak_hr_zone"], r["stroke_type"], r["distance_per_stroke"]])
    style_rows(ws_sp)
    ws_sp.freeze_panes = "A2"

    # --- PRs tab ---
    ws_pr = wb.create_sheet("PRs")
    prs = get_prs()
    style_header(ws_pr, ["Exercise", "Weight", "Unit", "Reps", "Date"], [24, 12, 8, 8, 14])
    for name, pr in sorted(prs.items()):
        ws_pr.append([name, pr["weight"], pr["unit"], pr["reps"], pr["date"]])
    style_rows(ws_pr)
    ws_pr.freeze_panes = "A2"

    # --- Rest Days tab ---
    ws_rd = wb.create_sheet("Rest Days")
    with db() as conn:
        rest_days_rows = conn.execute("SELECT date, reason FROM rest_days ORDER BY date DESC").fetchall()
    style_header(ws_rd, ["Date", "Reason"], [16, 24])
    for r in rest_days_rows:
        ws_rd.append([r["date"], r["reason"] or ""])
    style_rows(ws_rd)
    ws_rd.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"workouts_{date.today().strftime('%Y-%m-%d')}.xlsx"
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=filename)


@app.route("/sleep")
@login_required
def sleep():
    with db() as conn:
        logs = conn.execute(
            "SELECT * FROM sleep_logs ORDER BY date DESC"
        ).fetchall()
    # compute stats
    durations = [r["duration_min"] for r in logs if r["duration_min"]]
    qualities = [r["quality"] for r in logs if r["quality"]]
    avg_duration = round(sum(durations) / len(durations)) if durations else None
    avg_quality = round(sum(qualities) / len(qualities)) if qualities else None
    # weekly sleep debt vs nightly target, last 7 entries
    sleep_goal_min = setting_num(get_settings(), "sleep_goal_hours") * 60
    recent = [r["duration_min"] for r in logs[:7] if r["duration_min"]]
    weekly_debt = None
    if recent:
        weekly_debt = round((sleep_goal_min * len(recent) - sum(recent)) / 60, 1)
    # 30-day trend (last 30 logs)
    chart_logs = [r for r in logs[:30] if r["duration_min"]][::-1]
    chart_labels = [r["date"] for r in chart_logs]
    chart_durations = [round(r["duration_min"] / 60, 2) for r in chart_logs]
    chart_qualities = [r["quality"] for r in chart_logs]

    # weekly breakdown (Mon-Sun)
    from datetime import datetime as _dt, timedelta as _td
    weeks = {}
    for r in logs:
        d = _dt.strptime(r["date"], "%Y-%m-%d").date()
        week_start = d - _td(days=d.weekday())
        weeks.setdefault(week_start, []).append(r)
    weekly_stats = []
    for week_start in sorted(weeks.keys(), reverse=True)[:1]:
        entries = weeks[week_start]
        durs = [e["duration_min"] for e in entries if e["duration_min"]]
        quals = [e["quality"] for e in entries if e["quality"]]
        weekly_stats.append({
            "week_start": week_start.strftime("%Y-%m-%d"),
            "week_end": (week_start + _td(days=6)).strftime("%Y-%m-%d"),
            "nights": len(entries),
            "avg_duration": round(sum(durs) / len(durs)) if durs else None,
            "avg_quality": round(sum(quals) / len(quals)) if quals else None,
            "debt": round((sleep_goal_min * len(durs) - sum(durs)) / 60, 1) if durs else None,
        })

    quality_logs = [r for r in logs if r["quality"]]
    duration_logs = [r for r in logs if r["duration_min"]]
    best_quality_log = max(quality_logs, key=lambda r: r["quality"]) if quality_logs else None
    longest_duration_log = max(duration_logs, key=lambda r: r["duration_min"]) if duration_logs else None

    return render_template(
        "sleep.html",
        logs=logs,
        avg_duration=avg_duration,
        avg_quality=avg_quality,
        weekly_debt=weekly_debt,
        chart_labels=chart_labels,
        chart_durations=chart_durations,
        chart_qualities=chart_qualities,
        weekly_stats=weekly_stats,
        best_quality_log=best_quality_log,
        longest_duration_log=longest_duration_log,
    )


@app.route("/sleep/add", methods=["POST"])
@login_required
def add_sleep():
    data = request.json
    with db() as conn:
        conn.execute(
            "INSERT INTO sleep_logs (date, duration_min, quality) VALUES (?,?,?)",
            (
                data["date"],
                data.get("duration_min") or None,
                data.get("quality") or None,
            ),
        )
    return jsonify({"ok": True})



@app.route("/sleep/edit/<int:log_id>", methods=["POST"])
@login_required
def edit_sleep(log_id):
    data = request.json
    with db() as conn:
        conn.execute(
            "UPDATE sleep_logs SET date=?, duration_min=?, quality=? WHERE id=?",
            (data["date"], data.get("duration_min") or None, data.get("quality") or None, log_id),
        )
    return jsonify({"ok": True})


@app.route("/sleep/delete/<int:log_id>", methods=["POST"])
@login_required
def delete_sleep(log_id):
    with db() as conn:
        conn.execute("DELETE FROM sleep_logs WHERE id=?", (log_id,))
    return jsonify({"ok": True})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)
