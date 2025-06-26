import sqlite3

DB_PATH = "planning_manager.db"

def get_db_connection():
    return sqlite3.connect(DB_PATH)

def setup_user_table():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL DEFAULT 'user'
            )
        """)
        conn.commit()

def insert_admin_user():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
                ("admin", "admin123", "admin")
            )
            conn.commit()
            print("Admin user inserted.")
        except sqlite3.IntegrityError:
            print("Admin user already exists.")

# get credentials
def check_credentials(username, password):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        print(f"Testing credentials: {username=} {password=}")  # debug
        cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username.strip(), password.strip()))
        user = cursor.fetchone()
        print(f"Found user: {user}")  # debug
        return user is not None  # True si user trouvé, sinon False
#Get role
def get_role(username):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT role FROM users WHERE username = ?", (username.strip(),))
        result = cursor.fetchone()
        if result:
            return result[0]
        return None


#Create user
def create_user(username, password, role="user"):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        try:
            cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (username, password, role))
            conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False

def setup_shifts_table():
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS shifts (
                id INTEGER PRIMARY KEY,
                name TEXT,
                max_load INTEGER NOT NULL
            )
        """)
        conn.commit()

def insert_default_shifts():
    default_shifts = [
        (0, "Matin", 240),
        (1, "Après-midi", 300),
        (2, "Soir", 180)
    ]
    with get_db_connection() as conn:
        cursor = conn.cursor()
        for sid, name, max_load in default_shifts:
            cursor.execute("""
                INSERT OR IGNORE INTO shifts (id, name, max_load) VALUES (?, ?, ?)
            """, (sid, name, max_load))
        conn.commit()

def get_shift_max_load(shift_id):
    with get_db_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT max_load FROM shifts WHERE id = ?", (shift_id,))
        result = cursor.fetchone()
        if result:
            return result[0]
        return None  # ou une valeur par défaut, ex: 240





if __name__ == "__main__":
    setup_user_table()
    setup_shifts_table()
    insert_default_shifts()