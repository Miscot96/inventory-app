import sqlite3
import os
import datetime
from auth import hash_password

DB_PATH = os.environ.get("INVENTORY_DB_PATH", "inventory.db")


def get_conn():
    dbdir = os.path.dirname(DB_PATH)
    if dbdir and not os.path.exists(dbdir):
        os.makedirs(dbdir, exist_ok=True)

    conn = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    conn.execute("PRAGMA busy_timeout=5000;")
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        salt BLOB NOT NULL,
        pwd_hash BLOB NOT NULL,
        is_admin INTEGER NOT NULL DEFAULT 0,
        is_main_admin INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS branches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        description TEXT
    );

    CREATE TABLE IF NOT EXISTS materials (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        branch_id INTEGER NOT NULL,
        reference TEXT NOT NULL,
        description TEXT,
        UNIQUE(branch_id, reference),
        FOREIGN KEY(branch_id) REFERENCES branches(id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS material_colors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        material_id INTEGER NOT NULL,
        color TEXT NOT NULL,
        unit TEXT NOT NULL DEFAULT 'm',
        available_qty REAL NOT NULL DEFAULT 0,
        FOREIGN KEY(material_id) REFERENCES materials(id) ON DELETE CASCADE,
        UNIQUE(material_id, color)
    );

    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        branch_id INTEGER NOT NULL,
        name TEXT,
        product_name TEXT,
        output_qty REAL NOT NULL DEFAULT 1,
        output_unit TEXT NOT NULL DEFAULT 'pcs',
        created_by INTEGER,
        created_at TEXT NOT NULL,
        notes TEXT,
        FOREIGN KEY(branch_id) REFERENCES branches(id) ON DELETE CASCADE,
        FOREIGN KEY(created_by) REFERENCES users(id) ON DELETE SET NULL
    );

    CREATE TABLE IF NOT EXISTS order_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        material_color_id INTEGER NOT NULL,
        qty_per_piece REAL NOT NULL DEFAULT 0,
        needed_qty REAL NOT NULL DEFAULT 0,
        FOREIGN KEY(order_id) REFERENCES orders(id) ON DELETE CASCADE,
        FOREIGN KEY(material_color_id) REFERENCES material_colors(id) ON DELETE CASCADE
    );
    """)

    conn.commit()

    # users upgrades
    cur.execute("PRAGMA table_info(users)")
    user_columns = [row[1] for row in cur.fetchall()]

    if "is_main_admin" not in user_columns:
        cur.execute("ALTER TABLE users ADD COLUMN is_main_admin INTEGER NOT NULL DEFAULT 0")
        conn.commit()

    # material_colors upgrades
    cur.execute("PRAGMA table_info(material_colors)")
    mc_columns = [row[1] for row in cur.fetchall()]

    if "unit" not in mc_columns:
        cur.execute("ALTER TABLE material_colors ADD COLUMN unit TEXT NOT NULL DEFAULT 'm'")
        conn.commit()

    # orders upgrades
    cur.execute("PRAGMA table_info(orders)")
    order_columns = [row[1] for row in cur.fetchall()]

    if "name" not in order_columns:
        cur.execute("ALTER TABLE orders ADD COLUMN name TEXT")
        conn.commit()

    if "product_name" not in order_columns:
        cur.execute("ALTER TABLE orders ADD COLUMN product_name TEXT")
        conn.commit()

    if "output_qty" not in order_columns:
        cur.execute("ALTER TABLE orders ADD COLUMN output_qty REAL NOT NULL DEFAULT 1")
        conn.commit()
        if "piece_count" in order_columns:
            cur.execute("UPDATE orders SET output_qty = piece_count WHERE output_qty = 1")
            conn.commit()

    if "output_unit" not in order_columns:
        cur.execute("ALTER TABLE orders ADD COLUMN output_unit TEXT NOT NULL DEFAULT 'pcs'")
        conn.commit()

    # order_items upgrades
    cur.execute("PRAGMA table_info(order_items)")
    oi_columns = [row[1] for row in cur.fetchall()]

    if "qty_per_piece" not in oi_columns:
        cur.execute("ALTER TABLE order_items ADD COLUMN qty_per_piece REAL NOT NULL DEFAULT 0")
        conn.commit()

    cur.execute("SELECT COUNT(*) FROM users;")
    if cur.fetchone()[0] == 0:
        default_user = "admin"
        default_pass = "admin"
        salt, pwd_hash = hash_password(default_pass)

        cur.execute(
            "INSERT INTO users (username, salt, pwd_hash, is_admin, is_main_admin, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (default_user, salt, pwd_hash, 1, 1, datetime.datetime.utcnow().isoformat())
        )
        conn.commit()
        print("Created default admin user: username='admin' password='admin' - change it immediately")
    else:
        cur.execute("SELECT COUNT(*) FROM users WHERE is_main_admin = 1")
        if cur.fetchone()[0] == 0:
            cur.execute("UPDATE users SET is_main_admin = 1 WHERE username = 'admin'")
            conn.commit()

    conn.close()