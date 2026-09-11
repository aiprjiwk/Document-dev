import sqlite3
import hashlib
import uuid
import datetime
import os
import logging

logger = logging.getLogger("AuthService")

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "database", "ocr_system.db")

def get_db_connection():
    """Establish connection to SQLite database."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def hash_password(password: str, salt: str = None) -> tuple:
    """Generate SHA-256 hash for password with a unique salt."""
    if not salt:
        salt = uuid.uuid4().hex
    # Append salt to password and hash
    hashed = hashlib.sha256((password + salt).encode('utf-8')).hexdigest()
    return hashed, salt

def initialize_auth_db():
    """Initialize users table and seed default admin_master if not exists."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Create Users table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        salt TEXT NOT NULL,
        role TEXT NOT NULL,          -- 'admin_master', 'admin_department', 'operator_department'
        department TEXT NOT NULL,    -- User's department
        is_approved INTEGER NOT NULL DEFAULT 0, -- 0: Pending, 1: Approved, -1: Suspended
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        approved_by TEXT,
        approved_at TEXT
    )
    """)
    conn.commit()
    
    # 2. Check if admin_master exists, if not seed default 'admin'
    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin_master'")
    admin_count = cursor.fetchone()[0]
    
    if admin_count == 0:
        admin_username = "admin"
        admin_password = "admin123"
        hashed, salt = hash_password(admin_password)
        try:
            cursor.execute("""
            INSERT INTO users (username, password_hash, salt, role, department, is_approved)
            VALUES (?, ?, ?, 'admin_master', 'IT', 1)
            """, (admin_username, hashed, salt))
            conn.commit()
            logger.info("Seeded default admin_master user successfully.")
        except Exception as e:
            logger.error(f"Failed to seed default admin_master: {e}")
            
    conn.close()

def register_user(username: str, password: str, role: str, department: str) -> tuple:
    """Register a new user with pending status."""
    username = username.strip()
    if not username or not password or not role or not department:
        return False, "All fields are required."
        
    if role not in ["admin_department", "operator_department"]:
        return False, "Invalid role choice."
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        # Check if username exists
        cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
        if cursor.fetchone():
            conn.close()
            return False, f"Username '{username}' is already taken."
            
        # Hash password and save
        hashed, salt = hash_password(password)
        cursor.execute("""
        INSERT INTO users (username, password_hash, salt, role, department, is_approved)
        VALUES (?, ?, ?, ?, ?, 0)
        """, (username, hashed, salt, role, department))
        conn.commit()
        conn.close()
        return True, "Registration successful! Please wait for Admin Master approval."
    except Exception as e:
        conn.close()
        return False, f"Database Error: {str(e)}"

def verify_login(username: str, password: str) -> tuple:
    """Verify user login credentials and return status details."""
    username = username.strip()
    if not username or not password:
        return False, None, "Username and password cannot be empty."
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("SELECT * FROM users WHERE username = ?", (username,))
        user_row = cursor.fetchone()
        conn.close()
        
        if not user_row:
            return False, None, "Invalid username or password."
            
        user = dict(user_row)
        
        # Verify hash
        test_hash, _ = hash_password(password, user["salt"])
        if test_hash != user["password_hash"]:
            return False, None, "Invalid username or password."
            
        # Check approval status
        if user["is_approved"] == 0:
            return False, None, "Your account is pending approval by Admin Master."
        elif user["is_approved"] == -1:
            return False, None, "Your account has been suspended."
            
        return True, user, "Login successful."
    except Exception as e:
        if conn:
            conn.close()
        return False, None, f"Database Error: {str(e)}"

def get_all_users() -> list:
    """Retrieve all registered users sorted by role and creation date."""
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, username, role, department, is_approved, created_at, approved_by, approved_at FROM users ORDER BY is_approved ASC, created_at DESC")
    users = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return users

def update_user_status(user_id: int, status: int, approved_by: str) -> tuple:
    """Approve, suspend, or delete users."""
    if status not in [1, 0, -1]:
        return False, "Invalid status code."
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        # Check if user is admin_master (protect master admin from suspension/deletion)
        cursor.execute("SELECT role FROM users WHERE id = ?", (user_id,))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return False, "User not found."
            
        if row["role"] == "admin_master":
            conn.close()
            return False, "Admin Master role cannot be suspended, modified or deleted."
            
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("""
        UPDATE users 
        SET is_approved = ?, approved_by = ?, approved_at = ?
        WHERE id = ?
        """, (status, approved_by, now_str, user_id))
        conn.commit()
        conn.close()
        return True, "User status updated successfully."
    except Exception as e:
        conn.close()
        return False, f"Database Error: {str(e)}"

def delete_user(user_id: int) -> tuple:
    """Delete a user permanently (Admin Master only)."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT role FROM users WHERE id = ?", (user_id,))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return False, "User not found."
            
        if row["role"] == "admin_master":
            conn.close()
            return False, "Admin Master cannot be deleted."
            
        cursor.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
        conn.close()
        return True, "User deleted successfully."
    except Exception as e:
        conn.close()
        return False, f"Database Error: {str(e)}"
