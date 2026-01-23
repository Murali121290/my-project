import os
import psycopg2
from werkzeug.security import generate_password_hash

# Configuration matching docker-compose.yml defaults
DB_HOST = os.environ.get('DB_HOST', 'postgres')
DB_NAME = os.environ.get('DB_NAME', 'hub_db')
DB_USER = os.environ.get('DB_USER', 'hub_user')
DB_PASSWORD = os.environ.get('DB_PASSWORD', 'hub_secure_pass')

def reset_admin():
    print(f"Connecting to PostgreSQL at {DB_HOST}...")
    try:
        conn = psycopg2.connect(
            host=DB_HOST,
            database=DB_NAME,
            user=DB_USER,
            password=DB_PASSWORD
        )
        cur = conn.cursor()
        
        username = 'admin'
        password = 'admin123'
        # Generate hash compatible with the app
        hashed = generate_password_hash(password, method='pbkdf2:sha256')
        
        print(f"Checking for user '{username}'...")
        cur.execute("SELECT id FROM users WHERE username = %s", (username,))
        user = cur.fetchone()
        
        if user:
            print(f"User found (ID: {user[0]}). Updating password and ensuring admin privileges...")
            cur.execute("UPDATE users SET password = %s, is_admin = TRUE WHERE id = %s", (hashed, user[0]))
        else:
            print("User not found. Creating default admin user...")
            # Assuming 'role' column exists based on dumps, defaulting to 'ADMIN'
            cur.execute("INSERT INTO users (username, password, email, is_admin, role) VALUES (%s, %s, %s, %s, %s)",
                        (username, hashed, 'admin@example.com', True, 'ADMIN'))
        
        conn.commit()
        print("-" * 30)
        print("SUCCESS: Admin credentials updated.")
        print(f"Username: {username}")
        print(f"Password: {password}")
        print("-" * 30)
        
    except psycopg2.Error as e:
        print(f"Database Error: {e}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if 'conn' in locals() and conn:
            conn.close()

if __name__ == "__main__":
    reset_admin()
