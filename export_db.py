import sqlite3
import json
import os

DB_PATH = 'reference_validator.db'
JSON_OUTPUT = 'db_export.json'
SQL_OUTPUT = 'postgres_dump.sql'

def export_data():
    if not os.path.exists(DB_PATH):
        print(f"Database not found at {DB_PATH}")
        return

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    # Get all tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = [row['name'] for row in cursor.fetchall() if row['name'] != 'sqlite_sequence']

    data = {}
    
    with open(SQL_OUTPUT, 'w', encoding='utf-8') as sql_file:
        for table in tables:
            print(f"Exporting table: {table}")
            cursor.execute(f"SELECT * FROM {table}")
            rows = cursor.fetchall()
            
            # Serialize for JSON
            table_data = []
            for row in rows:
                table_data.append(dict(row))
            data[table] = table_data

            # Generate SQL INSERTs for Postgres
            if not rows:
                continue
                
            columns = rows[0].keys()
            col_str = ", ".join(columns)
            
            for row in rows:
                vals = []
                for key in columns:
                    val = row[key]
                    if val is None:
                        vals.append("NULL")
                    elif key == 'is_admin': # Handle boolean specifically
                        vals.append('TRUE' if val else 'FALSE')
                    elif isinstance(val, (int, float)):
                        vals.append(str(val))
                    else:
                        # Escape single quotes
                        cleaned = str(val).replace("'", "''")
                        vals.append(f"'{cleaned}'")
                
                val_str = ", ".join(vals)
                # Use ON CONFLICT DO NOTHING to handle duplicates (like admin user) gracefully
                sql = f"INSERT INTO {table} ({col_str}) VALUES ({val_str}) ON CONFLICT DO NOTHING;\n"
                sql_file.write(sql)

    conn.close()

    with open(JSON_OUTPUT, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, default=str)

    print(f"✅ Data exported to {JSON_OUTPUT}")
    print(f"✅ SQL dump created at {SQL_OUTPUT}")

if __name__ == "__main__":
    export_data()
