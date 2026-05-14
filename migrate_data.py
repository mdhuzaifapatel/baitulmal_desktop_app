import os
from sqlalchemy import create_engine, MetaData
from sqlalchemy.orm import sessionmaker
from models import Base
import database

# --- CONFIGURATION ---
# 1. First, set your DATABASE_URL environment variable to your Supabase URL
#    Example: set DATABASE_URL=postgresql://postgres:password@db.host.supabase.co:5432/postgres
# 2. Run this script: python migrate_data.py

LOCAL_DB_URL = "sqlite:///./baitulmal.db"
REMOTE_DB_URL = os.getenv("DATABASE_URL")

if not REMOTE_DB_URL or REMOTE_DB_URL.startswith("sqlite"):
    print("ERROR: Please set the DATABASE_URL environment variable to your Supabase URL.")
    print("Example (Windows): set DATABASE_URL=postgresql://user:pass@host:5432/postgres")
    exit(1)

# Fix for common postgres scheme issue
if REMOTE_DB_URL.startswith("postgres://"):
    REMOTE_DB_URL = REMOTE_DB_URL.replace("postgres://", "postgresql://", 1)

print(f"Connecting to Local SQLite...")
local_engine = create_engine(LOCAL_DB_URL)
LocalSession = sessionmaker(bind=local_engine)

print(f"Connecting to Remote Supabase...")
remote_engine = create_engine(REMOTE_DB_URL)
RemoteSession = sessionmaker(bind=remote_engine)

# Create tables on remote if they don't exist
print("Creating tables on Supabase...")
Base.metadata.create_all(bind=remote_engine)

def migrate():
    local_session = LocalSession()
    remote_session = RemoteSession()
    
    metadata = MetaData()
    metadata.reflect(bind=local_engine)
    
    # Order of migration to respect foreign keys
    # Adjust names based on your models.py
    tables = [
        "payment_modes", 
        "mads", 
        "registers", 
        "expense_categories",
        "receipts",
        "expenses",
        "contra_entries"
    ]
    
    try:
        for table_name in tables:
            if table_name not in metadata.tables:
                continue
                
            print(f"Migrating table: {table_name}...")
            table = metadata.tables[table_name]
            
            # Fetch all data from local
            rows = local_session.execute(table.select()).fetchall()
            
            if not rows:
                print(f"  - No data in {table_name}, skipping.")
                continue
            
            # Insert into remote
            # We convert rows to dicts for insertion
            data = [dict(row._mapping) for row in rows]
            
            # Use raw SQL to avoid model mismatches if any
            remote_session.execute(table.insert(), data)
            print(f"  - Successfully migrated {len(data)} rows.")
            
        remote_session.commit()
        print("\nSUCCESS: All data has been migrated to Supabase!")
        
    except Exception as e:
        remote_session.rollback()
        print(f"\nFAILED: Error during migration: {e}")
    finally:
        local_session.close()
        remote_session.close()

if __name__ == "__main__":
    confirm = input("This will push all data from baitulmal.db to Supabase. Continue? (y/n): ")
    if confirm.lower() == 'y':
        migrate()
    else:
        print("Cancelled.")
