from sqlalchemy.orm import Session
from database import SessionLocal, engine
from models import MadCategory, Register, PaymentMode, Base

# Create tables
Base.metadata.create_all(bind=engine)

MADS = [
    "Atiya",
    "Ambulance",
    "Charm",
    "Commercial Box",
    "Fitra",
    "Ijtemai Shadiyan",
    "Isale Sawab",
    "Residential Box",
    "Sadqa",
    "Special Donation",
    "Zakat",
]

REGISTERS = [
    "Monthly Trust Members Contribution",
    "Monthly Members Contribution",
    "Commercial Donation Box Collection",
    "Residential Donation Box Collection",
    "Bakrid Charm and Cash Collection",
    "Special Donation",
    "Ramzan Collection",
    "Ijtemai Shadiyan",
    "Ambulance Scheme"
]

PAYMENT_MODES = [
    "Cash",
    "BharatPe",
    "Axis Bank",
    "DCC Bank"
]

def seed_data():
    db = SessionLocal()
    
    try:
        # Check if data already exists
        if db.query(MadCategory).count() > 0:
            print("Data already seeded!")
            return
        
        # Seed Mad Categories
        for mad_name in MADS:
            mad = MadCategory(name=mad_name)
            db.add(mad)
        
        # Seed Payment Modes
        for mode_name in PAYMENT_MODES:
            mode = PaymentMode(name=mode_name)
            db.add(mode)
        
        # Seed Registers
        for register_name in REGISTERS:
            register = Register(name=register_name)
            db.add(register)
        
        db.commit()
        print("Data seeded successfully!")
        
    except Exception as e:
        print(f"Error seeding data: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    seed_data()