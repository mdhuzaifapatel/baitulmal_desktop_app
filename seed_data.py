from sqlalchemy.orm import Session
from database import SessionLocal, engine
from models import MadCategory, Register, PaymentMode, ExpenseCategory, Base

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

EXPENSE_CATEGORIES = [
    "Advertisement expenses",
    "Ambulance service",
    "Bank charges",
    "Fixed assets",
    "Fuel expenses",
    "General Expenditure",
    "Imdad to needy",
    "Mass marriages",
    "Miscellaneous",
    "Monthly pension for needy",
    "Printing and stationary",
    "Ration kit",
    "Repair and maintenance",
    "Staff salary",
    "Staff welfare expenses",
    "Travel expenses",
    "Vehicle service",
    "Wages"
]

def seed_data():
    db = SessionLocal()
    
    try:
        # Seed Mad Categories
        if db.query(MadCategory).count() == 0:
            for mad_name in MADS:
                db.add(MadCategory(name=mad_name))
            print("Mad Categories seeded.")
        
        # Seed Payment Modes
        if db.query(PaymentMode).count() == 0:
            for mode_name in PAYMENT_MODES:
                db.add(PaymentMode(name=mode_name))
            print("Payment Modes seeded.")
        
        # Seed Registers
        if db.query(Register).count() == 0:
            for register_name in REGISTERS:
                db.add(Register(name=register_name))
            print("Registers seeded.")
            
        # Seed Expense Categories
        if db.query(ExpenseCategory).count() == 0:
            for exp_name in EXPENSE_CATEGORIES:
                db.add(ExpenseCategory(name=exp_name))
            print("Expense Categories seeded.")
        
        db.commit()
        print("Seed check complete.")
        
    except Exception as e:
        print(f"Error seeding data: {e}")
        db.rollback()
    finally:
        db.close()

if __name__ == "__main__":
    seed_data()