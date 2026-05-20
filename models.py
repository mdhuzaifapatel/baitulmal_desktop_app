from sqlalchemy import Column, Integer, String, Float, ForeignKey, Date
from sqlalchemy.orm import relationship
from database import Base

class MadCategory(Base):
    __tablename__ = "mad_categories"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    initial_balance = Column(Float, default=0.0)
    
    def __repr__(self):
        return f"<MadCategory(name='{self.name}')>"

class Register(Base):
    __tablename__ = "registers"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    initial_balance = Column(Float, default=0.0)
    
    def __repr__(self):
        return f"<Register(name='{self.name}')>"

class PaymentMode(Base):
    __tablename__ = "payment_modes"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    initial_balance = Column(Float, default=0.0)
    
    def __repr__(self):
        return f"<PaymentMode(name='{self.name}')>"

class Receipt(Base):
    __tablename__ = "receipts"

    id = Column(Integer, primary_key=True)
    receipt_no = Column(String, nullable=False)
    date = Column(Date, nullable=False)
    name = Column(String, nullable=False)
    phone = Column(String, nullable=False)
    address = Column(String)
    amount = Column(Float, nullable=False)
    referred_by = Column(String)
    notes = Column(String)
    
    # Foreign keys
    mad_category_id = Column(Integer, ForeignKey("mad_categories.id"), nullable=False)
    register_id = Column(Integer, ForeignKey("registers.id"), nullable=False)
    payment_mode_id = Column(Integer, ForeignKey("payment_modes.id"), nullable=False)
    
    # Relationships
    mad_category = relationship("MadCategory")
    register = relationship("Register")
    payment_mode = relationship("PaymentMode")
    
class ContraEntry(Base):
    __tablename__ = "contra_entries"

    id = Column(Integer, primary_key=True)
    date = Column(Date, nullable=False)
    from_payment_mode_id = Column(Integer, ForeignKey("payment_modes.id"), nullable=False)
    to_payment_mode_id = Column(Integer, ForeignKey("payment_modes.id"), nullable=False)
    amount = Column(Float, nullable=False)
    notes = Column(String)
    
    # Relationships
    from_payment_mode = relationship("PaymentMode", foreign_keys=[from_payment_mode_id])
    to_payment_mode = relationship("PaymentMode", foreign_keys=[to_payment_mode_id])

    def __repr__(self):
        return f"<ContraEntry(amount={self.amount}, from={self.from_payment_mode_id}, to={self.to_payment_mode_id})>"

class ExpenseCategory(Base):
    __tablename__ = "expense_categories"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    
    def __repr__(self):
        return f"<ExpenseCategory(name='{self.name}')>"

class Expense(Base):
    __tablename__ = "expenses"

    id = Column(Integer, primary_key=True)
    voucher_no = Column(String, nullable=False)
    date = Column(Date, nullable=False)
    name = Column(String, nullable=False)
    phone = Column(String)
    amount = Column(Float, nullable=False)
    referred_by = Column(String)
    notes = Column(String)
    
    # Foreign keys
    mad_category_id = Column(Integer, ForeignKey("mad_categories.id"), nullable=False)
    expense_category_id = Column(Integer, ForeignKey("expense_categories.id"), nullable=False)
    payment_mode_id = Column(Integer, ForeignKey("payment_modes.id"), nullable=False)
    
    # Relationships
    mad_category = relationship("MadCategory")
    expense_category = relationship("ExpenseCategory")
    payment_mode = relationship("PaymentMode")

    def __repr__(self):
        return f"<Expense(voucher='{self.voucher_no}', amount={self.amount})>"

class MadInitialBalance(Base):
    __tablename__ = "mad_initial_balances"

    id = Column(Integer, primary_key=True)
    mad_category_id = Column(Integer, ForeignKey("mad_categories.id"), nullable=False)
    payment_mode_id = Column(Integer, ForeignKey("payment_modes.id"), nullable=False)
    amount = Column(Float, default=0.0)

    # Relationships
    mad_category = relationship("MadCategory")
    payment_mode = relationship("PaymentMode")

    def __repr__(self):
        return f"<MadInitialBalance(mad_id={self.mad_category_id}, payment_id={self.payment_mode_id}, amount={self.amount})>"

class Person(Base):
    __tablename__ = "people"

    id = Column(Integer, primary_key=True)
    name = Column(String, nullable=False)
    phone = Column(String)
    address = Column(String)
    person_type = Column(String, nullable=False) # 'Donor', 'Beneficiary', 'Staff'

    def __repr__(self):
        return f"<Person(name='{self.name}', type='{self.person_type}')>"