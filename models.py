from sqlalchemy import Column, Integer, String, Float, ForeignKey, Date
from sqlalchemy.orm import relationship
from database import Base

class MadCategory(Base):
    __tablename__ = "mad_categories"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    
    def __repr__(self):
        return f"<MadCategory(name='{self.name}')>"

class Register(Base):
    __tablename__ = "registers"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    
    def __repr__(self):
        return f"<Register(name='{self.name}')>"

class PaymentMode(Base):
    __tablename__ = "payment_modes"

    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    
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