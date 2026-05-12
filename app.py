from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import RedirectResponse, HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.sessions import SessionMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import func, desc, and_
from datetime import datetime, date, timedelta
import shutil
import os
from jinja2 import Environment, FileSystemLoader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import tkinter as tk
from tkinter import filedialog
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def prompt_save_path(title, default_name, file_types):
    try:
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        filepath = filedialog.asksaveasfilename(
            title=title,
            initialfile=default_name,
            filetypes=file_types
        )
        root.destroy()
        return filepath
    except Exception:
        return None

def prompt_open_path(title, file_types):
    try:
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.withdraw()
        filepath = filedialog.askopenfilename(
            title=title,
            filetypes=file_types
        )
        root.destroy()
        return filepath
    except Exception:
        return None
from database import SessionLocal, engine
from models import *
from seed_data import *

# Create tables
Base.metadata.create_all(bind=engine)

app = FastAPI(title="Baitulmal Donation Management System")

# Add session middleware for authentication
app.add_middleware(SessionMiddleware, secret_key="baitulmal_secret_key_2024", max_age=3600)

# Default admin credentials (plain text)
DEFAULT_ADMIN_USERNAME = "admin"
DEFAULT_ADMIN_PASSWORD = "admin"

# Direct Jinja2 environment
env = Environment(loader=FileSystemLoader(resource_path("templates")), autoescape=True)

def get_financial_year_range(fy_year: int):
    """
    Get the date range for a given financial year.
    Financial Year: April 1 to March 31
    Example: FY 2024-2025 means April 1, 2024 to March 31, 2025
    """
    start_date = date(fy_year, 4, 1)
    end_date = date(fy_year + 1, 3, 31)
    return start_date, end_date

def get_current_financial_year():
    """
    Get the current financial year (returns the starting year).
    Example: If current date is May 2024, returns 2024
    If current date is February 2025, returns 2024
    """
    today = date.today()
    if today.month >= 4:
        return today.year
    else:
        return today.year - 1

def get_date_range_for_filter(filter_type: str, start_date_str: str = None, end_date_str: str = None, fy: int = None):
    if filter_type == "current_month":
        today = date.today()
        actual_start_date = date(today.year, today.month, 1)
        if today.month == 12:
            actual_end_date = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            actual_end_date = date(today.year, today.month + 1, 1) - timedelta(days=1)
        title_suffix = f"{today.strftime('%B %Y')}"
    elif filter_type == "manual" and start_date_str and end_date_str:
        actual_start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
        actual_end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
        title_suffix = f"{actual_start_date.strftime('%d-%m-%Y')} to {actual_end_date.strftime('%d-%m-%Y')}"
    else:
        # Default to current FY
        if fy is None:
            fy = get_current_financial_year()
        actual_start_date, actual_end_date = get_financial_year_range(fy)
        title_suffix = f"FY {fy}-{fy+1}"
    
    return actual_start_date, actual_end_date, title_suffix

def get_balances(db, start_date, end_date):
    """Calculate opening, period, and closing balances."""
    opening_balance = db.query(func.sum(Receipt.amount)).filter(Receipt.date < start_date).scalar() or 0
    period_total = db.query(func.sum(Receipt.amount)).filter(
        and_(Receipt.date >= start_date, Receipt.date <= end_date)
    ).scalar() or 0
    closing_balance = opening_balance + period_total
    return opening_balance, period_total, closing_balance

# Static files
app.mount("/static", StaticFiles(directory=resource_path("static")), name="static")

def get_db():
    db = SessionLocal()
    try:
        return db
    except:
        db.close()
        raise

def format_indian_currency(amount):
    """Formats a number in Indian currency style (e.g., 1,00,000.00)"""
    if amount is None or amount == "":
        return "0.00"
    
    try:
        amount = round(float(amount), 2)
    except (ValueError, TypeError):
        return "0.00"
        
    s = f"{amount:.2f}"
    parts = s.split('.')
    main = parts[0]
    decimal = parts[1]
    
    is_negative = main.startswith('-')
    if is_negative:
        main = main[1:]
        
    if len(main) <= 3:
        res = main
    else:
        last_three = main[-3:]
        remaining = main[:-3]
        groups = []
        while len(remaining) > 2:
            groups.append(remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            groups.append(remaining)
        res = ",".join(reversed(groups)) + "," + last_three
    
    final = f"{res}.{decimal}"
    return f"-{final}" if is_negative else final

# Register the filter in Jinja environment
env.filters['indian_currency'] = format_indian_currency
env.globals['now'] = datetime.now

def verify_password(plain_password, stored_password):
    return plain_password == stored_password

def is_authenticated(request: Request):
    return request.session.get("authenticated", False)

def require_auth(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=303)
    return None

@app.get("/api/names")
def get_unique_names():
    db = SessionLocal()
    try:
        names = db.query(Receipt.name).distinct().all()
        return [name[0] for name in names]
    finally:
        db.close()

@app.get("/api/receipts")
def api_receipts(request: Request, fy: int = None, search: str = None, mad_id: int = None, register_id: int = None, payment_id: int = None):
    # Check if authenticated
    if not is_authenticated(request):
        return {"error": "Unauthorized"}
    db = get_db()
    try:
        if fy is None:
            fy = get_current_financial_year()
        start_date, end_date = get_financial_year_range(fy)
        
        query = db.query(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        )
        if search:
            search_term = f"%{search}%"
            query = query.filter(
                (Receipt.receipt_no.like(search_term)) | 
                (Receipt.name.like(search_term))
            )
        
        if mad_id:
            query = query.filter(Receipt.mad_category_id == mad_id)
        if register_id:
            query = query.filter(Receipt.register_id == register_id)
        if payment_id:
            query = query.filter(Receipt.payment_mode_id == payment_id)
            
        receipts = query.order_by(desc(Receipt.date)).all()
        
        return {
            "receipts": [
                {
                    "id": r.id,
                    "receipt_no": r.receipt_no,
                    "date": r.date.strftime('%d-%m-%Y'),
                    "name": r.name,
                    "phone": r.phone,
                    "address": r.address or '-',
                    "amount": r.amount,
                    "formatted_amount": format_indian_currency(r.amount),
                    "category": r.mad_category.name,
                    "register": r.register.name,
                    "payment": r.payment_mode.name
                } for r in receipts
            ]
        }
    finally:
        db.close()

@app.on_event("startup")
def startup():
    # Seed data on startup
    seed_data()

@app.get("/login")
def login_page(request: Request):
    # If already logged in, redirect to dashboard
    if is_authenticated(request):
        return RedirectResponse(url="/", status_code=303)
    template = env.get_template("login.html")
    html_content = template.render(request=request)
    return HTMLResponse(content=html_content)

@app.post("/login")
def login(request: Request, username: str = Form(...), password: str = Form(...)):
    # Verify credentials (plain text comparison)
    if username == DEFAULT_ADMIN_USERNAME and password == DEFAULT_ADMIN_PASSWORD:
        request.session["authenticated"] = True
        request.session["username"] = username
        return RedirectResponse(url="/", status_code=303)
    else:
        template = env.get_template("login.html")
        html_content = template.render(request=request, error="Invalid username or password")
        return HTMLResponse(content=html_content, status_code=401)

@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse(url="/login", status_code=303)

@app.get("/change-password")
def change_password_page(request: Request):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    template = env.get_template("change_password.html")
    html_content = template.render(request=request)
    return HTMLResponse(content=html_content)

@app.post("/change-password")
def change_password(
    request: Request,
    current_password: str = Form(...),
    new_password: str = Form(...),
    confirm_password: str = Form(...)
):
    global DEFAULT_ADMIN_PASSWORD
    
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    
    # Verify current password
    if current_password != DEFAULT_ADMIN_PASSWORD:
        template = env.get_template("change_password.html")
        html_content = template.render(request=request, error="Current password is incorrect")
        return HTMLResponse(content=html_content)
    
    # Check if new passwords match
    if new_password != confirm_password:
        template = env.get_template("change_password.html")
        html_content = template.render(request=request, error="New passwords do not match")
        return HTMLResponse(content=html_content)
    
    # Update password
    DEFAULT_ADMIN_PASSWORD = new_password
    
    template = env.get_template("change_password.html")
    html_content = template.render(request=request, success="Password changed successfully")
    return HTMLResponse(content=html_content)

@app.get("/")
def dashboard(request: Request, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        # Get financial year (default to current)
        if fy is None:
            fy = get_current_financial_year()
        
        # Get date range for the financial year
        start_date, end_date = get_financial_year_range(fy)
        
        # Get total receipts and amount for selected FY
        total_receipts = db.query(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        ).count()
        total_amount = db.query(func.sum(Receipt.amount)).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        ).scalar() or 0
        
        # Get recent receipts (all time)
        recent_receipts = db.query(Receipt).order_by(desc(Receipt.id)).limit(5).all()
        
        # Get Mad-wise totals for selected FY
        mad_totals = db.query(
            MadCategory.name,
            func.sum(Receipt.amount).label('total')
        ).join(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        ).group_by(MadCategory.name).all()
        
        # Get Register-wise totals for selected FY
        register_totals = db.query(
            Register.name,
            func.sum(Receipt.amount).label('total')
        ).join(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        ).group_by(Register.name).all()
        
        # Get available financial years (from existing receipts)
        available_years = db.query(
            func.extract('year', Receipt.date)
        ).distinct().order_by(func.extract('year', Receipt.date)).all()
        
        # Convert to list and determine financial years
        years = []
        for row in available_years:
            year = int(row[0])
            # Add both possible FYs for this calendar year
            years.extend([year, year - 1])
        
        # Remove duplicates and sort
        years = sorted(set(years))
        
        template = env.get_template("dashboard.html")
        html_content = template.render(
            request=request,
            total_receipts=total_receipts,
            total_amount=total_amount,
            recent_receipts=recent_receipts,
            mad_totals=mad_totals,
            register_totals=register_totals,
            current_fy=fy,
            available_years=years
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/add-receipt")
def add_receipt_page(request: Request):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        template = env.get_template("add_receipt.html")
        html_content = template.render(
            request=request,
            mad_categories=db.query(MadCategory).all(),
            registers=db.query(Register).all(),
            payment_modes=db.query(PaymentMode).all()
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.post("/add-receipt")
def add_receipt(
    request: Request,
    receipt_no: str = Form(...),
    date: str = Form(...),
    name: str = Form(...),
    phone: str = Form(...),
    address: str = Form(""),
    amount: float = Form(...),
    mad_category_id: int = Form(...),
    register_id: int = Form(...),
    payment_mode_id: int = Form(...),
    notes: str = Form("")
):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    
    try:
        # Validate phone number (exactly 10 digits)
        if not phone.isdigit() or len(phone) != 10:
            raise ValueError("Phone number must be exactly 10 digits")
        
        # Parse date
        receipt_date = datetime.strptime(date, "%Y-%m-%d").date()
        
        # Check if receipt_no already exists
        existing_receipt = db.query(Receipt).filter(Receipt.receipt_no == receipt_no).first()
        if existing_receipt:
            raise ValueError(f"Receipt number '{receipt_no}' already exists. Please check the receipt number.")
        
        # Create new receipt
        receipt = Receipt(
            receipt_no=receipt_no,
            date=receipt_date,
            name=name,
            phone=phone,
            address=address,
            amount=amount,
            mad_category_id=mad_category_id,
            register_id=register_id,
            payment_mode_id=payment_mode_id,
            notes=notes
        )
        
        db.add(receipt)
        db.commit()
        
        return RedirectResponse(url="/", status_code=303)
        
    except Exception as e:
        db.rollback()
        template = env.get_template("add_receipt.html")
        html_content = template.render(
            request=request,
            error=f"Error adding receipt: {str(e)}",
            mad_categories=db.query(MadCategory).all(),
            registers=db.query(Register).all(),
            payment_modes=db.query(PaymentMode).all()
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/edit-receipt/{receipt_id}")
def edit_receipt_page(request: Request, receipt_id: int):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        receipt = db.query(Receipt).filter(Receipt.id == receipt_id).first()
        if not receipt:
            return RedirectResponse(url="/receipts", status_code=303)
        
        template = env.get_template("edit_receipt.html")
        html_content = template.render(
            request=request,
            receipt=receipt,
            mad_categories=db.query(MadCategory).all(),
            registers=db.query(Register).all(),
            payment_modes=db.query(PaymentMode).all()
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.post("/edit-receipt/{receipt_id}")
def edit_receipt(
    request: Request,
    receipt_id: int,
    receipt_no: str = Form(...),
    date: str = Form(...),
    name: str = Form(...),
    phone: str = Form(...),
    address: str = Form(""),
    amount: float = Form(...),
    mad_category_id: int = Form(...),
    register_id: int = Form(...),
    payment_mode_id: int = Form(...),
    notes: str = Form("")
):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        receipt = db.query(Receipt).filter(Receipt.id == receipt_id).first()
        if not receipt:
            return RedirectResponse(url="/receipts", status_code=303)
        
        # Validate phone number (exactly 10 digits)
        if not phone.isdigit() or len(phone) != 10:
            raise ValueError("Phone number must be exactly 10 digits")
        
        # Parse date
        receipt_date = datetime.strptime(date, "%Y-%m-%d").date()
        
        # Check if receipt_no already exists and belongs to a different receipt
        existing_receipt = db.query(Receipt).filter(Receipt.receipt_no == receipt_no).first()
        if existing_receipt and existing_receipt.id != receipt_id:
            raise ValueError(f"Receipt number '{receipt_no}' already exists. Please check the receipt number.")
        
        # Update receipt
        receipt.receipt_no = receipt_no
        receipt.date = receipt_date
        receipt.name = name
        receipt.phone = phone
        receipt.address = address
        receipt.amount = amount
        receipt.mad_category_id = mad_category_id
        receipt.register_id = register_id
        receipt.payment_mode_id = payment_mode_id
        receipt.notes = notes
        
        db.commit()
        
        return RedirectResponse(url="/receipts", status_code=303)
        
    except Exception as e:
        db.rollback()
        receipt = db.query(Receipt).filter(Receipt.id == receipt_id).first()
        template = env.get_template("edit_receipt.html")
        html_content = template.render(
            request=request,
            receipt=receipt,
            error=f"Error updating receipt: {str(e)}",
            mad_categories=db.query(MadCategory).all(),
            registers=db.query(Register).all(),
            payment_modes=db.query(PaymentMode).all()
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/delete-receipt/{receipt_id}")
def delete_receipt_page(request: Request, receipt_id: int):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        receipt = db.query(Receipt).filter(Receipt.id == receipt_id).first()
        if not receipt:
            return RedirectResponse(url="/receipts", status_code=303)
        
        template = env.get_template("delete_confirm.html")
        html_content = template.render(
            request=request,
            receipt=receipt
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.post("/delete-receipt/{receipt_id}")
def delete_receipt(request: Request, receipt_id: int, password: str = Form(...)):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    
    # Verify password before allowing deletion
    if password != DEFAULT_ADMIN_PASSWORD:
        db = get_db()
        try:
            receipt = db.query(Receipt).filter(Receipt.id == receipt_id).first()
            template = env.get_template("delete_confirm.html")
            html_content = template.render(
                request=request,
                receipt=receipt,
                error="Incorrect password. Deletion cancelled."
            )
            return HTMLResponse(content=html_content)
        finally:
            db.close()
    
    db = get_db()
    try:
        receipt = db.query(Receipt).filter(Receipt.id == receipt_id).first()
        if receipt:
            db.delete(receipt)
            db.commit()
        
        return RedirectResponse(url="/receipts", status_code=303)
    finally:
        db.close()

@app.get("/receipts")
def receipts_page(request: Request, fy: int = None, search: str = None, mad_id: int = None, register_id: int = None, payment_id: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        # Get financial year (default to current)
        if fy is None:
            fy = get_current_financial_year()
        
        # Get date range for the financial year
        start_date, end_date = get_financial_year_range(fy)
        
        # Build query with F.Y. filter
        query = db.query(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        )
        
        # Add filters
        if search:
            search_term = f"%{search}%"
            query = query.filter(
                (Receipt.receipt_no.like(search_term)) | 
                (Receipt.name.like(search_term))
            )
        
        if mad_id:
            query = query.filter(Receipt.mad_category_id == mad_id)
        if register_id:
            query = query.filter(Receipt.register_id == register_id)
        if payment_id:
            query = query.filter(Receipt.payment_mode_id == payment_id)
            
        receipts = query.order_by(desc(Receipt.date)).all()
        
        # Calculate balances
        opening_balance, period_total, closing_balance = get_balances(db, start_date, end_date)
        
        # Get available financial years
        available_years = db.query(
            func.extract('year', Receipt.date)
        ).distinct().order_by(func.extract('year', Receipt.date)).all()
        
        years = []
        for row in available_years:
            year = int(row[0])
            years.extend([year, year - 1])
        years = sorted(set(years))
        if not years:
            years = [fy]

        # Get master data for filters
        mad_categories = db.query(MadCategory).all()
        registers = db.query(Register).all()
        payment_modes = db.query(PaymentMode).all()
        
        template = env.get_template("receipts.html")
        html_content = template.render(
            request=request,
            receipts=receipts,
            current_fy=fy,
            available_years=years,
            search=search or "",
            mad_id=mad_id,
            register_id=register_id,
            payment_id=payment_id,
            mad_categories=mad_categories,
            registers=registers,
            payment_modes=payment_modes,
            opening_balance=opening_balance,
            period_total=period_total,
            closing_balance=closing_balance
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/receipts/download")
def download_receipts_excel(request: Request, filter_type: str = "current_fy", start_date: str = None, end_date: str = None, fy: int = None, search: str = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        actual_start_date, actual_end_date, title_suffix = get_date_range_for_filter(filter_type, start_date, end_date, fy)
        
        # Build query with date filter
        query = db.query(Receipt).filter(
            and_(Receipt.date >= actual_start_date, Receipt.date <= actual_end_date)
        )
        
        # Add search filter if provided
        if search:
            search_term = f"%{search}%"
            query = query.filter(
                (Receipt.receipt_no.like(search_term)) | 
                (Receipt.name.like(search_term))
            )
        
        receipts = query.order_by(desc(Receipt.date)).all()
        
        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "All Receipts"
        
        # Title
        ws.merge_cells('A1:J1')
        ws['A1'] = f"All Receipts Report - {title_suffix}"
        if search:
            ws['A1'] = f"All Receipts Report - {title_suffix} (Search: {search})"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Headers
        headers = ['Receipt No', 'Date', 'Name', 'Phone', 'Address', 'Amount', 'Category', 'Register', 'Payment Mode', 'Notes']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
            cell.alignment = Alignment(horizontal='center')
        
        # Data
        for row_idx, receipt in enumerate(receipts, 4):
            ws.cell(row=row_idx, column=1, value=receipt.receipt_no)
            ws.cell(row=row_idx, column=2, value=receipt.date.strftime('%d-%m-%Y'))
            ws.cell(row=row_idx, column=3, value=receipt.name)
            ws.cell(row=row_idx, column=4, value=receipt.phone)
            ws.cell(row=row_idx, column=5, value=receipt.address or '')
            cell = ws.cell(row=row_idx, column=6, value=receipt.amount)
            cell.number_format = '#,##,##0.00'
            ws.cell(row=row_idx, column=7, value=receipt.mad_category.name if receipt.mad_category else '')
            ws.cell(row=row_idx, column=8, value=receipt.register.name if receipt.register else '')
            ws.cell(row=row_idx, column=9, value=receipt.payment_mode.name if receipt.payment_mode else '')
            ws.cell(row=row_idx, column=10, value=receipt.notes or '')
        
        # Calculate balances
        opening_balance, period_total, closing_balance = get_balances(db, actual_start_date, actual_end_date)
        
        # Total row
        total_row = len(receipts) + 4
        ws.cell(row=total_row, column=5, value="Total").font = Font(bold=True)
        cell = ws.cell(row=total_row, column=6, value=sum(r.amount for r in receipts))
        cell.font = Font(bold=True)
        cell.number_format = '#,##,##0.00'

        # Summary rows
        summary_row = total_row + 2
        ws.cell(row=summary_row, column=5, value="Opening Balance").font = Font(bold=True)
        ws.cell(row=summary_row, column=6, value=opening_balance).number_format = '#,##,##0.00'
        
        ws.cell(row=summary_row + 1, column=5, value="Period Total").font = Font(bold=True)
        ws.cell(row=summary_row + 1, column=6, value=period_total).number_format = '#,##,##0.00'
        
        ws.cell(row=summary_row + 2, column=5, value="Closing Balance").font = Font(bold=True)
        ws.cell(row=summary_row + 2, column=6, value=closing_balance).number_format = '#,##,##0.00'
        
        # Column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 30
        ws.column_dimensions['F'].width = 12
        ws.column_dimensions['G'].width = 15
        ws.column_dimensions['H'].width = 15
        ws.column_dimensions['I'].width = 15
        ws.column_dimensions['J'].width = 25
        
        # Save to temporary file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        search_suffix = f"_search_{search}" if search else ""
        default_filename = f"all_receipts_{title_suffix.replace(' ', '_')}{search_suffix}_{timestamp}.xlsx"
        
        filepath = prompt_save_path(
            "Save All Receipts Report",
            default_filename,
            [("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filepath:
            if not filepath.endswith('.xlsx'):
                filepath += '.xlsx'
            wb.save(filepath)
            return {"success": True, "message": f"Report saved successfully"}
        return {"success": False, "message": "Save cancelled by user."}
    except Exception as e:
        return {"success": False, "message": f"Error generating report: {str(e)}"}
    finally:
        db.close()

@app.get("/mad-report")
def mad_report_page(request: Request, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        # Get financial year (default to current)
        if fy is None:
            fy = get_current_financial_year()
        
        # Get date range for the financial year
        start_date, end_date = get_financial_year_range(fy)
        
        # Get Mad-wise totals for selected FY
        mad_report = db.query(
            MadCategory.id,
            MadCategory.name,
            func.count(Receipt.id).label('count'),
            func.sum(Receipt.amount).label('total')
        ).join(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        ).group_by(MadCategory.id, MadCategory.name).order_by(desc('total')).all()
        
        # Calculate balances
        opening_balance, period_total, closing_balance = get_balances(db, start_date, end_date)
        
        # Get available financial years (from existing receipts)
        available_years = db.query(
            func.extract('year', Receipt.date)
        ).distinct().order_by(func.extract('year', Receipt.date)).all()
        
        # Convert to list and determine financial years
        years = []
        for row in available_years:
            year = int(row[0])
            years.extend([year, year - 1])
        
        years = sorted(set(years))
        
        template = env.get_template("mad_report.html")
        html_content = template.render(
            request=request,
            mad_report=mad_report,
            current_fy=fy,
            available_years=years,
            opening_balance=opening_balance,
            period_total=period_total,
            closing_balance=closing_balance
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/mad-report/download")
def download_mad_report_excel(request: Request, filter_type: str = "current_fy", start_date: str = None, end_date: str = None, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        actual_start_date, actual_end_date, title_suffix = get_date_range_for_filter(filter_type, start_date, end_date, fy)
        
        # Get detailed Mad-wise data with balances
        mad_data = []
        categories = db.query(MadCategory).all()
        for cat in categories:
            opening = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.mad_category_id == cat.id, Receipt.date < actual_start_date)
            ).scalar() or 0
            period = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.mad_category_id == cat.id, Receipt.date >= actual_start_date, Receipt.date <= actual_end_date)
            ).scalar() or 0
            if opening > 0 or period > 0:
                mad_data.append({
                    "name": cat.name,
                    "opening": opening,
                    "period": period,
                    "closing": opening + period
                })
        
        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Mad-wise Report"
        
        # Define styles
        header_font_white = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Set column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        
        # Title
        ws.merge_cells('A1:D1')
        ws['A1'] = "Mad-wise Collection Report"
        ws['A1'].font = Font(bold=True, size=16)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Date Range
        ws.merge_cells('A2:D2')
        ws['A2'] = f"Date Range: {title_suffix}"
        ws['A2'].font = Font(bold=True, size=11)
        ws['A2'].alignment = Alignment(horizontal='center')
        
        # Headers
        headers = ["Category Name", "Opening Balance", "Period Total", "Closing Balance"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        
        # Data rows
        for row_idx, item in enumerate(mad_data, start=5):
            ws.cell(row=row_idx, column=1, value=item["name"]).border = border
            
            cell_o = ws.cell(row=row_idx, column=2, value=item["opening"])
            cell_o.border = border
            cell_o.number_format = '#,##,##0.00'
            
            cell_p = ws.cell(row=row_idx, column=3, value=item["period"])
            cell_p.border = border
            cell_p.number_format = '#,##,##0.00'
            
            cell_c = ws.cell(row=row_idx, column=4, value=item["closing"])
            cell_c.border = border
            cell_c.number_format = '#,##,##0.00'
        
        # Calculate balances
        opening_balance, period_total, closing_balance = get_balances(db, actual_start_date, actual_end_date)
        
        # Total row
        total_row = len(mad_data) + 5
        ws.cell(row=total_row, column=1, value="Grand Total").font = Font(bold=True)
        ws.cell(row=total_row, column=1).border = border
        
        cell_opening = ws.cell(row=total_row, column=2, value=sum(m["opening"] for m in mad_data))
        cell_opening.font = Font(bold=True)
        cell_opening.border = border
        cell_opening.number_format = '#,##,##0.00'
        
        cell_period = ws.cell(row=total_row, column=3, value=sum(m["period"] for m in mad_data))
        cell_period.font = Font(bold=True)
        cell_period.border = border
        cell_period.number_format = '#,##,##0.00'
        
        cell_closing = ws.cell(row=total_row, column=4, value=sum(m["closing"] for m in mad_data))
        cell_closing.font = Font(bold=True)
        cell_closing.border = border
        cell_closing.number_format = '#,##,##0.00'

        # Summary rows
        summary_row = total_row + 2
        ws.cell(row=summary_row, column=1, value="Opening Balance").font = Font(bold=True)
        ws.cell(row=summary_row, column=4, value=opening_balance).number_format = '#,##,##0.00'
        
        ws.cell(row=summary_row + 1, column=1, value="Period Total").font = Font(bold=True)
        ws.cell(row=summary_row + 1, column=4, value=period_total).number_format = '#,##,##0.00'
        
        ws.cell(row=summary_row + 2, column=1, value="Closing Balance").font = Font(bold=True)
        ws.cell(row=summary_row + 2, column=4, value=closing_balance).number_format = '#,##,##0.00'
        
        # Save to file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f"mad_report_{title_suffix.replace(' ', '_')}_{timestamp}.xlsx"
        
        filepath = prompt_save_path(
            "Save Mad-wise Report",
            default_filename,
            [("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filepath:
            if not filepath.endswith('.xlsx'):
                filepath += '.xlsx'
            wb.save(filepath)
            return {"success": True, "message": f"Report saved successfully"}
        return {"success": False, "message": "Save cancelled by user."}
    except Exception as e:
        return {"success": False, "message": f"Error generating report: {str(e)}"}
    finally:
        db.close()

@app.get("/register-report")
def register_report_page(request: Request, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        # Get financial year (default to current)
        if fy is None:
            fy = get_current_financial_year()
        
        # Get date range for the financial year
        start_date, end_date = get_financial_year_range(fy)
        
        # Get Register-wise totals for selected FY
        register_report = db.query(
            Register.id,
            Register.name,
            func.count(Receipt.id).label('count'),
            func.sum(Receipt.amount).label('total')
        ).join(Receipt).filter(
            and_(Receipt.date >= start_date, Receipt.date <= end_date)
        ).group_by(Register.id, Register.name).order_by(desc('total')).all()
        
        # Calculate balances
        opening_balance, period_total, closing_balance = get_balances(db, start_date, end_date)
        
        # Get available financial years (from existing receipts)
        available_years = db.query(
            func.extract('year', Receipt.date)
        ).distinct().order_by(func.extract('year', Receipt.date)).all()
        
        # Convert to list and determine financial years
        years = []
        for row in available_years:
            year = int(row[0])
            years.extend([year, year - 1])
        
        years = sorted(set(years))
        
        template = env.get_template("register_report.html")
        html_content = template.render(
            request=request,
            register_report=register_report,
            current_fy=fy,
            available_years=years,
            opening_balance=opening_balance,
            period_total=period_total,
            closing_balance=closing_balance
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/register-report/download")
def download_register_report_excel(request: Request, filter_type: str = "current_fy", start_date: str = None, end_date: str = None, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        actual_start_date, actual_end_date, title_suffix = get_date_range_for_filter(filter_type, start_date, end_date, fy)
        
        # Get detailed Register-wise data with balances
        register_data = []
        registers = db.query(Register).all()
        for reg in registers:
            opening = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.register_id == reg.id, Receipt.date < actual_start_date)
            ).scalar() or 0
            period = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.register_id == reg.id, Receipt.date >= actual_start_date, Receipt.date <= actual_end_date)
            ).scalar() or 0
            if opening > 0 or period > 0:
                register_data.append({
                    "name": reg.name,
                    "opening": opening,
                    "period": period,
                    "closing": opening + period
                })

        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Register-wise Report"
        
        # Define styles
        header_font_white = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Set column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        
        # Title
        ws.merge_cells('A1:D1')
        ws['A1'] = "Register-wise Collection Report"
        ws['A1'].font = Font(bold=True, size=16)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Date Range
        ws.merge_cells('A2:D2')
        ws['A2'] = f"Date Range: {title_suffix}"
        ws['A2'].font = Font(bold=True, size=11)
        ws['A2'].alignment = Alignment(horizontal='center')
        
        # Headers
        headers = ["Register Name", "Opening Balance", "Period Total", "Closing Balance"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        
        # Data rows
        for row_idx, item in enumerate(register_data, start=5):
            ws.cell(row=row_idx, column=1, value=item["name"]).border = border
            
            cell_o = ws.cell(row=row_idx, column=2, value=item["opening"])
            cell_o.border = border
            cell_o.number_format = '#,##,##0.00'
            
            cell_p = ws.cell(row=row_idx, column=3, value=item["period"])
            cell_p.border = border
            cell_p.number_format = '#,##,##0.00'
            
            cell_c = ws.cell(row=row_idx, column=4, value=item["closing"])
            cell_c.border = border
            cell_c.number_format = '#,##,##0.00'
        
        # Total row
        total_row = len(register_data) + 5
        ws.cell(row=total_row, column=1, value="Grand Total").font = Font(bold=True)
        ws.cell(row=total_row, column=1).border = border
        
        cell_opening = ws.cell(row=total_row, column=2, value=sum(r["opening"] for r in register_data))
        cell_opening.font = Font(bold=True)
        cell_opening.border = border
        cell_opening.number_format = '#,##,##0.00'
        
        cell_period = ws.cell(row=total_row, column=3, value=sum(r["period"] for r in register_data))
        cell_period.font = Font(bold=True)
        cell_period.border = border
        cell_period.number_format = '#,##,##0.00'
        
        cell_closing = ws.cell(row=total_row, column=4, value=sum(r["closing"] for r in register_data))
        cell_closing.font = Font(bold=True)
        cell_closing.border = border
        cell_closing.number_format = '#,##,##0.00'
        
        # Save to file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f"register_report_{title_suffix.replace(' ', '_')}_{timestamp}.xlsx"
        
        filepath = prompt_save_path(
            "Save Register-wise Report",
            default_filename,
            [("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filepath:
            if not filepath.endswith('.xlsx'):
                filepath += '.xlsx'
            wb.save(filepath)
            return {"success": True, "message": f"Report saved successfully"}
        return {"success": False, "message": "Save cancelled by user."}
    except Exception as e:
        return {"success": False, "message": f"Error generating report: {str(e)}"}
    finally:
        db.close()

@app.post("/backup")
def backup_database(request: Request):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    try:
        # Create backup filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"baitulmal_backup_{timestamp}.db"
        
        filepath = prompt_save_path(
            "Save Backup As",
            backup_filename,
            [("Database files", "*.db"), ("All files", "*.*")]
        )
        
        if not filepath:
            return {"success": False, "message": "Backup cancelled by user."}
            
        if not filepath.endswith('.db'):
            filepath += '.db'
            
        # Copy database file
        shutil.copy2("baitulmal.db", filepath)
        
        # Upload to Google Drive using OAuth2
        gdrive_status = ""
        try:
            from google_auth_oauthlib.flow import InstalledAppFlow
            from google.auth.transport.requests import Request as GoogleRequest
            import pickle
            
            SCOPES = ['https://www.googleapis.com/auth/drive.file']
            
            # Check for OAuth2 credentials file (client_secret.json from Google Cloud Console)
            CLIENT_SECRET_FILE = resource_path('client_secret.json')
            TOKEN_FILE = 'token.pickle'
            
            creds = None
            
            # Load existing token
            if os.path.exists(TOKEN_FILE):
                with open(TOKEN_FILE, 'rb') as token:
                    creds = pickle.load(token)
            
            # If no valid credentials, do OAuth2 flow
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(GoogleRequest())
                elif os.path.exists(CLIENT_SECRET_FILE):
                    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
                    creds = flow.run_local_server(port=0)
                    # Save token for future runs
                    with open(TOKEN_FILE, 'wb') as token:
                        pickle.dump(creds, token)
                else:
                    gdrive_status = " (Google Drive: client_secret.json not found - run setup first)"
                    return {"success": True, "message": f"Backup saved locally. {gdrive_status}"}
            
            # Build Drive service
            service = build('drive', 'v3', credentials=creds)
            
            # Upload file to specific folder
            BACKUP_FOLDER_ID = '1okT-RuqIumUji0P94_csx_WXmj0na_q0'
            file_metadata = {
                'name': os.path.basename(filepath),
                'parents': [BACKUP_FOLDER_ID]
            }
            media = MediaFileUpload(filepath, resumable=True)
            file = service.files().create(
                body=file_metadata,
                media_body=media,
                fields='id'
            ).execute()
            
            gdrive_status = " + uploaded to your Google Drive (Baitulmal Backups folder)"
            
        except Exception as gdrive_error:
            error_msg = str(gdrive_error)
            if "access_denied" in error_msg:
                gdrive_status = " (Google Drive: access denied - check OAuth credentials)"
            else:
                gdrive_status = f" (Google Drive: {str(gdrive_error)[:50]}...)"
        
        return {"success": True, "message": f"Backup saved locally. {gdrive_status}"}
    
    except Exception as e:
        return {"success": False, "message": f"Backup failed: {str(e)}"}

@app.post("/restore")
def restore_database(request: Request):
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    
    try:
        filepath = prompt_open_path(
            "Select Database Backup to Restore",
            [("Database files", "*.db"), ("All files", "*.*")]
        )
        
        if not filepath:
            return {"success": False, "message": "Restore cancelled by user."}
            
        # Copy selected file to baitulmal.db
        shutil.copy2(filepath, "baitulmal.db")
        
        return {"success": True, "message": f"Successfully restored from {os.path.basename(filepath)}"}
    except Exception as e:
        return {"success": False, "message": f"Restore failed: {str(e)}"}

@app.get("/payment-report")
def payment_report_page(request: Request, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        # Get financial year (default to current)
        if fy is None:
            fy = get_current_financial_year()
        
        # Get date range for the financial year
        start_date, end_date = get_financial_year_range(fy)
        
        # Get Payment Mode-wise data including Contra breakdown
        payment_modes = db.query(PaymentMode).all()
        payment_report_details = []
        for mode in payment_modes:
            # Receipts in period
            receipt_total = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.payment_mode_id == mode.id, Receipt.date >= start_date, Receipt.date <= end_date)
            ).scalar() or 0
            
            # Contra In in period
            contra_in = db.query(func.sum(ContraEntry.amount)).filter(
                and_(ContraEntry.to_payment_mode_id == mode.id, ContraEntry.date >= start_date, ContraEntry.date <= end_date)
            ).scalar() or 0
            
            # Contra Out in period
            contra_out = db.query(func.sum(ContraEntry.amount)).filter(
                and_(ContraEntry.from_payment_mode_id == mode.id, ContraEntry.date >= start_date, ContraEntry.date <= end_date)
            ).scalar() or 0
            
            # Opening balance for this mode
            opening_receipts = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.payment_mode_id == mode.id, Receipt.date < start_date)
            ).scalar() or 0
            opening_contra_in = db.query(func.sum(ContraEntry.amount)).filter(
                and_(ContraEntry.to_payment_mode_id == mode.id, ContraEntry.date < start_date)
            ).scalar() or 0
            opening_contra_out = db.query(func.sum(ContraEntry.amount)).filter(
                and_(ContraEntry.from_payment_mode_id == mode.id, ContraEntry.date < start_date)
            ).scalar() or 0
            mode_opening = opening_receipts + opening_contra_in - opening_contra_out
            
            mode_closing = mode_opening + receipt_total + contra_in - contra_out
            
            if receipt_total != 0 or contra_in != 0 or contra_out != 0 or mode_opening != 0:
                payment_report_details.append({
                    "id": mode.id,
                    "name": mode.name,
                    "opening": mode_opening,
                    "receipts": receipt_total,
                    "contra_in": contra_in,
                    "contra_out": contra_out,
                    "closing": mode_closing
                })
        
        # Calculate balances
        opening_balance, period_total, closing_balance = get_balances(db, start_date, end_date)
        
        # Get available financial years (from existing receipts)
        available_years = db.query(
            func.extract('year', Receipt.date)
        ).distinct().order_by(func.extract('year', Receipt.date)).all()
        
        # Convert to list and determine financial years
        years = []
        for row in available_years:
            year = int(row[0])
            years.extend([year, year - 1])
        
        years = sorted(set(years))
        
        template = env.get_template("payment_report.html")
        html_content = template.render(
            request=request,
            payment_report=payment_report_details,
            current_fy=fy,
            available_years=years,
            opening_balance=opening_balance,
            period_total=period_total,
            closing_balance=closing_balance
        )
        return HTMLResponse(content=html_content)
    finally:
        db.close()

@app.get("/payment-report/download")
def download_payment_report_excel(request: Request, filter_type: str = "current_fy", start_date: str = None, end_date: str = None, fy: int = None):
    # Check if authenticated
    auth_redirect = require_auth(request)
    if auth_redirect:
        return auth_redirect
    db = get_db()
    try:
        actual_start_date, actual_end_date, title_suffix = get_date_range_for_filter(filter_type, start_date, end_date, fy)
        
        # Get detailed Payment Mode-wise data with balances
        payment_data = []
        modes = db.query(PaymentMode).all()
        for mode in modes:
            opening = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.payment_mode_id == mode.id, Receipt.date < actual_start_date)
            ).scalar() or 0
            period = db.query(func.sum(Receipt.amount)).filter(
                and_(Receipt.payment_mode_id == mode.id, Receipt.date >= actual_start_date, Receipt.date <= actual_end_date)
            ).scalar() or 0
            if opening > 0 or period > 0:
                payment_data.append({
                    "name": mode.name,
                    "opening": opening,
                    "period": period,
                    "closing": opening + period
                })

        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Payment Report"
        
        # Define styles
        header_font_white = Font(bold=True, size=12, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Set column widths
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        
        # Title
        ws.merge_cells('A1:D1')
        ws['A1'] = "Payment Mode-wise Collection Report"
        ws['A1'].font = Font(bold=True, size=16)
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # Date Range
        ws.merge_cells('A2:D2')
        ws['A2'] = f"Date Range: {title_suffix}"
        ws['A2'].font = Font(bold=True, size=11)
        ws['A2'].alignment = Alignment(horizontal='center')
        
        # Headers
        headers = ["Payment Mode", "Opening Balance", "Period Total", "Closing Balance"]
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        
        # Data rows
        for row_idx, item in enumerate(payment_data, start=5):
            ws.cell(row=row_idx, column=1, value=item["name"]).border = border
            
            cell_o = ws.cell(row=row_idx, column=2, value=item["opening"])
            cell_o.border = border
            cell_o.number_format = '#,##,##0.00'
            
            cell_p = ws.cell(row=row_idx, column=3, value=item["period"])
            cell_p.border = border
            cell_p.number_format = '#,##,##0.00'
            
            cell_c = ws.cell(row=row_idx, column=4, value=item["closing"])
            cell_c.border = border
            cell_c.number_format = '#,##,##0.00'
        
        # Total row
        total_row = len(payment_data) + 5
        ws.cell(row=total_row, column=1, value="Grand Total").font = Font(bold=True)
        ws.cell(row=total_row, column=1).border = border
        
        cell_opening = ws.cell(row=total_row, column=2, value=sum(p["opening"] for p in payment_data))
        cell_opening.font = Font(bold=True)
        cell_opening.border = border
        cell_opening.number_format = '#,##,##0.00'
        
        cell_period = ws.cell(row=total_row, column=3, value=sum(p["period"] for p in payment_data))
        cell_period.font = Font(bold=True)
        cell_period.border = border
        cell_period.number_format = '#,##,##0.00'
        
        cell_closing = ws.cell(row=total_row, column=4, value=sum(p["closing"] for p in payment_data))
        cell_closing.font = Font(bold=True)
        cell_closing.border = border
        cell_closing.number_format = '#,##,##0.00'
        
        # Save to file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"payment_report_{title_suffix.replace(' ', '_')}_{timestamp}.xlsx"
        
        filepath = prompt_save_path(
            "Save Payment Report",
            default_filename,
            [("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filepath:
            if not filepath.endswith('.xlsx'):
                filepath += '.xlsx'
            wb.save(filepath)
            return {"success": True, "message": f"Report saved successfully"}
        return {"success": False, "message": "Save cancelled by user."}
    except Exception as e:
        return {"success": False, "message": f"Error generating report: {str(e)}"}
    finally:
        db.close()

@app.get("/contra")
def contra_page(request: Request, fy: int = None):
    auth_redirect = require_auth(request)
    if auth_redirect: return auth_redirect
    db = get_db()
    try:
        if fy is None: fy = get_current_financial_year()
        start_date, end_date = get_financial_year_range(fy)
        
        entries = db.query(ContraEntry).filter(
            and_(ContraEntry.date >= start_date, ContraEntry.date <= end_date)
        ).order_by(desc(ContraEntry.date)).all()
        
        payment_modes = db.query(PaymentMode).all()
        
        template = env.get_template("contra.html")
        return HTMLResponse(content=template.render(
            request=request,
            entries=entries,
            payment_modes=payment_modes,
            current_fy=fy
        ))
    finally:
        db.close()

@app.post("/add-contra")
def add_contra(
    request: Request,
    date: str = Form(...),
    from_mode_id: int = Form(...),
    to_mode_id: int = Form(...),
    amount: float = Form(...),
    notes: str = Form("")
):
    auth_redirect = require_auth(request)
    if auth_redirect: return auth_redirect
    
    db = get_db()
    try:
        entry = ContraEntry(
            date=datetime.strptime(date, "%Y-%m-%d").date(),
            from_payment_mode_id=from_mode_id,
            to_payment_mode_id=to_mode_id,
            amount=amount,
            notes=notes
        )
        db.add(entry)
        db.commit()
        return RedirectResponse(url="/contra", status_code=303)
    finally:
        db.close()

@app.get("/delete-contra/{contra_id}")
def delete_contra(request: Request, contra_id: int):
    auth_redirect = require_auth(request)
    if auth_redirect: return auth_redirect
    db = get_db()
    try:
        entry = db.query(ContraEntry).filter(ContraEntry.id == contra_id).first()
        if entry:
            db.delete(entry)
            db.commit()
        return RedirectResponse(url="/contra", status_code=303)
    finally:
        db.close()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
