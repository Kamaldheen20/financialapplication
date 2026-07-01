# ==========================
# IMPORTS
# ==========================
from math import ceil
from datetime import datetime, timedelta
from sqlalchemy.exc import IntegrityError
import os
from dotenv import load_dotenv
import io
import shutil
import webbrowser
from datetime import datetime
from threading import Timer

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file
)

from flask_login import (
    LoginManager,
    login_user,
    login_required,
    logout_user,
    current_user
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash
)

from openpyxl import Workbook
from models import (
    db,
    Admin,
    Customer,
    Payment
)

# ==========================
# APP CONFIG
# ==========================

load_dotenv()


app = Flask(__name__)

app.config["SECRET_KEY"] = os.getenv("SECRET_KEY")

app.config["SQLALCHEMY_DATABASE_URI"] = os.getenv("DATABASE_URL")

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db.init_app(app)

login_manager = LoginManager()

login_manager.init_app(app)

login_manager.login_view = "login"


@login_manager.user_loader
def load_user(user_id):

    return db.session.get(
Admin,
int(user_id)
)

# ==========================
# COMPANY SETTINGS MODEL
# ==========================

class CompanySettings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, nullable=False)
    company_name = db.Column(db.String(200))
    address = db.Column(db.String(500))
    phone = db.Column(db.String(50))


# ==========================
# PDF FONT HELPER
# Supports Tamil AND English (Latin) in the same PDF — including mixed
# names like "குmaரெசன்" (Tamil + Latin letters in one string).
#
# Root cause of the bug:
#   NotoSans-Regular does NOT contain Tamil glyphs (U+0B80–U+0BFF).
#   NotoSansTamil does NOT contain Latin glyphs.
#   Neither font alone can render mixed-script text.
#
# Fix — dual-font split-span approach:
#   1. Register NotoSansTamil  (for Tamil characters)
#   2. Register NotoSans       (for Latin/English characters)
#   3. _pdf_text(s) wraps every character in the correct <font> tag
#      so ReportLab Paragraph can render both scripts in one cell/line.
# ==========================

_FONT_CACHE = {}


def _find_font_file(candidates):
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


def _download_font(url, save_path):
    import urllib.request
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        print(f"[PDF font] Downloading {os.path.basename(save_path)} …")
        urllib.request.urlretrieve(url, save_path)
        print(f"[PDF font] Saved to {save_path}")
        return True
    except Exception as e:
        print(f"[PDF font] Download failed: {e}")
        return False


def _ensure_font(fname, github_subpath):
    """Locate font on disk across common paths, or auto-download it."""
    app_fonts = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    windir = os.environ.get("WINDIR", "C:\\Windows")

    candidates = [
        os.path.join(app_fonts, fname),
        os.path.join("C:\\FinanceManager", "fonts", fname),
        os.path.join(windir, "Fonts", fname),
        f"/usr/share/fonts/truetype/noto/{fname}",
        f"/usr/share/fonts/opentype/noto/{fname}",
        f"/Library/Fonts/{fname}",
        f"/System/Library/Fonts/{fname}",
    ]
    path = _find_font_file(candidates)
    if path:
        return path

    save_path = os.path.join(app_fonts, fname)
    url = (
        "https://github.com/googlefonts/noto-fonts/raw/main/"
        f"hinted/ttf/{github_subpath}"
    )
    return save_path if _download_font(url, save_path) else None


def _try_register(reg_name, path):
    """Register a TTFont. Returns True if already registered or success."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    try:
        pdfmetrics.getFont(reg_name)
        return True
    except KeyError:
        pass
    if path and os.path.exists(path):
        try:
            pdfmetrics.registerFont(TTFont(reg_name, path))
            return True
        except Exception as e:
            print(f"[PDF font] Cannot register {reg_name}: {e}")
    return False


def _register_pdf_fonts():
    """
    Register both NotoSansTamil and NotoSans with ReportLab.
    Returns (tamil_font, latin_font, bold_font, has_tamil).
    Always call this before building any PDF.
    """
    if "result" in _FONT_CACHE:
        return _FONT_CACHE["result"]

    # ── Tamil font (NotoSansTamil covers U+0B80–U+0BFF) ────────────────────
    tamil_path = _ensure_font(
        "NotoSansTamil-Regular.ttf",
        "NotoSansTamil/NotoSansTamil-Regular.ttf"
    )
    tamil_ok = _try_register("NotoSansTamil", tamil_path)
    if tamil_ok:
        _try_register("NotoSansTamil-Bold", tamil_path)  # reuse same file

    # ── Latin font (NotoSans covers A–Z, a–z, digits, punctuation) ─────────
    latin_path = _ensure_font(
        "NotoSans-Regular.ttf",
        "NotoSans/NotoSans-Regular.ttf"
    )
    latin_bold_path = _ensure_font(
        "NotoSans-Bold.ttf",
        "NotoSans/NotoSans-Bold.ttf"
    )
    latin_ok = _try_register("NotoSans", latin_path)
    if latin_ok:
        if not _try_register("NotoSans-Bold", latin_bold_path or latin_path):
            pass  # bold falls back to regular below

    if tamil_ok and latin_ok:
        print("[PDF font] Dual-font mode: NotoSansTamil + NotoSans (Tamil + Latin)")
        result = ("NotoSansTamil", "NotoSans", "NotoSans-Bold", True)
    elif tamil_ok:
        print("[PDF font] Tamil-only mode: NotoSansTamil (Latin may look basic)")
        result = ("NotoSansTamil", "NotoSansTamil", "NotoSansTamil-Bold", True)
    elif latin_ok:
        print("[PDF font] Latin-only mode: NotoSans (Tamil will not render)")
        result = ("NotoSans", "NotoSans", "NotoSans-Bold", False)
    else:
        print("[PDF font] WARNING: No Unicode fonts found. Using Helvetica.")
        result = ("Helvetica", "Helvetica", "Helvetica-Bold", False)

    _FONT_CACHE["result"] = result
    return result


def _pdf_text(text):
    """
    Wrap each character in the correct <font> tag so ReportLab Paragraph
    renders both Tamil and Latin characters correctly in the same string.

    Tamil Unicode block: U+0B80 – U+0BFF
    Everything else (Latin, digits, spaces, punctuation) uses the Latin font.
    """
    cached = _FONT_CACHE.get("result")
    if cached is None:
        _register_pdf_fonts()
        cached = _FONT_CACHE["result"]

    tamil_font = cached[0]
    latin_font = cached[1]

    # If no dual-font support, just return plain text
    if tamil_font == latin_font:
        return text

    result = []
    current_font = None
    for ch in str(text):
        cp = ord(ch)
        font = tamil_font if 0x0B80 <= cp <= 0x0BFF else latin_font
        if font != current_font:
            if current_font is not None:
                result.append("</font>")
            result.append(f'<font name="{font}">')
            current_font = font
        # Escape XML special characters
        if ch == "&":
            result.append("&amp;")
        elif ch == "<":
            result.append("&lt;")
        elif ch == ">":
            result.append("&gt;")
        else:
            result.append(ch)
    if current_font:
        result.append("</font>")
    return "".join(result)


# Backward-compat alias used throughout the rest of the file
def _register_tamil_font():
    tamil_font, latin_font, bold_font, has_tamil = _register_pdf_fonts()
    # Return (normal_font, bold_font, has_tamil) as before — callers use
    # the normal_font for body text; _pdf_text() handles per-char switching.
    return tamil_font, bold_font, has_tamil


# Pre-register at startup so the first PDF is fast
try:
    _register_pdf_fonts()
except Exception:
    pass


# ==========================
# CUSTOMER SORT HELPER
# Sorts customers by customer_id numerically (1, 2, 8, 22, 88, 888, 11112)
# instead of alphabetically (1, 11112, 2, 22, 8, 88, 888), since
# customer_id is stored as a string in the database.
# ==========================

def _sort_customers(customers):
    def _id_sort_key(c):
        try:
            return (0, int(c.customer_id))
        except (ValueError, TypeError):
            return (1, c.customer_id)
    return sorted(customers, key=_id_sort_key)


# ==========================
# HOME / REDIRECT
# ==========================

@app.route("/")
def home():
    return redirect(url_for("login"))


# ==========================
# LOGIN
# ==========================

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        admin = Admin.query.filter_by(username=username).first()

        if admin and check_password_hash(admin.password, password):
            login_user(admin)
            return redirect(url_for("dashboard"))

        flash("Invalid Username or Password")

    return render_template("login.html")


# ==========================
# REGISTER
# ==========================

@app.route("/register", methods=["GET", "POST"])
def register():
    

    if request.method == "POST":

        username = request.form["username"].strip()
        mobile = request.form["mobile"].strip()
        password = request.form["password"]

        # Check username
        if Admin.query.filter_by(username=username).first():
            flash("Username already exists", "danger")
            return redirect(url_for("register"))

        # Check mobile
        if Admin.query.filter_by(mobile=mobile).first():
            flash("Mobile number already registered", "danger")
            return redirect(url_for("register"))

        admin = Admin(
            username=username,
            mobile=mobile,
            password=generate_password_hash(password)
        )

        db.session.add(admin)
        db.session.commit()

        flash("Account Created Successfully", "success")
        return redirect(url_for("login"))
    return render_template("register.html")
#forget pass
@app.route("/forgot_password", methods=["GET", "POST"])
def forgot_password():

    if request.method == "POST":

        username = request.form["username"].strip()
        mobile = request.form["mobile"].strip()
        password = request.form["password"]
        confirm_password = request.form["confirm_password"]

        # Check password confirmation
        if password != confirm_password:
            flash("Passwords do not match.", "danger")
            return redirect(url_for("forgot_password"))

        # Find the user
        admin = Admin.query.filter_by(
            username=username,
            mobile=mobile
        ).first()

        if not admin:
            flash(
                "Invalid Username or Mobile Number.",
                "danger"
            )
            return redirect(url_for("forgot_password"))

        # Update password
        admin.password = generate_password_hash(password)

        db.session.commit()

        flash(
            "Password updated successfully. Please login.",
            "success"
        )

        return redirect(url_for("login"))
    return render_template("forgot_password.html")
# ==========================
# DASHBOARD
# ==========================

@app.route("/dashboard")
@login_required
def dashboard():
    customers = Customer.query.filter_by(user_id=current_user.id).all()
    customers = _sort_customers(customers)

    total_customers = len(customers)
    active_customers = len([c for c in customers if c.status == "Active"])
    closed_customers = len([c for c in customers if c.status == "Closed"])

    total_loan = sum(c.loan_amount for c in customers)
    total_paid = sum(c.total_paid for c in customers)
    total_balance = sum(c.remaining_balance for c in customers)

    today = datetime.now().strftime("%Y-%m-%d")
    current_month = datetime.now().strftime("%Y-%m")

    payments = Payment.query.filter_by(user_id=current_user.id).all()

    today_collection = sum(
        p.amount for p in payments
        if p.payment_date == today
    )

    month_collection = sum(
        p.amount for p in payments
        if p.payment_date.startswith(current_month)
    )

    return render_template(
        "dashboard.html",
        customers=customers,
        total_customers=total_customers,
        active_customers=active_customers,
        closed_customers=closed_customers,
        total_loan=total_loan,
        total_paid=total_paid,
        total_balance=total_balance,
        today_collection=today_collection,
        month_collection=month_collection
    )


# ==========================
# ADD CUSTOMER
# ==========================
@app.route("/add_customer", methods=["POST"])
@login_required
def add_customer():

    customer_id = request.form["customer_id"].strip()

    existing = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first()

    if existing:
        flash(f"Customer ID '{customer_id}' already exists.", "danger")
        return redirect(url_for("dashboard"))

    loan_amount = float(request.form["loan_amount"])
    daily_due = float(request.form["daily_due"])

    start_date = request.form.get("start_date", "")
    end_date = request.form.get("end_date", "")
    address = request.form.get("address", "").strip()

    if not end_date and start_date:
        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d")
            # Fixed 3-month loan term, regardless of loan amount / daily due
            month = start_dt.month - 1 + 3
            year = start_dt.year + month // 12
            month = month % 12 + 1
            day = min(start_dt.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
                                      31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
            end_dt = start_dt.replace(year=year, month=month, day=day)
            end_date = end_dt.strftime("%Y-%m-%d")
        except ValueError:
            end_date = ""

    customer = Customer(
        customer_id=customer_id,
        name=request.form["name"],
        mobile=request.form["mobile"],
        address=address,
        loan_amount=loan_amount,
        daily_due=daily_due,
        total_paid=0,
        remaining_balance=loan_amount,
        status="Active",
        start_date=start_date,
        end_date=end_date,
        user_id=current_user.id
    )

    try:
        db.session.add(customer)
        db.session.commit()
        flash("Customer Added Successfully", "success")

    except IntegrityError:
        db.session.rollback()
        flash(f"Customer ID '{customer_id}' already exists.", "danger")

    except Exception as e:
        db.session.rollback()
        flash(f"Error: {str(e)}", "danger")

    return redirect(url_for("dashboard"))
# ==========================
# ADD PAYMENT
# ==========================

@app.route("/add_payment", methods=["POST"])
@login_required
def add_payment():
    amount = float(request.form["amount"])
    customer_id = request.form["customer_id"]
    payment_date = request.form["payment_date"]

    customer = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first()

    # Check if a payment already exists for this customer on this date.
    # If it does, update it instead of inserting a duplicate row -
    # this prevents the amount from being double-counted when the
    # same id/date is submitted again (accidentally or otherwise).
    existing_payment = Payment.query.filter_by(
        customer_id=customer_id,
        payment_date=payment_date,
        user_id=current_user.id
    ).first()

    if existing_payment:
        old_amount = existing_payment.amount
        existing_payment.amount = amount
        amount_diff = amount - old_amount

        if customer:
            customer.total_paid += amount_diff

        flash("Existing payment for this date was updated (no duplicate created)")
    else:
        payment = Payment(
            customer_id=customer_id,
            payment_date=payment_date,
            amount=amount,
            user_id=current_user.id
        )
        db.session.add(payment)

        if customer:
            customer.total_paid += amount

        flash("Payment Added Successfully")

    if customer:
        customer.remaining_balance = customer.loan_amount - customer.total_paid

        if customer.remaining_balance <= 0:
            customer.remaining_balance = 0
            customer.status = "Closed"
        else:
            customer.status = "Active"

    db.session.commit()

    return redirect(url_for("dashboard"))


# ==========================
# COLLECTION SHEET
# ==========================

@app.route("/collection_sheet")
@login_required
def collection_sheet():
    selected_month = request.args.get(
        "month",
        datetime.now().strftime("%Y-%m")
    )

    customers = Customer.query.filter_by(user_id=current_user.id).all()
    payments = Payment.query.filter_by(user_id=current_user.id).all()

    customers = _sort_customers(customers)

    return render_template(
        "collection_sheet.html",
        customers=customers,
        payments=payments,
        selected_month=selected_month
    )


# ==========================
# DAILY REPORT
# ==========================

@app.route("/daily_report")
@login_required
def daily_report():
    selected_date = request.args.get(
        "date",
        datetime.now().strftime("%Y-%m-%d")
    )

    payments = Payment.query.filter_by(
        payment_date=selected_date,
        user_id=current_user.id
    ).all()

    total_collection = sum(payment.amount for payment in payments)

    return render_template(
        "daily_report.html",
        payments=payments,
        selected_date=selected_date,
        total_collection=total_collection
    )


# ==========================
# PENDING REPORT
# ==========================

@app.route("/pending_report")
@login_required
def pending_report():
    customers = Customer.query.filter(
        Customer.remaining_balance > 0,
        Customer.user_id == current_user.id
    ).all()

    total_pending = sum(
        customer.remaining_balance for customer in customers
    )

    return render_template(
        "pending_report.html",
        customers=customers,
        total_pending=total_pending
    )


# ==========================
# EXPORT DAILY REPORT EXCEL
# ==========================

@app.route("/export_daily_report_excel/<date>")
@login_required
def export_daily_report_excel(date):
    payments = Payment.query.filter_by(
        payment_date=date,
        user_id=current_user.id
    ).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"
    ws.append(["Date", "Customer ID", "Amount"])

    for payment in payments:
        ws.append([payment.payment_date, payment.customer_id, payment.amount])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"Daily_Report_{date}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# EXPORT DAILY REPORT PDF
# ==========================

@app.route("/export_daily_report_pdf/<date>")
@login_required
def export_daily_report_pdf(date):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.pagesizes import A4

    normal_font, bold_font, has_tamil = _register_tamil_font()

    payments = Payment.query.filter_by(
        payment_date=date,
        user_id=current_user.id
    ).all()

    total_collection = sum(p.amount for p in payments)

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    _reg = _FONT_CACHE.get("result", (None, "Helvetica", "Helvetica-Bold", False))
    base_font = _reg[1]   # NotoSans  (Latin)
    hdr_font  = _reg[2]   # NotoSans-Bold

    title_style = ParagraphStyle(
        "DailyTitle",
        parent=styles["Title"],
        fontName=hdr_font
    )
    normal_style = ParagraphStyle(
        "DailyNormal",
        parent=styles["Normal"],
        fontName=base_font
    )

    elements = []
    elements.append(Paragraph(_pdf_text(f"Daily Report - {date}"), title_style))
    elements.append(Spacer(1, 12))

    data = [["No", "Customer ID", "Date", "Amount"]]
    for i, p in enumerate(payments, start=1):
        data.append([
            Paragraph(_pdf_text(str(i)), normal_style),
            Paragraph(_pdf_text(str(p.customer_id)), normal_style),
            Paragraph(_pdf_text(p.payment_date), normal_style),
            Paragraph(_pdf_text(f"RS-{p.amount}"), normal_style)
        ])

    data.append([
        "",
        Paragraph(_pdf_text("TOTAL"), normal_style),
        "",
        Paragraph(_pdf_text(f"RS-{total_collection}"), normal_style)
    ])

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("BACKGROUND", (0, -1), (-1, -1), colors.lightyellow),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), hdr_font),
        ("FONTNAME", (0, -1), (-1, -1), hdr_font),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"Daily_Report_{date}.pdf",
        mimetype="application/pdf"
    )


# ==========================
# EXPORT MONTHLY COLLECTION EXCEL
# ==========================

@app.route("/export_collection/<month>")
@login_required
def export_collection(month):
    wb = Workbook()
    ws = wb.active
    ws.title = "Collection Sheet"

    headers = ["ID", "Name", "Loan"]
    for day in range(1, 32):
        headers.append(str(day))
    headers.extend(["Month Total", "Total Paid", "Balance", "Status"])
    ws.append(headers)

    customers = Customer.query.filter_by(user_id=current_user.id).all()
    customers = _sort_customers(customers)

    for customer in customers:
        row = [customer.customer_id, customer.name, customer.loan_amount]
        month_total = 0

        payments = Payment.query.filter_by(
            customer_id=customer.customer_id,
            user_id=current_user.id
        ).all()

        for day in range(1, 32):
            value = "-"
            for payment in payments:
                if payment.payment_date.startswith(month):
                    payment_day = int(payment.payment_date.split("-")[2])
                    if payment_day == day:
                        value = payment.amount
                        month_total += payment.amount
            row.append(value)

        row.extend([
            month_total,
            customer.total_paid,
            customer.remaining_balance,
            customer.status
        ])
        ws.append(row)

    # DAY TOTAL SUMMARY ROW
    all_payments = Payment.query.filter_by(user_id=current_user.id).all()

    month_total_all = sum(
        p.amount for p in all_payments
        if p.payment_date.startswith(month)
    )
    total_paid_all = sum(c.total_paid for c in customers)
    total_balance_all = sum(c.remaining_balance for c in customers)

    summary_row = ["DAY TOTAL", "", ""]
    for day in range(1, 32):
        day_total = sum(
            p.amount for p in all_payments
            if p.payment_date.startswith(month) and
            int(p.payment_date.split("-")[2]) == day
        )
        summary_row.append(day_total)

    summary_row.extend([month_total_all, total_paid_all, total_balance_all, "-"])
    ws.append(summary_row)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"Monthly_Collection_{month}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# EXPORT MONTHLY COLLECTION PDF
# (Tamil Language Rendering Fixed)
# ==========================

@app.route("/export_collection_pdf/<month>")
@login_required
def export_collection_pdf(month):
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.pagesizes import landscape, A3

    # Register Tamil-capable font
    normal_font, bold_font, has_tamil = _register_tamil_font()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3))
    styles = getSampleStyleSheet()

    _reg = _FONT_CACHE.get("result", (None, "Helvetica", "Helvetica-Bold", False))
    base_font = _reg[1]   # NotoSans  (Latin)
    hdr_font  = _reg[2]   # NotoSans-Bold

    title_style = ParagraphStyle(
        "MonthlyTitle",
        parent=styles["Title"],
        fontName=hdr_font
    )
    normal_style = ParagraphStyle(
        "MonthlyNormal",
        parent=styles["Normal"],
        fontName=base_font,
        fontSize=8
    )

    elements = []
    elements.append(
        Paragraph(_pdf_text(f"Monthly Collection Sheet - {month}"), title_style)
    )
    elements.append(Spacer(1, 10))

    headers = ["ID", "Name", "Loan"]
    for day in range(1, 32):
        headers.append(str(day))
    headers.extend(["Month Total", "Total Paid", "Balance", "Status"])

    data = [headers]

    customers = Customer.query.filter_by(user_id=current_user.id).all()
    customers = _sort_customers(customers)

    for customer in customers:
        # _pdf_text() switches font per-character: Tamil->NotoSansTamil, Latin->NotoSans
        name_cell = Paragraph(_pdf_text(customer.name), normal_style)
        row = [customer.customer_id, name_cell, customer.loan_amount]
        month_total = 0

        payments = Payment.query.filter_by(
            customer_id=customer.customer_id,
            user_id=current_user.id
        ).all()

        for day in range(1, 32):
            value = "-"
            for payment in payments:
                if payment.payment_date.startswith(month):
                    payment_day = int(payment.payment_date.split("-")[2])
                    if payment_day == day:
                        value = payment.amount
                        month_total += payment.amount
            row.append(value)

        row.extend([
            month_total,
            customer.total_paid,
            customer.remaining_balance,
            customer.status
        ])
        data.append(row)

    # DAY TOTAL ROW
    all_payments = Payment.query.filter_by(user_id=current_user.id).all()

    total_row = ["DAY TOTAL", "", ""]
    for day in range(1, 32):
        day_total = sum(
            p.amount for p in all_payments
            if p.payment_date.startswith(month) and
            int(p.payment_date.split("-")[2]) == day
        )
        total_row.append(day_total)

    month_total_all = sum(
        p.amount for p in all_payments
        if p.payment_date.startswith(month)
    )
    total_paid_all = sum(c.total_paid for c in customers)
    total_balance_all = sum(c.remaining_balance for c in customers)

    total_row.extend([month_total_all, total_paid_all, total_balance_all, "-"])
    data.append(total_row)

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("BACKGROUND", (0, -1), (-1, -1), colors.lightgreen),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), bold_font),
        ("FONTNAME", (0, -1), (-1, -1), bold_font),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"Monthly_Collection_{month}.pdf",
        mimetype="application/pdf"
    )


# ==========================
# CUSTOMER LEDGER
# ==========================

@app.route("/customer_ledger")
@login_required
def customer_ledger():
    customers = Customer.query.filter_by(user_id=current_user.id).all()
    customers = _sort_customers(customers)
    return render_template("customer_ledger.html", customers=customers)


# ==========================
# EXPORT CUSTOMER LEDGER EXCEL
# (Fixed: now filters by user_id)
# ==========================

@app.route("/export_customers")
@login_required
def export_customers():
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"

    headers = [
        "Customer ID", "Name", "Mobile", "Loan Amount",
        "Daily Due", "Total Paid", "Balance", "Status"
    ]
    ws.append(headers)

    # FIX: filter by current user (was missing user_id filter)
    customers = Customer.query.filter_by(user_id=current_user.id).all()
    customers = _sort_customers(customers)

    for customer in customers:
        ws.append([
            customer.customer_id,
            customer.name,
            customer.mobile,
            customer.loan_amount,
            customer.daily_due,
            customer.total_paid,
            customer.remaining_balance,
            customer.status
        ])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="customer_ledger.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# CUSTOMER DETAILS
# ==========================

@app.route("/customer/<customer_id>")
@login_required
def customer_details(customer_id):
    customer = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first_or_404()

    payments = Payment.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).all()

    # Calculate day-based fields from start_date and end_date
    total_days = "-"
    days_passed = "-"
    remaining_days = "-"

    try:
        if customer.start_date and customer.end_date:
            fmt = "%Y-%m-%d"
            start = datetime.strptime(str(customer.start_date), fmt)
            end   = datetime.strptime(str(customer.end_date),   fmt)
            today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

            # Inclusive counting: the start date itself counts as Day 1.
            total_days = (end - start).days + 1

            if today < start:
                # Loan hasn't started yet
                days_passed = 0
            else:
                days_passed = (today - start).days + 1
                days_passed = min(days_passed, total_days)

            remaining_days = max(0, total_days - days_passed)
    except (ValueError, TypeError):
        pass  # leave defaults as "-" if dates are missing / malformed

    return render_template(
        "customer_details.html",
        customer=customer,
        payments=payments,
        total_days=total_days,
        days_passed=days_passed,
        remaining_days=remaining_days
    )


# ==========================
# EDIT CUSTOMER
# ==========================

@app.route("/edit_customer/<customer_id>", methods=["GET", "POST"])
@login_required
def edit_customer(customer_id):
    customer = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first_or_404()

    if request.method == "POST":
        customer.name        = request.form["name"]
        customer.mobile      = request.form["mobile"]
        customer.address     = request.form.get("address", customer.address or "")
        customer.loan_amount = float(request.form["loan_amount"])
        customer.daily_due   = float(request.form["daily_due"])
        customer.end_date    = request.form.get("end_date", customer.end_date or "")
        customer.remaining_balance = customer.loan_amount - customer.total_paid

        db.session.commit()
        flash("Customer Updated Successfully")
        return redirect(url_for("customer_ledger"))

    return render_template("edit_customer.html", customer=customer)


# ==========================
# DELETE CUSTOMER
# ==========================

@app.route("/delete_customer/<customer_id>")
@login_required
def delete_customer(customer_id):
    customer = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first()

    if customer:
        Payment.query.filter_by(
            customer_id=customer_id,
            user_id=current_user.id
        ).delete()

        db.session.delete(customer)
        db.session.commit()
        flash("Customer Deleted Successfully")

    return redirect(url_for("customer_ledger"))


# ==========================
# EDIT PAYMENT
# ==========================

@app.route("/edit_payment/<int:payment_id>", methods=["GET", "POST"])
@login_required
def edit_payment(payment_id):
    payment = Payment.query.filter_by(
        id=payment_id,
        user_id=current_user.id
    ).first_or_404()

    customer = Customer.query.filter_by(
        customer_id=payment.customer_id,
        user_id=current_user.id
    ).first()

    if request.method == "POST":
        old_amount = payment.amount
        new_amount = float(request.form["amount"])
        new_date = request.form["payment_date"]

        # If moving this payment to a date that already has a different
        # payment for the same customer, merge into that one instead of
        # creating a second row for the same day.
        duplicate = Payment.query.filter(
            Payment.customer_id == payment.customer_id,
            Payment.payment_date == new_date,
            Payment.user_id == current_user.id,
            Payment.id != payment.id
        ).first()

        if duplicate:
            duplicate.amount += new_amount
            customer.total_paid = customer.total_paid - old_amount + new_amount
            db.session.delete(payment)
            flash("Merged into existing payment for that date")
        else:
            payment.payment_date = new_date
            payment.amount = new_amount
            customer.total_paid = customer.total_paid - old_amount + new_amount
            flash("Payment Updated Successfully")

        customer.remaining_balance = customer.loan_amount - customer.total_paid

        if customer.remaining_balance <= 0:
            customer.remaining_balance = 0
            customer.status = "Closed"
        else:
            customer.status = "Active"

        db.session.commit()
        return redirect(url_for("customer_details", customer_id=customer.customer_id))

    return render_template("edit_payment.html", payment=payment)


# ==========================
# DELETE PAYMENT
# ==========================

@app.route("/delete_payment/<int:payment_id>")
@login_required
def delete_payment(payment_id):
    payment = Payment.query.filter_by(
        id=payment_id,
        user_id=current_user.id
    ).first_or_404()

    customer = Customer.query.filter_by(
        customer_id=payment.customer_id,
        user_id=current_user.id
    ).first()

    customer.total_paid -= payment.amount
    customer.remaining_balance = customer.loan_amount - customer.total_paid

    if customer.remaining_balance <= 0:
        customer.remaining_balance = 0
        customer.status = "Closed"
    else:
        customer.status = "Active"

    db.session.delete(payment)
    db.session.commit()

    flash("Payment Deleted Successfully")
    return redirect(url_for("customer_details", customer_id=customer.customer_id))


# ==========================
# EXPORT SINGLE CUSTOMER EXCEL
# ==========================

@app.route("/export_customer/<customer_id>")
@login_required
def export_customer(customer_id):
    customer = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first_or_404()

    payments = Payment.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Customer Statement"

    ws.append(["Customer ID", customer.customer_id])
    ws.append(["Name", customer.name])
    ws.append(["Mobile", customer.mobile])
    ws.append(["Loan Amount", customer.loan_amount])
    ws.append(["Daily Due", customer.daily_due])
    ws.append(["Total Paid", customer.total_paid])
    ws.append(["Balance", customer.remaining_balance])
    ws.append(["Status", customer.status])
    ws.append([])
    ws.append(["Payment History"])
    ws.append(["Date", "Amount"])

    for payment in payments:
        ws.append([payment.payment_date, payment.amount])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"{customer.customer_id}_statement.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# CUSTOMER STATEMENT PDF
# (Tamil Language Rendering Fixed)
# ==========================

@app.route("/customer_statement_pdf/<customer_id>")
@login_required
def customer_statement_pdf(customer_id):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.pagesizes import A4

    customer = Customer.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).first_or_404()

    payments = Payment.query.filter_by(
        customer_id=customer_id,
        user_id=current_user.id
    ).all()

    # Register both Tamil and Latin fonts
    normal_font, bold_font, has_tamil = _register_tamil_font()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    # Use NotoSans (Latin) as base font; _pdf_text() switches to NotoSansTamil
    # per-character for Tamil Unicode ranges, so both scripts render correctly.
    # _register_tamil_font returns (tamil_font, bold_font, has_tamil);
    # retrieve the latin font name directly from the cache for the base style.
    _reg = _FONT_CACHE.get("result", (normal_font, normal_font, bold_font, has_tamil))
    base_font = _reg[1]   # latin_font (NotoSans)
    hdr_font  = _reg[2]   # bold font  (NotoSans-Bold)

    title_style = ParagraphStyle(
        "StmtTitle",
        parent=styles["Title"],
        fontName=hdr_font
    )
    normal_style = ParagraphStyle(
        "StmtNormal",
        parent=styles["Normal"],
        fontName=base_font
    )

    elements = []
    elements.append(Paragraph(_pdf_text("CUSTOMER STATEMENT"), title_style))
    elements.append(Spacer(1, 12))

    # _pdf_text() wraps Tamil chars with NotoSansTamil, Latin stays NotoSans
    info_lines = [
        _pdf_text(f"Customer ID: {customer.customer_id}"),
        _pdf_text(f"Name: {customer.name}"),
        _pdf_text(f"Mobile: {customer.mobile}"),
        _pdf_text(f"Loan Amount: RS-{customer.loan_amount}"),
        _pdf_text(f"Total Paid: RS-{customer.total_paid}"),
        _pdf_text(f"Balance: RS-{customer.remaining_balance}"),
    ]
    for line in info_lines:
        elements.append(Paragraph(line, normal_style))

    elements.append(Spacer(1, 15))

    data = [["No", "Date", "Amount"]]
    for index, payment in enumerate(payments, start=1):
        data.append([
            Paragraph(_pdf_text(str(index)), normal_style),
            Paragraph(_pdf_text(payment.payment_date), normal_style),
            Paragraph(_pdf_text(f"RS-{payment.amount}"), normal_style)
        ])

    table = Table(data)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), hdr_font),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{customer.customer_id}_statement.pdf",
        mimetype="application/pdf"
    )


# ==========================
# BACKUP DATABASE
# ==========================

@app.route("/backup_database")
@login_required
def backup_database():
    # NOTE: the app now runs on PostgreSQL (Supabase), so there is no local
    # .db file to send anymore. Export all of this user's data to Excel
    # instead, with one sheet per table, so backups actually work.
    wb = Workbook()

    ws_customers = wb.active
    ws_customers.title = "Customers"
    ws_customers.append([
        "Customer ID", "Name", "Mobile", "Address", "Loan Amount",
        "Daily Due", "Total Paid", "Remaining Balance", "Status",
        "Start Date", "End Date"
    ])
    customers = Customer.query.filter_by(user_id=current_user.id).all()
    for c in customers:
        ws_customers.append([
            c.customer_id, c.name, c.mobile, c.address, c.loan_amount,
            c.daily_due, c.total_paid, c.remaining_balance, c.status,
            c.start_date, c.end_date
        ])

    ws_payments = wb.create_sheet("Payments")
    ws_payments.append(["Customer ID", "Payment Date", "Amount"])
    payments = Payment.query.filter_by(user_id=current_user.id).all()
    for p in payments:
        ws_payments.append([p.customer_id, p.payment_date, p.amount])

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    backup_name = f"finance_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=backup_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================
# RESTORE DATABASE
# ==========================

@app.route("/restore_database", methods=["GET", "POST"])
@login_required
def restore_database():
    if request.method == "POST":
        backup_file = request.files.get("backup_file")

        if backup_file and backup_file.filename.endswith(".xlsx"):
            from openpyxl import load_workbook
            try:
                wb = load_workbook(backup_file, data_only=True)

                # Wipe this user's existing data before restoring
                Payment.query.filter_by(user_id=current_user.id).delete()
                Customer.query.filter_by(user_id=current_user.id).delete()

                if "Customers" in wb.sheetnames:
                    ws = wb["Customers"]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if not row or row[0] is None:
                            continue
                        (customer_id, name, mobile, address, loan_amount,
                         daily_due, total_paid, remaining_balance, status,
                         start_date, end_date) = row
                        db.session.add(Customer(
                            customer_id=str(customer_id),
                            name=name,
                            mobile=mobile,
                            address=address,
                            loan_amount=loan_amount or 0,
                            daily_due=daily_due or 0,
                            total_paid=total_paid or 0,
                            remaining_balance=remaining_balance or 0,
                            status=status or "Active",
                            start_date=start_date,
                            end_date=end_date,
                            user_id=current_user.id
                        ))

                if "Payments" in wb.sheetnames:
                    ws = wb["Payments"]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if not row or row[0] is None:
                            continue
                        customer_id, payment_date, amount = row
                        db.session.add(Payment(
                            customer_id=str(customer_id),
                            payment_date=str(payment_date),
                            amount=amount or 0,
                            user_id=current_user.id
                        ))

                db.session.commit()
                flash("Database Restored Successfully")
            except Exception as e:
                db.session.rollback()
                flash(f"Restore failed: {str(e)}", "danger")

            return redirect(url_for("dashboard"))

        flash("Please upload a valid .xlsx backup file")

    return render_template("restore_database.html")


# ==========================
# COMPANY SETTINGS
# ==========================

@app.route("/company_settings", methods=["GET", "POST"])
@login_required
def company_settings():
    settings = CompanySettings.query.filter_by(user_id=current_user.id).first()

    if not settings:
        settings = CompanySettings(user_id=current_user.id)
        db.session.add(settings)
        db.session.commit()

    if request.method == "POST":
        settings.company_name = request.form["company_name"]
        settings.address = request.form["address"]
        settings.phone = request.form["phone"]

        db.session.commit()
        flash("Settings Saved Successfully")
        return redirect(url_for("company_settings"))

    return render_template("company_settings.html", settings=settings)


# ==========================
# LOGOUT
# ==========================

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))


# ==========================
# DATABASE INIT
# ==========================

with app.app_context():
    db.create_all()

    admin = Admin.query.filter_by(username="admin").first()

    if not admin:

        admin = Admin(
            username="admin",
            mobile="9999999999",
            password=generate_password_hash("admin123")
        )

        db.session.add(admin)
        db.session.commit()


# ==========================
# RUN APP
# ==========================


    if __name__ == "__main__":
        app.run(host="0.0.0.0", port=5000)