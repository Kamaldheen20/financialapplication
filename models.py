from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin

db = SQLAlchemy()


# ==========================
# ADMIN TABLE
# ==========================

class Admin(UserMixin, db.Model):
    __tablename__ = "admins"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    username = db.Column(
        db.String(100),
        unique=True,
        nullable=False
    )

    password = db.Column(
        db.String(255),
        nullable=False
    )


# ==========================
# CUSTOMER TABLE
# ==========================

class Customer(db.Model):
    __tablename__ = "customers"
    __table_args__ = (
    db.UniqueConstraint(
        "customer_id",
        "user_id",
        name="uq_customer_user"
    ),
)

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    customer_id = db.Column(
    db.String(50),
    nullable=False
)
    address = db.Column(
        db.Text
    )

    name = db.Column(
        db.String(200),
        nullable=False
    )

    mobile = db.Column(
        db.String(20)
    )

    loan_amount = db.Column(
        db.Float,
        default=0
    )

    daily_due = db.Column(
        db.Float,
        default=0
    )

    total_paid = db.Column(
        db.Float,
        default=0
    )

    remaining_balance = db.Column(
        db.Float,
        default=0
    )

    status = db.Column(
        db.String(20),
        default="Active"
    )

    start_date = db.Column(
        db.String(50)
    )
    end_date = db.Column(
        db.String(50)
    )

    # Owner of this customer
    user_id = db.Column(
        db.Integer,
        db.ForeignKey("admins.id"),
        nullable=False
    )


# ==========================
# PAYMENT TABLE
# ==========================

class Payment(db.Model):
    __tablename__ = "payments"

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    customer_id = db.Column(
        db.String(50),
        nullable=False
    )

    payment_date = db.Column(
        db.String(50),
        nullable=False
    )

    amount = db.Column(
        db.Float,
        default=0
    )

    # Owner of this payment
    user_id = db.Column(
        db.Integer,
        db.ForeignKey("admins.id"),
        nullable=False
    )
    