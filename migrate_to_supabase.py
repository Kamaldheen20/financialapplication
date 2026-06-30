import sqlite3
from app import app
from models import db, Admin, Customer, Payment

# SQLite database
sqlite_conn = sqlite3.connect("finance.db")
sqlite_conn.row_factory = sqlite3.Row

cursor = sqlite_conn.cursor()

with app.app_context():

    print("Starting migration...")

    # --------------------------
    # ADMINS
    # --------------------------

    cursor.execute("SELECT * FROM admins")

    for row in cursor.fetchall():

        exists = Admin.query.filter_by(
            username=row["username"]
        ).first()

        if exists:
            continue

        admin = Admin(
            id=row["id"],
            username=row["username"],
            password=row["password"]
        )

        db.session.add(admin)

        db.session.commit()

        print("Admins migrated.")

        # --------------------------
        # CUSTOMERS
        # --------------------------

        cursor.execute("SELECT * FROM customers")

        for row in cursor.fetchall():

            exists = Customer.query.filter_by(
                customer_id=row["customer_id"]
            ).first()

            if exists:
                continue

            customer = Customer(
                id=row["id"],
                customer_id=row["customer_id"],
                address=row["address"],
                name=row["name"],
                mobile=row["mobile"],
                loan_amount=row["loan_amount"],
                daily_due=row["daily_due"],
                total_paid=row["total_paid"],
                remaining_balance=row["remaining_balance"],
                status=row["status"],
                start_date=row["start_date"],
                end_date=row["end_date"],
                user_id=row["user_id"]
            )

            db.session.add(customer)

            db.session.commit()

            print("Customers migrated.")

            # --------------------------
            # PAYMENTS
            # --------------------------

            cursor.execute("SELECT * FROM payments")

for row in cursor.fetchall():

    try:

        exists = Payment.query.filter_by(
            id=row["id"]
        ).first()

        if exists:
            continue

        payment = Payment(
            id=row["id"],
            customer_id=row["customer_id"],
            payment_date=row["payment_date"],
            amount=row["amount"],
            user_id=row["user_id"]
        )

        db.session.add(payment)
        db.session.commit()

    except Exception as e:

        db.session.rollback()

        print(f"Error migrating payment ID {row['id']}")
        print(e)
        break

    print("Payments migrated.")