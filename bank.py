import random
from datetime import datetime
import os
import csv
try:
    import pandas as pd
except Exception:
    pd = None


def save_banking_to_excel(banking, filename="bank_data.xlsx"):
    """Save banking dict to an Excel file using pandas with Sr. No."""
    if pd is None:
        return

    rows = []
    for acc_no, u in banking.items():
        rows.append({
            "acc_no": acc_no,
            "name": u.get("name", ""),
            "fathers_name": u.get("fathers_name", ""),
            "mother_name": u.get("mother_name", ""),
            "dob": u.get("dob", ""),
            "balance": u.get("balance", 0),
            # note: transactions are intentionally NOT saved to Excel
        })

    df = pd.DataFrame(rows)

    # ADD SR. NO. COLUMN
    df.insert(0, "sr_no", range(1, len(df) + 1))

    df.to_excel(filename, index=False)



def load_banking_from_excel(filename="bank_data.xlsx"):
    """Load banking dict from Excel file if it exists. Returns dict."""
    banking = {}
    if pd is None:
        return banking
    if not os.path.exists(filename):
        return banking
    try:
        df = pd.read_excel(filename, dtype={"acc_no": str})
    except Exception:
        return banking
    for _, row in df.iterrows():
        acc_no = str(row.get("acc_no", "")).strip()
        if not acc_no:
            continue
        # transactions may not exist in the Excel file; keep in-memory list empty
        tx_list = []
        banking[acc_no] = {
            "name": row.get("name", ""),
            "fathers_name": row.get("fathers_name", ""),
            "mother_name": row.get("mother_name", ""),
            "dob": row.get("dob", ""),
            "balance": float(row.get("balance", 0) or 0),
            "transactions": tx_list,
        }
    return banking

def generate_account_no(banking):
    while True:
        acc_no = str(random.randint(100000, 999999))
        if acc_no not in banking:
            return acc_no


def _ensure_csv_header(filename, headers):
    """Ensure CSV file exists with header row."""
    # If file doesn't exist, create it with the provided headers
    if not os.path.exists(filename):
        try:
            with open(filename, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
        except Exception:
            pass
        return

    # If file exists, ensure the header matches. If it doesn't, rewrite
    # the file with the requested header and backfill an `sr_no` column
    # for any existing data rows.
    try:
        with open(filename, mode="r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)
    except Exception:
        rows = []

    # If empty file or no rows, write the header
    if not rows:
        try:
            with open(filename, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
        except Exception:
            pass
        return

    existing_header = rows[0]
    # If header already matches exactly, nothing to do
    if existing_header == headers:
        return

    # Treat first row as header and the rest as data rows. Backfill sr_no
    data_rows = rows[1:]
    new_rows = [headers]
    for i, r in enumerate(data_rows, start=1):
        # Ensure row has enough columns for remaining headers
        needed = len(headers) - 1
        if len(r) < needed:
            r = r + [""] * (needed - len(r))
        else:
            r = r[:needed]
        new_rows.append([str(i)] + r)

    try:
        with open(filename, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(new_rows)
    except Exception:
        pass


def _next_csv_sr_no(filename):
    """Return next serial number for given CSV (1-based)."""
    if not os.path.exists(filename):
        return 1
    try:
        with open(filename, mode="r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            # skip header
            rows = list(reader)
            if not rows:
                return 1
            return max(1, len(rows))  # header + N rows -> next sr = len(rows)
    except Exception:
        return 1


def log_deposit(acc_no, amount, banking, filename="deposits.csv"):
    """Append a deposit record to deposits.csv."""
    date_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    user = banking.get(acc_no, {})
    name = user.get("name", "")
    balance = user.get("balance", 0)
    headers = ["sr_no", "datetime", "acc_no", "name", "amount", "balance_after"]
    _ensure_csv_header(filename, headers)
    try:
        sr = _next_csv_sr_no(filename)
        with open(filename, mode="a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([sr, date_time, acc_no, name, amount, balance])
    except Exception:
        pass


def log_withdrawal(acc_no, amount, banking, filename="withdrawals.csv"):
    """Append a withdrawal record to withdrawals.csv."""
    date_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    user = banking.get(acc_no, {})
    name = user.get("name", "")
    balance = user.get("balance", 0)
    headers = ["sr_no", "datetime", "acc_no", "name", "amount", "balance_after"]
    _ensure_csv_header(filename, headers)
    try:
        sr = _next_csv_sr_no(filename)
        with open(filename, mode="a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([sr, date_time, acc_no, name, amount, balance])
    except Exception:
        pass

def bank():
    print("Welcome to the bank!")
    # load existing data from Excel if available
    banking = load_banking_from_excel() if pd is not None else {}

    while True:
        print("\n===== Bank Menu =====")
        print("1. Create Account")
        print("2. Deposit Money")
        print("3. Withdraw Money")
        print("4. Check Balance")
        print("5. Statement")
        print("6. Update Details")
        print("7. Cheque Deposit")
        print("8. Check no. of Accounts")
        print("9. Exit")

        choice = input("Enter your choice (1-8): ")

        # 1. Create Account
        if choice == "1":

            # Name
            while True:
                name = input("Enter your name: ")
                if name.replace(" ", "").isalpha():
                    break
                else:
                    print("Please enter only characters.")

            # Father's Name
            while True:
                fathers_name = input("Enter your father's name: ")
                if fathers_name.replace(" ", "").isalpha():
                    break
                else:
                    print("Please enter only characters.")

            # Mother's Name
            while True:
                mother_name = input("Enter your mother's name: ")
                if mother_name.replace(" ", "").isalpha():
                    break
                else:
                    print("Please enter only characters.")

            # DOB
            while True:
                dob = input("Enter DOB (DD/MM/YYYY): ")
                try:
                    datetime.strptime(dob, "%d/%m/%Y")
                    break
                except ValueError:
                    print("Invalid date format.")

            acc_no = generate_account_no(banking)
            banking[acc_no] = {
                "name": name,
                "fathers_name": fathers_name,
                "mother_name": mother_name,
                "dob": dob,
                "balance": 0,
                "transactions": []
            }

            print("\nAccount created successfully!")
            print("Account Number:", acc_no)
            if pd is not None:
                save_banking_to_excel(banking)

        # 2. Deposit
        elif choice == "2":
            acc_no = input("Enter account number: ")
            if acc_no in banking:
                amount = float(input("Enter amount: "))
                if amount > 0:
                    banking[acc_no]["balance"] += amount
                    date_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    banking[acc_no]["transactions"].append(
                        f"{date_time} | Deposited ₹{amount}"
                    )
                    print("Balance:", banking[acc_no]["balance"])
                    if pd is not None:
                        save_banking_to_excel(banking)
                    # log deposit to CSV report
                    try:
                        log_deposit(acc_no, amount, banking)
                    except Exception:
                        pass
                else:
                    print("Amount must be positive.")
            else:
                print("Invalid account number.")

        # 3. Withdraw
        elif choice == "3":
            acc_no = input("Enter account number: ")
            if acc_no in banking:
                amount = float(input("Enter amount: "))
                if 0 < amount <= banking[acc_no]["balance"]:
                    banking[acc_no]["balance"] -= amount
                    date_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    banking[acc_no]["transactions"].append(
                        f"{date_time} | Withdrawn ₹{amount}"
                    )
                    print("Balance:", banking[acc_no]["balance"])
                    if pd is not None:
                        save_banking_to_excel(banking)
                    # log withdrawal to CSV report
                    try:
                        log_withdrawal(acc_no, amount, banking)
                    except Exception:
                        pass
                else:
                    print("Insufficient balance.")
            else:
                print("Invalid account number.")

        # 4. Balance
        elif choice == "4":
            acc_no = input("Enter account number: ")
            if acc_no in banking:
                print("Balance:", banking[acc_no]["balance"])
            else:
                print("Invalid account number.")

        # 5. Statement
        elif choice == "5":
            acc_no = input("Enter account number: ")
            if acc_no in banking:
                print("\nStatement")
                for t in banking[acc_no]["transactions"]:
                    print(t)
            else:
                print("Invalid account number.")

        # 6. Customer Details
        elif choice == "6":
            acc_no = input("Enter account number: ")
            if acc_no in banking:
                u = banking[acc_no]
                print("\n Customer Details")
                print("Name:", u["name"])
                print("Father:", u["fathers_name"])
                print("Mother:", u["mother_name"])
                print("DOB:", u["dob"])
                print("Balance:", u["balance"])
                if input("Update details? (y/n): ").strip().lower() == 'y':
                    new_name = input("New name (leave blank to keep): ")
                    if new_name.strip():
                        u['name'] = new_name.strip()
                    new_father = input("New father's name (leave blank to keep): ")
                    if new_father.strip():
                        u['fathers_name'] = new_father.strip()
                    new_mother = input("New mother's name (leave blank to keep): ")
                    if new_mother.strip():
                        u['mother_name'] = new_mother.strip()
                    new_dob = input("New DOB (DD/MM/YYYY) (leave blank to keep): ")
                    if new_dob.strip():
                        try:
                            datetime.strptime(new_dob, "%d/%m/%Y")
                            u['dob'] = new_dob.strip()
                        except ValueError:
                            print("Invalid date format; DOB not updated.")
                    if pd is not None:
                        save_banking_to_excel(banking)
            else:
                print("Account not found.")

        # 7. Cheque Deposit
        elif choice == "7":
            acc_no = input("Enter account number: ")
            if acc_no in banking:
                amount = float(input("Enter amount: "))
                if amount > 0:
                    banking[acc_no]["balance"] += amount
                    date_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
                    banking[acc_no]["transactions"].append(
                        f"{date_time} | Cheque Deposited ₹{amount}"
                    )
                    print("Balance:", banking[acc_no]["balance"])
                    if pd is not None:
                        save_banking_to_excel(banking)
                    # log cheque deposit to CSV report
                    try:
                        log_deposit(acc_no, amount, banking)
                    except Exception:
                        pass
                else:
                    print("Amount must be positive.")
            else:
                print("Invalid account number.")
        # 8. check no. of accounts
        elif choice == "8":
            if not banking:
                print(" No accounts found.")
            else:
                print("\n ALL ACCOUNT DETAILS")
                print("-" * 60)
            for acc_no, u in banking.items():
                print(f"Account No   : {acc_no}")
                print(f"Name         : {u['name']}")
                print(f"Father Name  : {u['fathers_name']}")
                print(f"Mother Name  : {u['mother_name']}")
                print(f"DOB          : {u['dob']}")
                print(f"Balance      : ₹{u['balance']}")
                print("-" * 60)

            
        # 9. Exit
        elif choice == "9":
            print("Thank you for using bank system ")
            if pd is not None:
                save_banking_to_excel(banking)
            break

        else:
            print("Invalid choice")

bank()
