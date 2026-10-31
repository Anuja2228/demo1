import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import webbrowser


# =========================================================
# EXCEL FILE
# =========================================================

FILE_NAME = "Monthly_Budget.xlsx"
SHEET_NAME = "Budget"

# Put your Power BI dashboard URL between the quotes.
POWER_BI_LINK = r"D:\power BI project\Monthly Budget planner\MonthlyBudgetPB.pbix"

COLUMNS = [
    "Budget ID",
    "Date",
    "Month",
    "Type",
    "Category",
    "Description",
    "Planned Amount",
    "Actual Amount",
    "Payment Mode"
]
# =========================================================
# CREATE EXCEL FILE IF IT DOES NOT EXIST
# =========================================================

def create_excel_file():
    if not os.path.exists(FILE_NAME):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(FILE_NAME, sheet_name=SHEET_NAME, index=False)

def create_excel_if_not_exists():
    """Create Excel file if it doesn't already exist."""

    if not os.path.exists(FILE_NAME):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(FILE_NAME, index=False)

def create_powerbi_if_not_exists():
    """Create PowerBI file if it doesn't already exist."""

    if not os.path.exists(POWER_BI_LINK):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(FILE_NAME, index=False)



# =========================================================
# GET NEXT BUDGET ID AUTOMATICALLY
# =========================================================

def get_next_budget_id():

    create_excel_file()

    try:
        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME)

        if df.empty:
            return "B001"

        ids = df["Budget ID"].astype(str)

        numbers = []

        for value in ids:
            try:
                number = int(value.replace("B", ""))
                numbers.append(number)
            except:
                pass

        if numbers:
            next_number = max(numbers) + 1
        else:
            next_number = 1

        return f"B{next_number:03d}"

    except Exception as e:
        messagebox.showerror("Error", f"Could not generate Budget ID:\n{e}")
        return "B001"


# =========================================================
# CLEAR FORM
# =========================================================

def clear_form():

    budget_id_var.set(get_next_budget_id())
    date_var.set("")
    month_var.set("")
    type_var.set("Expense")
    category_var.set("Food")
    description_var.set("")
    planned_var.set("")
    actual_var.set("")
    payment_var.set("UPI")

    date_entry.focus()


# =========================================================
# ADD RECORD
# =========================================================

def add_record():

    try:

        budget_id = budget_id_var.get()
        date = date_var.get()
        month = month_var.get()
        record_type = type_var.get()
        category = category_var.get()
        description = description_var.get()
        planned = planned_var.get()
        actual = actual_var.get()
        payment = payment_var.get()

        if date == "":
            messagebox.showwarning("Missing Data", "Please enter Date.")
            return

        if month == "":
            messagebox.showwarning("Missing Data", "Please enter Month.")
            return

        if description == "":
            messagebox.showwarning("Missing Data", "Please enter Description.")
            return

        if planned == "":
            planned = 0

        if actual == "":
            actual = 0

        planned = float(planned)
        actual = float(actual)

        create_excel_file()

        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME)

        new_record = pd.DataFrame([{
            "Budget ID": budget_id,
            "Date": date,
            "Month": month,
            "Type": record_type,
            "Category": category,
            "Description": description,
            "Planned Amount": planned,
            "Actual Amount": actual,
            "Payment Mode": payment
        }])

        df = pd.concat([df, new_record], ignore_index=True)

        df.to_excel(FILE_NAME, sheet_name=SHEET_NAME, index=False)

        messagebox.showinfo(
            "Success",
            f"Record {budget_id} added successfully!"
        )

        clear_form()

    except ValueError:
        messagebox.showerror(
            "Invalid Amount",
            "Planned Amount and Actual Amount must be numbers."
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================================================
# UPDATE RECORD
# =========================================================

def update_record():

    try:

        budget_id = budget_id_var.get()

        if budget_id == "":
            messagebox.showwarning(
                "Missing ID",
                "Please enter Budget ID."
            )
            return

        create_excel_file()

        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME)

        if budget_id not in df["Budget ID"].astype(str).values:

            messagebox.showwarning(
                "Not Found",
                f"Budget ID {budget_id} was not found."
            )

            return

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Date"
        ] = date_var.get()

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Month"
        ] = month_var.get()

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Type"
        ] = type_var.get()

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Category"
        ] = category_var.get()

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Description"
        ] = description_var.get()

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Planned Amount"
        ] = float(planned_var.get())

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Actual Amount"
        ] = float(actual_var.get())

        df.loc[
            df["Budget ID"].astype(str) == budget_id,
            "Payment Mode"
        ] = payment_var.get()

        df.to_excel(FILE_NAME, sheet_name=SHEET_NAME, index=False)

        messagebox.showinfo(
            "Success",
            f"Record {budget_id} updated successfully!"
        )

        clear_form()

    except ValueError:
        messagebox.showerror(
            "Invalid Amount",
            "Please enter valid numbers."
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))

# =========================================================
# DELETE RECORD
# =========================================================

def delete_record():

    try:

        budget_id = budget_id_var.get()

        if budget_id == "":
            messagebox.showwarning(
                "Missing ID",
                "Please enter Budget ID."
            )
            return

        create_excel_file()

        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME)

        if budget_id not in df["Budget ID"].astype(str).values:

            messagebox.showwarning(
                "Not Found",
                f"Budget ID {budget_id} was not found."
            )

            return

        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete {budget_id}?"
        )

        if confirm:

            df = df[
                df["Budget ID"].astype(str) != budget_id
            ]

            df.to_excel(
                FILE_NAME,
                sheet_name=SHEET_NAME,
                index=False
            )

            messagebox.showinfo(
                "Deleted",
                f"Record {budget_id} deleted successfully!"
            )

            clear_form()

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================================================
# VIEW RECORDS
# =========================================================

def view_records():

    create_excel_file()

    try:

        df = pd.read_excel(FILE_NAME, sheet_name=SHEET_NAME)

        if df.empty:
            messagebox.showinfo(
                "No Records",
                "There are no records in the Excel file."
            )
            return

        view_window = tk.Toplevel(root)

        view_window.title("Budget Records")
        view_window.geometry("1100x500")

        frame = tk.Frame(view_window)
        frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(
            frame,
            columns=COLUMNS,
            show="headings"
        )

        for column in COLUMNS:

            tree.heading(
                column,
                text=column
            )

            tree.column(
                column,
                width=120
            )

        for _, row in df.iterrows():

            tree.insert(
                "",
                "end",
                values=list(row)
            )

        scrollbar_y = ttk.Scrollbar(
            frame,
            orient="vertical",
            command=tree.yview
        )

        scrollbar_x = ttk.Scrollbar(
            frame,
            orient="horizontal",
            command=tree.xview
        )

        tree.configure(
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )

        tree.pack(
            side="top",
            fill="both",
            expand=True
        )

        scrollbar_y.pack(
            side="right",
            fill="y"
        )

        scrollbar_x.pack(
            side="bottom",
            fill="x"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))
# =========================================================
# SEARCH RECORD
# =========================================================

def search_record():

    budget_id = budget_id_var.get()

    if budget_id == "":
        messagebox.showwarning(
            "Search",
            "Enter Budget ID to search."
        )
        return

    try:

        df = pd.read_excel(
            FILE_NAME,
            sheet_name=SHEET_NAME
        )

        result = df[
            df["Budget ID"].astype(str) == budget_id
        ]

        if result.empty:

            messagebox.showinfo(
                "Not Found",
                f"No record found for {budget_id}."
            )

            return

        row = result.iloc[0]

        date_var.set(str(row["Date"]))
        month_var.set(str(row["Month"]))
        type_var.set(str(row["Type"]))
        category_var.set(str(row["Category"]))
        description_var.set(str(row["Description"]))
        planned_var.set(str(row["Planned Amount"]))
        actual_var.set(str(row["Actual Amount"]))
        payment_var.set(str(row["Payment Mode"]))

    except Exception as e:
        messagebox.showerror("Error", str(e))


# =========================================================
# OPEN EXCEL
# =========================================================

def open_excel():

    create_excel_file()

    try:
        os.startfile(FILE_NAME)

    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Could not open Excel:\n{e}"
        )


# =========================================================
# OPEN POWER BI
# =========================================================

def open_power_bi():

    create_powerbi_if_not_exists()

    try:
        os.startfile(os.path.abspath(POWER_BI_LINK))

    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Unable to open Power BI dashboard.\n\n{e}"
        )

# =========================================================
# MAIN WINDOW
# =========================================================

root = tk.Tk()

root.title("Monthly Budget Planner")
root.geometry("850x650")

root.resizable(False, False)


# =========================================================
# VARIABLES
# =========================================================

budget_id_var = tk.StringVar()
date_var = tk.StringVar()
month_var = tk.StringVar()
type_var = tk.StringVar(value="Expense")
category_var = tk.StringVar(value="Food")
description_var = tk.StringVar()
planned_var = tk.StringVar()
actual_var = tk.StringVar()
payment_var = tk.StringVar(value="UPI")


# =========================================================
# TITLE
# =========================================================

title = tk.Label(
    root,
    text="MONTHLY BUDGET PLANNER",
    font=("Arial", 24, "bold")
)

title.pack(pady=20)


# =========================================================
# FORM FRAME
# =========================================================

form_frame = tk.Frame(root)

form_frame.pack(pady=10)


# Budget ID
tk.Label(
    form_frame,
    text="Budget ID:",
    font=("Arial", 11)
).grid(row=0, column=0, padx=10, pady=8, sticky="e")

budget_id_entry = tk.Entry(
    form_frame,
    textvariable=budget_id_var,
    width=30
)

budget_id_entry.grid(
    row=0,
    column=1,
    padx=10,
    pady=8
)


# Date
tk.Label(
    form_frame,
    text="Date:",
    font=("Arial", 11)
).grid(row=1, column=0, padx=10, pady=8, sticky="e")

date_entry = tk.Entry(
    form_frame,
    textvariable=date_var,
    width=30
)

date_entry.grid(
    row=1,
    column=1,
    padx=10,
    pady=8
)


# Month
tk.Label(
    form_frame,
    text="Month:",
    font=("Arial", 11)
).grid(row=2, column=0, padx=10, pady=8, sticky="e")

month_combo = ttk.Combobox(
    form_frame,
    textvariable=month_var,
    values=[
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December"
    ],
    width=27,
    state="readonly"
)

month_combo.grid(
    row=2,
    column=1,
    padx=10,
    pady=8
)


# Type
tk.Label(
    form_frame,
    text="Type:",
    font=("Arial", 11)
).grid(row=3, column=0, padx=10, pady=8, sticky="e")

type_combo = ttk.Combobox(
    form_frame,
    textvariable=type_var,
    values=["Income", "Expense"],
    width=27,
    state="readonly"
)

type_combo.grid(
    row=3,
    column=1,
    padx=10,
    pady=8
)


# Category
tk.Label(
    form_frame,
    text="Category:",
    font=("Arial", 11)
).grid(row=4, column=0, padx=10, pady=8, sticky="e")

category_combo = ttk.Combobox(
    form_frame,
    textvariable=category_var,
    values=[
        "Food",
        "Groceries",
        "Transport",
        "Rent",
        "Electricity",
        "Mobile/Internet",
        "Education",
        "Shopping",
        "Entertainment",
        "Medical",
        "Savings",
        "Other"
    ],
    width=27,
    state="readonly"
)

category_combo.grid(
    row=4,
    column=1,
    padx=10,
    pady=8
)


# Description
tk.Label(
    form_frame,
    text="Description:",
    font=("Arial", 11)
).grid(row=5, column=0, padx=10, pady=8, sticky="e")

description_entry = tk.Entry(
    form_frame,
    textvariable=description_var,
    width=30
)

description_entry.grid(
    row=5,
    column=1,
    padx=10,
    pady=8
)


# Planned Amount
tk.Label(
    form_frame,
    text="Planned Amount:",
    font=("Arial", 11)
).grid(row=6, column=0, padx=10, pady=8, sticky="e")

planned_entry = tk.Entry(
    form_frame,
    textvariable=planned_var,
    width=30
)

planned_entry.grid(
    row=6,
    column=1,
    padx=10,
    pady=8
)


# Actual Amount
tk.Label(
    form_frame,
    text="Actual Amount:",
    font=("Arial", 11)
).grid(row=7, column=0, padx=10, pady=8, sticky="e")

actual_entry = tk.Entry(
    form_frame,
    textvariable=actual_var,
    width=30
)

actual_entry.grid(
    row=7,
    column=1,
    padx=10,
    pady=8
)


# Payment Mode
tk.Label(
    form_frame,
    text="Payment Mode:",
    font=("Arial", 11)
).grid(row=8, column=0, padx=10, pady=8, sticky="e")

payment_combo = ttk.Combobox(
    form_frame,
    textvariable=payment_var,
    values=[
        "Cash",
        "UPI",
        "Debit Card",
        "Credit Card",
        "Bank Transfer"
    ],
    width=27,
    state="readonly"
)

payment_combo.grid(
    row=8,
    column=1,
    padx=10,
    pady=8
)

# =========================================================
# BUTTON FRAME
# =========================================================

button_frame = tk.Frame(root)

button_frame.pack(pady=20)


tk.Button(
    button_frame,
    text="Add Record",
    width=15,
    command=add_record
).grid(row=0, column=0, padx=5, pady=5)


tk.Button(
    button_frame,
    text="Update",
    width=15,
    command=update_record
).grid(row=0, column=1, padx=5, pady=5)


tk.Button(
    button_frame,
    text="Delete",
    width=15,
    command=delete_record
).grid(row=0, column=2, padx=5, pady=5)


tk.Button(
    button_frame,
    text="Search",
    width=15,
    command=search_record
).grid(row=0, column=3, padx=5, pady=5)


tk.Button(
    button_frame,
    text="Clear",
    width=15,
    command=clear_form
).grid(row=1, column=0, padx=5, pady=5)


tk.Button(
    button_frame,
    text="View Records",
    width=15,
    command=view_records
).grid(row=1, column=1, padx=5, pady=5)


tk.Button(
    button_frame,
    text="Open Excel",
    width=15,
    command=open_excel
).grid(row=1, column=2, padx=5, pady=5)


tk.Button(
    button_frame,
    text="Open Power BI",
    width=15,
    command=open_power_bi
).grid(row=1, column=3, padx=5, pady=5)


# =========================================================
# START APPLICATION
# =========================================================

create_excel_file()

budget_id_var.set(get_next_budget_id())

root.mainloop()