import mysql.connector
from mysql.connector import Error
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# === CONFIGURATION ===
HOST = "localhost"
USER = "root"
PASSWORD = "" #mysql password
DATABASE = "" #Database name

# === CONNECT TO DATABASE ===
def connect_db():
    try:
        conn = mysql.connector.connect(
            host=HOST,
            port=4000,
            user=USER,
            password=PASSWORD,
            database=DATABASE
        )
        return conn
    except Error as e:
        messagebox.showerror("Database Error", str(e))
        return None

# === CREATE TABLE ===
def create_table():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS property_records (
        id INT AUTO_INCREMENT PRIMARY KEY,
        date DATE,
        time TIME,
        client VARCHAR(255),
        clerk VARCHAR(255),
        property_address VARCHAR(255),
        inv_type VARCHAR(255),
        status VARCHAR(255)
    )
    """)
    conn.commit()
    conn.close()

# === ADD RECORD ===
def add_record():
    now = datetime.now()
    date = now.date()
    time = now.time().replace(microsecond=0)

    values = [entry_client.get(), entry_clerk.get(), entry_address.get(), entry_type.get(), entry_status.get()]
    if not all(values):
        messagebox.showwarning("Input Error", "Please fill in all fields.")
        return

    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
    INSERT INTO property_records (date, time, client, clerk, property_address, inv_type, status)
    VALUES (%s, %s, %s, %s, %s, %s, %s)
    """, (date, time, *values))
    conn.commit()
    conn.close()
    messagebox.showinfo("Success", "Record added successfully!")
    fetch_all()

# === EDIT RECORD ===
def edit_record():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select a record", "Please select a record to edit.")
        return

    record_id = tree.item(selected[0])['values'][0]
    column = combo_column.get()
    new_value = entry_new_value.get()

    if not column or not new_value:
        messagebox.showwarning("Input Error", "Please select a column and enter a new value.")
        return

    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute(f"UPDATE property_records SET {column} = %s WHERE id = %s", (new_value, record_id))
    conn.commit()
    conn.close()
    messagebox.showinfo("Success", "Record updated.")
    fetch_all()

# === DELETE RECORD ===
def delete_record():
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("Select a record", "Please select a record to delete.")
        return

    record_id = tree.item(selected[0])['values'][0]

    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM property_records WHERE id = %s", (record_id,))
    conn.commit()
    conn.close()
    messagebox.showinfo("Deleted", "Record deleted successfully.")
    fetch_all()

# === FETCH ALL RECORDS ===
def fetch_all():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM property_records")
    rows = cursor.fetchall()
    update_tree(rows)
    conn.close()

# === SEARCH RECORDS ===
def search_records():
    client = search_client.get()
    clerk = search_clerk.get()
    address = search_address.get()

    query = """
    SELECT * FROM property_records
    WHERE client LIKE %s AND clerk LIKE %s AND property_address LIKE %s
    """
    values = [f"%{client}%", f"%{clerk}%", f"%{address}%"]

    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute(query, values)
    rows = cursor.fetchall()
    update_tree(rows)
    conn.close()

# === UPDATE TREEVIEW ===
def update_tree(rows):
    for row in tree.get_children():
        tree.delete(row)
    for row in rows:
        tree.insert("", tk.END, values=row)

# === GUI SETUP ===
root = tk.Tk()
root.title("Property Record Manager")

create_table()

# === Input Frame ===
input_frame = tk.LabelFrame(root, text="Add New Record")
input_frame.pack(fill="x", padx=10, pady=5)

tk.Label(input_frame, text="Client").grid(row=0, column=0)
entry_client = tk.Entry(input_frame)
entry_client.grid(row=0, column=1)

tk.Label(input_frame, text="Clerk").grid(row=0, column=2)
entry_clerk = tk.Entry(input_frame)
entry_clerk.grid(row=0, column=3)

tk.Label(input_frame, text="Address").grid(row=1, column=0)
entry_address = tk.Entry(input_frame)
entry_address.grid(row=1, column=1)

tk.Label(input_frame, text="Inventory Type").grid(row=1, column=2)
entry_type = tk.Entry(input_frame)
entry_type.grid(row=1, column=3)

tk.Label(input_frame, text="Status").grid(row=2, column=0)
entry_status = tk.Entry(input_frame)
entry_status.grid(row=2, column=1)

tk.Button(input_frame, text="Add Record", command=add_record).grid(row=2, column=3, padx=5, pady=5)

# === Edit/Delete Frame ===
edit_frame = tk.LabelFrame(root, text="Edit / Delete Record")
edit_frame.pack(fill="x", padx=10, pady=5)

tk.Label(edit_frame, text="Column").grid(row=0, column=0)
combo_column = ttk.Combobox(edit_frame, values=["client", "clerk", "property_address", "inv_type", "status"])
combo_column.grid(row=0, column=1)

tk.Label(edit_frame, text="New Value").grid(row=0, column=2)
entry_new_value = tk.Entry(edit_frame)
entry_new_value.grid(row=0, column=3)

tk.Button(edit_frame, text="Edit Record", command=edit_record).grid(row=0, column=4)
tk.Button(edit_frame, text="Delete Record", command=delete_record).grid(row=0, column=5)

# === Search Frame ===
search_frame = tk.LabelFrame(root, text="Search Records")
search_frame.pack(fill="x", padx=10, pady=5)

tk.Label(search_frame, text="Client").grid(row=0, column=0)
search_client = tk.Entry(search_frame)
search_client.grid(row=0, column=1)

tk.Label(search_frame, text="Clerk").grid(row=0, column=2)
search_clerk = tk.Entry(search_frame)
search_clerk.grid(row=0, column=3)

tk.Label(search_frame, text="Address").grid(row=0, column=4)
search_address = tk.Entry(search_frame)
search_address.grid(row=0, column=5)

tk.Button(search_frame, text="Search", command=search_records).grid(row=0, column=6)
tk.Button(search_frame, text="Fetch All", command=fetch_all).grid(row=0, column=7)

# === Treeview ===
tree = ttk.Treeview(root, columns=("ID", "Date", "Time", "Client", "Clerk", "Address", "Type", "Status"), show="headings")
for col in tree["columns"]:
    tree.heading(col, text=col)
    tree.column(col, width=100)
tree.pack(fill="both", expand=True, padx=10, pady=10)

fetch_all()
root.mainloop()
