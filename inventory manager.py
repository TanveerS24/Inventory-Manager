import mysql.connector
from mysql.connector import Error
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# === CONFIGURATION ===
HOST = "localhost"
USER = "root"
PASSWORD = "tanveer"
DATABASE = "inventory"

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

# === DELETE RECORD WITH CONFIRMATION ===
def delete_record(record_id):
    confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this record?")
    if confirm:
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM property_records WHERE id = %s", (record_id,))
        conn.commit()
        conn.close()
        messagebox.showinfo("Deleted", "Record deleted successfully.")
        fetch_all()

# === EDIT RECORD POPUP ===
def open_edit_popup(record):
    popup = tk.Toplevel(root)
    popup.title("Edit Record")

    fields = ["client", "clerk", "property_address", "inv_type", "status"]
    entries = {}

    for idx, field in enumerate(fields):
        tk.Label(popup, text=field.capitalize()).grid(row=idx, column=0, padx=5, pady=5)
        entry = tk.Entry(popup)
        entry.grid(row=idx, column=1, padx=5, pady=5)
        entries[field] = entry

    def save_changes():
        updates = {}
        for field, entry in entries.items():
            value = entry.get().strip()
            if value:
                updates[field] = value
        if updates:
            conn = connect_db()
            cursor = conn.cursor()
            for col, val in updates.items():
                cursor.execute(f"UPDATE property_records SET {col} = %s WHERE id = %s", (val, record[0]))
            conn.commit()
            conn.close()
            messagebox.showinfo("Updated", "Record updated successfully.")
            popup.destroy()
            fetch_all()
        else:
            messagebox.showinfo("No Changes", "No values entered for update.")

    tk.Button(popup, text="Save", command=save_changes).grid(row=len(fields), column=0, columnspan=2, pady=10)

# === FETCH AND DISPLAY RECORDS ===
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

    reset_action_buttons()

# === Handle Treeview Selection ===
selected_record = None

def on_row_select(event):
    global selected_record
    selected_item = tree.selection()
    if selected_item:
        selected_record = tree.item(selected_item[0])['values']
        btn_edit.config(state="normal", command=lambda: open_edit_popup(selected_record))
        btn_delete.config(state="normal", command=lambda: delete_record(selected_record[0]))
    else:
        reset_action_buttons()

def reset_action_buttons():
    global selected_record
    selected_record = None
    btn_edit.config(state="disabled", command=None)
    btn_delete.config(state="disabled", command=None)

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
tree.bind("<<TreeviewSelect>>", on_row_select)

# === Action Buttons ===
action_frame = tk.LabelFrame(root, text="Selected Row Actions")
action_frame.pack(fill="x", padx=10, pady=5)

btn_edit = tk.Button(action_frame, text="Edit", state="disabled")
btn_edit.pack(side="left", padx=10)

btn_delete = tk.Button(action_frame, text="Delete", state="disabled")
btn_delete.pack(side="left", padx=10)

# === Initial Load ===
fetch_all()
root.mainloop()
