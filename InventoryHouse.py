import mysql.connector
from mysql.connector import Error
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT as WDO
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT as WDPA
import os

#print(os.path.join(os.path.dirname(mysql.connector.__file__), 'locales'))

# === CONFIGURATION ===
HOST = "localhost"
USER = "root"
PASSWORD = ""
DATABASE = "inventory"

CLERK_OPTIONS = ["Tom Tyrrel", "Kevin Crack", "Bill West"]
STATUS_OPTIONS = ["Inspected", "Audio Recorded"]

import mysql.connector
from mysql.connector import Error

def check_db():
    connection = None
    try:
        connection = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password=''
        )

        if connection and connection.is_connected():
            cursor = connection.cursor()
            cursor.execute("CREATE DATABASE IF NOT EXISTS inventory")
            return True
    except Error as e:
        print("Error while connecting to MySQL:", e)
    finally:
        if connection and connection.is_connected():
            cursor.close()
            connection.close()

def connect_db():
    try:
        check_db()
        conn = mysql.connector.connect(
            host=HOST,
            port=3306,
            user=USER,
            password=PASSWORD,
            database=DATABASE
        )
        return conn
    except Error as e:
        messagebox.showerror("Database Error", str(e))
        return None

def create_table():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS property_records (
        id INT AUTO_INCREMENT PRIMARY KEY,
        date DATE,
        clerk VARCHAR(255),
        property_address VARCHAR(255),
        client VARCHAR(255),
        inv_type VARCHAR(255),
        status VARCHAR(255)
    )
    """)
    conn.commit()
    conn.close()

def add_record_popup():
    popup = tk.Toplevel(root)
    popup.title("Add New Record")

    tk.Label(popup, text="Clerk").grid(row=0, column=0)
    clerk_cb = ttk.Combobox(popup, values=CLERK_OPTIONS, state="readonly")
    clerk_cb.grid(row=0, column=1)

    tk.Label(popup, text="Address").grid(row=1, column=0)
    addr_entry = tk.Entry(popup)
    addr_entry.grid(row=1, column=1)

    tk.Label(popup, text="Client").grid(row=2, column=0)
    client_entry = tk.Entry(popup)
    client_entry.grid(row=2, column=1)

    tk.Label(popup, text="Inventory Type").grid(row=3, column=0)
    inv_entry = tk.Entry(popup)
    inv_entry.grid(row=3, column=1)

    tk.Label(popup, text="Status").grid(row=4, column=0)
    status_cb = ttk.Combobox(popup, values=STATUS_OPTIONS, state="readonly")
    status_cb.grid(row=4, column=1)

    def submit():
        values = [clerk_cb.get(), addr_entry.get(), client_entry.get(), inv_entry.get(), status_cb.get()]
        if not all(values):
            messagebox.showwarning("Input Error", "Please fill in all fields.")
            return
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO property_records (date, clerk, property_address, client, inv_type, status)
        VALUES (%s, %s, %s, %s, %s, %s)
        """, (datetime.now().date(), *values))
        conn.commit()
        conn.close()
        popup.destroy()
        fetch_all()
        messagebox.showinfo("Success", "Record added successfully!")

    tk.Button(popup, text="Add Record", command=submit).grid(row=5, columnspan=2, pady=10)

def search_popup():
    popup = tk.Toplevel(root)
    popup.title("Search Records")

    tk.Label(popup, text="Client").grid(row=0, column=0)
    client_entry = tk.Entry(popup)
    client_entry.grid(row=0, column=1)

    tk.Label(popup, text="Clerk").grid(row=1, column=0)
    clerk_cb = ttk.Combobox(popup, values=CLERK_OPTIONS, state="readonly")
    clerk_cb.grid(row=1, column=1)

    tk.Label(popup, text="Address").grid(row=2, column=0)
    address_entry = tk.Entry(popup)
    address_entry.grid(row=2, column=1)

    def search():
        client = client_entry.get()
        clerk = clerk_cb.get()
        address = address_entry.get()
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("""
        SELECT * FROM property_records
        WHERE client LIKE %s AND clerk LIKE %s AND property_address LIKE %s
        """, (f"%{client}%", f"%{clerk}%", f"%{address}%"))
        rows = cursor.fetchall()
        conn.close()
        formatted = [(i+1, row[0], row[1].strftime("%d-%m-%Y"), *row[2:]) for i, row in enumerate(rows)]
        update_tree(formatted)
        btn_clear_filters.config(state="normal")
        popup.destroy()

    tk.Button(popup, text="Search", command=search).grid(row=3, columnspan=2, pady=10)

def fetch_all():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("select * from property_records;")
    rows = cursor.fetchall()
    conn.close()
    formatted = [(i+1, row[0], row[1].strftime("%d-%m-%Y"), *row[2:]) for i, row in enumerate(rows)]
    update_tree(formatted)

def update_tree(rows):
    for item in tree.get_children():
        tree.delete(item)
    for row in rows:
        tree.insert("", tk.END, values=row)
    reset_action_buttons()

def delete_record(record_id):
    if messagebox.askyesno("Confirm", "Delete this record?"):
        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM property_records WHERE id = %s", (record_id,))
        conn.commit()
        conn.close()
        fetch_all()

def open_edit_popup(record):
    popup = tk.Toplevel(root)
    popup.title("Edit Record")

    is_completed = record[7] == "completed"  # Status value is in column index 7

    # Dropdown for Clerk
    tk.Label(popup, text="Clerk").grid(row=0, column=0, padx=5, pady=5)
    clerk_cb = ttk.Combobox(popup, values=CLERK_OPTIONS, state="readonly")
    clerk_cb.set(record[3])
    clerk_cb.grid(row=0, column=1, padx=5, pady=5)

    # Entry for Address
    tk.Label(popup, text="Address").grid(row=1, column=0, padx=5, pady=5)
    address_entry = tk.Entry(popup)
    address_entry.insert(0, record[4])
    address_entry.grid(row=1, column=1, padx=5, pady=5)

    # Entry for Client
    tk.Label(popup, text="Client").grid(row=2, column=0, padx=5, pady=5)
    client_entry = tk.Entry(popup)
    client_entry.insert(0, record[5])
    client_entry.grid(row=2, column=1, padx=5, pady=5)

    # Entry for Inventory Type
    tk.Label(popup, text="Inventory Type").grid(row=3, column=0, padx=5, pady=5)
    inv_entry = tk.Entry(popup)
    inv_entry.insert(0, record[6])
    inv_entry.grid(row=3, column=1, padx=5, pady=5)

    # Dropdown for Status
    tk.Label(popup, text="Status").grid(row=4, column=0, padx=5, pady=5)
    status_cb = ttk.Combobox(popup, values=STATUS_OPTIONS + ["completed"], state="readonly" if not is_completed else "disabled")
    status_cb.set(record[7])
    status_cb.grid(row=4, column=1, padx=5, pady=5)

    def save_changes():
        clerk = clerk_cb.get()
        address = address_entry.get().strip()
        client = client_entry.get().strip()
        inv_type = inv_entry.get().strip()
        status = record[7] if is_completed else status_cb.get()

        if not all([clerk, address, client, inv_type, status]):
            messagebox.showwarning("Input Error", "Please fill in all fields.")
            return

        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("""
            UPDATE property_records 
            SET clerk = %s, property_address = %s, client = %s, inv_type = %s, status = %s
            WHERE id = %s
        """, (clerk, address, client, inv_type, status, record[1]))
        conn.commit()
        conn.close()

        popup.destroy()
        fetch_all()
        messagebox.showinfo("Updated", "Record updated successfully.")

    tk.Button(popup, text="Save", command=save_changes).grid(row=5, columnspan=2, pady=10)

def paste_photos(record_id):
    source_folder = filedialog.askdirectory(title="Select Image Folder")
    if not source_folder:
        return
    try:
        template_path = r""
        output_file = "Photo_gallery.docx"
        images_per_page = 8
        images_per_row = 4
        image_width = Cm(5.85)
        image_height = Cm(6.11)

        doc = Document(template_path)

        landscape_section = doc.add_section(start_type=1)
        landscape_section.orientation = WDO.LANDSCAPE
        landscape_section.page_width, landscape_section.page_height = (
            landscape_section.page_height, landscape_section.page_width
        )

        for idx, sec in enumerate(doc.sections):
            if idx != 0:
                sec.orientation = WDO.LANDSCAPE
                sec.page_width, sec.page_height = sec.page_height, sec.page_width

        image_files = sorted([
            f for f in os.listdir(source_folder)
            if f.lower().endswith(('.png', '.jpg', '.jpeg'))
        ])

        photo_counter = 1

        for start in range(0, len(image_files), images_per_page):
            if start != 0:
                doc.add_page_break()

            heading = doc.add_paragraph()
            heading.alignment = WDPA.LEFT
            heading.paragraph_format.space_after = Pt(12)
            heading_run = heading.add_run("PHOTO INDEX")
            heading_run.bold = True
            heading_run.font.size = Pt(20)

            table = doc.add_table(rows=4, cols=images_per_row)
            table.autofit = True

            for i in range(images_per_page):
                idx = start + i
                if idx >= len(image_files):
                    break

                row = (i // images_per_row) * 2
                col = i % images_per_row
                img_path = os.path.join(source_folder, image_files[idx])

                img_cell = table.cell(row, col)
                img_para = img_cell.paragraphs[0]
                img_para.paragraph_format.space_after = Pt(0)
                img_para.alignment = WDPA.CENTER
                run = img_para.add_run()
                run.add_picture(img_path, width=image_width, height=image_height)

                caption_cell = table.cell(row + 1, col)
                caption_para = caption_cell.paragraphs[0]
                caption_para.alignment = WDPA.CENTER
                caption_para.paragraph_format.space_before = Pt(10)
                caption_para.paragraph_format.space_after = Pt(10)
                caption_run = caption_para.add_run(f"Photo {photo_counter:03d}")
                caption_run.font.size = Pt(12)

                photo_counter += 1

        if os.path.exists(output_file):
            os.remove(output_file)

        doc.save(output_file)

        conn = connect_db()
        cursor = conn.cursor()
        cursor.execute("UPDATE property_records SET status = %s WHERE id = %s", ("completed", record_id))
        conn.commit()
        conn.close()
        fetch_all()
        messagebox.showinfo("Success", "Photo document saved and status updated.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def clear_filters():
    fetch_all()
    btn_clear_filters.config(state="disabled")

# === GUI SETUP ===
root = tk.Tk()
root.title("InventoryHouse")
root.configure(bg="#9ECAD6")
root.geometry("1150x700")

title_label = tk.Label(root, text="InventoryHouse", font=("Helvetica", 24, "bold"), bg="#9ECAD6", foreground='red')
title_label.pack(pady=10)

# Treeview + Scrollbar
style = ttk.Style()
style.theme_use('clam') 
style.configure("Treeview.Heading", background="#98A1BC", foreground="#555879", font=('Segoe UI', 15, 'bold'))
tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
tree_y_scroll = tk.Scrollbar(tree_frame, orient="vertical")
tree_y_scroll.pack(side="right", fill="y")
tree_x_scroll = tk.Scrollbar(tree_frame, orient="horizontal", command=lambda *args: tree.xview(*args))
tree_x_scroll.pack(side="bottom", fill="x")

tree = ttk.Treeview(tree_frame, columns=("No", "ID", "Date", "Clerk", "Address", "Client", "Type", "Status"),
                    show="headings", yscrollcommand=tree_y_scroll.set, xscrollcommand=tree_x_scroll.set)

tree.tag_configure('oddrow', background='white', font=('Segoe UI', 12))
tree.tag_configure('evenrow', background='lightgray', font=('Segoe UI', 12))

def update_tree(rows):
    for item in tree.get_children():
        tree.delete(item)
    for i, row in enumerate(rows):
        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
        tree.insert("", tk.END, values=row, tags=(tag,))
    reset_action_buttons()

tree_y_scroll.config(command=tree.yview)
tree_x_scroll.config(command=tree.xview)
for col in tree["columns"]:
    tree.heading(col, text=col)
    if col == "ID":
        tree.column(col, width=0, stretch=False)
    else:
        tree.column(col, width=120, anchor="center")
tree.pack(fill="both", expand=True)
tree.column("No", width=30, anchor="center")
tree.column("ID", width=0, stretch=False)
tree.column("Date", width=70, anchor="center")
tree.column("Clerk", width=80, anchor="center")
tree.column("Address", width=350, anchor="center", stretch=True)
tree.column("Client", width=100, anchor="center")
tree.column("Type", width=70, anchor="center")
tree.column("Status", width=90, anchor="center")
tree.bind("<<TreeviewSelect>>", lambda e: on_row_select(e))

# Table Actions
action_frame = tk.LabelFrame(root, text="Table Actions", bg="#DFF6FF", fg="#3C5B6F", font=("Arial", 17, "bold"))
action_frame.pack(fill="x", padx=10, pady=5)

tk.Button(action_frame, text="Add Record", bg="#748DAE", fg="#2C3E50", font=("Arial", 14, "bold"), command=add_record_popup).pack(side="left", padx=5)
tk.Button(action_frame, text="Search", bg="#748DAE", fg="#2C3E50", font=("Arial", 14, "bold"), command=search_popup).pack(side="left", padx=5)

btn_clear_filters = tk.Button(action_frame, text="Clear Filters", state="disabled", bg="#748DAE", fg="#2C3E50", font=("Arial", 14, "bold"), command=lambda: clear_filters())
btn_clear_filters.pack(side="left", padx=5)

# Row Actions
row_action_frame = tk.LabelFrame(root, text="Selected Row Actions", bg="#DFF6FF", fg="#3C5B6F", font=("Arial", 17, "bold"))
row_action_frame.pack(fill="x", padx=10, pady=5)

btn_edit = tk.Button(row_action_frame, text="Edit", state="disabled", bg="#748DAE", fg="#2C3E50", font=("Arial", 14, "bold"))
btn_edit.pack(side="left", padx=5)

btn_delete = tk.Button(row_action_frame, text="Delete", state="disabled", bg="#748DAE", fg="#2C3E50", font=("Arial", 14, "bold"))
btn_delete.pack(side="left", padx=5)

btn_photos = tk.Button(row_action_frame, text="Paste Photos", state="disabled", bg="#748DAE", fg="#2C3E50", font=("Arial", 14, "bold"))
btn_photos.pack(side="left", padx=5)

selected_record = None
def on_row_select(event):
    global selected_record
    selected_item = tree.selection()
    if selected_item:
        selected_record = tree.item(selected_item[0])['values']
        btn_edit.config(state="normal", command=lambda: open_edit_popup(selected_record))

        # Disable Paste Photos if status is 'completed'
        if selected_record[7] == "completed":
            btn_photos.config(state="disabled")
        else:
            btn_photos.config(state="normal", command=lambda: paste_photos(selected_record[1]))

        btn_delete.config(state="normal", command=lambda: delete_record(selected_record[1]))
    else:
        reset_action_buttons()

def reset_action_buttons():
    btn_edit.config(state="disabled")
    btn_delete.config(state="disabled")
    btn_photos.config(state="disabled")

# === Start App ===
create_table()
fetch_all()
root.mainloop()
