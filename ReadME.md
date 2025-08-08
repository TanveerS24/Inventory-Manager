📋 InventoryHouse

InventoryHouse is a desktop-based inventory and property record management system built using Python, Tkinter, and MySQL. It allows clerks to add, search, view, edit, and manage inventory/property records efficiently through an intuitive GUI.

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

✅ Features

🧾 Record Management
- Add new inventory records using a popup form with dropdowns for Clerk and Status.
- View all records in a well-formatted Table with auto-numbering and alternating row colors.
- Hide internal record ID from user view while maintaining it internally for editing/deletion.

🔍 Search & Filter
- Search for records based on:
  - Client
  - Clerk
  - Property Address
- Supports partial matches using SQL `LIKE %...%`.
- Clear search filters using the "Clear Filters" button.

🛠️ Edit & Delete
- Select a row to:
  - Edit details using a popup (with dropdowns for Clerk and Status).
  - Delete the record after confirmation.
- Status field becomes non-editable if already marked as completed.

🖼️ Paste Photos
- Generate a Photo Index Document (.docx) by selecting a folder of images.
- Inserts photos in a 4x2 grid layout with captions (Photo 001, Photo 002, ...).
- Automatically marks the status as **completed** after successful photo document creation.

📦 Database & Schema
- Automatically creates the MySQL database (`inventory`) and table (`property_records`) if they do not exist.
- Uses a local MySQL server.

🎨 User Interface
- Soft color theme with clear layout:
  - Title at the top
  - Table view in the center
  - **Table Actions** (Add, Search, Clear Filters) at the bottom
  - **Row Actions** (Edit, Delete, Paste Photos) enabled on selection
- Scrollbars included for better navigation.

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

🛠️ Technologies Used
- Python (Tkinter for GUI)
- MySQL (via `mysql-connector-python`)
- python-docx (for Word document generation)
- Pillow (for image support if needed)
- OS, datetime (core Python libraries)

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

⚙️ Requirements

Install the required Python libraries:
```bash
pip install mysql-connector-python python-docx
```

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

🚀 Getting Started

1. Clone or download the project.
2. Run 
	'''ALTER USER 'root'@'localhost' IDENTIFIED WITH caching_sha2_password BY 'your_password';
	   FLUSH PRIVILEGES;
	''' 
   in MySQL
3. Make sure to add in your root, user, password in 'inventory house.py' and change the template path
4. Find the locale using python in line 12
5. Re-bundle the app if necessary using the command 
	"""pyinstaller --onefile --noconsole ^
	   More? --add-data "<your locale>" ^
	   More? InventoryHouse.py"""
6. Your app will be located in 'dist' folder
6. Start managing your inventory records!

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

🧑‍💻 Author

Tanveer S. 
Built with ❤️ using Python.