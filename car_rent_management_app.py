import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from PIL import Image, ImageTk
import tkinter.messagebox as messagebox
import openpyxl
import shutil
import os

current_image_path = None
selected_serial_number = None


def browse_image():
    global current_image_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
    if file_path:
        current_image_path = file_path
        load_image(file_path)
    else:
        hide_preview()


def save_customer():
    customer_name = name_entry.get()
    pending_amount = pending_entry.get()
    advance_amount = advance_entry.get()
    start_date = start_entry.get()
    end_date = end_entry.get()

    # Check if any field is empty
    if not (customer_name and pending_amount and advance_amount and start_date and end_date):
        messagebox.showerror("Error", "Please fill in all the fields.")
        return

    save_dir = "Customer Details Folder"
    excel_file_path = os.path.join(save_dir, "customer_details.xlsx")
    images_dir = "Car Images Folder"

    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    if not os.path.exists(images_dir):
        os.makedirs(images_dir)

    if os.path.isfile(excel_file_path):
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Customer Details"
        sheet["A1"] = "Serial Number"
        sheet["B1"] = "Customer Name"
        sheet["C1"] = "Pending Amount"
        sheet["D1"] = "Advance Amount"
        sheet["E1"] = "Start Date"
        sheet["F1"] = "End Date"

    # Get the existing serial number or start from 1
    serial_number = sheet.max_row if sheet.max_row > 1 else 1

    row = (serial_number, customer_name, pending_amount,
           advance_amount, start_date, end_date)
    sheet.append(row)

    workbook.save(excel_file_path)

    if current_image_path:
        image_file_name = f"{customer_name}_{serial_number}.png"
        shutil.copyfile(current_image_path, os.path.join(
            images_dir, image_file_name))
        messagebox.showinfo("Success", "Customer details saved successfully.")
    else:
        messagebox.showwarning("Warning", "No image selected.")

    name_entry.delete(0, tk.END)
    pending_entry.delete(0, tk.END)
    advance_entry.delete(0, tk.END)
    start_entry.delete(0, tk.END)
    end_entry.delete(0, tk.END)
    hide_preview()

    # Update table with saved customer details
    update_table()


def update_table():
    excel_file_path = os.path.join(
        "Customer Details Folder", "customer_details.xlsx")
    images_dir = "Car Images Folder"

    if not os.path.isfile(excel_file_path):
        return

    # Clear existing items in the table
    table_view.delete(*table_view.get_children())

    # Load customer details from Excel
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    # Configure the table columns
    table_view.column("#0", width=100, anchor="center")
    table_view.column("Serial Number", width=100, anchor="center")
    table_view.column("Customer Name", width=150, anchor="center")
    table_view.column("Pending Amount", width=150, anchor="center")
    table_view.column("Advance Amount", width=150, anchor="center")
    table_view.column("Start Date", width=150, anchor="center")
    table_view.column("End Date", width=150, anchor="center")

    # Add details to the table
    for row in sheet.iter_rows(min_row=2, values_only=True):
        serial_number = row[0]
        customer_name = row[1]
        pending_amount = row[2]
        advance_amount = row[3]
        start_date = row[4]
        end_date = row[5]

        # Add details to the table
        table_view.insert("", tk.END, values=(
            serial_number, customer_name, pending_amount, advance_amount, start_date, end_date))

    # Configure the table to allow editing only for rows with the "editable" tag
    table_view.tag_configure(
        "editable", background="#D3D3D3", foreground="black")


def load_image(file_path):
    global image_preview
    image = Image.open(file_path)
    image = image.resize((200, 200))
    image_preview = ImageTk.PhotoImage(image)
    preview_label.configure(image=image_preview)
    preview_label.image = image_preview
    show_preview()


def show_preview():
    preview_label.grid(row=6, column=0, columnspan=2, padx=10, pady=5)


def hide_preview():
    preview_label.grid_remove()


# Create the main window
root = tk.Tk()
root.title("Car Rental Management System")

# Set a custom background color
root.configure(bg="#90EE90")

# Create a frame for the customer details
details_frame = tk.Frame(root, bg="#f2f2f2")
details_frame.pack(padx=10, pady=10)

# Customer Details
name_label = tk.Label(details_frame, text="Customer Name:",
                      font=("Arial", 12), bg="#f2f2f2")
name_label.grid(row=0, column=0, sticky="w")
name_entry = tk.Entry(details_frame, font=("Arial", 12))
name_entry.grid(row=0, column=1, padx=10)

pending_label = tk.Label(
    details_frame, text="Pending Amount:", font=("Arial", 12), bg="#f2f2f2")
pending_label.grid(row=1, column=0, sticky="w")
pending_entry = tk.Entry(details_frame, font=("Arial", 12))
pending_entry.grid(row=1, column=1, padx=10)

advance_label = tk.Label(
    details_frame, text="Advance Amount:", font=("Arial", 12), bg="#f2f2f2")
advance_label.grid(row=2, column=0, sticky="w")
advance_entry = tk.Entry(details_frame, font=("Arial", 12))
advance_entry.grid(row=2, column=1, padx=10)

start_label = tk.Label(details_frame, text="Start Date:",
                       font=("Arial", 12), bg="#f2f2f2")
start_label.grid(row=3, column=0, sticky="w")
start_entry = tk.Entry(details_frame, font=("Arial", 12))
start_entry.grid(row=3, column=1, padx=10)

end_label = tk.Label(details_frame, text="End Date:",
                     font=("Arial", 12), bg="#f2f2f2")
end_label.grid(row=4, column=0, sticky="w")
end_entry = tk.Entry(details_frame, font=("Arial", 12))
end_entry.grid(row=4, column=1, padx=10)

# Car Image
image_label = tk.Label(details_frame, text="Car Image:",
                       font=("Arial", 12), bg="#f2f2f2")
image_label.grid(row=5, column=0, sticky="w")
browse_button = tk.Button(details_frame, text="Browse",
                          font=("Arial", 12), command=browse_image)
browse_button.grid(row=5, column=1, pady=5)

preview_label = tk.Label(details_frame, width=200,
                         height=200, bg="white", relief="solid")
preview_label.grid(row=6, column=0, columnspan=2, padx=10, pady=5)
hide_preview()

# Save Button
save_button = tk.Button(root, text="Save", font=(
    "Arial", 12), command=save_customer)
save_button.pack(pady=10)

# Table
style = ttk.Style()
style.theme_use("clam")

style.configure("Treeview", background="#f2f2f2",
                foreground="black", rowheight=25, fieldbackground="#f2f2f2")
style.map("Treeview", background=[("selected", "#347083")])

scrollbar = ttk.Scrollbar(root)
scrollbar.pack(side="right", fill="y")

table_view = ttk.Treeview(root, columns=("Serial Number", "Customer Name", "Pending Amount", "Advance Amount",
                          "Start Date", "End Date"), show="headings", style="Treeview", yscrollcommand=scrollbar.set)
table_view.heading("Serial Number", text="Serial Number")
table_view.heading("Customer Name", text="Customer Name")
table_view.heading("Pending Amount", text="Pending Amount")
table_view.heading("Advance Amount", text="Advance Amount")
table_view.heading("Start Date", text="Start Date")
table_view.heading("End Date", text="End Date")

scrollbar.config(command=table_view.yview)
table_view.pack(expand=True, fill="both")

# Update table with saved customer details
update_table()

# Start the GUI main loop
root.mainloop()
