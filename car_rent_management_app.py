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
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
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

    row = (serial_number, customer_name, pending_amount, advance_amount, start_date, end_date)
    sheet.append(row)

    workbook.save(excel_file_path)

    if current_image_path:
        image_file_name = f"{customer_name}_{serial_number}.png"
        shutil.copyfile(current_image_path, os.path.join(images_dir, image_file_name))
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

def show_edit_dialog(serial_number):
    global selected_serial_number
    selected_serial_number = serial_number

    # Create a new Toplevel window for editing
    edit_window = tk.Toplevel(root)
    edit_window.title("Edit Customer Details")

    # Set a custom background color
    edit_window.configure(bg="#f2f2f2")

    # Create a frame for the edit fields
    edit_frame = tk.Frame(edit_window, bg="#f2f2f2")
    edit_frame.pack(padx=10, pady=10)

    # Get the existing values of the selected row
    excel_file_path = os.path.join("Customer Details Folder", "customer_details.xlsx")
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    existing_values = None
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] == serial_number:
            existing_values = row
            break

    if existing_values:
        # Extract the existing values
        existing_customer_name = existing_values[1]
        existing_pending_amount = existing_values[2]
        existing_advance_amount = existing_values[3]
        existing_start_date = existing_values[4]
        existing_end_date = existing_values[5]

        # Customer Details
        name_label = tk.Label(edit_frame, text="Customer Name:", font=("Arial", 12), bg="#f2f2f2")
        name_label.grid(row=0, column=0, sticky="w")
        name_entry = tk.Entry(edit_frame, font=("Arial", 12))
        name_entry.insert(tk.END, existing_customer_name)
        name_entry.grid(row=0, column=1, padx=10)

        pending_label = tk.Label(edit_frame, text="Pending Amount:", font=("Arial", 12), bg="#f2f2f2")
        pending_label.grid(row=1, column=0, sticky="w")
        pending_entry = tk.Entry(edit_frame, font=("Arial", 12))
        pending_entry.insert(tk.END, existing_pending_amount)
        pending_entry.grid(row=1, column=1, padx=10)

        advance_label = tk.Label(edit_frame, text="Advance Amount:", font=("Arial", 12), bg="#f2f2f2")
        advance_label.grid(row=2, column=0, sticky="w")
        advance_entry = tk.Entry(edit_frame, font=("Arial", 12))
        advance_entry.insert(tk.END, existing_advance_amount)
        advance_entry.grid(row=2, column=1, padx=10)

        start_label = tk.Label(edit_frame, text="Start Date:", font=("Arial", 12), bg="#f2f2f2")
        start_label.grid(row=3, column=0, sticky="w")
        start_entry = tk.Entry(edit_frame, font=("Arial", 12))
        start_entry.insert(tk.END, existing_start_date)
        start_entry.grid(row=3, column=1, padx=10)

        end_label = tk.Label(edit_frame, text="End Date:", font=("Arial", 12), bg="#f2f2f2")
        end_label.grid(row=4, column=0, sticky="w")
        end_entry = tk.Entry(edit_frame, font=("Arial", 12))
        end_entry.insert(tk.END, existing_end_date)
        end_entry.grid(row=4, column=1, padx=10)

        # Save Button
        save_button = tk.Button(edit_window, text="Save", font=("Arial", 12), command=update_customer)
        save_button.pack(pady=10)

        # Function to update the customer details
        def update_customer():
            updated_customer_name = name_entry.get()
            updated_pending_amount = pending_entry.get()
            updated_advance_amount = advance_entry.get()
            updated_start_date = start_entry.get()
            updated_end_date = end_entry.get()

            # Check if any field is empty
            if not (updated_customer_name and updated_pending_amount and updated_advance_amount and updated_start_date and updated_end_date):
                messagebox.showerror("Error", "Please fill in all the fields.")
                return

            # Update the values in the Excel sheet
            existing_values[1] = updated_customer_name
            existing_values[2] = updated_pending_amount
            existing_values[3] = updated_advance_amount
            existing_values[4] = updated_start_date
            existing_values[5] = updated_end_date

            workbook.save(excel_file_path)

            # Update the table
            update_table()

            messagebox.showinfo("Success", "Customer details updated successfully.")
            edit_window.destroy()

    else:
        messagebox.showerror("Error", f"Customer details not found for serial number: {serial_number}")
        edit_window.destroy()

def delete_selected_row(event):
    selected_item = table_view.selection()
    if selected_item:
        serial_number = table_view.item(selected_item, "values")[0]
        delete_row(serial_number)

def delete_row(serial_number):
    excel_file_path = os.path.join("Customer Details Folder", "customer_details.xlsx")
    images_dir = "Car Images Folder"

    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    # Find the row with the given serial number
    delete_row_index = None
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == serial_number:
            delete_row_index = index
            break

    if delete_row_index:
        # Get the corresponding customer name
        customer_name = row[1]

        # Delete the row from the Excel sheet
        sheet.delete_rows(delete_row_index)

        # Delete the corresponding image file
        image_file_name = f"{customer_name}_{serial_number}.png"
        image_path = os.path.join(images_dir, image_file_name)
        if os.path.isfile(image_path):
            os.remove(image_path)

        # Save the updated Excel sheet
        workbook.save(excel_file_path)

        # Update the table
        update_table()

    else:
        messagebox.showerror("Error", f"Customer details not found for serial number: {serial_number}")

def update_table():
    excel_file_path = os.path.join("Customer Details Folder", "customer_details.xlsx")
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
    table_view.column("Edit", width=100, anchor="center")
    table_view.column("Delete", width=100, anchor="center")

    # Add edit and delete buttons to the table
    table_view.heading("#0", text="", anchor="center")
    table_view.heading("Serial Number", text="Serial Number", anchor="center")
    table_view.heading("Customer Name", text="Customer Name", anchor="center")
    table_view.heading("Pending Amount", text="Pending Amount", anchor="center")
    table_view.heading("Advance Amount", text="Advance Amount", anchor="center")
    table_view.heading("Start Date", text="Start Date", anchor="center")
    table_view.heading("End Date", text="End Date", anchor="center")
    table_view.heading("Edit", text="Edit", anchor="center")
    table_view.heading("Delete", text="Delete", anchor="center")

    # Iterate through rows and add details to the table
    for row in sheet.iter_rows(min_row=2, values_only=True):
        serial_number = row[0]
        customer_name = row[1]
        pending_amount = row[2]
        advance_amount = row[3]
        start_date = row[4]
        end_date = row[5]

        # Create edit and delete buttons as text values
        edit_button = "Edit"
        delete_button = "Delete"

        # Add details and buttons to the table
        table_view.insert("", tk.END, values=(serial_number, customer_name, pending_amount, advance_amount, start_date, end_date, edit_button, delete_button))

    # Configure the table to allow editing only for rows with the "editable" tag
    table_view.tag_configure("editable", background="#D3D3D3", foreground="black")

def load_image(file_path):
    global image_preview
    image = Image.open(file_path)
    image = image.resize((200, 200))
    image_preview = ImageTk.PhotoImage(image)
    preview_label.configure(image=image_preview)
    preview_label.image = image_preview
    show_preview()

def edit_row(serial_number):
    excel_file_path = os.path.join(
        "Customer Details Folder", "customer_details.xlsx")
    images_dir = "Car Images Folder"

    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    # Find the row with the given serial number
    edit_row_index = None
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == serial_number:
            edit_row_index = index
            break

    if edit_row_index:
        # Get the existing values of the row
        existing_values = sheet[edit_row_index]
        existing_customer_name = existing_values[1].value
        existing_pending_amount = existing_values[2].value
        existing_advance_amount = existing_values[3].value
        existing_start_date = existing_values[4].value
        existing_end_date = existing_values[5].value

        # Open a new window for editing the row data
        edit_window = tk.Toplevel()
        edit_window.title("Edit Customer Details")

        # Create and place labels and entry fields for editing
        name_label = tk.Label(edit_window, text="Customer Name:", font=("Arial", 12))
        name_label.grid(row=0, column=0, sticky="w")
        name_entry = tk.Entry(edit_window, font=("Arial", 12))
        name_entry.grid(row=0, column=1, padx=10)
        name_entry.insert(tk.END, existing_customer_name)

        pending_label = tk.Label(edit_window, text="Pending Amount:", font=("Arial", 12))
        pending_label.grid(row=1, column=0, sticky="w")
        pending_entry = tk.Entry(edit_window, font=("Arial", 12))
        pending_entry.grid(row=1, column=1, padx=10)
        pending_entry.insert(tk.END, existing_pending_amount)

        advance_label = tk.Label(edit_window, text="Advance Amount:", font=("Arial", 12))
        advance_label.grid(row=2, column=0, sticky="w")
        advance_entry = tk.Entry(edit_window, font=("Arial", 12))
        advance_entry.grid(row=2, column=1, padx=10)
        advance_entry.insert(tk.END, existing_advance_amount)

        start_label = tk.Label(edit_window, text="Start Date:", font=("Arial", 12))
        start_label.grid(row=3, column=0, sticky="w")
        start_entry = tk.Entry(edit_window, font=("Arial", 12))
        start_entry.grid(row=3, column=1, padx=10)
        start_entry.insert(tk.END, existing_start_date)

        end_label = tk.Label(edit_window, text="End Date:", font=("Arial", 12))
        end_label.grid(row=4, column=0, sticky="w")
        end_entry = tk.Entry(edit_window, font=("Arial", 12))
        end_entry.grid(row=4, column=1, padx=10)
        end_entry.insert(tk.END, existing_end_date)

        # Save Button
        save_button = tk.Button(edit_window, text="Save", font=("Arial", 12),
                                command=lambda: save_edited_row(serial_number, name_entry.get(), pending_entry.get(),
                                                               advance_entry.get(), start_entry.get(), end_entry.get(),
                                                               excel_file_path, images_dir, edit_window))
        save_button.grid(row=5, column=0, columnspan=2, pady=10)

    else:
        messagebox.showerror("Error", f"Customer details not found for serial number: {serial_number}")


def save_edited_row(serial_number, customer_name, pending_amount, advance_amount, start_date, end_date,
                    excel_file_path, images_dir, edit_window):
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    # Find the row with the given serial number
    edit_row_index = None
    for index, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] == serial_number:
            edit_row_index = index
            break

    if edit_row_index:
        # Update the values of the row
        sheet.cell(row=edit_row_index, column=2).value = customer_name
        sheet.cell(row=edit_row_index, column=3).value = pending_amount
        sheet.cell(row=edit_row_index, column=4).value = advance_amount
        sheet.cell(row=edit_row_index, column=5).value = start_date
        sheet.cell(row=edit_row_index, column=6).value = end_date

        # Save the updated Excel sheet
        workbook.save(excel_file_path)

        # Close the edit window
        edit_window.destroy()

        # Update the table
        update_table()

    else:
        messagebox.showerror("Error", f"Customer details not found for serial number: {serial_number}")

def show_preview():
    preview_label.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

def hide_preview():
    preview_label.grid_remove()

# Create the main window
root = tk.Tk()
root.title("Car Rental Management System")

# Set a custom background color
root.configure(bg="#f2f2f2")

# Create a frame for the customer details
details_frame = tk.Frame(root, bg="#f2f2f2")
details_frame.pack(padx=10, pady=10)

# Customer Details
name_label = tk.Label(details_frame, text="Customer Name:", font=("Arial", 12), bg="#f2f2f2")
name_label.grid(row=0, column=0, sticky="w")
name_entry = tk.Entry(details_frame, font=("Arial", 12))
name_entry.grid(row=0, column=1, padx=10)

pending_label = tk.Label(details_frame, text="Pending Amount:", font=("Arial", 12), bg="#f2f2f2")
pending_label.grid(row=1, column=0, sticky="w")
pending_entry = tk.Entry(details_frame, font=("Arial", 12))
pending_entry.grid(row=1, column=1, padx=10)

advance_label = tk.Label(details_frame, text="Advance Amount:", font=("Arial", 12), bg="#f2f2f2")
advance_label.grid(row=2, column=0, sticky="w")
advance_entry = tk.Entry(details_frame, font=("Arial", 12))
advance_entry.grid(row=2, column=1, padx=10)

start_label = tk.Label(details_frame, text="Start Date:", font=("Arial", 12), bg="#f2f2f2")
start_label.grid(row=3, column=0, sticky="w")
start_entry = tk.Entry(details_frame, font=("Arial", 12))
start_entry.grid(row=3, column=1, padx=10)

end_label = tk.Label(details_frame, text="End Date:", font=("Arial", 12), bg="#f2f2f2")
end_label.grid(row=4, column=0, sticky="w")
end_entry = tk.Entry(details_frame, font=("Arial", 12))
end_entry.grid(row=4, column=1, padx=10)

# Car Image
image_label = tk.Label(details_frame, text="Car Image:", font=("Arial", 12), bg="#f2f2f2")
image_label.grid(row=5, column=0, sticky="w")
browse_button = tk.Button(details_frame, text="Browse", font=("Arial", 12), command=browse_image)
browse_button.grid(row=5, column=1, pady=5)

preview_label = tk.Label(details_frame, width=200, height=200, bg="white", relief="solid")
preview_label.grid(row=6, column=0, columnspan=2, padx=10, pady=5)
hide_preview()

# Save Button
save_button = tk.Button(root, text="Save", font=("Arial", 12), command=save_customer)
save_button.pack(pady=10)

# Table
style = ttk.Style()
style.theme_use("clam")

style.configure("Treeview", background="#D3D3D3", foreground="black", rowheight=25, fieldbackground="#D3D3D3")
style.map("Treeview", background=[("selected", "#347083")])

scrollbar = ttk.Scrollbar(root)
scrollbar.pack(side="right", fill="y")

table_view = ttk.Treeview(root, columns=("Serial Number", "Customer Name", "Pending Amount", "Advance Amount", "Start Date", "End Date", "Edit", "Delete"), show="headings", style="Treeview", yscrollcommand=scrollbar.set)
table_view.heading("Serial Number", text="Serial Number")
table_view.heading("Customer Name", text="Customer Name")
table_view.heading("Pending Amount", text="Pending Amount")
table_view.heading("Advance Amount", text="Advance Amount")
table_view.heading("Start Date", text="Start Date")
table_view.heading("End Date", text="End Date")
table_view.heading("Edit", text="Edit")
table_view.heading("Delete", text="Delete")

scrollbar.config(command=table_view.yview)
table_view.pack(expand=True, fill="both")

# Bind the double click event to edit the selected row
table_view.bind("<Double-1>", lambda event: show_edit_dialog(table_view.item(table_view.selection(), "values")[0]))

# Bind the delete key press event to delete the selected row
table_view.bind("<Delete>", delete_selected_row)

# Update table with saved customer details
update_table()

# Start the GUI main loop
root.mainloop()
