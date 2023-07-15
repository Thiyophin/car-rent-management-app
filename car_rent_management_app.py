import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import tkinter.messagebox as messagebox
import openpyxl
import shutil
import os


current_image_path = None
serial_number = 1

def browse_image():
    global current_image_path
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
    if file_path:
        current_image_path = file_path
        load_image(file_path)
    else:
        hide_preview()

import tkinter.messagebox as messagebox

def save_customer():
    global serial_number

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

    row = (serial_number, customer_name, pending_amount, advance_amount, start_date, end_date)
    sheet.append(row)
    serial_number += 1

    workbook.save(excel_file_path)

    if current_image_path:
        image_file_name = f"{customer_name}_{serial_number - 1}.png"
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





def load_image(file_path):
    global image_preview
    image = Image.open(file_path)
    image = image.resize((200, 200))
    image_preview = ImageTk.PhotoImage(image)
    preview_label.configure(image=image_preview)
    preview_label.image = image_preview
    show_preview()


def show_preview():
    preview_label.grid(row=5, column=1, padx=10, pady=5)


def hide_preview():
    preview_label.grid_remove()


# Create the main window
root = tk.Tk()
root.title("Car Rental Management System")

# Set a custom background color
root.configure(bg="#f2f2f2")

# Customer Details
name_label = tk.Label(root, text="Customer Name:",
                      font=("Arial", 12), bg="#f2f2f2")
name_label.grid(row=0, column=0, sticky="w", padx=10, pady=5)
name_entry = tk.Entry(root, font=("Arial", 12))
name_entry.grid(row=0, column=1, padx=10, pady=5)

pending_label = tk.Label(root, text="Pending Amount:",
                         font=("Arial", 12), bg="#f2f2f2")
pending_label.grid(row=1, column=0, sticky="w", padx=10, pady=5)
pending_entry = tk.Entry(root, font=("Arial", 12))
pending_entry.grid(row=1, column=1, padx=10, pady=5)

advance_label = tk.Label(root, text="Advance Amount:",
                         font=("Arial", 12), bg="#f2f2f2")
advance_label.grid(row=2, column=0, sticky="w", padx=10, pady=5)
advance_entry = tk.Entry(root, font=("Arial", 12))
advance_entry.grid(row=2, column=1, padx=10, pady=5)

start_label = tk.Label(root, text="Start Date:",
                       font=("Arial", 12), bg="#f2f2f2")
start_label.grid(row=3, column=0, sticky="w", padx=10, pady=5)
start_entry = tk.Entry(root, font=("Arial", 12))
start_entry.grid(row=3, column=1, padx=10, pady=5)

end_label = tk.Label(root, text="End Date:", font=("Arial", 12), bg="#f2f2f2")
end_label.grid(row=4, column=0, sticky="w", padx=10, pady=5)
end_entry = tk.Entry(root, font=("Arial", 12))
end_entry.grid(row=4, column=1, padx=10, pady=5)

# Car Image
image_label = tk.Label(root, text="Car Image:",
                       font=("Arial", 12), bg="#f2f2f2")
image_label.grid(row=5, column=0, sticky="w", padx=10, pady=5)

preview_label = tk.Label(root, width=200, height=200,
                         bg="white", relief="solid")
preview_label.grid(row=5, column=1, padx=10, pady=5)
hide_preview()

browse_button = tk.Button(root, text="Browse", font=(
    "Arial", 12), command=browse_image)
browse_button.grid(row=6, column=0, columnspan=2, pady=10)

# Save Button
save_button = tk.Button(root, text="Save", font=(
    "Arial", 12), command=save_customer)
save_button.grid(row=7, columnspan=2, pady=10)

# Start the GUI main loop
root.mainloop()
