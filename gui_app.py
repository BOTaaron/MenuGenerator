import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xlsx_loader import XLSXLoader
from exclusion_manager import ExclusionManager
from pdf_generator import PDFGenerator


class MenuGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Menu Generator")
        self.root.geometry("1920x1080")

        # Instantiate the main components
        self.xlsx_loader = XLSXLoader()
        self.exclusion_manager = ExclusionManager()
        self.pdf_generator = PDFGenerator()

        # Define options for dropdowns
        self.departments = ["Draught Beer", "Wine", "Spirits", "Soft Drinks", "Snacks"]
        self.price_levels = ["125ml", "250ml", "Bottle", "Half Pint", "Pint"]

        # Top frame for buttons
        top_frame = tk.Frame(self.root, padx=20, pady=20)
        top_frame.pack(anchor="ne", fill="x")

        # Buttons
        self.open_button = tk.Button(top_frame, text="Open File", command=self.open_file, font=("Arial", 16))
        self.open_button.pack(side="right", padx=10)

        self.generate_button = tk.Button(top_frame, text="Generate", command=self.generate_menu, font=("Arial", 16))
        self.generate_button.pack(side="right", padx=10)

        # Content area (for exclusions and item previews)
        self.content_frame = tk.Frame(self.root, padx=20, pady=20)
        self.content_frame.pack(fill="both", expand=True)

        # Treeview for displaying table
        self.tree = ttk.Treeview(self.content_frame, columns=("Department", "Price Level", "Item Name", "Price"),
                                 show="headings")
        self.tree.heading("Department", text="Department")
        self.tree.heading("Price Level", text="Price Level")
        self.tree.heading("Item Name", text="Item Name")
        self.tree.heading("Price", text="Price")
        self.tree.column("Department", width=200)
        self.tree.column("Price Level", width=150)
        self.tree.column("Item Name", width=300)
        self.tree.column("Price", width=100)
        self.tree.pack(fill="both", expand=True)

        # Enable double-click editing
        self.tree.bind("<Double-1>", self.edit_cell)

    def open_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.xlsx_loader.load_data(file_path)
                messagebox.showinfo("File Loaded", "File has been successfully loaded and processed!")
                self.display_data()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")

    def display_data(self):
        # Clear previous data
        for i in self.tree.get_children():
            self.tree.delete(i)

        # Insert new data
        data = self.xlsx_loader.get_data()
        if not data:
            messagebox.showwarning("Warning", "No data has been loaded!")
        for department, price_levels in data.items():
            for price_level, items in price_levels.items():
                print(items)
                for item in items:
                    self.tree.insert("", "end", values=(department, price_level, item["name"], item["price"]))

    def edit_cell(self, event):
        selected_item = self.tree.selection()[0]
        col = self.tree.identify_column(event.x)
        column_name = self.tree.heading(col)["text"]

        # Get current value of the selected cell
        current_value = self.tree.item(selected_item, "values")[self.tree["columns"].index(column_name)]

        if column_name == "Department" or column_name == "Price Level":
            # Use a dropdown for Department and Price Level
            options = self.departments if column_name == "Department" else self.price_levels
            edit_window = tk.Toplevel(self.root)
            edit_window.title(f"Edit {column_name}")

            # Dropdown for selection
            tk.Label(edit_window, text=f"{column_name}:").pack()
            dropdown = ttk.Combobox(edit_window, values=options)
            dropdown.pack()
            dropdown.set(current_value)

            def save_value():
                new_value = dropdown.get()
                current_values = list(self.tree.item(selected_item, "values"))
                current_values[self.tree["columns"].index(column_name)] = new_value
                self.tree.item(selected_item, values=current_values)
                edit_window.destroy()

            save_button = tk.Button(edit_window, text="Save", command=save_value)
            save_button.pack()

        else:
            # For other columns, use an entry field
            edit_window = tk.Toplevel(self.root)
            edit_window.title(f"Edit {column_name}")
            tk.Label(edit_window, text=f"{column_name}:").pack()
            entry = tk.Entry(edit_window)
            entry.pack()
            entry.insert(0, current_value)

            def save_value():
                new_value = entry.get()
                current_values = list(self.tree.item(selected_item, "values"))
                current_values[self.tree["columns"].index(column_name)] = new_value
                self.tree.item(selected_item, values=current_values)
                edit_window.destroy()

            save_button = tk.Button(edit_window, text="Save", command=save_value)
            save_button.pack()

    def generate_menu(self):
        # Apply exclusions and generate menu PDF
        filtered_data = self.exclusion_manager.apply_exclusions(self.xlsx_loader.get_data())
        self.pdf_generator.create_pdf(filtered_data)
        messagebox.showinfo("Menu Generated", "Menu PDF has been generated successfully!")


# Run the app
