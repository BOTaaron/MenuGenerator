import pandas as pd
import re
from tkinter import messagebox

class XLSXLoader:
    def __init__(self):
        self.departments = {}

    """
    Read a Excel spreadsheet to pull the data needed to generate the menu.
    """

    def load_data(self, file_path):
        if not file_path:
            raise ValueError("File path cannot be empty or None.")
        try:
            print(f"Attempting to load file: {file_path}")  # Debugging line
            df = pd.read_excel(file_path, header=None, engine="openpyxl")
            self.departments = self.parse_departments(df)
            messagebox.showinfo("Success", "Data loaded!")
        except FileNotFoundError:
            messagebox.showerror("Error", "File not found.")
        except ValueError as ve:
            messagebox.showerror("Error", f"Value error: {ve}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    """
    CES Intelligence POS exports a spreadsheet containing prices, product names, and categories.
    For example, 'Department: Draught Beer - 8 items'.
    This will be parsed to add items into a category used to generate a menu. 
    Regex is used to strip the department name to categorise items and prices depending on their type, such as bottled beers or spirits.
    """

    def parse_departments(self, df):
        departments = {}
        current_department = None
        current_price_level = None
        price_levels = {}

        for index, row in df.iterrows():
            cell_a = row[0] if not pd.isna(row[0]) else None
            cell_b = row[1] if not pd.isna(row[1]) else None
            cell_c = row[2] if not pd.isna(row[2]) else None
            cell_d = row[3] if not pd.isna(row[3]) else None
            cell_e = row[4] if not pd.isna(row[4]) else None



            # Ignore rows with "Price Band:" in column A
            if cell_a and "Price Band:" in cell_a:
                continue


            # Parse price level from column B
            if pd.notna(cell_b) and "Price Level:" in cell_b:
                print(cell_b)
                price_level_match = re.search(r"Price Level:\s*(\d+)", cell_b)
                if price_level_match:
                    current_price_level = int(price_level_match.group(1))
                    print(current_price_level)
                    volume = self.map_price_level_to_volume(current_price_level, current_department)
                    if current_department:
                        price_levels[current_price_level] = volume
                        departments[current_department][volume] = []
                    print(f"Set price level {current_price_level} with volume {volume} for department {current_department}")
                continue

            # Ignore rows with "Group:" in column C
            if pd.notna(cell_c) and "Group:" in cell_c:
                print("Skipping row due to 'Group' in column C")
                continue

            # Parse department from column D
            if pd.notna(cell_d):
                if "Department:" in cell_d:
                    department_match = re.search(r"Department:\s*(.*?)\s*-\s*\d+\s*items", cell_d)
                    if department_match:
                        current_department = department_match.group(1)
                        if current_department != "Misc":  # Skip "Misc" department
                            departments[current_department] = {}
                            price_levels = {}
                            print(f"Found department: {current_department}")
                        else:
                            print("No department match found in cell D")
                    continue

                # Parse item row if in column D and not a department or code
                elif current_department and current_price_level and not cell_d.isnumeric():
                    item = {
                        "name": cell_e,  # Assuming description is in column D
                        "price": row[10]  # Assuming price is in column K
                    }
                    departments[current_department][price_levels[current_price_level]].append(item)
                    print(f"Added item: {item} to {current_department} under {price_levels[current_price_level]}")

        print("Parsed departments: ", departments)
        return departments

    """
    Maps the pricing level of each item to the correct category. 
    In the current setup for example, 'Price Level: 2 - 1 item' is used for half pints and large glasses of wine.
    This will map the correct price to each item depending on its size.
    
    """

    def map_price_level_to_volume(self, price_level, department):
        if price_level == 1:
            return "Bottle" if department == "Wine" else "Pint" if department == "Draught Beer" else "Standard"
        elif price_level == 2:
            return "250ml Glass" if department == "Wine" else "Half Pint" if department == "Draught Beer" else None
        elif price_level == 3 and department == "Wine":
            return "125ml Glass"
        return f"Price Level {price_level}"

    def get_data(self):
        return self.departments