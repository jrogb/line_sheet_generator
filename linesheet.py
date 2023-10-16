import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from wsgiref.handlers import format_date_time
from docxtpl import DocxTemplate
import pandas as pd
import os
import datetime


class LineSheetApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Line Sheet Generator")

        self.frame = ttk.Frame(root)
        self.frame.grid(row=0, column=0, padx=10, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Create a combo box for Destination
        self.destination_label = ttk.Label(self.frame, text="Destination:")
        self.destination_combobox = ttk.Combobox(self.frame, values=[
            "Advance", "Cancam", "Headland", "Jumaras", "Vectura",
            "Wavelengths", "Nexgistix", "Copperzone", "Africlan", "Other"
        ])

        # Create an entry for Fleet Number/Store
        self.fleet_label = ttk.Label(self.frame, text="Fleet Number/Store:")
        self.fleet_entry = ttk.Entry(self.frame)

        # Create an entry for Stock Code
        self.stock_code_label = ttk.Label(self.frame, text="Stock Code:")
        self.stock_code_entry = ttk.Entry(self.frame)

        # Create an entry for QTY
        self.qty_label = ttk.Label(self.frame, text="QTY:")
        self.qty_entry = ttk.Entry(self.frame)

        # Create an entry for PO Number
        self.po_label = ttk.Label(self.frame, text="PO Number:")
        self.po_entry = ttk.Entry(self.frame)

        # Buttons for adding and deleting stock items
        self.add_stock_button = ttk.Button(self.frame, text="Add Stock Item", command=self.add_stock_item)
        self.delete_stock_button = ttk.Button(self.frame, text="Delete Stock Item", command=self.delete_stock_item)

        # Create a frame to contain Generate and New buttons
        button_frame = ttk.Frame(self.frame)
        button_frame.grid(row=0, column=2, rowspan=7, padx=10, pady=10)

        self.generate_button = ttk.Button(button_frame, text="Generate", command=self.generate_line_sheet)
        self.new_button = ttk.Button(button_frame, text="New", command=self.clear_inputs)

        # Create a Treeview for displaying line sheet items
        self.line_sheet_tree = ttk.Treeview(self.frame, columns=("Destination", "Fleet Number/Store", "Stock Code", "Description", "QTY", "PO Number"), show="headings")

        # Set Treeview headings
        current_datetime = datetime.datetime.now()
        formatted_datetime = current_datetime.strftime("%d-%m-%y")

        self.tree_view_label = ttk.Label(self.frame, text=f"LINE SHEET: {formatted_datetime}")
        self.line_sheet_tree.heading("Destination", text="Destination")
        self.line_sheet_tree.heading("Fleet Number/Store", text="Fleet Number/Store")
        self.line_sheet_tree.heading("Stock Code", text="Stock Code")
        self.line_sheet_tree.heading("Description", text="Description")
        self.line_sheet_tree.heading("QTY", text="QTY")
        self.line_sheet_tree.heading("PO Number", text="PO Number")

        # Set Treeview columns width
        self.line_sheet_tree.column("Destination", width=100)
        self.line_sheet_tree.column("Fleet Number/Store", width=100)
        self.line_sheet_tree.column("Stock Code", width=100)
        self.line_sheet_tree.column("Description", width=200)
        self.line_sheet_tree.column("QTY", width=50)
        self.line_sheet_tree.column("PO Number", width=100)

        # Create a list to store line sheet items
        self.line_sheet_items = []    

        # Read product info from an Excel file
        self.product_info_df = pd.read_excel("product_info.xlsx")

        # Place widgets in the grid
        self.destination_label.grid(row=0, column=0, padx=10, pady=10)
        self.destination_combobox.grid(row=0, column=1, padx=10, pady=10)

        self.fleet_label.grid(row=1, column=0, padx=10, pady=10)
        self.fleet_entry.grid(row=1, column=1, padx=10, pady=10)

        self.stock_code_label.grid(row=2, column=0, padx=10, pady=10)
        self.stock_code_entry.grid(row=2, column=1, padx=10, pady=10)

        self.qty_label.grid(row=3, column=0, padx=10, pady=10)
        self.qty_entry.grid(row=3, column=1, padx=10, pady=10)

        self.po_label.grid(row=4, column=0, padx=10, pady=10)
        self.po_entry.grid(row=4, column=1, padx=10, pady=10)

        self.add_stock_button.grid(row=5, column=0, padx=10, pady=10)
        self.delete_stock_button.grid(row=5, column=1, padx=10, pady=10)

        self.generate_button.grid(row=0, column=0, padx=10, pady=10)
        self.new_button.grid(row=1, column=0, padx=10, pady=10)

        self.tree_view_label.grid(row=0, column=3)
        self.line_sheet_tree.grid(row=1, column=3, rowspan=7, padx=10, pady=10)

    df = pd.DataFrame(columns=['Stock Code','Description','Destination', 'Fleet Number / Store', 'QTY', 'PO Number'])

    def add_stock_item(self):
        destination = self.destination_combobox.get()
        fleet_number = self.fleet_entry.get()
        stock_code = self.stock_code_entry.get()
        qty = self.qty_entry.get()
        po_number = self.po_entry.get()
        stock_description = self.get_stock_description(stock_code)

        if not stock_description:
            messagebox.showerror("Stock Description Not Found", f"Stock code {stock_code} not found in product_info.xlsx")
            return

        if destination.strip() and fleet_number.strip() and stock_code.strip() and qty.strip() and po_number.strip():
            self.line_sheet_tree.insert("", "end", values=(destination, fleet_number, stock_code, stock_description, qty, po_number))
            self.stock_code_entry.delete(0, tk.END)
            self.qty_entry.delete(0, tk.END)
            self.line_sheet_items.append((destination, fleet_number, stock_code, stock_description, qty, po_number))

    def get_stock_description(self, stock_code):
        description = self.product_info_df.loc[
            self.product_info_df['Stock Code'] == stock_code, 'Description'].values
        if len(description) > 0:
            return description[0]
        else:
            return None

    def delete_stock_item(self):
        selected_item = self.line_sheet_tree.selection()
        if selected_item:
            item = self.line_sheet_tree.item(selected_item)
            stock_code = item['values'][2]
            self.line_sheet_tree.delete(selected_item)
            self.line_sheet_items = [item for item in self.line_sheet_items if item[2] != stock_code]

    def clear_inputs(self):
        self.destination_combobox.set("")  # Clear Destination selection
        self.fleet_entry.delete(0, tk.END)
        self.stock_code_entry.delete(0, tk.END)
        self.qty_entry.delete(0, tk.END)
        self.po_entry.delete(0, tk.END)
        self.line_sheet_tree.delete(*self.line_sheet_tree.get_children())  # Clear Treeview

    def generate_line_sheet(self):
        current_datetime = datetime.datetime.now()
        formatted_datetime = current_datetime.strftime("%d-%m-%y : %H:%M")
        destination = self.destination_combobox.get()
        fleet_number = self.fleet_entry.get()
        po_number = self.po_entry.get()

        if not destination.strip() or not fleet_number.strip() or not self.line_sheet_items:
            messagebox.showerror("Error", "Please fill in Destination and Fleet Number/Store, and add at least one stock item.")
            return

        try:
            doc = DocxTemplate("linesheet_template.docx")

            # Extract line sheet items from the Treeview
            current_datetime = datetime.datetime.now()
            formatted_datetime1 = current_datetime.strftime("%d-%m-%y : %H:%M")
            line_sheet_items = []

            for item in self.line_sheet_tree.get_children():
                values = self.line_sheet_tree.item(item, 'values')
                line_sheet_items.append(values)

            context = {
                "date": formatted_datetime1,
                "destination": destination,
                "fleet_number": fleet_number,
                "line_sheet_items": line_sheet_items,
                "po_number": po_number
            }

            doc.render(context)
            line_sheet_filename = f"{fleet_number}_line_sheet.docx"
            doc.save(line_sheet_filename)

            # Open the generated line sheet with the default application for DOCX files
            os.startfile(line_sheet_filename)

        except Exception as e:
            messagebox.showerror("Error Generating Line Sheet", f"An error occurred while generating the Line Sheet: {e}")

def main():
    root = tk.Tk()
    app = LineSheetApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
