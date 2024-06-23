import pandas as pd
import tkinter as tk
from tkinter import ttk

# Load the data dictionary
file_path = r'C:\Users\BryanAgas\Downloads\Database Data Dictionary.xlsx'
sheet_name = 'Redshift Data Dictionary'
data_dict = pd.read_excel(file_path, sheet_name=sheet_name)

# Function to get the column description
def get_column_description(column_name):
    result = data_dict[data_dict['Columns'].str.contains(column_name, case=False, na=False)]
    return result[['Columns', 'Column Description', 'Tables', 'Schema', 'Database']]

# Function to display results in the treeview
def display_results(results):
    for i in tree.get_children():
        tree.delete(i)
    for index, row in results.iterrows():
        tree.insert("", "end", values=(row['Columns'], row['Column Description'], row['Tables'], row['Schema'], row['Database']))

# Function to handle the search
def search():
    column_name = entry.get()
    results = get_column_description(column_name)
    display_results(results)

# Create the main application window
root = tk.Tk()
root.title("Data Dictionary Search")

# Create the input field and search button
tk.Label(root, text="Enter Column Name:").grid(row=0, column=0, padx=10, pady=10)
entry = tk.Entry(root)
entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Search", command=search).grid(row=0, column=2, padx=10, pady=10)

# Create the treeview for displaying results
tree = ttk.Treeview(root, columns=("Column Name", "Description", "Table", "Schema", "Database"), show="headings")
tree.heading("Column Name", text="Column Name")
tree.heading("Description", text="Description")
tree.heading("Table", text="Table")
tree.heading("Schema", text="Schema")
tree.heading("Database", text="Database")
tree.grid(row=1, column=0, columnspan=3)

# Add a scrollbar
scrollbar = ttk.Scrollbar(root, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.grid(row=1, column=3, sticky='ns')

root.mainloop()