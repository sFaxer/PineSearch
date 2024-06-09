# Importing necessary modules
import tkinter as tk
from tkinter import ttk, messagebox, Scrollbar, Entry, Button, END, filedialog
import pandas as pd
import os

# Function to search Excel and return results
def search_excel(excel_path, query, search_by_customer=False):
    try:
        # Read all sheets in the Excel file
        xl = pd.ExcelFile(excel_path, engine='openpyxl')

        # Initialize variables to store the sum of units, cost, and the search results
        total_units = 0
        total_cost = 0
        daily_units = {}  # Dictionary to store units by each day
        daily_cost = {}  # Dictionary to store cost by each day
        daily_routes_count = {}  # Dictionary to store route count by each day
        result = pd.DataFrame()

        # Determine which column to search based on the toggle button
        search_column = 1 if search_by_customer else 11  # Column B (customer) or Column L (transporter)

        # Iterate through all sheet names
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(excel_path, sheet_name, engine='openpyxl')

            # Check if the query is present in the selected column of the current sheet (partial match)
            sheet_result = df[df.iloc[:, search_column].astype(str).str.contains(query, case=False, na=False)]

            # Filter out entries with 0.0 units
            sheet_result = sheet_result[sheet_result.iloc[:, 18] != 0]  # Assuming Column S is at position 18

            # Add the results to the overall result DataFrame
            result = pd.concat([result, sheet_result])

            # Calculate the sum of units for the current sheet
            sheet_units_sum = sheet_result.iloc[:, 18].sum()  # Assuming Column S is at position 18
            sheet_cost_sum = sheet_result.iloc[:, 19].sum()   # Assuming Column T is at position 19

            # Extract unique routes for the current sheet and count each unique route only once
            unique_routes_count = sheet_result.iloc[:, 10].nunique()  # Assuming Column K is at position 10
            daily_routes_count[sheet_name] = unique_routes_count

            # Check if the sheet has non-zero units
            if sheet_units_sum > 0:
                # Add the daily sum to the total
                daily_units[sheet_name] = sheet_units_sum
                daily_cost[sheet_name] = sheet_cost_sum

                # Add the daily sum to the total
                total_units += sheet_units_sum
                total_cost += sheet_cost_sum

        return result, total_units, total_cost, daily_units, daily_cost, daily_routes_count
    except Exception as e:
        return str(e), 0, 0, {}, {}, {}

# Function to save search results to a text file
def save_to_text(daily_units, daily_cost, daily_routes_count, total_units, total_cost, filename):
    try:
        # Get the current script's directory and construct the path to the text file
        script_dir = os.path.dirname(os.path.abspath(__file__))
        text_path = os.path.join(script_dir, filename)

        # Save daily units, cost, and route count to a text file
        with open(text_path, 'w') as file:
            for day, units in daily_units.items():
                cost = daily_cost[day]
                route_count = daily_routes_count[day]
                file.write(f"{day}: {units:.2f} units, Cost: €{cost:.2f}, Unique Route Count: {route_count}\n")
            file.write(f"\nTotal Units: {total_units:.2f}\n")
            file.write(f"Total Cost: €{total_cost:.2f}\n")

        return text_path
    except Exception as e:
        return str(e)

# Function to handle file selection
def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
    if file_path:
        file_path_entry.delete(0, END)
        file_path_entry.insert(0, file_path)

# Function to handle search button click
def on_search_click():
    query = entry.get()
    excel_path = file_path_entry.get()

    if not excel_path:
        messagebox.showinfo("Search Result", "No file selected.")
        return

    # Perform the search
    result, total_units, total_cost, daily_units, daily_cost, daily_routes_count = search_excel(excel_path, query)

    if daily_units:
        # Save the daily units, cost, and route count to a text file
        filename = "search_result.txt"
        saved_path = save_to_text(daily_units, daily_cost, daily_routes_count, total_units, total_cost, filename)

        # Display a message with the path to the saved text file
        message = f"Daily units, cost, and route count saved to:\n{saved_path}"
        messagebox.showinfo("Search Result", message)

        # Create and place the Treeview widget for displaying results in a table format
        columns = ["Day", "Units", "Cost", "Unique Route Count"]
        result_treeview = ttk.Treeview(root, columns=columns, show="headings", selectmode="browse")

        # Add column headings
        for col in columns:
            result_treeview.heading(col, text=col)

        # Add data to the Treeview
        for day, units in daily_units.items():
            cost = daily_cost[day]
            route_count = daily_routes_count[day]
            result_treeview.insert("", "end", values=[day, f"{units:.2f}", f"€{cost:.2f}", route_count])

        result_treeview.grid(row=2, column=0, columnspan=4, padx=10, pady=10)

        # Add scrollbar to the Treeview
        treeview_scrollbar = Scrollbar(root, command=result_treeview.yview)
        treeview_scrollbar.grid(row=2, column=4, sticky="nsew")
        result_treeview.configure(yscrollcommand=treeview_scrollbar.set)

    else:
        messagebox.showinfo("Search Result", "No matching records found.")

# Create the main window
root = tk.Tk()
root.title("PineSearch")  # Set the title

# Configure the title bar
root.configure(bg="red")  # Set background color to red for the title bar

# Get the screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Set the window size to the screen width and 80% of the screen height
root.geometry(f"{screen_width}x{int(screen_height * 0.8)}")

# Create and place the search input box
entry = tk.Entry(root, width=60, bg="black", fg="white", insertbackground="white", font=("Courier", 12))  # Increased width and set colors
entry.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

# Create and place the choose file button
choose_file_button = Button(root, text="Choose Excel File", command=choose_file, bg="black", fg="white", font=("Courier", 12))
choose_file_button.grid(row=1, column=0, padx=10, pady=10)

# Create and place the file path entry
file_path_entry = Entry(root, width=60, bg="black", fg="white", font=("Courier", 12))
file_path_entry.grid(row=1, column=1, columnspan=2, padx=10, pady=10)

# Create and place the search button
search_button = tk.Button(root, text="Search", command=on_search_click, bg="black", fg="white", font=("Courier", 12))  # Set colors and font size
search_button.grid(row=0, column=3, padx=10, pady=10)

# Create and place the listbox for displaying results
result_listbox = ttk.Treeview(root, columns=("Day", "Units", "Cost", "Unique Route Count"), show="headings")
result_listbox.grid(row=2, column=0, columnspan=4, padx=10, pady=10)

# Add scrollbar to the listbox
scrollbar = Scrollbar(root, command=result_listbox.yview)
scrollbar.grid(row=2, column=4, sticky="nsew")
result_listbox.configure(yscrollcommand=scrollbar.set)

# Run the GUI main loop
root.mainloop()