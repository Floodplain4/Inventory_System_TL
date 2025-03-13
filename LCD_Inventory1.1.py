import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel
import csv
import os
from datetime import datetime

# Define the file name for the log
log_file = 'lcd_log.csv'

# Function to initialize the log file if it doesn't exist
def initialize_log():
    if not os.path.exists(log_file):
        with open(log_file, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Work Order', 'Serial Number', 'Status', 'Notes', 'Timestamp'])

# Function to add a new entry to the log
def add_entry(work_order, serial_number, status, notes):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([work_order, serial_number, status, notes, timestamp])

# Function to update the status of an entry in the log
def update_status(work_order, serial_number, new_status):
    rows = []
    updated = False
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(log_file, mode='r') as file:
        reader = csv.reader(file)
        header = next(reader)  # Read the header row
        rows.append(header)  # Keep the header in the rows
        for row in reader:
            print(f"Checking row: {row}")  # Debugging output
            if len(row) == 5:
                print(f"Row details - Work Order: {row[0]}, Serial Number: {row[1]}, Status: {row[2]}")  # More detailed output
                if row[0] == work_order and row[1] == serial_number:
                    print(f"Match found. Updating status from {row[2]} to {new_status}")
                    row[2] = new_status
                    row[4] = timestamp
                    updated = True
                    print(f"Updated row: {row}")
            rows.append(row)

    if updated:
        with open(log_file, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(rows)
        print("Log file updated")
        display_log()  # Refresh log in UI
        update_dashboard()  # Refresh dashboard
    else:
        print("No matching entry found to update")

# Function to display the log entries in the treeview
def display_log():
    for i in tree.get_children():
        tree.delete(i)
    with open(log_file, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for row in reader:
            tree.insert('', 'end', values=row)

# Function to handle adding a new entry
def handle_add_entry():
    work_order = entry_work_order.get()
    serial_number = entry_serial_number.get()
    status = combo_status.get()
    notes = text_notes.get("1.0", tk.END).strip()
    
    # Validate work order and serial number lengths
    if len(work_order) > 10:
        messagebox.showwarning("Input Error", "Work Order number must be 10 digits or less.")
        return
    if len(serial_number) > 8:
        messagebox.showwarning("Input Error", "Serial Number must be 8 digits or less.")
        return
    
    if work_order and serial_number and status:
        add_entry(work_order, serial_number, status, notes)
        display_log()
        update_dashboard()
        entry_work_order.delete(0, tk.END)
        entry_serial_number.delete(0, tk.END)
        combo_status.set('')
        text_notes.delete("1.0", tk.END)
    else:
        messagebox.showwarning("Input Error", "Please fill in all fields.")

# Function to handle updating the status of an entry
def handle_update_status():
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        if len(item['values']) == 5:
            work_order, serial_number, _, _, _ = item['values']
            new_status = combo_update_status.get()
            print(f"Selected Work Order: {work_order}, Serial Number: {serial_number}, New Status: {new_status}")  # Debugging output
            if new_status:
                update_status(work_order, serial_number, new_status)
                # Update the selected item's status in the treeview
                tree.item(selected_item, values=(work_order, serial_number, new_status, item['values'][3], item['values'][4]))
                combo_update_status.set('')
            else:
                messagebox.showwarning("Input Error", "Please select a new status.")
        else:
            messagebox.showwarning("Data Error", "Selected entry does not have the correct number of columns.")
    else:
        messagebox.showwarning("Selection Error", "Please select an entry to update.")

def handle_delete_entry():
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        print(f"Selected item values: {item['values']}")  # Print the selected item values
        if len(item['values']) == 5:
            # Convert work_order and serial_number to strings
            work_order = str(item['values'][0])  # Convert to string
            serial_number = str(item['values'][1])  # Convert to string
            print(f"Attempting to delete entry: Work Order: '{work_order}', Serial Number: '{serial_number}'")

            # Read the current log file
            rows = []
            entry_found = False  # Flag to check if the entry was found
            try:
                with open(log_file, mode='r') as file:
                    reader = csv.reader(file)
                    header = next(reader)  # Read the header row
                    rows.append(header)  # Keep the header in the rows
                    for row in reader:
                        # Print the row values and their types for debugging
                        print(f"Row values: {row}, Types: {[type(value) for value in row]}")

                        # Ensure all values are treated as strings
                        row = [str(value) for value in row]  # Convert all values to strings
                        row_work_order = row[0].strip()  # Strip whitespace
                        row_serial_number = row[1].strip()  # Strip whitespace

                        print(f"Comparing: '{row_work_order}' with '{work_order.strip()}' and '{row_serial_number}' with '{serial_number.strip()}'")

                        # Only keep rows that do not match the selected entry
                        if not (row_work_order == work_order.strip() and row_serial_number == serial_number.strip()):
                            rows.append(row)
                        else:
                            entry_found = True  # Mark that we found the entry to delete
                            print(f"Entry matched and will be deleted: {row}")

                if entry_found:
                    # Write back the updated log
                    with open(log_file, mode='w', newline='') as file:
                        writer = csv.writer(file)
                        writer.writerows(rows)
                        print("Log file updated after deletion.")

                    # Remove the entry from the treeview
                    tree.delete(selected_item)

                    # Update dashboard stats
                    update_dashboard()
                else:
                    print("No matching entry found to delete.")
                    messagebox.showwarning("Deletion Error", "No matching entry found to delete.")
            except Exception as e:
                print(f"An error occurred while processing the log file: {e}")
                messagebox.showerror("File Error", f"An error occurred: {e}")
        else:
            messagebox.showwarning("Data Error", "Selected entry does not have the correct number of columns.")
    else:
        messagebox.showwarning("Selection Error", "Please select an entry to delete.")
        
# Function to handle editing an entry
def handle_edit_entry():
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        if len(item['values']) == 5:
            # Retrieve values from the selected item
            work_order, serial_number, status, notes, _ = item['values']
            
            # Ensure work_order and serial_number are strings
            work_order = str(work_order)
            serial_number = str(serial_number)
            
            edit_window = Toplevel(root)
            edit_window.title("Edit Entry")
            
            # Create and place the entry fields
            ttk.Label(edit_window, text="Work Order:").grid(row=0, column=0, padx=5, pady=5)
            edit_work_order = ttk.Entry(edit_window)
            edit_work_order.grid(row=0, column=1, padx=5, pady=5)
            edit_work_order.insert(0, work_order)
            
            ttk.Label(edit_window, text="Serial Number:").grid(row=1, column=0, padx=5, pady=5)
            edit_serial_number = ttk.Entry(edit_window)
            edit_serial_number.grid(row=1, column=1, padx=5, pady=5)
            edit_serial_number.insert(0, serial_number)
            
            ttk.Label(edit_window, text="Status:").grid(row=2, column=0, padx=5, pady=5)
            edit_status = ttk.Combobox(edit_window, values=["Ordered", "Pending", "Replaced", "Returned"])
            edit_status.grid(row=2, column=1, padx=5, pady=5)
            edit_status.set(status)
            
            ttk.Label(edit_window, text="Notes:").grid(row=3, column=0, padx=5, pady=5)
            edit_notes = tk.Text(edit_window, height=4)
            edit_notes.grid(row=3, column=1, padx=5, pady=5)
            edit_notes.insert(tk.END, notes)
            
            # Save button function
            def save_edits():
                new_work_order = edit_work_order.get().strip()  # Strip any leading/trailing spaces
                new_serial_number = edit_serial_number.get().strip()  # Strip any leading/trailing spaces
                new_status = edit_status.get()
                new_notes = edit_notes.get("1.0", tk.END).strip()
                
                rows = []
                with open(log_file, mode='r') as file:
                    reader = csv.reader(file)
                    header = next(reader)  # Read the header row
                    rows.append(header)  # Keep the header in the rows
                    for row in reader:
                        # Convert row values to strings to avoid AttributeError
                        row_work_order = str(row[0]).strip()
                        row_serial_number = str(row[1]).strip()
                        
                        if row_work_order == work_order.strip() and row_serial_number == serial_number.strip():
                            # Update the row with new values
                            row = [new_work_order, new_serial_number, new_status, new_notes, row[4]]  # Keep the original timestamp
                        rows.append(row)
                
                with open(log_file, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerows(rows)
                
                display_log()  # Refresh the displayed log
                update_dashboard()  # Refresh the dashboard
                edit_window.destroy()  # Close the edit window
            
            # Create and place the save button
            btn_save = ttk.Button(edit_window, text="Save", command=save_edits)
            btn_save.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        else:
            messagebox.showwarning("Data Error", "Selected entry does not have the correct number of columns.")
    else:
        messagebox.showwarning("Selection Error", "Please select an entry to edit.")

# Function to handle refreshing the log without removing returned entries
def handle_refresh():
    display_log()
    update_dashboard()

# Function to export log to CSV
def export_to_csv():
    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv', filetypes=[("CSV files", "*.csv")])
    if export_file_path:
        with open(export_file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Work Order', 'Serial Number', 'Status', 'Notes', 'Timestamp'])
            with open(log_file, mode='r') as log_file_read:
                reader = csv.reader(log_file_read)
                next(reader)  # Skip header row
                for row in reader:
                    writer.writerow(row)
        messagebox.showinfo("Export Successful", f"Log exported successfully to {export_file_path}")

# Function to import log from CSV
def import_from_csv():
    import_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if import_file_path:
        with open(import_file_path, mode='r') as file:
            reader = csv.reader(file)
            next(reader)  # Skip header row
            for row in reader:
                add_entry(row[0], row[1], row[2], row[3])
        display_log()
        update_dashboard()
        messagebox.showinfo("Import Successful", f"Log imported successfully from {import_file_path}")

# Function to search log entries and highlight the closest matching entry
def search_log():
    query = entry_search.get().strip().lower()
    if not query:
        return  # If search is empty, do nothing

    closest_match = None
    for item in tree.get_children():
        values = tree.item(item, 'values')
        if query in values[0].lower() or query in values[1].lower():
            closest_match = item
            break  # Stop at the first match

    if closest_match:
        tree.selection_set(closest_match)
        tree.focus(closest_match)
        tree.see(closest_match)  # Ensure it's visible in the view
        
# Function to update dashboard statistics
def update_dashboard():
    print("Updating dashboard...")
    total_entries = 0
    ordered_count = 0
    pending_count = 0
    replaced_count = 0
    returned_count = 0
    
    with open(log_file, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for row in reader:
            print(row)  # Print each row to verify data
            total_entries += 1
            if row[2] == "Ordered":
                ordered_count += 1
            elif row[2] == "Pending":
                pending_count += 1
            elif row[2] == "Replaced":
                replaced_count += 1
            elif row[2] == "Returned":
                returned_count += 1
    
    print(f"Total Entries: {total_entries}")
    print(f"Ordered: {ordered_count}")
    print(f"Pending: {pending_count}")
    print(f"Replaced: {replaced_count}")
    print(f"Returned: {returned_count}")
    
    lbl_total_entries.config(text=f"Total Entries: {total_entries}")
    lbl_ordered_count.config(text=f"Ordered: {ordered_count}")
    lbl_pending_count.config(text=f"Pending: {pending_count}")
    lbl_replaced_count.config(text=f"Replaced: {replaced_count}")
    lbl_returned_count.config(text=f"Returned: {returned_count}")
    root.update_idletasks()

# Initialize the log file
initialize_log()

# Create the main window
root = tk.Tk()
root.title("LCD Tracking System")
root.geometry("1009x743")
root.configure(bg="#f0f0f0")

style = ttk.Style()
style.configure("TLabel", background="#f0f0f0")
style.configure("TButton", background="#4CAF50", foreground="black")
style.configure("TCombobox", background="white")
style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))

# Create and place the search bar
frame_search = ttk.LabelFrame(root, text="Search", padding=(10, 5))
frame_search.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
ttk.Label(frame_search, text="Search:").grid(row=0, column=0, padx=5, pady=5)
entry_search = ttk.Entry(frame_search)
entry_search.grid(row=0, column=1, padx=5, pady=5)

btn_search = ttk.Button(frame_search, text="Search", command=search_log)
btn_search.grid(row=0, column=2, padx=5, pady=5)

# Create and place the refresh button and other buttons
frame_buttons = ttk.Frame(root)
frame_buttons.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

btn_refresh = ttk.Button(frame_buttons, text="Refresh Log", command=handle_refresh)
btn_refresh.grid(row=0, column=0, padx=5, pady=5)

btn_export = ttk.Button(frame_buttons, text="Export to CSV", command=export_to_csv)
btn_export.grid(row=0, column=1, padx=5, pady=5)

btn_import = ttk.Button(frame_buttons, text="Import from CSV", command=import_from_csv)
btn_import.grid(row=0, column=2, padx=5, pady=5)

# Create and place widgets for adding a new entry
frame_add_entry = ttk.LabelFrame(root, text="Add New Entry", padding=(10, 5))
frame_add_entry.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

ttk.Label(frame_add_entry, text="Work Order:").grid(row=0, column=0, padx=5, pady=5)
entry_work_order = ttk.Entry(frame_add_entry)
entry_work_order.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame_add_entry, text="Serial Number:").grid(row=1, column=0, padx=5, pady=5)
entry_serial_number = ttk.Entry(frame_add_entry)
entry_serial_number.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame_add_entry, text="Status:").grid(row=2, column=0, padx=5, pady=5)
combo_status = ttk.Combobox(frame_add_entry, values=["Ordered", "Pending", "Replaced", "Returned"])
combo_status.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(frame_add_entry, text="Notes:").grid(row=3, column=0, padx=5, pady=5)
text_notes = tk.Text(frame_add_entry, height=4)
text_notes.grid(row=3, column=1, padx=5, pady=5)

# Add a scrollbar to the text widget
scrollbar_notes = ttk.Scrollbar(frame_add_entry, orient=tk.VERTICAL, command=text_notes.yview)
text_notes.configure(yscroll=scrollbar_notes.set)
scrollbar_notes.grid(row=3, column=2, sticky='ns')

btn_add_entry = ttk.Button(frame_add_entry, text="Add Entry", command=handle_add_entry)
btn_add_entry.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

# Create and place widgets for updating the status of an entry
frame_update_status = ttk.LabelFrame(root, text="Update Status", padding=(10, 5))
frame_update_status.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

ttk.Label(frame_update_status, text="New Status:").grid(row=0, column=0, padx=5, pady=5)
combo_update_status = ttk.Combobox(frame_update_status, values=["Ordered", "Pending", "Replaced", "Returned"])
combo_update_status.grid(row=0, column=1, padx=5, pady=5)

btn_update_status = ttk.Button(frame_update_status, text="Update Status", command=handle_update_status)
btn_update_status.grid(row=0, column=2, padx=5, pady=5)

btn_delete_entry = ttk.Button(frame_update_status, text="Delete Entry", command=handle_delete_entry)
btn_delete_entry.grid(row=0, column=3, padx=5, pady=5)

btn_edit_entry = ttk.Button(frame_update_status, text="Edit Entry", command=handle_edit_entry)
btn_edit_entry.grid(row=0, column=4, padx=5, pady=5)

# Create and place a treeview widget for displaying the log entries
frame_tree = ttk.Frame(root)
frame_tree.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

tree_columns = ("Work Order", "Serial Number", "Status", "Notes", "Timestamp")
tree = ttk.Treeview(frame_tree, columns=tree_columns, show='headings')
for col in tree_columns:
    tree.heading(col, text=col)
tree.pack(fill=tk.BOTH, expand=True)

# Add a scrollbar to the treeview
scrollbar = ttk.Scrollbar(frame_tree, orient=tk.VERTICAL, command=tree.yview)
tree.configure(yscroll=scrollbar.set)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Create and place dashboard for summary statistics
frame_dashboard = ttk.LabelFrame(root, text="Dashboard", padding=(10, 5))
frame_dashboard.grid(row=0, column=1, rowspan=4, padx=10, pady=10, sticky="nsew")

lbl_total_entries = ttk.Label(frame_dashboard, text="Total Entries: 0", font=("Helvetica", 10, "bold"))
lbl_total_entries.grid(row=0, column=0, padx=5, pady=5, sticky="w")

lbl_ordered_count = ttk.Label(frame_dashboard, text="Ordered: 0", font=("Helvetica", 10, "bold"))
lbl_ordered_count.grid(row=1, column=0, padx=5, pady=5, sticky="w")

lbl_pending_count = ttk.Label(frame_dashboard, text="Pending: 0", font=("Helvetica", 10, "bold"))
lbl_pending_count.grid(row=2, column=0, padx=5, pady=5, sticky="w")

lbl_replaced_count = ttk.Label(frame_dashboard, text="Replaced: 0", font=("Helvetica", 10, "bold"))
lbl_replaced_count.grid(row=3, column=0, padx=5, pady=5, sticky="w")

lbl_returned_count = ttk.Label(frame_dashboard, text="Returned: 0", font=("Helvetica", 10, "bold"))
lbl_returned_count.grid(row=4, column=0, padx=5, pady=5, sticky="w")

# Display the initial log entries and update dashboard
display_log()
update_dashboard()

# Run the application
root.mainloop()