import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Toplevel
import csv
import os
import re
import pyperclip
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
def update_status(selected_items, new_status):
    rows = []
    updated = False
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, mode='r') as file:
        reader = csv.reader(file)
        header = next(reader)  # Read the header row
        rows.append(header)  # Keep the header in the rows
        for row in reader:
            if len(row) == 5:
                # Check if the current row matches any of the selected items
                for item in selected_items:
                    work_order = str(tree.item(item)['values'][0]).strip()
                    serial_number = str(tree.item(item)['values'][1]).strip()
                    if row[0] == work_order and row[1] == serial_number:
                        row[2] = new_status  # Update the status
                        row[4] = timestamp  # Update the timestamp
                        updated = True
                        break  # No need to check other items if we found a match
            rows.append(row)
    if updated:
        with open(log_file, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(rows)
        display_log()  # Refresh log in UI
        update_dashboard()  # Refresh dashboard


# Variable to track visibility of complete entries
show_complete_entries = True

# Function to display the log entries in the treeview
def display_log():
    for i in tree.get_children():
        tree.delete(i)
    with open(log_file, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for row in reader:
            if show_complete_entries or row[2] != "Complete":  # Check the status
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
    else:
        messagebox.showinfo("No Match", "No matching entries found.")

# Function to handle updating the status of an entry
def handle_update_status():
    selected_items = tree.selection()
    if selected_items:
        new_status = combo_update_status.get()
        if new_status:
            update_status(selected_items, new_status)  # Pass the selected items and new status
            combo_update_status.set('')  # Clear the status combobox
        else:
            messagebox.showwarning("Input Error", "Please select a new status.")
    else:
        messagebox.showwarning("Selection Error", "Please select at least one entry to update.")

# Initialize sorting order for each column
sort_order = {col: False for col in ["Work Order", "Serial Number", "Status", "Notes", "Timestamp"]}
 
def sort_treeview(column_index):
     global sort_order
     rows = [tree.item(item)['values'] for item in tree.get_children()]
     column = tree_columns[column_index]
     sort_order[column] = not sort_order[column]
     rows.sort(key=lambda x: x[column_index], reverse=sort_order[column])
     
     for item in tree.get_children():
         tree.delete(item)
     for row in rows:
         tree.insert('', 'end', values=row)
 
def handle_delete_entry():
     selected_items = tree.selection()
     if selected_items:
         if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected entry/entries?"):
             rows = []
             with open(log_file, mode='r') as file:
                 reader = csv.reader(file)
                 header = next(reader)
                 rows.append(header)
                 for row in reader:
                     if len(row) == 5:
                         row_work_order = row[0].strip()
                         row_serial_number = row[1].strip()
                         if not any((row_work_order == str(tree.item(item)['values'][0]).strip() and
                                      row_serial_number == str(tree.item(item)['values'][1]).strip()) for item in selected_items):
                             rows.append(row)
 
             with open(log_file, mode='w', newline='') as file:
                 writer = csv.writer(file)
                 writer.writerows(rows)
 
             for item in selected_items:
                 tree.delete(item)
 
             update_dashboard()

# Function to copy the selected serial number to the clipboard
def copy_serial_number():
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        if len(item['values']) == 5:
            serial_number = item['values'][1]  # Get the serial number
            root.clipboard_clear()  # Clear the clipboard
            root.clipboard_append(serial_number)  # Append the serial number to the clipboard
            messagebox.showinfo("Copied", f"Serial Number '{serial_number}' copied to clipboard.")
        else:
            messagebox.showwarning("Selection Error", "Selected entry does not have the correct number of columns.")
    else:
        messagebox.showwarning("Selection Error", "Please select an entry to copy.")

# Function to copy the selected work order number to the clipboard
def copy_work_order_number():
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        if len(item['values']) == 5:
            work_order = item['values'][0]  # Get the work order number
            root.clipboard_clear()  # Clear the clipboard
            root.clipboard_append(work_order)  # Append the work order number to the clipboard
            messagebox.showinfo("Copied", f"Work Order '{work_order}' copied to clipboard.")
        else:
            messagebox.showwarning("Selection Error", "Selected entry does not have the correct number of columns.")
    else:
        messagebox.showwarning("Selection Error", "Please select an entry to copy.")

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
            edit_status = ttk.Combobox(edit_window, values=["Ordered", "Pending", "Replaced", "Returned", "Complete"])
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

# Function to handle deleting an entry
def handle_delete_entry():
    selected_items = tree.selection()
    if selected_items:
        if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected entry/entries?"):
            rows = []
            with open(log_file, mode='r') as file:
                reader = csv.reader(file)
                header = next(reader)
                rows.append(header)
                for row in reader:
                    if len(row) == 5:
                        row_work_order = row[0].strip()
                        row_serial_number = row[1].strip()
                        if not any((row_work_order == str(tree.item(item)['values'][0]).strip() and
                                     row_serial_number == str(tree.item(item)['values'][1]).strip()) for item in selected_items):
                            rows.append(row)

            with open(log_file, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerows(rows)

            for item in selected_items:
                tree.delete(item)

            update_dashboard()
    else:
        messagebox.showwarning("Selection Error", "Please select at least one entry to delete.")

def parse_lenovo_clipboard_minimal(clipboard_text):
    """Extract Work Order and Serial Number from Lenovo clipboard text."""
    data = {
        "WO": None,
        "Serial": None
    }
    # Match Work Order Number (WO: followed by 8-10 digits)
    wo_match = re.search(r"WO:\s*(\d{8,10})", clipboard_text)
    if wo_match:
        data["WO"] = wo_match.group(1)
    # Match Serial Number (look for "Serial Number" line followed by serial on next line)
    serial_match = re.search(r"Serial\s*Number\s*[\s:\-]*\n([A-Z0-9]{7,10})", clipboard_text, re.IGNORECASE)
    if not serial_match:
        # Alternative match if the above doesn't work - looks for Serial Number then optional : or space then the serial
        serial_match = re.search(r"Serial\s*Number\s*[:\s]*([A-Z0-9]{7,10})", clipboard_text, re.IGNORECASE)
    
    if serial_match:
        data["Serial"] = serial_match.group(1)
    return data

def fill_from_clipboard(work_order, serial_number, status, notes):
    try:
        clipboard_text = pyperclip.paste()
        parsed = parse_lenovo_clipboard_minimal(clipboard_text)

        if parsed["WO"]:
            work_order.delete(0, tk.END)
            work_order.insert(0, parsed["WO"])

        if parsed["Serial"]:
            serial_number.delete(0, tk.END)
            serial_number.insert(0, parsed["Serial"])

        status.set("Ordered")
        notes.delete("1.0", tk.END)

        messagebox.showinfo("Success", "Fields populated from clipboard.")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to process clipboard: {str(e)}")


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

# Function to update dashboard statistics
def update_dashboard():
    total_entries = 0
    ordered_count = 0
    pending_count = 0
    replaced_count = 0
    returned_count = 0
    complete_count = 0  # New count for "Complete"
    
    with open(log_file, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for row in reader:
            total_entries += 1
            if row[2] == "Ordered":
                ordered_count += 1
            elif row[2] == "Pending":
                pending_count += 1
            elif row[2] == "Replaced":
                replaced_count += 1
            elif row[2] == "Returned":
                returned_count += 1
            elif row[2] == "Complete":  # Count for "Complete"
                complete_count += 1
    
    lbl_total_entries.config(text=f"Total Entries: {total_entries}")
    lbl_ordered_count.config(text=f"Ordered: {ordered_count}")
    lbl_pending_count.config(text=f"Pending: {pending_count}")
    lbl_replaced_count.config(text=f"Replaced: {replaced_count}")
    lbl_returned_count.config(text=f"Returned: {returned_count}")
    lbl_complete_count.config(text=f"Complete: {complete_count}")  # Update label for "Complete"
    root.update_idletasks()

# Initialize the log file
initialize_log()

# Create the main window
root = tk.Tk()
root.title("LCD Tracking System")
root.geometry("1009x743")
root.configure(bg="#f0f0f0")

# Bind the Control-C key combination to the copy_serial_number function
root.bind('<Control-c>', lambda event: copy_serial_number())

# Configure styles
style = ttk.Style()
style.configure("TLabel", background="#f0f0f0", font=("Helvetica", 12))
style.configure("TButton", background="#4CAF50", foreground="black", font=("Helvetica", 10, "bold"))
style.configure("TCombobox", background="white", font=("Helvetica", 10))
style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"))

# Button hover effects
def on_enter(e):
    e.widget['background'] = '#45a049'  # Darker green on hover

def on_leave(e):
    e.widget['background'] = '#4CAF50'  # Original color

# Tooltips class
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        if self.tooltip_window is not None:
            return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, background="white", relief="solid", borderwidth=1)
        label.pack()

    def hide_tooltip(self, event):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

# Create and place the search bar
frame_search = ttk.LabelFrame(root, text="Search", padding=(10, 5))
frame_search.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
ttk.Label(frame_search, text="Search:").grid(row=0, column=0, padx=5, pady=5)
entry_search = ttk.Entry(frame_search)
entry_search.grid(row=0, column=1, padx=5, pady=5)

# Bind the Enter key to the search_log function
entry_search.bind('<Return>', lambda event: search_log())

btn_search = ttk.Button(frame_search, text="Search", command=search_log)
btn_search.grid(row=0, column=2, padx=5, pady=5)
ToolTip(btn_search, "Search for entries")

# Create and place the refresh button and other buttons
frame_buttons = ttk.Frame(root)
frame_buttons.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

btn_refresh = ttk.Button(frame_buttons, text="Refresh Log", command=handle_refresh)
btn_refresh.grid(row=0, column=0, padx=5, pady=5)
ToolTip(btn_refresh, "Refresh the log entries")

btn_export = ttk.Button(frame_buttons, text="Export to CSV", command=export_to_csv)
btn_export.grid(row=0, column=1, padx=5, pady=5)
ToolTip(btn_export, "Export log to CSV file")

btn_import = ttk.Button(frame_buttons, text="Import from CSV", command=import_from_csv)
btn_import.grid(row=0, column=2, padx=5, pady=5)
ToolTip(btn_import, "Import log from CSV file")

# Add the toggle button for complete entries
btn_toggle_complete = ttk.Button(frame_buttons, text="Hide Complete Entries", command=lambda: toggle_complete_entries())
btn_toggle_complete.grid(row=0, column=3, padx=5, pady=5)
ToolTip(btn_toggle_complete, "Show/Hide complete entries")

# Function to toggle visibility of complete entries
def toggle_complete_entries():
    global show_complete_entries
    show_complete_entries = not show_complete_entries  # Toggle the state
    btn_toggle_complete.config(text="Show Complete Entries" if not show_complete_entries else "Hide Complete Entries")
    display_log()  # Refresh the log display

# Add the LCD Script button and functionality
btn_lcd_script = ttk.Button(frame_buttons, text="LCD Script", command=lambda: open_lcd_script_window())
btn_lcd_script.grid(row=0, column=4, padx=5, pady=5)
ToolTip(btn_lcd_script, "Open LCD script for easy copying")

# Add the Paste Lenovo Info button
btn_paste_info = ttk.Button(frame_buttons, text="Paste Lenovo Info", command=lambda: fill_from_clipboard(entry_work_order, entry_serial_number, combo_status, text_notes))
btn_paste_info.grid(row=0, column=5, padx=5, pady=5)
ToolTip(btn_paste_info, "Paste Work Order and Serial Number from clipboard")

def select_all_text(event):
    widget = event.widget
    widget.after(1, lambda: widget.tag_add('sel', '1.0', 'end'))

def open_lcd_script_window():
    script_window = Toplevel(root)
    script_window.title("LCD Script")
    script_window.geometry("450x220")
    script_window.resizable(False, False)

    # First text box
    text_box_1 = tk.Text(script_window, height=4, wrap=tk.WORD)
    text_box_1.insert(tk.END, "Student dropped laptop which cracked the LCD.")
    text_box_1.configure(state='normal')
    text_box_1.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
    text_box_1.focus_set()
    text_box_1.bind('<FocusIn>', select_all_text)

    # Second text box
    text_box_2 = tk.Text(script_window, height=4, wrap=tk.WORD)
    text_box_2.insert(tk.END, "Replaced LCD. Laptop boots and displays image correctly.")
    text_box_2.configure(state='normal')
    text_box_2.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
    text_box_2.bind('<FocusIn>', select_all_text)

    # Close button
    btn_close = ttk.Button(script_window, text="Close", command=script_window.destroy)
    btn_close.pack(pady=10)
    # Bind Enter key on close button to invoke its command
    btn_close.bind('<Return>', lambda event: script_window.destroy())

    # List of widgets for tab order
    focusable_widgets = [text_box_1, text_box_2, btn_close]

    # Bind the Tab key to switch focus between text boxes and the close button
    def focus_next_widget(event):
        current_widget = event.widget
        try:
            next_index = (focusable_widgets.index(current_widget) + 1) % len(focusable_widgets)
        except ValueError:
            next_index = 0
        focusable_widgets[next_index].focus_set()
        return "break"  # Prevent default tab behavior

    for widget in focusable_widgets:
        widget.bind('<Tab>', focus_next_widget)

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
combo_status = ttk.Combobox(frame_add_entry, values=["Ordered", "Pending", "Replaced", "Returned", "Complete"])
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
btn_add_entry.bind("<Enter>", on_enter)
btn_add_entry.bind("<Leave>", on_leave)
ToolTip(btn_add_entry, "Add a new entry to the log")

# Create and place widgets for updating the status of an entry
frame_update_status = ttk.LabelFrame(root, text="Update Status", padding=(10, 5))
frame_update_status.grid(row=3, column=0, padx=10, pady=10, sticky="ew")

ttk.Label(frame_update_status, text="New Status:").grid(row=0, column=0, padx=5, pady=5)
combo_update_status = ttk.Combobox(frame_update_status, values=["Ordered", "Pending", "Replaced", "Returned", "Complete"])
combo_update_status.grid(row=0, column=1, padx=5, pady=5)

btn_update_status = ttk.Button(frame_update_status, text="Update Status", command=handle_update_status)
btn_update_status.grid(row=0, column=2, padx=5, pady=5)
btn_update_status.bind("<Enter>", on_enter)
btn_update_status.bind("<Leave>", on_leave)
ToolTip(btn_update_status, "Update the status of the selected entry")

btn_delete_entry = ttk.Button(frame_update_status, text="Delete Entry", command=handle_delete_entry)
btn_delete_entry.grid(row=0, column=3, padx=5, pady=5)
btn_delete_entry.bind("<Enter>", on_enter)
btn_delete_entry.bind("<Leave>", on_leave)
ToolTip(btn_delete_entry, "Delete the selected entry")

btn_edit_entry = ttk.Button(frame_update_status, text="Edit Entry", command=handle_edit_entry)
btn_edit_entry.grid(row=0, column=4, padx=5, pady=5)
btn_edit_entry.bind("<Enter>", on_enter)
btn_edit_entry.bind("<Leave>", on_leave)
ToolTip(btn_edit_entry, "Edit the selected entry")

# Create and place a treeview widget for displaying the log entries
frame_tree = ttk.Frame(root)
frame_tree.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

tree_columns = ("Work Order", "Serial Number", "Status", "Notes", "Timestamp")
tree = ttk.Treeview(frame_tree, columns=tree_columns, show='headings')
for col in tree_columns:
    tree.heading(col, text=col)
tree.pack(fill=tk.BOTH, expand=True)

# Bind the sorting function to each column header
for index, col in enumerate(tree_columns):
    tree.heading(col, text=col, command=lambda idx=index: sort_treeview(idx))

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

lbl_complete_count = ttk.Label(frame_dashboard, text="Complete: 0", font=("Helvetica", 10, "bold"))  # New label for "Complete"
lbl_complete_count.grid(row=5, column=0, padx=5, pady=5, sticky="w")

# Display the initial log entries and update dashboard
display_log()
update_dashboard()

# Create a context menu
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy Serial Number", command=copy_serial_number)
context_menu.add_command(label="Copy Work Order Number", command=copy_work_order_number)
context_menu.add_command(label="Edit Entry", command=handle_edit_entry)
context_menu.add_command(label="Delete Entry", command=handle_delete_entry)

# Function to show the context menu
def show_context_menu(event):
    selected_item = tree.selection()
    if selected_item:
        context_menu.post(event.x_root, event.y_root)

# Bind the right-click event to the treeview
tree.bind("<Button-3>", show_context_menu)

# Run the application
root.mainloop()
