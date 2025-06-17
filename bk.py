import pandas as pd
from tkinter import *
from tkinter import messagebox, simpledialog, ttk
from datetime import datetime # For date validation

# Initialize Excel file and DataFrame
EXCEL_FILE = 'trading_book.xlsx'

def init_excel_file():
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Date', 'Ticker', 'Trade_Type', 'Quantity', 'Price', 'Total', 'Notes'])
        df.to_excel(EXCEL_FILE, index=False)

def load_data():
    try:
        return pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        messagebox.showerror("Error", "Excel file not found. Initializing a new one.")
        init_excel_file()
        return pd.read_excel(EXCEL_FILE)

def add_record(date, ticker, trade_type, quantity, price, notes):
    try:
        total = quantity * price  # Calculate total trade value
        new_record = pd.DataFrame({'Date': [date], 'Ticker': [ticker], 'Trade_Type': [trade_type],
                                   'Quantity': [quantity], 'Price': [price], 'Total': [total], 'Notes': [notes]})
        df = load_data()
        df = pd.concat([df, new_record], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to add record: {e}")
        return False

def edit_record(index, date, ticker, trade_type, quantity, price, notes):
    df = load_data()
    if 0 <= index < len(df):
        try:
            df.at[index, 'Date'] = date
            df.at[index, 'Ticker'] = ticker
            df.at[index, 'Trade_Type'] = trade_type
            df.at[index, 'Quantity'] = quantity
            df.at[index, 'Price'] = price
            df.at[index, 'Total'] = quantity * price  # Update total trade value
            df.at[index, 'Notes'] = notes
            df.to_excel(EXCEL_FILE, index=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit record: {e}")
            return False
    else:
        messagebox.showerror("Error", "Invalid index for editing.")
        return False


def delete_record(index):
    df = load_data()
    if 0 <= index < len(df):
        try:
            df = df.drop(index).reset_index(drop=True)
            df.to_excel(EXCEL_FILE, index=False)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete record: {e}")
            return False
    else:
        messagebox.showerror("Error", "Invalid index for deletion.")
        return False

# --- UI Functions ---

def validate_input(date_str, quantity_str, price_str, trade_type):
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Validation Error", "Date must be in YYYY-MM-DD format.")
        return False, None, None

    try:
        quantity = float(quantity_str)
        price = float(price_str)
        if quantity <= 0 or price <= 0:
            messagebox.showerror("Validation Error", "Quantity and Price must be positive numbers.")
            return False, None, None
    except ValueError:
        messagebox.showerror("Validation Error", "Quantity and Price must be valid numbers.")
        return False, None, None

    if trade_type.lower() not in ['buy', 'sell']:
        messagebox.showerror("Validation Error", "Trade Type must be 'Buy' or 'Sell'.")
        return False, None, None
        
    return True, quantity, price

def add_edit_form(is_edit=False, record_index=None, current_data=None, update_callback=None):
    form_window = Toplevel()
    form_window.title("Edit Record" if is_edit else "Add Record")

    labels = ['Date (YYYY-MM-DD):', 'Ticker:', 'Trade Type (Buy/Sell):', 'Quantity:', 'Price:', 'Notes:']
    entries = {}

    for i, label_text in enumerate(labels):
        Label(form_window, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="w")
        entry = Entry(form_window)
        entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        entries[label_text.replace(':', '').strip()] = entry

    if is_edit and current_data:
        entries['Date (YYYY-MM-DD)'].insert(0, current_data['Date'])
        entries['Ticker'].insert(0, current_data['Ticker'])
        entries['Trade Type (Buy/Sell)'].insert(0, current_data['Trade_Type'])
        entries['Quantity'].insert(0, str(current_data['Quantity']))
        entries['Price'].insert(0, str(current_data['Price']))
        entries['Notes'].insert(0, current_data['Notes'])

    def save_action():
        date = entries['Date (YYYY-MM-DD)'].get()
        ticker = entries['Ticker'].get()
        trade_type = entries['Trade Type (Buy/Sell)'].get()
        quantity_str = entries['Quantity'].get()
        price_str = entries['Price'].get()
        notes = entries['Notes'].get()

        is_valid, quantity, price = validate_input(date, quantity_str, price_str, trade_type)
        if not is_valid:
            return

        if is_edit:
            if edit_record(record_index, date, ticker, trade_type, quantity, price, notes):
                messagebox.showinfo("Success", "Record edited successfully.")
                if update_callback:
                    update_callback()
                form_window.destroy()
        else:
            if add_record(date, ticker, trade_type, quantity, price, notes):
                messagebox.showinfo("Success", "Record added successfully.")
                form_window.destroy()

    Button(form_window, text="Save", command=save_action).grid(row=len(labels), column=0, padx=5, pady=10)
    Button(form_window, text="Cancel", command=form_window.destroy).grid(row=len(labels), column=1, padx=5, pady=10)

    form_window.grab_set() # Make this window modal
    root.wait_window(form_window)


def show_records():
    records_window = Toplevel()
    records_window.title("Trading Records")
    records_window.geometry("900x600")

    df = load_data()
    # Ensure 'Date' column is string for consistent display and search
    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d') if pd.api.types.is_datetime64_any_dtype(df['Date']) else df['Date'].astype(str)

    # Search and Filter Frame
    control_frame = Frame(records_window)
    control_frame.pack(pady=10, fill='x')

    Label(control_frame, text="Search:").pack(side=LEFT, padx=5)
    search_entry = Entry(control_frame, width=30)
    search_entry.pack(side=LEFT, padx=5)

    Label(control_frame, text="Filter by Type:").pack(side=LEFT, padx=5)
    trade_type_filter = ttk.Combobox(control_frame, values=["All", "Buy", "Sell"], state="readonly", width=10)
    trade_type_filter.set("All")
    trade_type_filter.pack(side=LEFT, padx=5)

    # Treeview for structured display
    tree_frame = Frame(records_window)
    tree_frame.pack(expand=True, fill='both', padx=10, pady=10)

    tree_scroll_y = Scrollbar(tree_frame, orient="vertical")
    tree_scroll_y.pack(side="right", fill="y")
    tree_scroll_x = Scrollbar(tree_frame, orient="horizontal")
    tree_scroll_x.pack(side="bottom", fill="x")

    tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set, selectmode="browse")
    tree.pack(expand=True, fill='both')

    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)

    tree["columns"] = list(df.columns)
    tree["show"] = "headings" # Hide the default first column (index)

    for col in df.columns:
        tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))
        tree.column(col, width=120, anchor="center")
    
    # Add a hidden column for the original index, useful for editing/deleting
    tree["columns"] = ("#0",) + tuple(df.columns)
    tree.heading("#0", text="Index", command=lambda : treeview_sort_column(tree, "#0", False))
    tree.column("#0", width=50, anchor="center")

    def populate_tree(data_frame):
        # Clear existing entries
        for item in tree.get_children():
            tree.delete(item)
        
        for idx, row in data_frame.iterrows():
            # Insert with original index in the hidden #0 column
            tree.insert("", "end", iid=str(idx), text=str(idx), values=row.tolist())

    def treeview_sort_column(tv, col, reverse):
        # Handle sorting for the "Index" column (which is #0)
        if col == "#0":
            l = [(int(tv.set(k, col)), k) for k in tv.get_children('')]
        else:
            l = [(tv.set(k, col), k) for k in tv.get_children('')]

        l.sort(reverse=reverse)

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # reverse sort next time
        tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

    def apply_filters_and_search():
        current_df = load_data()
        current_df['Date'] = current_df['Date'].dt.strftime('%Y-%m-%d') if pd.api.types.is_datetime64_any_dtype(current_df['Date']) else current_df['Date'].astype(str)

        search_term = search_entry.get().lower()
        filter_type = trade_type_filter.get()

        filtered_df = current_df.copy()

        if filter_type != "All":
            filtered_df = filtered_df[filtered_df['Trade_Type'].str.lower() == filter_type.lower()]

        if search_term:
            filtered_df = filtered_df[filtered_df.apply(lambda row: row.astype(str).str.lower().str.contains(search_term).any(), axis=1)]
        
        populate_tree(filtered_df)

    search_entry.bind("<KeyRelease>", lambda event: apply_filters_and_search())
    trade_type_filter.bind("<<ComboboxSelected>>", lambda event: apply_filters_and_search())

    # Initial population
    populate_tree(df)

    # Edit and Delete Buttons
    action_frame = Frame(records_window)
    action_frame.pack(pady=10)

    def edit_selected_record():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a record to edit.")
            return
        
        # The iid of the treeview item is the original DataFrame index
        selected_index = int(tree.item(selected_item[0], "text")) # Get the text from the hidden #0 column

        # Get current data for pre-filling the form
        df_current = load_data()
        current_record_data = df_current.iloc[selected_index].to_dict()
        
        add_edit_form(is_edit=True, record_index=selected_index, current_data=current_record_data, update_callback=apply_filters_and_search)

    def delete_selected_record():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a record to delete.")
            return

        selected_index = int(tree.item(selected_item[0], "text")) # Get the text from the hidden #0 column
        
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete record at index {selected_index}?"):
            if delete_record(selected_index):
                messagebox.showinfo("Success", "Record deleted successfully.")
                apply_filters_and_search() # Refresh the view

    Button(action_frame, text="Edit Selected", command=edit_selected_record).pack(side=LEFT, padx=5)
    Button(action_frame, text="Delete Selected", command=delete_selected_record).pack(side=LEFT, padx=5)


def gui_main():
    global root # Make root accessible globally for add_edit_form
    root = Tk()
    root.title('Trading Bookkeeping System')
    root.geometry("400x250")

    # Center the window
    root.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() / 2) - (root.winfo_width() / 2)
    y = root.winfo_y() + (root.winfo_height() / 2) - (root.winfo_height() / 2)
    root.geometry(f'+{int(x)}+{int(y)}')


    Label(root, text="Trading Bookkeeping", font=("Arial", 16)).pack(pady=20)

    Button(root, text='Add New Record', command=lambda: add_edit_form(update_callback=show_records)).pack(pady=5, fill='x', padx=50)
    Button(root, text='Show All Records', command=show_records).pack(pady=5, fill='x', padx=50)

    root.mainloop()

if __name__ == '__main__':
    init_excel_file()
    gui_main()