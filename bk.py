import pandas as pd
from tkinter import *
from tkinter import messagebox, simpledialog, ttk, filedialog
from datetime import datetime
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns # For nicer plots

# For PDF Export
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# Initialize Excel file and DataFrame
EXCEL_FILE = 'trading_book.xlsx'
# Global variable for undo/redo stack
undo_stack = []
redo_stack = []
MAX_UNDO_HISTORY = 10 # Limit undo history to prevent excessive memory usage

# Global variable for decimal precision settings
decimal_precision = {
    'quantity': 8,
    'price': 2,
    'total': 2,
    'pnl': 2,
    'avg_buy_price': 2
}

# Global variables to track Toplevel windows
show_records_window = None
summary_window = None
settings_window = None # Track settings window too

def init_excel_file():
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Date', 'Ticker', 'Trade_Type', 'Quantity', 'Price', 'Total', 'Notes'])
        df.to_excel(EXCEL_FILE, index=False)

def load_data():
    try:
        df = pd.read_excel(EXCEL_FILE)
        # Ensure 'Date' is correctly parsed on load for internal consistency
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date'])
        return df
    except FileNotFoundError:
        messagebox.showerror("Error", "Excel file not found. Initializing a new one.")
        init_excel_file()
        return pd.read_excel(EXCEL_FILE)

def save_data(df, record_undo=True):
    """Saves DataFrame to Excel and manages undo/redo stack."""
    if record_undo:
        # Save current state before modification
        current_df_state = load_data()
        undo_stack.append(current_df_state.copy())
        if len(undo_stack) > MAX_UNDO_HISTORY:
            undo_stack.pop(0) # Remove oldest state if history limit exceeded
        redo_stack.clear() # Any new action clears the redo stack

    df.to_excel(EXCEL_FILE, index=False)

def undo_last_action():
    global show_records_window, summary_window # Declare global usage for modification
    if undo_stack:
        # Save current state to redo stack before undoing
        current_df_state = load_data()
        redo_stack.append(current_df_state.copy())

        # Load the previous state from undo stack
        previous_df_state = undo_stack.pop()
        previous_df_state.to_excel(EXCEL_FILE, index=False)
        messagebox.showinfo("Undo", "Last action undone.")

        # Refresh any open windows that display data
        # Check if the window object itself is not None before calling winfo_exists()
        if show_records_window is not None and show_records_window.winfo_exists():
            show_records_window.destroy() # Re-open to refresh
            show_records()
        if summary_window is not None and summary_window.winfo_exists():
            summary_window.destroy() # Re-open to refresh
            show_portfolio_summary()
    else:
        messagebox.showinfo("Undo", "No more actions to undo.")

def redo_last_undo():
    global show_records_window, summary_window # Declare global usage for modification
    if redo_stack:
        # Save current state to undo stack before redoing
        current_df_state = load_data()
        undo_stack.append(current_df_state.copy())

        # Load the next state from redo stack
        next_df_state = redo_stack.pop()
        next_df_state.to_excel(EXCEL_FILE, index=False)
        messagebox.showinfo("Redo", "Last undo redone.")

        # Refresh any open windows that display data
        # Check if the window object itself is not None before calling winfo_exists()
        if show_records_window is not None and show_records_window.winfo_exists():
            show_records_window.destroy()
            show_records()
        if summary_window is not None and summary_window.winfo_exists():
            summary_window.destroy()
            show_portfolio_summary()
    else:
        messagebox.showinfo("Redo", "No more actions to redo.")


def add_record(date, ticker, trade_type, quantity, price, notes):
    try:
        total = quantity * price
        new_record = pd.DataFrame({'Date': [date], 'Ticker': [ticker], 'Trade_Type': [trade_type],
                                   'Quantity': [quantity], 'Price': [price], 'Total': [total], 'Notes': [notes]})
        df = load_data()
        df = pd.concat([df, new_record], ignore_index=True)
        save_data(df, record_undo=True)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to add record: {e}")
        return False

def edit_record(index, date, ticker, trade_type, quantity, price, notes):
    df = load_data()
    if 0 <= index < len(df):
        try:
            # Record state BEFORE editing
            save_data(df.copy(), record_undo=True) # Manually record undo for edit

            df.at[index, 'Date'] = date
            df.at[index, 'Ticker'] = ticker
            df.at[index, 'Trade_Type'] = trade_type
            df.at[index, 'Quantity'] = quantity
            df.at[index, 'Price'] = price
            df.at[index, 'Total'] = quantity * price
            df.at[index, 'Notes'] = notes
            df.to_excel(EXCEL_FILE, index=False) # Direct save as save_data already recorded
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
            # Record state BEFORE deleting
            save_data(df.copy(), record_undo=True) # Manually record undo for delete

            df = df.drop(index).reset_index(drop=True)
            df.to_excel(EXCEL_FILE, index=False) # Direct save as save_data already recorded
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete record: {e}")
            return False
    else:
        messagebox.showerror("Error", "Invalid index for deletion.")
        return False

# --- UI Functions ---

def center_window(window): # Removed 'parent' argument
    window.update_idletasks()
    
    # Get screen width and height
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Get window width and height
    window_width = window.winfo_width()
    window_height = window.winfo_height()

    # Calculate x and y coordinates for the window to be centered on the screen
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)

    window.geometry(f'+{x}+{y}')


def validate_input(date_str, quantity_str, price_str, trade_type):
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Validation Error", "Date must be inYYYY-MM-DD format.")
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
    form_window = Toplevel(root)
    form_window.title("Edit Record" if is_edit else "Add Record")
    form_window.geometry("350x300")
    center_window(form_window) # Changed

    # Bind the close protocol for the form window
    form_window.protocol("WM_DELETE_WINDOW", lambda: on_toplevel_closing(form_window))

    labels_text = ['Date:', 'Ticker:', 'Trade Type:', 'Quantity:', 'Price:', 'Notes:']
    entries = {}

    for i, label_text in enumerate(labels_text):
        Label(form_window, text=label_text).grid(row=i, column=0, padx=5, pady=5, sticky="w")

        if label_text == 'Date:':
            entry = DateEntry(form_window, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        elif label_text == 'Trade Type:':
            entry = ttk.Combobox(form_window, values=["Buy", "Sell"], state="readonly")
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
        else:
            entry = Entry(form_window)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
            
        entries[label_text.replace(':', '').strip()] = entry

    if is_edit and current_data:
        # Check if it's already a pandas Timestamp (datetime64)
        if pd.api.types.is_datetime64_any_dtype(current_data['Date']):
            date_obj = current_data['Date'].to_pydatetime()
            entries['Date'].set_date(date_obj)
        else:
            # If it's not a datetime64 (likely a string from Excel),
            # try to parse it with the full format including time part
            try:
                date_obj = datetime.strptime(str(current_data['Date']), '%Y-%m-%d %H:%M:%S')
                entries['Date'].set_date(date_obj)
            except ValueError:
                # If even that fails (e.g., date is corrupted or different format)
                # try parsing just the date part (first 10 characters for YYYY-MM-DD)
                try:
                    date_string_only = str(current_data['Date'])[:10]
                    date_obj = datetime.strptime(date_string_only, '%Y-%m-%d')
                    entries['Date'].set_date(date_obj)
                except ValueError:
                    messagebox.showwarning("Date Error", "Could not parse existing date for Date Picker. Please verify format in Excel.")
                    entries['Date'].delete(0, END) # Clear the field if unparseable
                    entries['Date'].insert(0, str(current_data['Date'])) # Insert original string

        entries['Ticker'].insert(0, str(current_data['Ticker']))
        entries['Trade Type'].set(str(current_data['Trade_Type']))
        entries['Quantity'].insert(0, str(current_data['Quantity']))
        entries['Price'].insert(0, str(current_data['Price']))
        entries['Notes'].insert(0, str(current_data['Notes']))

    def save_action():
        date = entries['Date'].get()
        ticker = entries['Ticker'].get()
        trade_type = entries['Trade Type'].get()
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
                if update_callback: # For consistency, refresh records after add too
                    update_callback()
                form_window.destroy()

    Button(form_window, text="Save", command=save_action).grid(row=len(labels_text), column=0, padx=5, pady=10)
    Button(form_window, text="Cancel", command=form_window.destroy).grid(row=len(labels_text), column=1, padx=5, pady=10)

    form_window.grab_set()
    root.wait_window(form_window)


def show_records():
    global show_records_window
    if show_records_window and show_records_window.winfo_exists():
        show_records_window.lift()
        return

    show_records_window = Toplevel(root)
    show_records_window.title("Trading Records")
    show_records_window.geometry("1000x600") # Increased width from 900 to 1000
    center_window(show_records_window) # Changed

    # Bind the close protocol for this Toplevel window
    show_records_window.protocol("WM_DELETE_WINDOW", lambda: on_toplevel_closing(show_records_window))


    df = load_data()
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date'])
    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')

    # Search and Filter Frame
    control_frame = Frame(show_records_window)
    control_frame.pack(pady=10, fill='x')

    Label(control_frame, text="Search:").pack(side=LEFT, padx=5)
    search_entry = Entry(control_frame, width=30)
    search_entry.pack(side=LEFT, padx=5)

    Label(control_frame, text="Filter by Type:").pack(side=LEFT, padx=5)
    trade_type_filter = ttk.Combobox(control_frame, values=["All", "Buy", "Sell"], state="readonly", width=10)
    trade_type_filter.set("All")
    trade_type_filter.pack(side=LEFT, padx=5)

    Button(control_frame, text="Export CSV", command=export_records_csv).pack(side=RIGHT, padx=5)

    # Treeview for structured display
    tree_frame = Frame(show_records_window)
    tree_frame.pack(expand=True, fill='both', padx=10, pady=10)

    tree_scroll_y = Scrollbar(tree_frame, orient="vertical")
    tree_scroll_y.pack(side="right", fill="y")
    tree_scroll_x = Scrollbar(tree_frame, orient="horizontal")
    tree_scroll_x.pack(side="bottom", fill="x")

    display_columns = df.columns.tolist()    
    
    tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set,    
                         selectmode="browse", columns=display_columns)
    tree.pack(expand=True, fill='both')

    tree_scroll_y.config(command=tree.yview)
    tree_scroll_x.config(command=tree.xview)

    tree.heading("#0", text="Index", command=lambda : treeview_sort_column(tree, "#0", False))
    tree.column("#0", width=50, anchor="center")

    column_widths = {
        'Date': 100, # YYYY-MM-DD
        'Ticker': 80,
        'Trade_Type': 70,
        'Quantity': 120, # Increased for cryptocurrencies
        'Price': 100,
        'Total': 120,
        'Notes': 150 # Increased for potentially longer notes
    }

    for col in display_columns: # Iterate over the actual data columns
        tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))
        tree.column(col, width=column_widths.get(col, 100), anchor="center") # Default to 100 if not specified
    
    def populate_tree(data_frame):
        for item in tree.get_children():
            tree.delete(item)
            
        for idx, row in data_frame.iterrows():
            display_values = []
            for col_name in display_columns:
                value = row[col_name]
                if col_name == 'Quantity':
                    display_values.append(f"{value:.{decimal_precision['quantity']}f}")
                elif col_name == 'Price':
                    display_values.append(f"{value:.{decimal_precision['price']}f}")
                elif col_name == 'Total':
                    display_values.append(f"{value:.{decimal_precision['total']}f}")
                elif col_name == 'Notes' and pd.isna(value):
                    display_values.append("")
                else:
                    display_values.append(str(value))
            
            tree.insert("", "end", iid=str(idx), text=str(idx), values=display_values)

    def treeview_sort_column(tv, col, reverse):
        if col == "#0":
            l = [(int(tv.set(k, col)), k) for k in tv.get_children('')]
        else:
            l = [(tv.set(k, col), k) for k in tv.get_children('')]
            
            try:
                l.sort(key=lambda t: float(t[0].replace(',', '').strip()) if isinstance(t[0], str) and t[0].replace('.', '', 1).replace(',', '').strip().replace('-', '').isdigit() else t[0], reverse=reverse)
            except ValueError:
                l.sort(reverse=reverse)

        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

    def apply_filters_and_search():
        current_df = load_data()
        current_df['Date'] = pd.to_datetime(current_df['Date'], errors='coerce')
        current_df = current_df.dropna(subset=['Date'])
        current_df['Date'] = current_df['Date'].dt.strftime('%Y-%m-%d')

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

    populate_tree(df)

    # Edit and Delete Buttons
    action_frame = Frame(show_records_window)
    action_frame.pack(pady=10)

    def edit_selected_record():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a record to edit.")
            return
        
        selected_index = int(tree.item(selected_item[0], "text"))

        df_current = load_data()
        current_record_data = df_current.iloc[selected_index].to_dict()
        
        add_edit_form(is_edit=True, record_index=selected_index, current_data=current_record_data, update_callback=apply_filters_and_search)

    def delete_selected_record():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a record to delete.")
            return

        selected_index = int(tree.item(selected_item[0], "text"))
        
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete record at index {selected_index}?"):
            if delete_record(selected_index):
                messagebox.showinfo("Success", "Record deleted successfully.")
                apply_filters_and_search()

    Button(action_frame, text="Edit Selected", command=edit_selected_record).pack(side=LEFT, padx=5)
    Button(action_frame, text="Delete Selected", command=delete_selected_record).pack(side=LEFT, padx=5)


def export_records_csv():
    df = load_data()
    if df.empty:
        messagebox.showinfo("Export", "No records to export.")
        return
    
    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                                             title="Export Records as CSV")
    if file_path:
        try:
            # Ensure date is string format for CSV export consistency
            df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
            df.to_csv(file_path, index=False)
            messagebox.showinfo("Export Success", "Records exported to CSV successfully!")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export records: {e}")

# --- Analytical Functions ---

def calculate_realized_pnl(df):
    pnl = {}
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df_cleaned = df.dropna(subset=['Date']).copy() # Use .copy() to avoid SettingWithCopyWarning
    
    df_cleaned = df_cleaned.sort_values(by='Date').reset_index(drop=True)

    for ticker in df_cleaned['Ticker'].unique():
        ticker_trades = df_cleaned[df_cleaned['Ticker'] == ticker].copy()
        
        ticker_trades['Quantity'] = pd.to_numeric(ticker_trades['Quantity'], errors='coerce')
        ticker_trades['Price'] = pd.to_numeric(ticker_trades['Price'], errors='coerce')
        ticker_trades = ticker_trades.dropna(subset=['Quantity', 'Price'])

        buys = ticker_trades[ticker_trades['Trade_Type'].str.lower() == 'buy']
        sells = ticker_trades[ticker_trades['Trade_Type'].str.lower() == 'sell']
        
        realized_pnl = 0
        
        buy_queue = [] # Stores (quantity, price) for buys (FIFO)

        for index, row in buys.iterrows():
            buy_queue.append({'quantity': row['Quantity'], 'price': row['Price']})
        
        for index, row in sells.iterrows():
            sell_quantity = row['Quantity']
            sell_price = row['Price']
            
            while sell_quantity > 0 and buy_queue:
                buy_data = buy_queue[0]
                buy_quantity = buy_data['quantity']
                buy_price = buy_data['price']
                
                if sell_quantity >= buy_quantity:
                    realized_pnl += (sell_price - buy_price) * buy_quantity
                    sell_quantity -= buy_quantity
                    buy_queue.pop(0)
                else:
                    realized_pnl += (sell_price - buy_price) * sell_quantity
                    buy_data['quantity'] -= sell_quantity
                    sell_quantity = 0

        pnl[ticker] = realized_pnl
    return pnl

def calculate_cumulative_pnl_per_ticker(df):
    """Calculates cumulative P&L for each ticker over time."""
    df_cleaned = df.dropna(subset=['Date', 'Quantity', 'Price']).copy()
    df_cleaned['Date'] = pd.to_datetime(df_cleaned['Date'])
    df_cleaned = df_cleaned.sort_values(by='Date')

    cumulative_pnl_data = {}

    for ticker in df_cleaned['Ticker'].unique():
        ticker_trades = df_cleaned[df_cleaned['Ticker'] == ticker].copy()
        
        # Initialize holdings for FIFO calculation for this ticker
        buy_queue = [] # Stores (quantity, price) for buys
        daily_pnl = {}
        current_cumulative_pnl = 0

        for index, row in ticker_trades.iterrows():
            trade_date = row['Date'].strftime('%Y-%m-%d')
            quantity = row['Quantity']
            price = row['Price']
            trade_type = row['Trade_Type'].lower()

            if trade_type == 'buy':
                buy_queue.append({'quantity': quantity, 'price': price})
            elif trade_type == 'sell':
                sell_quantity = quantity
                sell_price = price
                
                while sell_quantity > 0 and buy_queue:
                    buy_data = buy_queue[0]
                    buy_quantity_in_lot = buy_data['quantity']
                    buy_price_in_lot = buy_data['price']
                    
                    if sell_quantity >= buy_quantity_in_lot:
                        pnl_from_lot = (sell_price - buy_price_in_lot) * buy_quantity_in_lot
                        current_cumulative_pnl += pnl_from_lot
                        sell_quantity -= buy_quantity_in_lot
                        buy_queue.pop(0)
                    else:
                        pnl_from_lot = (sell_price - buy_price_in_lot) * sell_quantity
                        current_cumulative_pnl += pnl_from_lot
                        buy_data['quantity'] -= sell_quantity
                        sell_quantity = 0
            
            # Record cumulative P&L for this date
            daily_pnl[trade_date] = current_cumulative_pnl
        
        # Convert daily_pnl to a pandas Series, then reindex to fill missing dates
        # This ensures a continuous time series for plotting
        if daily_pnl:
            pnl_series = pd.Series(daily_pnl)
            # Create a full date range from the first trade to the last
            idx = pd.date_range(start=pnl_series.index.min(), end=pnl_series.index.max())
            pnl_series.index = pd.to_datetime(pnl_series.index)
            pnl_series = pnl_series.reindex(idx, method='ffill') # Forward-fill missing dates
            pnl_series = pnl_series.fillna(0) # Fill initial NaNs if no trades on first day
            cumulative_pnl_data[ticker] = pnl_series
    
    return cumulative_pnl_data


def get_current_holdings(df):
    holdings = {}
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df_cleaned = df.dropna(subset=['Date']).copy() # Use .copy() to avoid SettingWithCopyWarning
    
    df_cleaned = df_cleaned.sort_values(by='Date').reset_index(drop=True)

    for ticker in df_cleaned['Ticker'].unique():
        ticker_trades = df_cleaned[df_cleaned['Ticker'] == ticker].copy()
        
        ticker_trades['Quantity'] = pd.to_numeric(ticker_trades['Quantity'], errors='coerce')
        ticker_trades['Price'] = pd.to_numeric(ticker_trades['Price'], errors='coerce')
        ticker_trades = ticker_trades.dropna(subset=['Quantity', 'Price'])

        net_quantity = 0
        buy_queue = []    

        for index, row in ticker_trades.iterrows():
            if row['Trade_Type'].lower() == 'buy':
                net_quantity += row['Quantity']
                buy_queue.append({'quantity': row['Quantity'], 'price': row['Price']})
            elif row['Trade_Type'].lower() == 'sell':
                sell_quantity = row['Quantity']
                while sell_quantity > 0 and buy_queue:
                    buy_data = buy_queue[0]
                    buy_quantity_in_lot = buy_data['quantity']
                    
                    if sell_quantity >= buy_quantity_in_lot:
                        sell_quantity -= buy_quantity_in_lot
                        net_quantity -= buy_quantity_in_lot
                        buy_queue.pop(0)
                    else:
                        buy_data['quantity'] -= sell_quantity
                        net_quantity -= sell_quantity
                        sell_quantity = 0

        if net_quantity > 0:
            remaining_value = sum(item['quantity'] * item['price'] for item in buy_queue)
            average_price = remaining_value / net_quantity if net_quantity != 0 else 0
            holdings[ticker] = {'quantity': net_quantity, 'average_buy_price': average_price}
        elif net_quantity == 0:
            holdings[ticker] = {'quantity': 0, 'average_buy_price': 0}
            
    return {k: v for k, v in holdings.items() if v['quantity'] > 0}

def calculate_performance_metrics(df):
    total_buy_value = df[df['Trade_Type'].str.lower() == 'buy']['Total'].sum()
    total_sell_value = df[df['Trade_Type'].str.lower() == 'sell']['Total'].sum()
    
    realized_pnl = calculate_realized_pnl(df.copy())
    total_realized_pnl = sum(realized_pnl.values())

    # Total Return / ROI
    # Simplified ROI: (Total Sell Value - Total Buy Value) / Total Buy Value
    # More accurately, you'd track cash inflow/outflow. This is a basic version.
    if total_buy_value > 0:
        total_roi = (total_sell_value - total_buy_value) / total_buy_value * 100
    else:
        total_roi = 0.0

    # Win/Loss Rate
    # Assuming each 'buy' followed by 'sell' for the same ticker is a "trade"
    # This simplified calculation counts each ticker's realized P&L as one outcome
    win_trades = sum(1 for pnl in realized_pnl.values() if pnl > 0)
    loss_trades = sum(1 for pnl in realized_pnl.values() if pnl < 0)
    total_closed_trades = win_trades + loss_trades
    
    win_rate = (win_trades / total_closed_trades * 100) if total_closed_trades > 0 else 0.0

    # Average Profit/Loss per Trade (only for closed trades)
    avg_profit_per_trade = (sum(p for p in realized_pnl.values() if p > 0) / win_trades) if win_trades > 0 else 0.0
    avg_loss_per_trade = (sum(abs(p) for p in realized_pnl.values() if p < 0) / loss_trades) if loss_trades > 0 else 0.0

    return {
        'total_realized_pnl': total_realized_pnl,
        'total_roi': total_roi,
        'win_rate': win_rate,
        'avg_profit_per_trade': avg_profit_per_trade,
        'avg_loss_per_trade': avg_loss_per_trade
    }

def show_portfolio_summary():
    global summary_window
    if summary_window and summary_window.winfo_exists():
        summary_window.lift()
        return

    summary_window = Toplevel(root)
    summary_window.title("Portfolio Summary")
    summary_window.geometry("700x750") # Make window taller for more content
    center_window(summary_window) # Changed

    # Bind the close protocol for this Toplevel window
    summary_window.protocol("WM_DELETE_WINDOW", lambda: on_toplevel_closing(summary_window))


    df = load_data()
    
    # --- Performance Metrics ---
    metrics = calculate_performance_metrics(df.copy())
    metrics_frame = LabelFrame(summary_window, text="Performance Metrics", padx=10, pady=10)
    metrics_frame.pack(pady=10, padx=10, fill='x')

    Label(metrics_frame, text="Total Realized P&L:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
    pnl_color = "green" if metrics['total_realized_pnl'] >= 0 else "red"
    Label(metrics_frame, text=f"{metrics['total_realized_pnl']:.{decimal_precision['pnl']}f}", fg=pnl_color).grid(row=0, column=1, sticky="e", padx=5, pady=2)

    Label(metrics_frame, text="Total ROI:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
    roi_color = "green" if metrics['total_roi'] >= 0 else "red"
    Label(metrics_frame, text=f"{metrics['total_roi']:.2f}%", fg=roi_color).grid(row=1, column=1, sticky="e", padx=5, pady=2)

    Label(metrics_frame, text="Win Rate:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
    Label(metrics_frame, text=f"{metrics['win_rate']:.2f}%").grid(row=2, column=1, sticky="e", padx=5, pady=2)
    
    Label(metrics_frame, text="Avg. Profit per Win:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
    Label(metrics_frame, text=f"{metrics['avg_profit_per_trade']:.{decimal_precision['pnl']}f}").grid(row=3, column=1, sticky="e", padx=5, pady=2)

    Label(metrics_frame, text="Avg. Loss per Loss:").grid(row=4, column=0, sticky="w", padx=5, pady=2)
    Label(metrics_frame, text=f"{metrics['avg_loss_per_trade']:.{decimal_precision['pnl']}f}").grid(row=4, column=1, sticky="e", padx=5, pady=2)


    # --- Current Holdings ---
    current_holdings = get_current_holdings(df.copy())
    
    holdings_frame = LabelFrame(summary_window, text="Current Holdings", padx=10, pady=10)
    holdings_frame.pack(pady=10, padx=10, fill='x')

    if current_holdings:
        ttk.Label(holdings_frame, text="Ticker", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=2)
        ttk.Label(holdings_frame, text="Quantity", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(holdings_frame, text="Avg. Buy Price", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=2)

        for i, (ticker, data) in enumerate(current_holdings.items()):
            ttk.Label(holdings_frame, text=ticker).grid(row=i+1, column=0, padx=5, pady=2, sticky="w")
            ttk.Label(holdings_frame, text=f"{data['quantity']:.{decimal_precision['quantity']}f}").grid(row=i+1, column=1, padx=5, pady=2)
            ttk.Label(holdings_frame, text=f"{data['average_buy_price']:.{decimal_precision['avg_buy_price']}f}").grid(row=i+1, column=2, padx=5, pady=2)
    else:
        Label(holdings_frame, text="No current holdings to display.").pack()

    # --- Charts ---
    chart_notebook = ttk.Notebook(summary_window)
    chart_notebook.pack(pady=10, padx=10, fill='both', expand=True)

    # Pie Chart for Allocation
    pie_chart_frame = Frame(chart_notebook)
    chart_notebook.add(pie_chart_frame, text="Portfolio Allocation")
    
    if current_holdings:
        labels = [ticker for ticker in current_holdings.keys()]
        sizes = [data['quantity'] * data['average_buy_price'] for data in current_holdings.values()] # Value-based allocation

        if sum(sizes) > 0:
            fig, ax = plt.subplots(figsize=(5, 4))
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 8})
            ax.axis('equal')
            ax.set_title("Portfolio Allocation by Value", fontsize=10)
            fig.tight_layout() # Adjust layout to prevent labels overlapping

            canvas = FigureCanvasTkAgg(fig, master=pie_chart_frame)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill='both', expand=True)
            canvas.draw()
        else:
            Label(pie_chart_frame, text="No positive portfolio value to display chart.").pack(expand=True)
    else:
        Label(pie_chart_frame, text="No holdings to generate allocation chart.").pack(expand=True)

    # P&L Over Time Chart (Line Chart - Total)
    pnl_over_time_frame = Frame(chart_notebook)
    chart_notebook.add(pnl_over_time_frame, text="Total P&L Over Time")

    if not df.empty:
        daily_trades = df.copy()
        daily_trades['Date'] = pd.to_datetime(daily_trades['Date']) # Ensure datetime type for grouping
        daily_trades['Trade_Value'] = daily_trades.apply(lambda row: row['Total'] if row['Trade_Type'].lower() == 'sell' else -row['Total'], axis=1)
        
        # Calculate overall cumulative P&L for plotting
        overall_daily_pnl_df = daily_trades.groupby('Date')['Trade_Value'].sum().to_frame()
        overall_daily_pnl_df['Cumulative_P&L'] = overall_daily_pnl_df['Trade_Value'].cumsum()

        # Reindex to ensure continuous dates for smoother plotting
        if not overall_daily_pnl_df.empty:
            idx = pd.date_range(start=overall_daily_pnl_df.index.min(), end=overall_daily_pnl_df.index.max())
            overall_daily_pnl_df = overall_daily_pnl_df.reindex(idx, method='ffill').fillna(0) # Forward-fill and fill initial NaNs with 0

            fig_pnl, ax_pnl = plt.subplots(figsize=(5, 4))
            sns.lineplot(x=overall_daily_pnl_df.index, y=overall_daily_pnl_df['Cumulative_P&L'], ax=ax_pnl)
            ax_pnl.set_title("Total Cumulative P&L Over Time", fontsize=10)
            ax_pnl.set_xlabel("Date", fontsize=8)
            ax_pnl.set_ylabel("Cumulative P&L", fontsize=8)
            ax_pnl.tick_params(axis='x', rotation=45, labelsize=7)
            ax_pnl.tick_params(axis='y', labelsize=7)
            ax_pnl.grid(True)
            fig_pnl.tight_layout()

            canvas_pnl = FigureCanvasTkAgg(fig_pnl, master=pnl_over_time_frame)
            canvas_pnl_widget = canvas_pnl.get_tk_widget()
            canvas_pnl_widget.pack(fill='both', expand=True)
            canvas_pnl.draw()
        else:
            Label(pnl_over_time_frame, text="Not enough data to generate Total P&L over time chart.").pack(expand=True)
    else:
        Label(pnl_over_time_frame, text="No data to generate Total P&L over time chart.").pack(expand=True)


    # Trade Volume Over Time (Bar Chart)
    volume_over_time_frame = Frame(chart_notebook)
    chart_notebook.add(volume_over_time_frame, text="Trade Volume")

    if not df.empty:
        df_volume = df.copy()
        df_volume['Date'] = pd.to_datetime(df_volume['Date'])
        # Aggregate quantity by date
        daily_volume = df_volume.groupby('Date')['Quantity'].sum()

        if not daily_volume.empty:
            fig_vol, ax_vol = plt.subplots(figsize=(5, 4))
            # Corrected line to address FutureWarning
            sns.barplot(x=daily_volume.index, y=daily_volume.values, ax=ax_vol, hue=daily_volume.index, palette="viridis", legend=False)
            ax_vol.set_title("Trade Volume Over Time", fontsize=10)
            ax_vol.set_xlabel("Date", fontsize=8)
            ax_vol.set_ylabel("Total Quantity Traded", fontsize=8)
            ax_vol.tick_params(axis='x', rotation=45, labelsize=7)
            ax_vol.tick_params(axis='y', labelsize=7)
            ax_vol.grid(axis='y', linestyle='--')
            fig_vol.tight_layout()

            canvas_vol = FigureCanvasTkAgg(fig_vol, master=volume_over_time_frame)
            canvas_vol_widget = canvas_vol.get_tk_widget()
            canvas_vol_widget.pack(fill='both', expand=True)
            canvas_vol.draw()
        else:
            Label(volume_over_time_frame, text="Not enough data to generate Trade Volume chart.").pack(expand=True)
    else:
        Label(volume_over_time_frame, text="No data to generate Trade Volume chart.").pack(expand=True)

    # Ticker Specific Cumulative P&L Chart (Line Chart) - UPDATED
    ticker_cumulative_pnl_frame = Frame(chart_notebook)
    chart_notebook.add(ticker_cumulative_pnl_frame, text="Ticker Cumulative P&L")

    if not df.empty:
        cumulative_pnl_per_ticker = calculate_cumulative_pnl_per_ticker(df.copy())
        
        if cumulative_pnl_per_ticker:
            fig_ticker_cum_pnl, ax_ticker_cum_pnl = plt.subplots(figsize=(5, 4))
            
            # Use distinct markers/linestyles if many tickers for better differentiation
            for ticker, pnl_series in cumulative_pnl_per_ticker.items():
                sns.lineplot(x=pnl_series.index, y=pnl_series.values, ax=ax_ticker_cum_pnl, label=ticker, marker='o', markersize=3) # Added marker
            
            ax_ticker_cum_pnl.set_title("Cumulative P&L per Ticker Over Time", fontsize=10)
            ax_ticker_cum_pnl.set_xlabel("Date", fontsize=8)
            ax_ticker_cum_pnl.set_ylabel("Cumulative P&L", fontsize=8)
            ax_ticker_cum_pnl.tick_params(axis='x', rotation=45, labelsize=7)
            ax_ticker_cum_pnl.tick_params(axis='y', labelsize=7)
            ax_ticker_cum_pnl.grid(True)
            ax_ticker_cum_pnl.legend(fontsize=7, loc='upper left')
            fig_ticker_cum_pnl.tight_layout()

            canvas_ticker_cum_pnl = FigureCanvasTkAgg(fig_ticker_cum_pnl, master=ticker_cumulative_pnl_frame)
            canvas_ticker_cum_pnl_widget = canvas_ticker_cum_pnl.get_tk_widget()
            canvas_ticker_cum_pnl_widget.pack(fill='both', expand=True)
            canvas_ticker_cum_pnl.draw()
        else:
            Label(ticker_cumulative_pnl_frame, text="No cumulative P&L to display for tickers.").pack(expand=True)
    else:
        Label(ticker_cumulative_pnl_frame, text="No data to generate Ticker Cumulative P&L chart.").pack(expand=True)


def open_settings_window():
    global settings_window # Declare global to assign to it
    if settings_window and settings_window.winfo_exists():
        settings_window.lift()
        return

    settings_window = Toplevel(root)
    settings_window.title("Settings")
    settings_window.geometry("300x250")
    center_window(settings_window) # Changed
    settings_window.protocol("WM_DELETE_WINDOW", lambda: on_toplevel_closing(settings_window))


    Label(settings_window, text="Set Decimal Precision:").pack(pady=10)

    precision_labels = ['Quantity:', 'Price:', 'Total:', 'P&L:', 'Avg. Buy Price:']
    precision_keys = ['quantity', 'price', 'total', 'pnl', 'avg_buy_price']
    spinboxes = {}

    for i, label_text in enumerate(precision_labels):
        frame = Frame(settings_window)
        frame.pack(fill='x', padx=10, pady=2)
        Label(frame, text=label_text, width=15, anchor='w').pack(side='left')
        
        key = precision_keys[i]
        spinbox = Spinbox(frame, from_=0, to_=10, width=5, justify='center')
        spinbox.set(decimal_precision[key]) # Set current value
        spinbox.pack(side='right')
        spinboxes[key] = spinbox
    
    def save_settings():
        for key, spinbox in spinboxes.items():
            try:
                decimal_precision[key] = int(spinbox.get())
            except ValueError:
                messagebox.showerror("Input Error", f"Invalid input for {key} precision. Must be an integer.")
                return
        messagebox.showinfo("Settings Saved", "Decimal precision settings updated.")
        settings_window.destroy()

    Button(settings_window, text="Save Settings", command=save_settings).pack(pady=10)

# --- Exit Handling Functions ---
def on_toplevel_closing(toplevel_window):
    """Handles the closing of Toplevel windows."""
    if messagebox.askokcancel("Close Window", "Are you sure you want to close this window?"):
        # If it's the records or summary window, clear the global reference
        global show_records_window, summary_window, settings_window
        if toplevel_window == show_records_window:
            show_records_window = None
        elif toplevel_window == summary_window:
            summary_window = None
        elif toplevel_window == settings_window: # Added for settings window
            settings_window = None
        toplevel_window.destroy()

def on_closing():
    """Handles the closing of the main application window."""
    if messagebox.askokcancel("Quit", "Do you want to quit the application?"):
        root.quit() # This will terminate the entire Tkinter application


# --- Main Application Window ---
init_excel_file()

root = Tk()
root.title("Trading Book Manager")
root.geometry("400x300")
root.resizable(False, False) # Disable resizing for a fixed layout

# Bind the close protocol for the main window
root.protocol("WM_DELETE_WINDOW", on_closing)


# Main buttons
Button(root, text="Add Record", command=lambda: add_edit_form(update_callback=show_records)).pack(pady=10, fill='x', padx=50)
Button(root, text="Show Records", command=show_records).pack(pady=10, fill='x', padx=50)
Button(root, text="Show Portfolio Summary", command=show_portfolio_summary).pack(pady=10, fill='x', padx=50)
Button(root, text="Undo Last Action", command=undo_last_action).pack(pady=10, fill='x', padx=50)
Button(root, text="Redo Last Undo", command=redo_last_undo).pack(pady=10, fill='x', padx=50)
Button(root, text="Settings", command=open_settings_window).pack(pady=10, fill='x', padx=50)

# Add an Exit button
Button(root, text="Exit", command=on_closing, bg="red", fg="white").pack(pady=20, fill='x', padx=50)


root.mainloop()