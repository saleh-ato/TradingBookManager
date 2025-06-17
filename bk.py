import pandas as pd
from tkinter import *
from tkinter import messagebox, simpledialog, ttk
from datetime import datetime
from tkcalendar import DateEntry
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

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
        total = quantity * price
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
            df.at[index, 'Total'] = quantity * price
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

def center_window(window, parent):
    window.update_idletasks()
    parent_x = parent.winfo_x()
    parent_y = parent.winfo_y()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()

    window_width = window.winfo_width()
    window_height = window.winfo_height()

    x = parent_x + (parent_width // 2) - (window_width // 2)
    y = parent_y + (parent_height // 2) - (window_height // 2)

    window.geometry(f'+{x}+{y}')


def validate_input(date_str, quantity_str, price_str, trade_type):
    try:
        datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Validation Error", "Date must be in McClellan-MM-DD format.")
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
    center_window(form_window, root)

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
        # Ensure date is a valid datetime object for DateEntry
        try:
            date_obj = datetime.strptime(current_data['Date'], '%Y-%m-%d')
            entries['Date'].set_date(date_obj)
        except ValueError:
            messagebox.showwarning("Date Error", "Could not parse existing date for Date Picker. Please verify format.")
            entries['Date'].delete(0, END) # Clear the field if unparseable
            entries['Date'].insert(0, current_data['Date']) # Insert original string
            
        entries['Ticker'].insert(0, current_data['Ticker'])
        entries['Trade Type'].set(current_data['Trade_Type'])
        entries['Quantity'].insert(0, str(current_data['Quantity']))
        entries['Price'].insert(0, str(current_data['Price']))
        entries['Notes'].insert(0, current_data['Notes'])

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
                form_window.destroy()

    Button(form_window, text="Save", command=save_action).grid(row=len(labels_text), column=0, padx=5, pady=10)
    Button(form_window, text="Cancel", command=form_window.destroy).grid(row=len(labels_text), column=1, padx=5, pady=10)

    form_window.grab_set()
    root.wait_window(form_window)


def show_records():
    records_window = Toplevel(root)
    records_window.title("Trading Records")
    records_window.geometry("900x600")
    center_window(records_window, root)


    df = load_data()
    # Ensure 'Date' column is string for consistent display and search
    # Use errors='coerce' to turn unparseable dates into NaT
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.dropna(subset=['Date']) # Drop rows where Date conversion failed
    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d') # Convert to string for display


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
    tree["show"] = "headings"

    for col in df.columns:
        tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(tree, _col, False))
        tree.column(col, width=120, anchor="center")
    
    tree["columns"] = ("#0",) + tuple(df.columns)
    tree.heading("#0", text="Index", command=lambda : treeview_sort_column(tree, "#0", False))
    tree.column("#0", width=50, anchor="center")

    def populate_tree(data_frame):
        for item in tree.get_children():
            tree.delete(item)
        
        for idx, row in data_frame.iterrows():
            # Format Quantity and Price for display in Treeview
            display_values = row.tolist()
            if 'Quantity' in data_frame.columns:
                q_idx = data_frame.columns.get_loc('Quantity')
                display_values[q_idx] = f"{row['Quantity']:.8f}" # Format Quantity to 8 decimal places
            if 'Price' in data_frame.columns:
                p_idx = data_frame.columns.get_loc('Price')
                display_values[p_idx] = f"{row['Price']:.2f}" # Keep Price to 2 decimal places (or adjust as needed)
            if 'Total' in data_frame.columns:
                t_idx = data_frame.columns.get_loc('Total')
                display_values[t_idx] = f"{row['Total']:.2f}" # Keep Total to 2 decimal places

            tree.insert("", "end", iid=str(idx), text=str(idx), values=display_values)

    def treeview_sort_column(tv, col, reverse):
        if col == "#0":
            l = [(int(tv.set(k, col)), k) for k in tv.get_children('')]
        else:
            # Handle potential non-string values during sorting for data columns
            try:
                l = [(tv.set(k, col), k) for k in tv.get_children('')]
                # Attempt to convert to numeric if possible for better sorting
                # Adjust for potentially formatted strings (remove commas) before converting to float
                l.sort(key=lambda t: float(t[0].replace(',', '')) if isinstance(t[0], str) and t[0].replace('.', '', 1).replace(',', '').isdigit() else t[0], reverse=reverse)
            except ValueError:
                # Fallback to string comparison if not numeric
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
    action_frame = Frame(records_window)
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

# --- Analytical Functions ---

def calculate_realized_pnl(df):
    pnl = {}
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df_cleaned = df.dropna(subset=['Date'])
    
    df_cleaned = df_cleaned.sort_values(by='Date').reset_index(drop=True)

    for ticker in df_cleaned['Ticker'].unique():
        ticker_trades = df_cleaned[df_cleaned['Ticker'] == ticker].copy()
        
        ticker_trades['Quantity'] = pd.to_numeric(ticker_trades['Quantity'], errors='coerce')
        ticker_trades['Price'] = pd.to_numeric(ticker_trades['Price'], errors='coerce')
        ticker_trades = ticker_trades.dropna(subset=['Quantity', 'Price'])

        buys = ticker_trades[ticker_trades['Trade_Type'].str.lower() == 'buy']
        sells = ticker_trades[ticker_trades['Trade_Type'].str.lower() == 'sell']
        
        realized_pnl = 0
        
        buy_queue = [] # Stores (quantity, price) for buys

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

def get_current_holdings(df):
    holdings = {}
    
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df_cleaned = df.dropna(subset=['Date'])
    
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
            # Handle potential division by zero if net_quantity somehow becomes 0 here
            average_price = remaining_value / net_quantity if net_quantity != 0 else 0
            holdings[ticker] = {'quantity': net_quantity, 'average_buy_price': average_price}
        elif net_quantity == 0:
            holdings[ticker] = {'quantity': 0, 'average_buy_price': 0}
            
    return {k: v for k, v in holdings.items() if v['quantity'] > 0}


def show_portfolio_summary():
    summary_window = Toplevel(root)
    summary_window.title("Portfolio Summary")
    summary_window.geometry("700x500")
    center_window(summary_window, root)

    df = load_data()
    
    # Realized P&L
    realized_pnl = calculate_realized_pnl(df.copy()) # Pass a copy to avoid modifying the original DataFrame in place
    
    pnl_frame = LabelFrame(summary_window, text="Realized Profit/Loss per Ticker", padx=10, pady=10)
    pnl_frame.pack(pady=10, padx=10, fill='x')

    if realized_pnl:
        total_realized_pnl = sum(realized_pnl.values())
        Label(pnl_frame, text="--- Total Realized P&L ---", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=2)
        total_color = "green" if total_realized_pnl >= 0 else "red"
        Label(pnl_frame, text=f"{total_realized_pnl:,.2f}", fg=total_color, font=("Arial", 10, "bold")).grid(row=0, column=1, sticky="e", padx=5, pady=2)

        Label(pnl_frame, text="Ticker", font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="w", padx=5, pady=2)
        Label(pnl_frame, text="P&L", font=("Arial", 9, "bold")).grid(row=1, column=1, sticky="e", padx=5, pady=2)

        row_num = 2
        for ticker, pnl_value in realized_pnl.items():
            color = "green" if pnl_value >= 0 else "red"
            Label(pnl_frame, text=ticker).grid(row=row_num, column=0, sticky="w", padx=5, pady=2)
            Label(pnl_frame, text=f"{pnl_value:,.2f}", fg=color).grid(row=row_num, column=1, sticky="e", padx=5, pady=2)
            row_num += 1
    else:
        Label(pnl_frame, text="No realized P&L to display yet.").pack()

    # Current Holdings
    current_holdings = get_current_holdings(df.copy()) # Pass a copy
    
    holdings_frame = LabelFrame(summary_window, text="Current Holdings", padx=10, pady=10)
    holdings_frame.pack(pady=10, padx=10, fill='x')

    if current_holdings:
        ttk.Label(holdings_frame, text="Ticker", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=2)
        ttk.Label(holdings_frame, text="Quantity", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=2)
        ttk.Label(holdings_frame, text="Avg. Buy Price", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=2)

        for i, (ticker, data) in enumerate(current_holdings.items()):
            ttk.Label(holdings_frame, text=ticker).grid(row=i+1, column=0, padx=5, pady=2, sticky="w")
            # --- MODIFIED LINE HERE ---
            # Format Quantity to 8 decimal places for crypto
            ttk.Label(holdings_frame, text=f"{data['quantity']:.8f}").grid(row=i+1, column=1, padx=5, pady=2)
            # --- END OF MODIFIED LINE ---
            ttk.Label(holdings_frame, text=f"{data['average_buy_price']:,.2f}").grid(row=i+1, column=2, padx=5, pady=2)
    else:
        Label(holdings_frame, text="No current holdings to display.").pack()

    # Asset Allocation Pie Chart
    chart_frame = Frame(summary_window)
    chart_frame.pack(pady=10, padx=10, fill='both', expand=True)

    if current_holdings:
        labels = [ticker for ticker in current_holdings.keys()]
        sizes = [data['quantity'] * data['average_buy_price'] for data in current_holdings.values()] # Value-based allocation

        if sum(sizes) > 0: # Only plot if there's actual positive value to allocate
            fig, ax = plt.subplots(figsize=(5, 4))
            ax.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
            ax.set_title("Portfolio Allocation by Value")

            canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill='both', expand=True)
            canvas.draw()
        else:
            Label(chart_frame, text="No positive portfolio value to display chart.").pack()
    else:
        Label(chart_frame, text="No holdings to generate chart.").pack()


def gui_main():
    global root
    root = Tk()
    root.title('Trading Bookkeeping System')
    root.geometry("400x250")
    root.resizable(False, False)

    root.update_idletasks()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = root.winfo_width()
    window_height = root.winfo_height()

    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f'+{x}+{y}')

    Label(root, text="Trading Bookkeeping", font=("Arial", 16, "bold")).pack(pady=20)

    Button(root, text='Add New Record', command=lambda: add_edit_form(update_callback=show_records)).pack(pady=5, fill='x', padx=50)
    Button(root, text='Show All Records', command=show_records).pack(pady=5, fill='x', padx=50)
    Button(root, text='Portfolio Summary', command=show_portfolio_summary).pack(pady=5, fill='x', padx=50)


    root.mainloop()

if __name__ == '__main__':
    init_excel_file()
    gui_main()