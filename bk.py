import pandas as pd
from tkinter import *
from tkinter import messagebox, simpledialog, ttk

# Initialize Excel file and DataFrame
EXCEL_FILE = 'trading_book.xlsx'

def init_excel_file():
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Date', 'Ticker', 'Trade_Type', 'Quantity', 'Price', 'Total', 'Notes'])
        df.to_excel(EXCEL_FILE, index=False)

def load_data():
    return pd.read_excel(EXCEL_FILE)

def add_record(date, ticker, trade_type, quantity, price, notes):
    total = quantity * price  # Calculate total trade value
    new_record = pd.DataFrame({'Date': [date], 'Ticker': [ticker], 'Trade_Type': [trade_type],
                               'Quantity': [quantity], 'Price': [price], 'Total': [total], 'Notes': [notes]})
    df = load_data()
    df = pd.concat([df, new_record], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

def edit_record(index, date, ticker, trade_type, quantity, price, notes):
    df = load_data()
    if 0 <= index < len(df):
        df.at[index, 'Date'] = date
        df.at[index, 'Ticker'] = ticker
        df.at[index, 'Trade_Type'] = trade_type
        df.at[index, 'Quantity'] = quantity
        df.at[index, 'Price'] = price
        df.at[index, 'Total'] = quantity * price  # Update total trade value
        df.at[index, 'Notes'] = notes
        df.to_excel(EXCEL_FILE, index=False)

def delete_record(index):
    df = load_data()
    if 0 <= index < len(df):
        df = df.drop(index).reset_index(drop=True)
        df.to_excel(EXCEL_FILE, index=False)

def show_records():
    df = load_data()

    records_window = Toplevel()
    records_window.title("Trading Records")

    # Sorting dropdown
    sort_column = StringVar(records_window)
    sort_column.set(df.columns[0])  # Default to first column
    combobox = ttk.Combobox(records_window, textvariable=sort_column, values=list(df.columns), state="readonly")

    sort_order = BooleanVar(records_window)  # False = Ascending, True = Descending

    def update_table():
        sorted_df = df.sort_values(by=sort_column.get(), ascending=not sort_order.get())
        tree.delete(*tree.get_children())  # Clear old records

        for _, row in sorted_df.iterrows():
            tree.insert("", "end", values=row.tolist())

    Label(records_window, text="Sort by:").pack()
    combobox.pack()
    #OptionMenu(records_window, sort_column, *df.columns).pack()
    Checkbutton(records_window, text="Descending", variable=sort_order).pack()
    Button(records_window, text="Apply Sort", command=update_table).pack()


    # Treeview for structured display
    tree = ttk.Treeview(records_window)
    tree["columns"] = list(df.columns)

    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120)

    for _, row in df.iterrows():
        tree.insert("", "end", values=row.tolist())

    tree.pack(expand=True, fill='both')



def gui_main():
    root = Tk()
    root.title('Trading Bookkeeping System')

    Button(root, text='Add Record', command=add_record_form).pack()
    Button(root, text='Delete Record', command=delete_record_form).pack()
    Button(root, text='Edit Record', command=edit_record_form).pack()
    Button(root, text='Show Records', command=show_records).pack()

    root.mainloop()

def add_record_form():
    date = simpledialog.askstring("Input", "Enter date (YYYY-MM-DD):")
    ticker = simpledialog.askstring("Input", "Enter Ticker:")
    trade_type = simpledialog.askstring("Input", "Enter Trade Type (Buy/Sell):")
    quantity = simpledialog.askfloat("Input", "Enter Quantity:")
    price = simpledialog.askfloat("Input", "Enter Price:")
    notes = simpledialog.askstring("Input", "Enter Notes:")
    add_record(date, ticker, trade_type, quantity, price, notes)
    messagebox.showinfo("Success", "Record added successfully.")

def delete_record_form():
    index = simpledialog.askinteger("Input", "Enter index to delete:")
    delete_record(index)
    messagebox.showinfo("Success", "Record deleted successfully.")

def edit_record_form():
    index = simpledialog.askinteger("Input", "Enter index to edit:")
    if index is not None:
        date = simpledialog.askstring("Input", "Enter new date (YYYY-MM-DD):")
        ticker = simpledialog.askstring("Input", "Enter new Ticker:")
        trade_type = simpledialog.askstring("Input", "Enter new Trade Type (Buy/Sell):")
        quantity = simpledialog.askfloat("Input", "Enter new Quantity:")
        price = simpledialog.askfloat("Input", "Enter new Price:")
        notes = simpledialog.askstring("Input", "Enter new Notes:")
        edit_record(index, date, ticker, trade_type, quantity, price, notes)
        messagebox.showinfo("Success", "Record edited successfully.")

if __name__ == '__main__':
    init_excel_file()
    gui_main()
