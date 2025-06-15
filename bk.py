import pandas as pd
from tkinter import *
from tkinter import messagebox, simpledialog

# Initialize Excel file and DataFrame
EXCEL_FILE = 'trading_book.xlsx'

def init_excel_file():
    # Create an Excel file with headers if it doesn't exist
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Date', 'Ticker', 'Trade_Type', 'Quantity', 'Price', 'Notes'])
        df.to_excel(EXCEL_FILE, index=False)

def load_data():
    return pd.read_excel(EXCEL_FILE)

def add_record(date, ticker, trade_type, quantity, price, notes):
    new_record = pd.DataFrame({'Date': [date], 'Ticker': [ticker], 'Trade_Type': [trade_type],
                               'Quantity': [quantity], 'Price': [price], 'Notes': [notes]})
    df = load_data()
    df = pd.concat([df, new_record], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

def delete_record(index):
    df = load_data()
    if 0 <= index < len(df):
        df = df.drop(index).reset_index(drop=True)
        df.to_excel(EXCEL_FILE, index=False)

def edit_record(index, date, ticker, trade_type, quantity, price, notes):
    df = load_data()
    if 0 <= index < len(df):
        df.at[index, 'Date'] = date
        df.at[index, 'Ticker'] = ticker
        df.at[index, 'Trade_Type'] = trade_type
        df.at[index, 'Quantity'] = quantity
        df.at[index, 'Price'] = price
        df.at[index, 'Notes'] = notes
        df.to_excel(EXCEL_FILE, index=False)

def show_records():
    df = load_data()
    records_window = Toplevel()
    records_window.title("Trading Records")
    
    text = Text(records_window)
    text.insert(END, df.to_string(index=False))
    text.pack()

def gui_main():
    root = Tk()
    root.title('Trading Bookkeeping System')
    
    # Buttons for functionalities
    Button(root, text='Add Record', command=lambda: add_record_form()).pack()
    Button(root, text='Delete Record', command=lambda: delete_record_form()).pack()
    Button(root, text='Edit Record', command=lambda: edit_record_form()).pack()
    Button(root, text='Show Records', command=show_records).pack()
    
    root.mainloop()

def add_record_form():
    # Simple form to add a record
    date = simpledialog.askstring("Input", "Enter date (YYYY-MM-DD):")
    ticker = simpledialog.askstring("Input", "Enter Ticker:")
    trade_type = simpledialog.askstring("Input", "Enter Trade Type (Buy/Sell):")
    quantity = simpledialog.askinteger("Input", "Enter Quantity:")
    price = simpledialog.askfloat("Input", "Enter Price:")
    notes = simpledialog.askstring("Input", "Enter Notes:")
    add_record(date, ticker, trade_type, quantity, price, notes)
    messagebox.showinfo("Success", "Record added successfully.")

def delete_record_form():
    index = simpledialog.askinteger("Input", "Enter index to delete (starts from 0):")
    delete_record(index)
    messagebox.showinfo("Success", "Record deleted successfully.")

def edit_record_form():
    index = simpledialog.askinteger("Input", "Enter index to edit (starts from 0):")
    if index is not None:
        date = simpledialog.askstring("Input", "Enter new date (YYYY-MM-DD):")
        ticker = simpledialog.askstring("Input", "Enter new Ticker:")
        trade_type = simpledialog.askstring("Input", "Enter new Trade Type (Buy/Sell):")
        quantity = simpledialog.askinteger("Input", "Enter new Quantity:")
        price = simpledialog.askfloat("Input", "Enter new Price:")
        notes = simpledialog.askstring("Input", "Enter new Notes:")
        edit_record(index, date, ticker, trade_type, quantity, price, notes)
        messagebox.showinfo("Success", "Record edited successfully.")

if __name__ == '__main__':
    init_excel_file()
    gui_main()
