import os

import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd


def get_revenue(sheet):
    return "${:,.2f}".format(sheet.sum())


def get_row_count(sheet):
    return len(sheet.index)


def print_invoices(sheet):
    print("Invoices:\n")
    for invoice in sheet.tolist():
        print("\"" + invoice + "\"")


def open_file():
    file = fd.askopenfilename(
        initialdir='/',
        title='Open Excel File',
        filetypes=[('Excel File', '*.xlsx')]
    )
    df = pd.read_excel(file, sheet_name=-1)  # create df of spreadsheet

    file_name = os.path.basename(file)
    total_invoices = get_row_count(df)
    total_revenue = get_revenue(df.iloc[:, -1].sum())

    # display file
    file_txt = tk.Text(root)
    file_txt.insert(tk.END, file_name)
    file_txt.config(state='disabled')
    file_txt.pack()

    # display total invoices
    total_invoices_txt = tk.Text(root)
    total_invoices_txt.insert(tk.END, total_invoices)
    total_invoices_txt.config(state=tk.DISABLED)
    total_invoices_txt.pack()

    # display total revenue
    total_revenue_txt = tk.Text(root)
    total_revenue_txt.insert(tk.END, total_revenue)
    total_revenue_txt.config(state=tk.DISABLED)
    total_revenue_txt.pack()

    print_invoices(df)


if __name__ == '__main__':
    root = tk.Tk()
    root.title('RTS Upload Manager')
    root.geometry('600x400+50+50')

    button = tk.Button(root, text="Open", command=open_file)
    button.pack(side=tk.TOP, pady=5)

    root.mainloop()
