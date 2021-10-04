import os

import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd


def get_revenue(sheet):
    return "${:,.2f}".format(sheet.sum())


def get_row_count(sheet):
    return len(sheet.index)


def get_invoices(sheet):
    invoices = sheet['Invoice#'].to_string(index=False).split('\n')
    print(len(invoices))
    rt = ''
    for i in invoices:
        rt += '\"{}.pdf\" '.format(i)
    print(rt)
    return rt


def open_file():
    file = fd.askopenfilename(
        initialdir='/',
        title='Open Excel File',
        filetypes=[('Excel File', '*.xlsx')]
    )
    df = pd.read_excel(file, sheet_name=-1)  # create df of spreadsheet

    invoice_count = get_row_count(df)
    total_revenue = get_revenue(df.iloc[:, -1].sum())
    invoices = get_invoices(df)

    # display file
    file_txt = tk.Text(root, height=5, width=52)
    file_txt.insert(tk.END, os.path.basename(file))
    file_txt.config(state='disabled')
    file_txt.grid(row=0, column=1)

    # display total invoices
    invoice_count_txt = tk.Text(root, height=5, width=52)
    invoice_count_txt.insert(tk.END, invoice_count)
    invoice_count_txt.config(state=tk.DISABLED)
    invoice_count_txt.grid(row=1, column=1)

    # display total revenue
    total_revenue_txt = tk.Text(root, height=5, width=52)
    total_revenue_txt.insert(tk.END, total_revenue)
    total_revenue_txt.config(state=tk.DISABLED)
    total_revenue_txt.grid(row=2, column=1)

    # display total invoices
    invoice_list_txt = tk.Text(root, height=5, width=52)
    invoice_list_txt.insert(tk.END, invoices)
    invoice_list_txt.config(state=tk.DISABLED)
    invoice_list_txt.grid(row=3, column=1)


if __name__ == '__main__':
    root = tk.Tk()
    root.title('RTS Upload Manager')
    root.geometry('600x400+50+50')

    button = tk.Button(root, text="Open", command=open_file)
    button.grid(row=1, column=0, pady=2)

    root.mainloop()
