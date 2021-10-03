from main2 import get_row_count, get_revenue, print_invoices

import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd


# def run_report(filename):
#     df = pd.read_excel(filename, sheet_name=-1)  # create df of spreadsheet
#     file = "File:", os.path.basename(filename)
#     total_invoices = "Total Invoices:", get_row_count(df)
#     total_revenue = "Total Revenue:", get_revenue(df.iloc[:, -1])
#     # print_invoices(df.iloc[:, 1])
#     global result
#     result = total_invoices
#

def open_file():
    file = fd.askopenfilename(
        initialdir='/',
        title='Open Excel File',
        filetypes=[('Excel File', '*.xlsx')]
    )
    report.insert(tk.END, file)
    file = open(file, 'r', encoding="utf8")
    data = file.read()
    report.insert(tk.END, data)
    file.close()


if __name__ == '__main__':
    root = tk.Tk()
    root.title('RTS Upload Manager')
    root.geometry('500x300')

    report = tk.Text(root, height=20, width=40)

    button = tk.Button(root, text="Open", command=open_file)
    button.grid(column=1, row=1)

    # T.pack()
    # T.insert(tk.END, "Just a text Widget\nin two lines\n")  # return our report here

    root.mainloop()
