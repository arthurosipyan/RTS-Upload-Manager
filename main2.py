# ClientNo (0) | Invoice# (1) | DebtorNo (2) | DebtorName (3) | Pono (4) | InvDate (5) | InvAmt (6)

import os
from tkinter import Tk  # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import pandas as pd


def get_revenue(sheet):
    return "${:,.2f}".format(sheet.sum())


def get_row_count(sheet):
    return len(sheet.index)


def print_invoices(sheet):
    print("Invoices:\n")
    for invoice in sheet.tolist():
        print("\"" + invoice + "\"")


if __name__ == '__main__':
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    df = pd.read_excel(filename, sheet_name=-1)  # create df of spreadsheet

    file = os.path.basename(filename)
    total_invoices = get_row_count(df)
    total_revenue = get_revenue(df.iloc[:, -1])

    print("File:", file)
    print("Total Invoices:", total_invoices)
    print("Total Revenue:", total_revenue)
    # print_invoices(df.iloc[:, 1])
