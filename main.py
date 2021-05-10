# read spreadsheet and save pro #'s column with no more than 25 rows at a time

# Example of spreadsheet column format:
# ClientNo (0) | Invoice# (1) | DebtorNo (2) | DebtorName (3) | Pono (4) | InvDate (5) | InvAmt (6)

# example output: "OT1900" "OT1901" "OT1903"
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import pandas as pd


def total_invoices(file):
    return len(rts_upload(file, 'Invoice#'))


def total_revenue(file):
    revenue = 0
    for x in rts_upload(file, 'InvAmt'):
        revenue += x
    return revenue


def invoices(file):
    upload_files = ""
    for file in rts_upload(file, 'Invoice#'):
        upload_files += '\"' + file + '\" '
    return upload_files


def rts_upload(file, column):
    df = pd.read_excel(file)
    return df[column].tolist()


if __name__ == '__main__':
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    print("File:", filename)
    print("Total Invoices:", total_invoices(filename))
    print("Total Revenue:", total_revenue(filename))
    print("Invoices:", invoices(filename))
