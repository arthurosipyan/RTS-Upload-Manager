# ClientNo (0) | Invoice# (1) | DebtorNo (2) | DebtorName (3) | Pono (4) | InvDate (5) | InvAmt (6)

from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import pandas as pd


def df_column(file, column):
    df = pd.read_excel(file)
    return df[column]


def invoices(file):
    upload_files = ""
    count = 0
    for file in df_column(file, 'Invoice#').tolist():
        if count <= 30:
            upload_files += '\"' + file + '\" '
            count += 1
        else:
            upload_files += '\n'
            count = 0
    return upload_files


if __name__ == '__main__':
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
    print("File:", filename)
    print("Total Invoices:", len(df_column(filename, 'Invoice#').tolist()))
    print("Total Revenue:", "${:,.2f}".format(df_column(filename, 'InvAmt').sum()))
    print(invoices(filename))

# TODO:
# - Tkinter GUI
# - open file button
# - display results
# - copy to clipboard button and indicate once copied
