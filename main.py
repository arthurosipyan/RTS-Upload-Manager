# read spreadsheet and save pro #'s column with no more than 25 rows at a time

# Example of spreadsheet column format:
# ClientNo (0) | Invoice# (1) | DebtorNo (2) | DebtorName (3) | Pono (4) | InvDate (5) | InvAmt (6)

# intended output: "OT1900" "OT1901" "OT1903"
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
import pandas as pd

# Method to get total count of invoices

# Method to get total amount


def rts_upload():
    Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
    filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file

    print(filename)
    df = pd.read_excel(filename)

    return df['Invoice#'].tolist()


if __name__ == '__main__':
    upload_files = ""
    count = 0
    for file in rts_upload():
        upload_files += '\"' + file + '\" '
        count += 1
    print("Total Invoices: ", count)
    print(upload_files)
