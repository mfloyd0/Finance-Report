import  Formatter
import Data_Insertion
from dateutil.relativedelta import relativedelta
from datetime import datetime
import tkinter
from tkinter import filedialog


if __name__ == '__main__':

    previous_month = datetime.today() - relativedelta(months=1)

    month = previous_month.strftime("%B")
    #month = 'February'

    report_file = f'{month} Report.xlsx'
    #report_file = 'Report Template.xlsx'

    print("Choose the last month bank statement")
    tkinter.Tk().withdraw()
    bankFile = filedialog.askopenfilename()



    Formatter.formatFile(report_file)

    Data_Insertion.insertData(bankFile, report_file)

    print("done")




