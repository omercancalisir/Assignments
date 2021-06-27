# C:\Program Files\CaseWare IDEA\IDEA\Lib\site-packages dizininde bulunan IDEALib' in import edilmesi işlemi
import IDEALib.IDEALib as ideaLib
from pandastable import Table
from tkinter import *


# table oluşturma kodu
class TestApp(Frame):
    """Basic test frame for the table"""

    def __init__(self, parent=None, dbname=None, dataframe=None):
        self.parent = parent
        self.dbname = dbname
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('600x400+200+100')
        self.main.title('Customers Over Credit Limit ')
        f = Frame(self.main)
        f.pack(fill=BOTH, expand=1)

        data = dataframe

        df = data
        self.table = pt = Table(f, dataframe=df,
                                showtoolbar=True, showstatusbar=True)
        pt.show()
        return


class ConnectionIdea:

    def __init__(self):
        self.dbName = ""
        self.idea = ""

    def connectIdea(self, dbname):
        self.dbName = dbname
        # Connect to IDEA
        self.idea = ideaLib.idea_client()

    def exportData(self):
        # Export IDEA database
        dataframe = ideaLib.idea2py(database=self.dbName, client=self.idea)
        return dataframe


def dataExtraction(dbname):
    con = ConnectionIdea()
    con.connectIdea(dbname)
    dataframe = con.exportData()
    return dataframe


class Program:

    def __init__(self):
        self.dataframe = ""
        self.custOverCreditLimit = ""

    def run(self):
        dbname = "CMF-BT.IMD"
        self.dataframe = dataExtraction(dbname)
        self.custOverCreditLimit = self.controlExceeding()

    def controlExceeding(self):
        customersOverTheLimit = self.dataframe.groupby(['CUSTNO', 'CREDIT_LIM']).sum().dropna().reset_index()
        for index, row in customersOverTheLimit.iterrows():
            if (row["CREDIT_LIM"] + row["AMOUNT"]) >= 0:
                customersOverTheLimit = customersOverTheLimit.drop(index)
        customersOverTheLimit['AMOUNT'] = customersOverTheLimit['AMOUNT'].apply(lambda x: '{:.2f}'.format(x))
        return customersOverTheLimit


if __name__ == '__main__':
    program = Program()
    program.run()

    app = TestApp(dataframe=program.custOverCreditLimit)
    app.mainloop()
