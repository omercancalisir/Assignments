# C:\Program Files\CaseWare IDEA\IDEA\Lib\site-packages dizininde bulunan İDEALib' in import edilmesi işlemi
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
        self.main.title('Table app')
        f = Frame(self.main)
        f.pack(fill=BOTH, expand=1)
        data = dataframe
        df = data
        self.table = pt = Table(f, dataframe=df,
                                showtoolbar=True, showstatusbar=True)
        pt.show()

        return


class UserDefined:

    def __init__(self):
        self.datafr = ""

    def main(self):
        option = self.selectOptions()
        columnname = self.getColumnName()
        criteria = self.getCondition()
        if option == "1":
            dbName = "Client Master File-Database.IMD"
            self.customCluster(columnname, criteria, dbName)

        elif option == "2":
            dbName = "Sample-Bank Transactions.IMD"
            self.customCluster(columnname, criteria, dbName)
        else:
            print("invalid option selected")

    def customCluster(self, columnname, criteria, dbName):
        idea = ideaLib.idea_client()
        try:
            dbName = dbName
            dataframe = ideaLib.idea2py(database=dbName, client=idea)
            dataTypeObj = dataframe[columnname].dtype.name

            if dataTypeObj == 'object' or dataTypeObj == 'category':
                dataframe = dataframe.loc[dataframe[columnname] == criteria]
                self.datafr = dataframe
            elif dataTypeObj == 'int64' or dataTypeObj == 'float64':
                criteria = criteria.split()
                if criteria[0] == ">":
                    dataframe = dataframe.loc[dataframe[columnname] > int(criteria[1])]
                    self.datafr = dataframe
                elif criteria[0] == "<":
                    dataframe = dataframe.loc[dataframe[columnname] < int(criteria[1])]
                    self.datafr = dataframe
                elif criteria[0] == ">=":
                    dataframe = dataframe.loc[dataframe[columnname] >= int(criteria[1])]
                    self.datafr = dataframe
                elif criteria[0] == "<=":
                    dataframe = dataframe.loc[dataframe[columnname] <= int(criteria[1])]
                    self.datafr = dataframe

        finally:
            idea.RefreshFileExplorer()
            """
            task = None
            db = None
            idea = None
            """
    def openDatabase(self, dbName, idea):
        db = idea.opendatabase(dbName)
        return db

    def selectOptions(self):
        print("1. Client Master File-Database.IMD")
        print("2. Sample-Bank Transactions.IMD")
        print("İslem yapmak istediğiniz veritabanını seçiniz :", end="")
        userinput = input()
        return userinput

    def getColumnName(self):
        print("Column Name :", end="")
        userinput = input()
        return userinput

    def getCondition(self):
        print("Condition :", end="")
        userinput = input()
        return userinput


if __name__ == "__main__":
    program = UserDefined()
    program.main()

    app = TestApp(dataframe=program.datafr)
    app.mainloop()
