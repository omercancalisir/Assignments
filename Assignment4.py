# C:\Program Files\CaseWare IDEA\IDEA\Lib\site-packages dizininde bulunan İDEALib' in import edilmesi işlemi
import IDEALib.IDEALib as ideaLib
from pandastable import Table

from tkinter import *
from tkinter import filedialog

import os


# table oluşturma kodu
class TestApp(Frame):
    """Basic test frame for the table"""

    def __init__(self, parent=None, dbname=None):
        self.parent = parent
        self.dbname = dbname
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('600x400+200+100')
        self.main.title('Table app')
        f = Frame(self.main)
        f.pack(fill=BOTH, expand=1)

        con = ConnectionIdea()
        con.connectIdea(dbname)

        data = con.exportData()
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


class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Assignment 4")
        self.minsize(640, 400)
        self.label = Label(self, text="Open A File").pack()
        self.button()

    def button(self):
        self.button = Button(self.label, text="Browse A File", command=self.fileDialog).pack()

    def fileDialog(self):
        # full path of the file
        self.filename = filedialog.askopenfilename(title="Select A File")
        # just the name of file
        self.filename = os.path.basename(self.filename)
        self.label2 = Label(self.label, text=self.filename).pack()
        self.showtable()

    def showtable(self):
        TestApp(dbname=self.filename).mainloop()


if __name__ == '__main__':
    root = Root()
    root.mainloop()
