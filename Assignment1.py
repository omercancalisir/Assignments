import win32com.client as win32comclient

if __name__ == "__main__":
    idea = win32comclient.Dispatch(dispatch="Idea.IdeaClient")
    new_filename = idea.UniqueFileName("New-Bank Transaction")
    try:
        db = idea.opendatabase("Sample-Bank Transactions.IMD")
        task = db.Extraction()
        task.IncludeAllFields
        task.AddExtraction(new_filename, "", "AMOUNT < 0")
        task.PerformTask(1, db.Count)
        task = None
    finally:
        idea.RefreshFileExplorer()
        task = None
        db = None
        idea = None
