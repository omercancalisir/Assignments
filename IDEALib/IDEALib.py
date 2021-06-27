# region Imports
import win32com.client as win32com
import logging
import pandas
import os.path as path
import tempfile
import numpy
import csv
import winreg
import time
import locale
from winreg import *
# endregion
__version__ = "1.0.0"

# region Setup
logging.basicConfig(
    filename="IDEALib.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Use appropriate delimiter and decimal separator values based on set locale
DELIMITER = ','
locale.setlocale(locale.LC_ALL, '')
DECIMAL_SEPARATOR = locale.localeconv()["decimal_point"]
if locale.localeconv()["decimal_point"] is not '.':
    DELIMITER = ';'



# endregion

# region Classes
class _SingletonIdeaClient:
    __instance = None
    __client = None

    @staticmethod
    def get_instance():
        if _SingletonIdeaClient.__instance is None:
            _SingletonIdeaClient()
        return _SingletonIdeaClient.__instance

    @staticmethod
    def get_client():
        if _SingletonIdeaClient.__instance is None:
            _SingletonIdeaClient()
        return _SingletonIdeaClient.__client

    def __init__(self):
        if _SingletonIdeaClient.__instance is not None:
            msg = "This class is a singleton and already has an instance"
            logging.warning(msg)
            raise Exception(msg)
        else:
            _SingletonIdeaClient.__instance = self
            _SingletonIdeaClient.__client = _connect_to_idea()


# endregion

# region Helper Functions
def _get_keys_by_value(dict,value):
    dictAsList = dict.items()
    keys = []
    for item in dictAsList:
        if item[1] is value:
            keys.append(item[0])
    return keys

'''
Reads IDEA registry to find out if it is ASCII or Unicode
Returns the appropriate file extension
'''
def _get_db_extension():
    subKey = "SOFTWARE\\CaseWare IDEA\\CaseWare IDEA\\InstallInfo"
    name = "AppStandard"
    try:
        registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subKey, 0,winreg.KEY_READ) 
        value, regtype = winreg.QueryValueEx(registry_key, name)
        winreg.CloseKey(registry_key)
    except WindowsError:
        return None
    
    if value == "ASCII":
        return ".imd"

    if value =="UNICODE":
        return ".idm"

    return None

# endregion

# region Connect To IDEA

# Returns IDEA Client object
def _connect_to_idea():
    try:
        idea = win32com.Dispatch(dispatch="Idea.IdeaClient")
    except:
        msg = "Unable to connect to IDEA."
        logging.error(msg)
        raise Exception(msg)
    return idea


def idea_client():
    logging.info("idea_client: Function Called")
    return _SingletonIdeaClient.get_client()


# endregion

# region Get Data From IDEA
def _export_database_from_IDEA(db,client,tempPath):    
    exportPath = path.join(tempPath,"tempExport.del")     
    task = db.ExportDatabase()    
    task.IncludeAllFields()
    task.IncludeFieldNames ="TRUE"
    eqn = ""
    task.Separators(DELIMITER,DECIMAL_SEPARATOR)

    task.PerformTask(exportPath,"Database","DEL UTF-8",1,db.Count,eqn)

    return exportPath    

## Goes through the database and returns a list of all the date columns, time columns,
## and a dictionary of all the other columns paired with their respective pandas type     
##
## IDEA Type             => Panda Type
## ---------------------------------------
## Character < 2500 vals => category
## Other Character       => object
## Numeric w/ 0 decimals => int64
## Numeric w/ decimals   => float64
## Date                  => datetime64
## Time                  => timedelta
## Multistate            => int8
## Boolean               => bool
def _map_database_col_types(db,client):
    # Define constants
    WI_VIRT_CHAR  = 0
    WI_VIRT_DATE  = 2
    WI_VIRT_NUM   = 1
    WI_VIRT_TIME  = 13
    
    WI_BOOL       = 10
    WI_MULTISTATE = 9

    WI_EDIT_CHAR  = 7
    WI_EDIT_DATE  = 8
    WI_EDIT_NUM   = 6
    WI_EDIT_TIME  = 12
    
    WI_DATE_FIELD = 5
    WI_CHAR_FIELD = 3
    WI_NUM_FIELD  = 4
    WI_TIME_FIELD = 11

    DATES = [WI_VIRT_DATE, WI_EDIT_DATE, WI_DATE_FIELD]
    TIMES = [WI_VIRT_TIME, WI_EDIT_TIME, WI_TIME_FIELD]
    CHARS = [WI_VIRT_CHAR, WI_EDIT_CHAR, WI_CHAR_FIELD]
    NUMS  = [WI_VIRT_NUM, WI_EDIT_NUM, WI_NUM_FIELD]
        
    tableDef = db.TableDef()    
    numCols = tableDef.Count
    columnPairs={}
    dates = []
    times = []

    for i in range(numCols):
        col = tableDef.GetFieldAt(i+1)        
        if col.Type in CHARS:
            columnPairs[col.Name] = object
        elif col.Type in NUMS:
            columnPairs[col.Name] = numpy.int64 if (col.Decimals is 0) else numpy.float64            
        elif col.Type  in TIMES:            
            times.append(col.Name) 
        elif col.Type in DATES:            
            dates.append(col.Name)
        elif col.Type is WI_BOOL:
            columnPairs[col.Name] = bool
        elif col.Type is WI_MULTISTATE:
            columnPairs[col.Name] = numpy.int8
    
    return columnPairs, dates, times


def _import_csv_as_dataframe(csvPath,colMapping,dateMapping):
    dataframe = pandas.read_csv(csvPath, dtype=colMapping, parse_dates=dateMapping, 
                                infer_datetime_format=True, encoding='UTF–8', sep=DELIMITER, decimal=DECIMAL_SEPARATOR, quotechar='"', quoting = csv.QUOTE_NONNUMERIC)
    return dataframe

# Converts the time fields from datetime to time only. Prevents data from being defaulted to date of import
def _clean_imported_times(df,times):
    for column in times:
         df[column] = pandas.to_timedelta(df[column])
    return None

# Check if it imported correctly, if not try to parse it again in more detail
def _clean_imported_dates(df,dates):
    for column in dates:
        if 'datetime64' not in df[column].dtype.name:
            df[column]=pandas.to_datetime(df[column])
    return None


# Based on type conversion, set any columns that were CHARACTERS and had less that 250 unique values to type category
def _convert_characters_to_categories(df,characters):
    UNIQUE_VALUE_THRESHHOLD = 2500
    for col in characters:
        if df[col].nunique() < UNIQUE_VALUE_THRESHHOLD:
            df[col] = df[col].astype('category')
    return None


'''
Takes an IDEA database file and imports it into a Panadas Dataframe

Parametes:
database - Path to the database, if only a file is given it will assume that it is the current working directory, if empty it will use current database.
header   - Boolean that determines if the database should be exported with first row as field names, defualts to True.
client   - COM Client object of IDEA, if not supplied it will create a new COM connection and use that.

Returns: Dataframe of the IDEA database
'''
def idea2py(database=None, client=None):
    logging.info("idea2py: Function Called")
    EMPTY = ""    
    if client is None:
        client = idea_client()        
    
    if database is None:
        try:
            database = client.CurrentDatabase().Name            
        except:            
            database = client.CommonDialogs().FileExplorer()            
            if database is EMPTY:
                msg = "You must select an IDEA database."
                logging.warning(msg)
                return None
            database = database.replace('/','\\')        
    
    root,ext = path.splitext(database)
    if ext is None or (ext.lower() != ".imd" and ext.lower() != ".idm"):
        ideaExtension = _get_db_extension()
        if ideaExtension is None:
            msg = "Error reading the IDEA registry."
            logging.error(msg)
            return None        
        database = root+ideaExtension

    try:
        db = client.OpenDatabase(database)
    except:
        msg = "An error occurred while opening the {} database.".format(database)
        logging.error(msg)
        return None

    if db.Count is 0:
        msg = "The {} database has no records.".format(database)
        logging.error(msg)
        return None
    logging.info("idea2py: Parameters verified.")      
    tempDir = tempfile.TemporaryDirectory()
    tempDirPath = tempDir.name
    try:
        tempPath = _export_database_from_IDEA(db,client,tempDirPath)
        logging.info("idea2py: IDEA Database exported to CSV.")

        mapping,dates,times = _map_database_col_types(db,client)
        logging.info("idea2py: IDEA column types mapped.")

        dataframe = _import_csv_as_dataframe(tempPath,mapping,dates)
        logging.info("idea2py: CSV imported as a dataframe.")
        
        # clean up the database
        _clean_imported_times(dataframe,times)
        _clean_imported_dates(dataframe,dates)
        logging.info("idea2py: Converted date and time columns to proper data type.")

        characters = _get_keys_by_value(mapping,object)
        _convert_characters_to_categories(dataframe,characters)
        logging.info("idea2py: Eligible character columns converted to categories.")

    except Exception as e:
        msg = "idea2py: IDEA database {} could not be imported.".format(database)
        logging.error(msg)
        msg = "Issue: {}".format(e)
        logging.error(msg)
        return None

    # Clean up resources
    tempDir.cleanup()
    logging.info("idea2py: Successful")
    return dataframe


# endregion

# region Send Data To IDEA
'''
Converts an IDEA column to a type that requires a mask(date/time)
'''
def _convert_masked_column(colName,mask,type,tableDef,tableMgt):
    field = tableDef.NewField()

    field.Name = colName
    field.Description = ""
    field.Type = type
    field.Equation = mask

    tableMgt.ReplaceField(colName,field)
    tableMgt.PerformTask()


'''
Converts imported date and time columns into actual date and time columns within idea
'''
def _convert_idea_columns(db,columnMap,client):
    # Define constants
    WI_TIME_FIELD = 11
    WI_DATE_FIELD = 5
    WI_BOOL       = 10
    WI_MULTISTATE = 9

    TIMEMASK = "HH:MM:SS"
    DATEMASK = "YYYY-MM-DD"
    
    tableMgt = db.TableManagement()
    tableDef = db.tableDef()
    columnMap = columnMap.items()
    for mapping in columnMap:        
        colName = mapping[0].upper()
        colType = mapping[1]
        if colType == "time":
            _convert_masked_column(colName,TIMEMASK,WI_TIME_FIELD,tableDef,tableMgt)
            
        if colType == "date":
            _convert_masked_column(colName,DATEMASK,WI_DATE_FIELD,tableDef,tableMgt)
                

'''
Imports a csv into IDEA 
'''
def _import_csv_into_idea(csvPath,tempPath,databaseName,client):
    rdfPath = path.join(tempPath,"temp_definition.rdf")    
    UTF8 = 2

    rdfTask = client.NewCsvDefinition()
    rdfTask.DefinitionFilePath = rdfPath
    rdfTask.CsvFilePath = csvPath
    rdfTask.FieldDelimiter = DELIMITER
    rdfTask.TextEncapsulator  = '"'
    rdfTask.CsvFileEncoding = UTF8
    rdfTask.FirstRowIsFieldNames = True

    dbObj = None
    try:
        client.SaveCSVDefinitionFile(rdfTask)
        client.ImportUTF8DelimFile(csvPath,databaseName,True,"",rdfPath,True)    
        dbObj = client.OpenDatabase(databaseName)        
    except:
        msg = "Error importing database."
        logging.error(msg)

    return dbObj
    

'''
Exports the dataframe to csv and returns the path to the file
'''
def _export_dataframe_to_csv(df,tempPath):
    exportPath = path.join(tempPath,"temp_export.csv")
    df.to_csv(path_or_buf=exportPath, sep=DELIMITER, index=False, header=True, encoding='utf-8', quotechar='"', decimal='.', quoting=csv.QUOTE_NONNUMERIC)

    # This is a work-around to how pandas works currently. When we upgrade pandas to the latest/newer version, this needs to be retested and probably refactored. Bug #6364
    if DECIMAL_SEPARATOR is not '.':
        df = pandas.read_csv(exportPath, sep=DELIMITER, quoting=csv.QUOTE_NONE, encoding='utf-8')
        df.to_csv(open(exportPath, 'w'), sep=DELIMITER, index=False, header=True, encoding='utf-8', decimal=DECIMAL_SEPARATOR, quoting=csv.QUOTE_NONE)
    return exportPath


'''
Takes in a bool and returns 1 for True, 0 for False.
'''
def _clean_boolean_values(bool):
    return int(bool)


'''
Takes in a timedelta and returns it in the format of HH:MM:SS.
'''
def _clean_timedelta_values(x):
    ts = x.total_seconds()
    hours, remainder = divmod(ts, 3600)
    minutes, seconds = divmod(remainder, 60)
    return ('{:02d}:{:02d}:{:02d}').format(int(hours), int(minutes), int(seconds)) 


'''
Cleans the incoming dataframe.

Splits up datetime fields into name_DATE, name_TIME, this is ignored if all the times are empty.
Converts all boolean to 1s for True, 0s for False.
Converts all the timedelta fields to the format of HH:MM:SS.

Returns the cleaned dataframe and a mapping of the columns.
'''
def _clean_dataframe_for_export(df):
    types = df.dtypes
    columns = df.columns
    toDrop = []
    map = {}
    
    for col in range(len(columns)):
        columnType = str(types[col])
        columnName = columns[col]
        
        if "datetime64" in columnType:
            logging.info("clean dataframe: Splitting datetime fields for column {}.".format(columnName))            
            colDate = pandas.to_datetime(df[columnName], errors='raise').dt.date
            colTime = pandas.to_datetime(df[columnName], errors='raise').dt.time
            
            logging.info("clean dataframe: Time number of unique values: {}.".format(colTime.nunique()))
            logging.info("clean dataframe: Time first value: {}.".format(str(colTime.head(1))))
            importTime = False
            if((colTime.nunique() > 1) or (colTime.nunique() is 1 and "00:00:00" not in str(colTime.head(1)))):
                logging.info("clean dataframe: Checking if format of {} is representing a datetime field.".format(colTime))
                importTime = True
            
            if importTime:
                logging.info("clean dataframe: Setting name_DATE portion.")
                dateHeader = "{}_DATE".format(columnName)            
                df.insert(col,dateHeader,colDate)
                map[dateHeader]="date"
                
                logging.info("clean dataframe: Setting name_TIME portion.")
                timeHeader = "{}_TIME".format(columnName)
                df.insert(col+1,timeHeader,colTime)
                map[timeHeader]="time"
                
                logging.info("clean dataframe: Datetime field new column name being: {}".format(columnName))
                toDrop.append(columnName)
            else:
                logging.info("clean dataframe: Determined {} column is not a datetime field".format(columnName))
                df[columnName]=colDate
                map[columnName]="date"
        else:
            if columnType == "bool":
                logging.info("clean dataframe: Converting {} to parseable boolean values".format(columnName))
                map[columnName]= "boolean"
                df[columnName] = numpy.vectorize(_clean_boolean_values)(df[columnName])
                 
            if columnType == "int8":
                 logging.info("clean dataframe: Converting int8 {} to parseable value".format(columnName))
                 map[columnName]= "multistate"

            if "timedelta" in columnType:
                logging.info("clean dataframe: Converting timeDelta {} to parseable value".format(columnName))
                map[columnName]= "time"
                df[columnName] = df[columnName].apply(_clean_timedelta_values)
    df = df.drop(toDrop,axis=1)
    return df,map



'''
Takes in a dataframe,database name and idea client
Creates an IDEA database with the same name and data
Returns the IDEA database object if successful
'''
def py2idea(dataframe, databaseName, client=None, createUniqueFile = False):
    logging.info("py2idea: Function Called")
    if client is None:
        client = idea_client()
    
    if databaseName is None:
        msg = "Missing database name."
        logging.warning(msg)
        return None

    if dataframe is None:
        msg = "Missing dataframe."
        logging.warning(msg)
        return None
    
    if len(dataframe) is 0:
        msg = "The dataframe has no records."
        logging.warning(msg)
        return None       
    
    root,ext = path.splitext(databaseName)
    if ext is None or (ext.lower() != ".imd" and ext.lower() != ".idm"):
        ideaExtension = _get_db_extension()
        if ideaExtension is None:
            msg = "Error reading the IDEA registry."
            logging.error(msg)
            return None        
        databaseName = root+ideaExtension
    
   
    if databaseName.count("\\") == 0:
        # Assume it to be in the current working folder 
        workingDir = client.WorkingDirectory
        databaseName = path.join(workingDir,databaseName)

    if createUniqueFile:
        databaseName = client.UniqueFileName(databaseName)
    elif path.exists(databaseName):
        msg = "IDEA database {} already exists".format(databaseName)
        logging.error(msg)
        return None
    
    tempDir = tempfile.TemporaryDirectory()
    tempPath = tempDir.name
    logging.info("py2idea: Parameters verified")
    try: 
        dataframe,mapping = _clean_dataframe_for_export(dataframe)
        logging.info("py2idea: Dataframe values cleaned for export.")
        csvPath = _export_dataframe_to_csv(dataframe,tempPath)
        db = _import_csv_into_idea(csvPath,tempPath,databaseName,client)
        logging.info("py2idea: CSV imported into IDEA.")

        _convert_idea_columns(db,mapping,client)
        logging.info("py2idea: IDEA columns converted.")
    except Exception as e:
        msg = "py2idea: IDEA database {} could not be created.".format(databaseName)
        logging.error(msg)
        msg = "Issue: {}".format(e)
        logging.error(msg)
        return None
    
    tempDir.cleanup()
    logging.info("py2idea: Successful")
    return db
# endregion

logging.info("")
logging.info("----IDEALib.py Loaded----")