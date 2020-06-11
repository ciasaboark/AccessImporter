import csv, time, json, logging, traceback
from exceptions import ImportException, FileFormatException

import pyodbc
import hashlib

#The path to the access database file
TABLE_NAME = "Table1"


#The imported data will be considered stale after 4 days
MAX_AGE_MS = 1000 * 60 * 60 * 24 * 4

logger = logging.getLogger('watcher.importer')

class Importer:
    def __init__(self, database, filename):
        self.database = database
        self.filename = filename

    def sha1(self):
        hash_sha1 = hashlib.sha1()
        with open (self.filename, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_sha1.update(chunk)
        return hash_sha1.hexdigest()

    def begin_import(self):
        logger.info("Beginning import of file {0}".format(self.filename))
        logger.debug("Sha1 hash of file: {0}".format(self.sha1()))
    
        
        #pull the data from the import file.  This will be returned as a 2D array
        importData = self.import_data()
        self.insert_rows(importData)
        
        
    def import_data(self):
        logger.debug('Reading records from Table1')
        conn = pyodbc.connect('Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};'.format(self.filename))
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM [Table1]')
        result = cursor.fetchall()
        logger.debug('Read {0} records'.format(len(result)))
        return result
    
    def insert_rows(self, importData):
        with pyodbc.connect('Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};'.format(self.database)) as conn:
            with conn.cursor() as cursor:
                cursor.executemany('INSERT INTO [Table1] (id, first, last) VALUES (?, ?, ?)', importData)
            



def to_dict(obj):
    return json.loads(json.dumps(obj, default=lambda o: o.__dict__))
