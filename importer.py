import csv, time, json, logging, traceback, os, datetime
import xlrd
from xlrd import open_workbook
from exceptions import ImportException, FileFormatException

import pyodbc
import hashlib

#The path to the access database file
TABLE_NAME = "WeeklyShipments"


#The imported data will be considered stale after 4 days
MAX_AGE_MS = 1000 * 60 * 60 * 24 * 4

logger = logging.getLogger('importer')

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
        logger.info("Sha1 hash of file: {0}".format(self.sha1()))
    
        #pull the data from the import file.  This will be returned as a 2D array
        importData = self.import_data()
        self.insert_rows(importData)
        
        
    def import_data(self):
        logger.info('Reading records')
        path = os.path.normpath(self.filename)
        book = open_workbook(path)
        sheet = book.sheet_by_name('Page1_2')

        # Rows and columns are indexed starting at 0.  Skip row 0 since this is a title row

        # Pull the column names from row 1
        keys = [sheet.cell(1, col).value for col in range(sheet.ncols)]

        data = []

        # Pull the row data from row 2 through the end of the sheet (skipping the trailing footer row)
        for row in range(2, sheet.nrows -1):
            d = {
                keys[col]: sheet.cell(row, col).value 
                for col in range(sheet.ncols)
            }
            # Excel stores the date as a number, convert back to a string
            # Convert to a datetime date string in ISO format (YYYY-MM-DD)
            dt_xl = d.get('Transaction Date')
            logger.debug("read date as {}".format(dt_xl))

            dt = xlrd.xldate_as_datetime(dt_xl, book.datemode)
            logger.debug("As datetime {}".format(dt))
            d['Transaction DateTime'] = dt

            dts = dt.date().isoformat()
            logger.debug("As date string {}".format(dts))

            d['Transaction Date String'] = dts
            data.append(d)

            if row == 2:
                logger.info("File appears to be for DC {}".format(d.get('DC Id')))
        
        logger.info('Read {0} records'.format(len(data)))
        return data
    
    def insert_rows(self, importData):
        # Insert all rows into the database. If we encounter an exception the transaction will be rolled
        # back to the time the database connection was established
        logger.info("Writing records into database")
        with pyodbc.connect('Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={0};'.format(self.database)) as conn:
            with conn.cursor() as cursor:
                for row in importData:
                    try:
                        cursor.execute("INSERT INTO [{0}] ([DC ID], [DC Name], [Store ID], [Store Name], [Address], [City], [State],"
                            "[Zip], [Transaction Date], [Container Type], [Container Qty]) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)".format(TABLE_NAME),
                            row.get('DC Id'),
                            row.get('DC Name'),
                            row.get('Store Id'),
                            row.get('Store Name'),
                            row.get('Address'),
                            row.get('City'),
                            row.get('State'),
                            row.get('Zip'),
                            row.get('Transaction DateTime'),
                            row.get('Container Type'),
                            row.get('Container Qty')
                        )
                    except Exception as e:
                        logger.error("Error importing row {}".format(row))
                        raise e
                logger.info('Wrote {0} records'.format(len(importData)))
            



def to_dict(obj):
    return json.loads(json.dumps(obj, default=lambda o: o.__dict__))
