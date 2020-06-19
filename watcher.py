# -*- coding: utf-8 -*-

import sys, os, time, logging, shutil, re
from watchdog.observers.polling import PollingObserver
from watchdog import events
from pathlib import Path
from importer import Importer
from exceptions import ImportException, FileFormatException
from time import sleep
import struct
import pyodbc
import configargparse
import textwrap
import traceback
from pprint import pformat
from datetime import datetime
from pathlib import Path
from logging.handlers import TimedRotatingFileHandler
from logging.handlers import NTEventLogHandler
import win32serviceutil, win32service
import win32event
import servicemanager
import configargparse
import winreg
import types

VERSION = "0.0.7"

#The interval (in seconds) between sleep cycles. A manual check is performed on each wake
SLEEP_INTERVAL = 60 * 10

REGISTRY_KEY_NAME = 'SOFTWARE\\ContainerTracking'

IMPORT_FILE_TYPES = [".xls", ".xlsx"]

#Keep the logger and the configuration as global variables
logger = None
opts = types.SimpleNamespace()

def read_key(key, name: str, def_val):
    """Read the current value for key 'name' from the Windows registry.
    
    The default value will be used if the registry could not be opened
    or if the value name does not exist for the given key
    
    Args:
        key: the base key to pull values from
        name (str): the value name to pull from the base key
        def_val (str): a default value to use if an error occurs reading from the registry
    
    Returns:
        str: The value read from the registry or def_val
    """
    logger.info("checking for key {}".format(name))
    val = def_val
    try:
        val = winreg.QueryValueEx(key, name)[0]
    except FileNotFoundError as e:
        logger.warning("Unable to read config option '{0}' from registry, using default value of '{1}'"
            .format(name, def_val))
        write_key(key, name, val)
        
    return val

def write_key(key, name, val):
    """Write a value into the Windows registry.

    Args:
        key: The base key to write values into.
        name (str): The name of the value to write.
        val (str): The value to write.

    """
    logger.info("Inserting value '{0}' into registry as 'HKLM\\{1}\{2}'".format(def_val, REGISTRY_KEY_NAME, name))
    try:
        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, val)
    except Exception as e:
        logger.error("Unable to write value '{0}' to value name '{1}' in key 'HKLM\\{2}'".format(val, name, REGISTRY_KEY_NAME))
        logger.exception(e)

def write_default_opts(key):
    """Write default values to the Windows registry if they do not already exist
    
    Args:
        key: The base key to write values into.    
    """
    write_default(key, 'watch', "C:\\import")
    write_default(key, 'archive', "C:\import\\archive")
    write_default(key, 'errors', "C:\\import\\archive\\errors")
    write_default(key, 'database', "C:\\db\\database.accdb")
    write_default(key, 'log_file', "C:\\db\\logs\\watcher.log")

def write_default(key, name: str, value: str):
    """Write a value to the Windows registry if it does not exist
    
    Args:
        key: The base key to write values into.
        name (str): The name of the value to write.
        val (str): The value to write.
    """
    if read_key(key, name, None) == None:
        write_key(key, name, value)

# Set up the logger.  We won't know what file to log to yet, so we will start by logging
# only warnings and above to the Windows event log
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Log warnings and above to the windows event log
eh = NTEventLogHandler('container_tracking_importer')
eh.setLevel(logging.WARNING)
eh.setFormatter(formatter)
logger.addHandler(eh)


class Watcher(win32serviceutil.ServiceFramework):
    """The Windows service responsible for monitoring the import directory for Excel files.
    The watcher will use a combination of a polling observer and timed manual wakes to search
    for new Excel files. Any files found will be passed off to the Importer for processing.
    """

    _svc_name_ = 'container_tracking_importer'
    _svc_display_name_ = 'Container Tracking Importer'
    _svc_description_ = 'Import container counts from Excel spreadsheets to an Access database. (v{})'.format(VERSION)

    def __init__(self, args):
        """Set up our run-time options and switch to the rotating log
        """

        # We need to read the log_file setting before anything else so we can begin logging to a
        # file as soon as possible
        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, REGISTRY_KEY_NAME)
        write_default_opts(key)
        opts.log_file = read_key(key, 'log_file', "C:\\db\\logs\\watcher.log")

        try:
            fh = TimedRotatingFileHandler(opts.log_file, when='d', interval=1, backupCount=7, encoding='utf-8')
            fh.setLevel(logging.INFO)
            fh.setFormatter(formatter)
            logger.addHandler(fh)

            # remove the windows event log handler
            logger.removeHandler(eh)
        except Exception as e:
            logger.error("Unable to initialize log file: '{}'. Verify the parent folder exists and that you have write access.".format(e))
            sys.exit(1)

        logger.info("Service initializing")
        logger.info("Reading settings from HKEY_LOCAL_MACHINE\\SOFTWARE\\ContainerTracking\\")

        opts.watch = read_key(key, 'watch', "C:\\import")
        opts.archive = read_key(key, 'archive', "C:\import\\archive")
        opts.errors = read_key(key, 'errors', "C:\\import\\archive\\errors")
        opts.database = read_key(key, 'database', "C:\\db\\database.accdb")

        logger.info("Running with options: {}".format(opts))
        winreg.CloseKey(key)

        #Keep track of the last time the service woke (in epoch seconds)
        self.last_wake = -1

        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        
    def SvcStop(self):
        """Shut down the service
        This will be called by automatically by Windows
        """
        logger.info("Service asked to stop")
        self.run = False
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """Start the service
        This will be called automatically be Windows
        """
        logger.info("Service is starting")
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_,''))
        self.run = True
        self.start()
        self.main()

    def stop(self):
        """Perform any last steps to shut down the service
        There is not much to do here other than to ask the file watcher to
        stop.
        """
        self.isRunning = False
        self.observer = None
        self.handler = None
        self.observer.stop()
        self.observer.join()

    def start(self):
        """Perform required initialization before the main loop can begin
        """
        self.isRunning = True
        logger.info("╔═════════════════════════════════════════╗")
        logger.info("║ Container Tracking Importer             ║")
        logger.info("╠═════════════════════════════════════════╣")
        logger.info("║ Version {:<8}                        ║".format(VERSION))
        logger.info("║                                         ║")
        logger.info("╚═════════════════════════════════════════╝")

        # The polling observer is heavier than the other observers, but was more reliable during testing
        self.observer = PollingObserver()
        event_handler = Handler()
        
        # Store the list of watched directories as a set so we can avoid duplicates
        self.watched_directories = set()

        # Do some tests to make sure we can work with the settings provided
        self.do_integrety_tests()
        
        # Schedule the polling observer and start observing
        for d in self.watched_directories:
            self.observer.schedule(event_handler, d, recursive=False)
        self.observer.start()

        logger.info("Watchdog running. Monitoring new files in paths {0}".format(pformat(self.watched_directories, indent=1, width=80, depth=None, compact=False)))
        logger.info("Completed imports will be moved into '{0}'".format(opts.archive))  
        logger.info("Failed imports will be moved into '{0}'".format(opts.errors))

    def do_integrety_tests(self):
        """Check that we can work with the settings given. Terminate the process otherwise
        """

        # Do we have an Access database driver installed?
        # Microsoft has two main versions. A old driver that can only access .mdb
        # databases, and a newer driver that can read .mdb and .accdb databases.
        
        drivers = [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]
        if len(drivers) == 0:
            logger.error('Unable to find any Access database drivers installed')
            log_access_driver_error()
            sys.exit(1)
        elif 'Microsoft Access Driver (*.mdb, *.accdb)' not in drivers:
            #We have at least one Access driver installed, but it is the older version
            logger.error('Unable to find the updated MS Access database driver')
            log_access_driver_error()
            sys.exit(1)
        else:
            logger.info("Using driver: Microsoft Access Driver (*.mdb, *.accdb)")


        # Filter down the list of directories to watch to only the ones that exist
        exists = os.path.exists(opts.watch)
        isFile = os.path.isfile(opts.watch)
        if not exists:
            logger.warning("Will not watch for new files in '{0}'. Directory does not exist.".format(opts.watch))
            sys.exit(1)
        elif isFile:
            logger.warning("Will not watch for new files in '{0}'. Path is a file, not a directory.".format(opts.watch))
            sys.exit(1)
        else:
            logger.debug("Adding '{0}' to the watched directories list".format(opts.watch))
            self.watched_directories.add(opts.watch)

        # Error out if we have no directories to watch
        if len(self.watched_directories) > 0:
            logger.debug("Found at least one directory to watch for imported data")
        else:
            logger.error("Could not find any directories to watch.  Exiting...")
            sys.exit(1)
        
        # Check if the archive directory exists
        if not os.path.exists(opts.archive) or os.path.isfile(opts.archive):
            logger.error("The archive directory '{0}' does not exist".format(opts.archive))
            sys.exit(1)

        # Check if the errors archive directory exists
        if not os.path.exists(opts.errors) or os.path.isfile(opts.errors):
            logger.error("The errors archive directory '{0}' does not exist".format(opts.errors))
            sys.exit(1)

        # Check if the database we are importing into exists. Defer checking whether this is a valid Access database until later
        if not os.path.exists(opts.database) or not os.path.isfile(opts.database):
            logger.error("The database file '{0}' does not exist".format(opts.database))
            sys.exit(1)

        # Make sure we aren't being asked to move imported data into the same folder we are watching.
        # Use os.path.sameFile() to resolve relative paths back to absolute ones
        arkPath = os.path.normpath(opts.archive)
        errPath = os.path.normpath(opts.errors)
        for path in self.watched_directories:
            realPath = os.path.normpath(path)
            if (os.path.samefile(realPath, arkPath) or os.path.samefile(realPath, errPath)):
                logger.error("The archive directory '{0}' or the error directory {1} is in the list of directories to watch. This "
                "will result in an endless loop.".format(opts.archive, opts.errors))
                logger.error("Choose a different archive directory or error directory or remove this directory from the list of watched directories")
                sys.exit(1)
    
    def main(self):
        """The main event loop
        """
        try:
            while True:
                # Pause here for one second to check if we have been asked to stop
                win32event.WaitForSingleObject(self.hWaitStop, 1000)

                # The service has been asked to stop.
                if not self.isRunning:
                    logger.info("Exiting...")
                    sys.exit(0)

                # Wake after every sleep interval to check for any Excel files that need to be imported
                # The manual check is needed to handle importing files that could not be imported earlier
                # or that already existed when the service started
                now = time.time()
                if (self.last_wake == -1 or now - self.last_wake > SLEEP_INTERVAL):
                    logger.info("🌞🥱 Checking for any files to manually import")
                    self.last_wake = now
                    self.manual_import()
                    logger.info("Back to sleep 😴")
        except Exception as e:
            logger.debug("Error with watchdog.  Exiting...")
            logger.exception(e)
            sys.exit(1)           
            

    def manual_import(self):
        """ Check for any pre-existing database files to import.  Since 
        these already exist we won't get a on_created() callback in the Handler """

        logger.debug("Performing manual check for database files to import")
        to_import = []
        for d in self.watched_directories:
            dir = os.path.normpath(d)
            logger.debug("Checking directory '{}'".format(dir))
                        
            for f in os.listdir(dir):
                logger.debug("Found child '{}'".format(f))
                full_path = os.path.join(d, f)
                if os.path.isfile(full_path):
                    logger.debug("Child is a file {}".format(f))
                    to_import.append(full_path)

        if len(to_import) == 0:
            logger.info("No files found to import")
        else:
            logger.info("Found {} existing files in the watched directories.".format(len(to_import)))
            for f in to_import:
                check_file(f)

    
    

def check_file(path: str):
    """
    Check the file path to see if it contains an Excel file.
    This check is done strictly by the file extension
    """
    logger.info("Checking file '{0}'".format(path))
    isDb = list(filter(path.endswith, IMPORT_FILE_TYPES))
    if isDb:
        #begin import
        logger.info("File {} looks like an Excel file, beginning import".format(path))
        import_file(path)
    else:
        logger.info("File {} does not look like an Excel file, skipping...".format(path))
    
    
def test_permissions(file):
    ''' Returns true if the file can be read and modified, false otherwise '''
    try:
        path = os.path.normpath(file)
        os.rename(path, path)
        return True
    except:
        return False

    
def import_file(file: str):
    logger.info("╭╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼")
    logger.info("╽")
    logger.info("╽ Beginning import of {}".format(file))
    logger.info("╽")
    logger.info("╰╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼")

    #should the file be moved to the archive directory after the import?
    moveFile = 'Skipped'

    try:
        # Before we begin the import check that we have read and write access to the file.
        #+ If the file is open in another program we will not be able to archive the file later
        writable = test_permissions(file)
        if not writable:
            raise ImportException("Could not establish read and write access to file. Skipping import")

        importer = Importer(opts.database, file)
        importer.begin_import()
        moveFile = 'Archived'
    except ImportException as e:
        logger.warning("Unable to import data from file: {0}".format(e.message))
        logger.warning("This is a recoverable error. The import will be attempted again later")
        logging.exception(e, exc_info=True)
    except FileFormatException as e:
        logger.error("Unable to import data from file: {0}".format(e.message))
        logger.error("This is a non-recoverable error. The import will not be attempted again later.")
        moveFile = 'Error'
        logging.exception(e, exc_info=True)
    except pyodbc.Error as e:
        # Check if this a known SQL error code. Some SQL errors we may be able to recover from. Others will
        #+ require the file to be moved to the errors directory
        logger.error("Unable to import data from file: {0}".format(str(e)))
        errCode = e.args[0]
        connErrPat = re.compile("^0800[12347]$") # Pattern to match a connection error code

        # Check for the unrecoverable errors first
        if (errCode == '23000'):                       # Key constraint violation. Trying to import data twice
            logger.error("Unable to import data due to an unrecoverable SQL error.")
            moveFile = 'Error'
        elif errCode == 'HY000':                       # General error. Could occur due to missing field value, because the file was locked, or because the Access database is corrupted.
            logger.error("Since this was a database error this may be recoverable.  The import will be attempted again later")
        logging.exception(e, exc_info=True)
    except Exception as e:
        logger.error("Unable to import data from file for an unknown reason")
        logger.error("This is a non-recoverable error. The import will not be attempted again later.")
        moveFile = 'Error'
        logging.exception(e, exc_info=True)
    finally:
        if moveFile != 'Skipped':
            try:
                archive_directory = None
                if moveFile == 'Archived':
                    archive_directory = opts.archive
                elif moveFile == 'Error':
                    archive_directory = opts.errors

                postfix = datetime.today().strftime('%Y-%m-%d-%H-%M-%S-%f')

                newName = "{0}.{1}".format(
                    os.path.basename(file),
                    postfix
                )

                newPath = os.path.join(archive_directory, newName)
                logger.info("Archiving file as '{0}'.".format(newPath))
                shutil.move(file, newPath)
            except Exception as e:
                logger.error("Unable to move file '{0}'->'{1}: {2}".format(file, newPath, e))
                logger.exception(e)

    
    logger.info("╭╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼")
    logger.info("╽")
    logger.info("╽ Finished import of {}".format(file))
    logger.info("╽ Status: {}".format(moveFile))
    logger.info("╽")
    logger.info("╰╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼╼")     
        
            

class Handler(events.FileSystemEventHandler):
    """
    The handler is only interested in file creation events.  The Watcher will decide which
    of these files triggers an import
    """
    def on_created(self, event):
        if not event.is_directory:    
            src = event.src_path
            logger.info("File created: '{0}'".format(src))
            check_file(event.src_path)
        


def log_to_file():
    # Rotate the log files every day, keep 7 backups
    fh = TimedRotatingFileHandler("watcher.log", when='d', interval=1, backupCount=7, encoding='utf-8')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(fh)

def log_access_driver_error():
    #check if we are running in the 32 or 64 bit version of Python
    bitVer = 8 * struct.calcsize("P")
    logger.info("╔═════════════════════════════════════════════════════════════════════╗")
    logger.info("║                                                                     ║")
    logger.info("║ Download the latest ACE driver                                      ║")
    logger.info("║ https://www.microsoft.com/en-US/download/details.aspx?id=13255      ║")
    logger.info("║                                                                     ║")
    logger.info("║ You will need to download the {0}bit version of the driver to work   ║".format(bitVer))
    logger.info("║ with this version of Python                                         ║")
    logger.info("║                                                                     ║")
    logger.info("╚═════════════════════════════════════════════════════════════════════╝")

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(Watcher)

