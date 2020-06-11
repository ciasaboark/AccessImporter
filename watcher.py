# -*- coding: utf-8 -*-

import sys, os, time, logging, shutil
from watchdog.observers.polling import PollingObserver
from watchdog import events
from pathlib import Path
from importer import Importer
from exceptions import ImportException, FileFormatException
from daemon import Daemon
from time import sleep
import struct
import pyodbc
import configargparse
import textwrap
import traceback
from pprint import pformat
from datetime import datetime
from pathlib import Path

VERSION = "0.0.1"

#The interval (in seconds) between sleep cycles. A manual check is performed on each wake
SLEEP_INTERVAL = 1 * 60 * 10

IMPORT_FILE_TYPES = [".accdb"]

logger = None
watcher = None

class Watcher(Daemon):
    def __init__(self, pidfile, args, parser):
        self.observer = PollingObserver()
        #store the list of watched directories as a set so we can avoid duplicates
        self.watched_directories = set()
        self.args = args
        self.parser = parser
        super(Watcher, self).__init__(pidfile)

    def run(self):
        logger.info("**************************************")
        logger.info("* MS Access database importer starting")
        logger.info("* Version {0}".format(VERSION))
        logger.info("**************************************")

        logger.info("Running with config:")
        logger.info(pformat(self.args, indent=1, width=80, depth=None, compact=False))

        event_handler = Handler()

        # #Convert the paths into the OS specific formatting.
        # # This allows paths to be defined in the config file as C:/import
        # # and be auto converted to C:\import
        # for index, p in enumerate(args.watch):
        #     converted_path = Path(p)
        #     args.watch[index] = converted_path
        # args.archive = Path(args.archive)
        # args.database = Path(args.database)

        #Quick test to see if we have the MS Access driver installed
        #Get a list of all installed Access DB drivers
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

        #Filter down the list of directories to watch to only the ones that exist
        for folder in self.args.watch:
            exists = os.path.exists(folder)
            isFile = os.path.isfile(folder)
            if not exists:
                logger.warning("Will not watch for new files in '{0}'. Directory does not exist.".format(folder))
            elif isFile:
                logger.warning("Will not watch for new files in '{0}'. Path is a file, not a directory.".format(folder))
            else:
                logger.debug("Adding '{0}' to the watched directories list".format(folder))
                self.watched_directories.add(folder)
                self.observer.schedule(event_handler, folder, recursive=False)
        
        #Check if the archive directory exists
        if not os.path.exists(self.args.archive) or os.path.isfile(self.args.archive):
            logger.error("The archive directory '{0}' does not exist".format(self.args.archive))
            sys.exit(1)

        #Check if the import to database exists. Defer checking whether this is a valid Access database until later
        if not os.path.exists(self.args.database) or not os.path.isfile(self.args.database):
            logger.error("The database file '{0}' does not exist".format(self.args.database))
            sys.exit(1)

        #Error out if we have no directories to watch
        if len(self.watched_directories) == 0:
            logger.debug("Could not find any directories to watch.  Exiting...")
            sys.exit(1)

        #Make sure we aren't being asked to move imported data into the same folder we are watching.
        #Use os.path.sameFile() to make sure we resolve relative paths back
        arkPath = os.path.normpath(self.args.archive)
        for path in self.watched_directories:
            realPath = os.path.normpath(path)
            if (os.path.samefile(realPath, arkPath)):
                logger.error("The archive directory '{0}' is in the list of directories to watch. This "
                "will result in an endless loop.".format(self.args.archive))
                logger.error("Choose a different archive directory or remove this directory from the list of watched directories")
                sys.exit(1)

        self.manual_import()

        #  Fire up the observer and begin the main loop
        self.observer.start()
        logger.info("Watchdog running. Monitoring new files in paths {0}".format(pformat(self.watched_directories, indent=1, width=80, depth=None, compact=False)))
        logger.info("Completed imports will be moved into '{0}'".format(self.args.archive))  

        try:
            while True:
                #sleep for 10 minute intervals before triggering a manual check for any database
                # files that need to be imported.  The on_created() callback will still fire during
                # this interval.  The manual check is needed to handle importing files that failed
                # during the on_created() callback (for example if the database is locked)
                time.sleep(SLEEP_INTERVAL)
                self.manual_import()

        except:
            self.observer.stop()
            logger.debug("Error with watchdog.  Exiting...")

        self.observer.join()

    def manual_import(self):
        #Everything looks OK.  Check for any pre-existing database files to import.  Since
        # these already exist we won't get a on_created() callback in the Handler
        logger.debug("Performing manual check for database files to import")
        to_import = []
        for d in self.args.watch:
            for file in os.listdir(d):
                full_path = os.path.join(d, file)
                if os.path.isfile(full_path):
                    to_import.append(full_path)

        if len(to_import) == 0:
            logger.debug("No files found to import")
        else:
            logger.debug("Found existing files in the watched directories.")
            for f in to_import:
                self.check_file(f)

    def check_file(self, path: str):
        """
        Check the file path to see if it contains an Access database file.  This check is done
        strictly by the file extension
        """
        logger.info("Checking file '{0}'".format(path))
        isDb = list(filter(path.endswith, IMPORT_FILE_TYPES))
        if isDb:
            #begin import
            logger.debug("File {} looks like an Access database, beginning import".format(path))
            self.import_file(path)
        else:
            logger.debug("File {} does not look like an Access database, skipping...".format(path))
    

    def import_file(self, file: str):
        importer = Importer(self.args.database, file)

        #should the file be moved to the archive directory after the import?
        moveFile = False
        try:
            importer.begin_import()
            moveFile = True
        except ImportException as e:
            logger.warning("Unable to import data from file: {0}".format(e.message))
            logger.warning("This is a recoverable error. The import will be attempted again later")
            logging.exception(e, exc_info=True)
        except FileFormatException as e:
            logger.error("Unable to import data from file: {0}".format(e.message))
            logger.error("This is a non-recoverable error. The import will not be attempted again later.")
            moveFile = True
            logging.exception(e, exc_info=True)
        except pyodbc.Error as e:
            logger.error("Unable to import data from file: {0}".format(str(e)))
            logger.error("Since this was a database error this may be recoverable.  The import will be attempted again later")
            logging.exception(e, exc_info=True)
        except Exception as e:
            logger.error("Unable to import data from file for an unknown reason")
            logger.error("This is a non-recoverable error. The import will not be attempted again later.")
            moveFile = True
            logging.exception(e, exc_info=True)
        finally:
            if moveFile:
                postfix = datetime.today().strftime('%Y-%m-%d-%H-%M-%S-%f')

                newName = "{0}.{1}".format(
                    os.path.basename(file),
                    postfix
                )
                newPath = os.path.join(self.args.archive, newName)
                logger.debug("Archiving file as '{0}'.".format(newPath))
                shutil.move(file, newPath)
        
            

class Handler(events.FileSystemEventHandler):
    """
    The handler is only interested in file creation events.  The Watcher will decide which
    of these files triggers an import
    """
    def on_created(self, event):
        if not event.is_directory:    
            src = event.src_path
            logger.debug("File created: '{0}'".format(src))
            watcher.check_file(event.src_path)
        


def log_to_file():
    fh = logging.FileHandler("watcher.log")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(fh)

def log_access_driver_error():
    #check if we are running in the 32 or 64 bit version of Python
    bitVer = 8 * struct.calcsize("P")
    logger.info("-----------------------------------------------------------------------")
    logger.info("|                                                                     |")
    logger.info("| Download the latest ACE driver                                      |")
    logger.info("| https://www.microsoft.com/en-US/download/details.aspx?id=13255      |")
    logger.info("|                                                                     |")
    logger.info("| You will need to download the {0}bit version of the driver to work   |".format(bitVer))
    logger.info("| with this version of Python                                         |")
    logger.info("|                                                                     |")
    logger.info("-----------------------------------------------------------------------")

if __name__ == '__main__':
    logger = logging.getLogger('watcher')
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    p = configargparse.ArgParser(
        description='Microsoft Access database importer',
        epilog=textwrap.dedent('''\
         Additional information:
             
         ''')
    )
    p.add_argument('-c', '--config', help='<Required> The configuration file to use.', is_config_file=True)
    p.add('-w', '--watch', required=True, help='<Required> The directory to watch for new database files. Use additional -w arguments to watch multiple directories. At least one of the watched directories must exist.', action='append')
    p.add('-a', '--archive', required=True, help='<Required> The archive directory to hold imported database files. This directory must already exist.')
    p.add('-d', '--database', required=True, help='<Required> The path of the Access database file to import into. This directory must already exist')
    p.add('action', choices=['start', 'stop', 'restart'], nargs="?",  help='<Optional> Run the program as a daemon.  If omitted the program will run in the foreground.')

    args = p.parse_args()

    watcher = Watcher('/tmp/returns-watcher.pid', args, p)
    if args.action is not None:
        log_to_file()
        if 'start' == args.action:
            watcher.start()
        elif 'stop' == args.action:
            watcher.stop()
        elif 'restart' == args.action:
            watcher.restart()
        else:
            print("Unknown command: '{}'".format(args.action))
            print("usage: {} start|stop|restart".format(args.action))
            sys.exit(1)
        sys.exit(0)
    else:
        # print("usage: %s start|stop|restart" % sys.argv[0])
        print("Running in interactive mode")
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        watcher.run()
        sys.exit(2)



