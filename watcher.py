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
from pprint import pformat

VERSION = "0.0.1"

IMPORT_FILE_TYPES = [".accdb"]

logger = None


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

        # pp = pprint.PrettyPrinter(indent=4)
        logger.info("Running with config:")
        logger.info(pformat(self.args, indent=1, width=80, depth=None, compact=False))

        event_handler = Handler()

        #Quick test to see if we have the MS Access driver installed
        #Get a list of all installed Access DB drivers
        drivers = [x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]
        if len(drivers) == 0:
            logger.error('Unable to find any Access database drivers installed')
            log_access_driver_error()
            raise Exception
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

        #Everything looks OK.  Fire up the observer and begin the main loop
        self.observer.start()
        logger.info("Watchdog running. Monitoring new files in paths {0}".format(pformat(self.watched_directories, indent=1, width=80, depth=None, compact=False)))
        logger.info("Completed imports will be moved into '{0}'".format(self.args.archive))  

        try:
            while True:
                time.sleep(5)

        except:
            self.observer.stop()
            logger.debug("Error with watchdog.  Exiting...")

        self.observer.join()

    @staticmethod
    def check_file(path: str):
        """
        Check the file path to see if it contains an Access database file.  This check is done
        strictly by the file extension
        """
        isDb = list(filter(path.endswith, IMPORT_FILE_TYPES))
        if isDb:
            #begin import
            logger.debug("File {} looks like an Access database, beginning import".format(path))
            Watcher.import_file(path)
        else:
            logger.debug("File {} does not look like an Access database, skipping...".format(path))
    

    @staticmethod
    def import_file(file: str, archive_dir: str):
        importer = Importer(file)

        #should the file be moved to the archive directory after the import?
        moveFile = False
        try:
            importer.begin_import()
            moveFile = True
        except ImportException as e:
            logger.warning("Unable to import data from file: {0}".format(e.message))
            logger.warning("This is a recoverable error. The import will be attempted again later")
            logging.error(traceback.format_exc())
        except FileFormatException as e:
            logger.error("Unable to import data from file: {0}".format(e.message))
            logger.error("This is a non-recoverable error. The import will not be attempted again later.")
            logging.error(traceback.format_exc())
        except Exception as e:
            logger.error("Unable to import data from file for an unknown reason")
            logger.error("This is a non-recoverable error. The import will not be attempted again later.")
            logging.error(traceback.format_exc())
        finally:
            if moveFile:
                logger.debug("Moving file {0} to archive")
                shutil.move(file, archive_dir)
        
            

class Handler(events.FileSystemEventHandler):
    """
    The handler is only interested in file creation events.  The Watcher will decide which
    of these files triggers an import
    """
    def on_created(self, event):
        if not event.is_directory:    
            src = event.src_path
            logger.debug("File created: '{0}'".format(src))
            Watcher.check_file(event.src_path)
        


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
    p.add_argument('-c', '--config', help='<Required> The configuration file to use.', required=True, is_config_file=True)
    p.add('-w', '--watch', required=True, help='<Required> The directory to watch for new database files. Use additional -w arguments to watch multiple directories.', action='append')
    p.add('-a', '--archive', required=True, help='<Required> The archive directory to hold imported database files')
    p.add('-d', '--database', required=True, help='<Required> The path of the Access database file to import into')
    

    args = p.parse_args()

    daemon = Watcher('/tmp/returns-watcher.pid', args, p)
    if len(sys.argv) == 2:
        log_to_file()
        if 'start' == sys.argv[1]:
            daemon.start()
        elif 'stop' == sys.argv[1]:
            daemon.stop()
        elif 'restart' == sys.argv[1]:
            daemon.restart()
        else:
            print("Unknown command: '{}'".format(sys.argv[1]))
            print("usage: {} start|stop|restart".format(sys.argv[0]))
            sys.exit(1)
        sys.exit(0)
    else:
        # print("usage: %s start|stop|restart" % sys.argv[0])
        print("Running in interactive mode")
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        daemon.run()
        sys.exit(2)



