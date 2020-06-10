import sys, os, time, logging
from watchdog.observers.polling import PollingObserver
from watchdog import events
from pathlib import Path
from importer import Importer
from daemon import Daemon
from time import sleep



DIRECTORIES_TO_WATCH = ["/tmp/", "/foobar", "/tmp/foobar"]
ARCHIVE_DIR = "/tmp/old/"
IMPORT_FILE_TYPES = [".accdb"]

logger = None


class Watcher(Daemon):
    watched_directories = []

    def __init__(self, pidfile):
        self.observer = PollingObserver()
        self.watched_directoires = []
        super(Watcher, self).__init__(pidfile)

    def run(self):
        event_handler = Handler()

        for folder in DIRECTORIES_TO_WATCH:
            exists = os.path.exists(folder)
            isFile = os.path.isfile(folder)
            if not exists:
                logger.warning("Skipping {0}. Directory does not exist.".format(folder))
            elif isFile:
                logger.warning("Skipping {0}. This is a file, not a directory.".format(folder))
            else:
                logger.debug("Adding {0} to the watched directories list".format(folder))
                self.watched_directories.append(folder)
                self.observer.schedule(event_handler, folder, recursive=False)

        if (len(self.watched_directories)) > 0:
            self.observer.start()
            folders = "\n"
            for f in self.watched_directories:
                folders += "\t● '{0}'\n".format(f)
            logger.info("Watching running.  Monitoring changes on paths {}".format(folders))
            logger.info("Completed imports will be moved into '{0}'".format(ARCHIVE_DIR))
        else:
            logger.debug("Could not find any directories to watch.  Exiting...")
            raise Exception

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
    def import_file(file: str):
        importer = Importer(file)
        try:
            importer.begin_import()
        except Exception as e:
            logger.error("Caught exception importing file: {}".format(e))
            

class Handler(events.FileSystemEventHandler):
    """
    The handler is only interested in file creation events.  The Watcher will decide which
    of these files triggers an import
    """
    def on_created(self, event):
        if not event.is_directory:    
            src = event.src_path
            logger.debug("File created: '{0}'".format(src)
            Watcher.check_file(event.src_path)
        


def log_to_file():
    fh = logging.FileHandler("watcher.log")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(fh)

if __name__ == '__main__':
    logger = logging.getLogger('watcher')
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    daemon = Watcher('/tmp/returns-watcher.pid')
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


