class ImportException(Exception):
    """Raised when a file could not be imported, but the import process may be tried again later"""
    def __init__(self, message=""):
        self.message = message

class FileFormatException(Exception):
    """Raised when the import file is the wrong format"""
    
    def __init__(self, message=""):
        self.message = message