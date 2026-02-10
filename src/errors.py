
class Error(Exception):
    """Base class for exceptions in this module."""
    pass


class ParsingError(Error):
    """Exception raised for errors in the parsing process."""

    def __init__(self, message: str):
        self.message = message
        super().__init__(self.message)


class FileReadError(Error):
    """Exception raised for errors in reading files."""

    def __init__(self, message: str):
        self.message = message
        super().__init__(self.message)


class InvalidFileError(Error):
    """Exception raised for invalid files at reading time."""
    def __init__(self, message: str):
        self.message = message
        super().__init__(self.message)

class FileWriteError(Error):
    """Exception raised for errors in writing files."""

    def __init__(self, message: str):
        self.message = message
        super().__init__(self.message)

class ValidationError(Error):
    """Exception raised for validation errors."""

    def __init__(self, message: str):
        self.message = message
        super().__init__(self.message)
