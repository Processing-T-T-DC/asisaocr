
from src.errors import FileReadError


class FileReader:

    type ReadingTarget = str

    target: ReadingTarget
    
    def set_target(self, target: ReadingTarget) -> None:
        """Set the target for reading."""
        self.target = target

    def read(self) -> bytes | FileReadError:
        """Read the target model and return its content or an error."""
        try:
            # Simulate reading and parsing logic
            with open(self.target, "rb") as file:

                return file.read()
            
        except Exception as e:
            return FileReadError(message=str(e))

