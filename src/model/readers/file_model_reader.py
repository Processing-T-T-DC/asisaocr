
from pathlib import Path
from src.errors import FileReadError, InvalidFileError
from src.model.model import File, Parser
from src.model.parsers.excel_model_parser import ExcelModelParser
from src.model.parsers.html_model_parser import HTMLModelParser
from src.model.parsers.pdf_model_parser import PDFModelParser


class FileReader:

    type ReadingTarget = str

    target: ReadingTarget
    
    def set_target(self, target: ReadingTarget) -> None:
        """Set the target for reading."""
        self.target = target

    def read(self) -> tuple[Parser, File] | FileReadError:
        """Read the target model and return its content or an error."""
        try:

            file = Path(self.target)
            extension = file.suffix.lower()
            parser: Parser

            if extension in {".xls", ".xlsx", ".xlsm", ".xlsb"}:
                if file.name.startswith("~$"):
                    return FileReadError(message="WARNING: This is an excel lock file. Skipping.")
                
                parser = ExcelModelParser()
            elif extension == ".pdf":
                parser = PDFModelParser({})
            elif extension in {".html", ".htm"}:
                parser = HTMLModelParser()
            else:
                return FileReadError(message="Type of file not known.")

            with open(self.target, "rb") as file:
                return (parser, File(file.read(), self.target))
            
        except Exception as e:
            return FileReadError(message=str(e))

