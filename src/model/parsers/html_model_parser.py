from io import BytesIO
from src.errors import Error, ParsingError
from src.model.model import File, Model, ParseResult, ParsedFile, Parser, Section, WritableFile, Writer
from bs4 import BeautifulSoup

from src.model.models.PIA_model import PIA_Model
from src.model.writers.excel_writer import ExcelWriter

class HTMLModelParser(Parser):


    def _parse_main_title(self, soup: BeautifulSoup) -> str:
        """Extract the main title from the text."""

        return ""

    def _parse_subsection(self, soup: BeautifulSoup) -> Section:
        """Extract the main title from the text."""
        return Section()

    def _parse_sections(self, soup: BeautifulSoup) -> list[Section]:
        """Extract sections from the textlines."""

        sections: list[Section] = []
        return sections
        

    def parse(self, file: File) -> ParseResult | ParsingError:
        """Parse the excel data and return a ParsedFile or an Error."""

        soup: BeautifulSoup = BeautifulSoup(file.data, 'html.parser')


        parsed_file = ParsedFile()

        parsed_file.title = self._parse_main_title(soup)
        parsed_file.sections = self._parse_sections(soup)
        
        writable_file: WritableFile | None
        model: Model | None = None
        writer: Writer | None = None
        template: str | None = None

        
        if True: # Placeholder condition for selecting model. TODO: implement actual logic
            model = PIA_Model()
            parsed_file = model.process(file.data, file.filename, "output/pia_master.xlsx")
            template = "templates/PIA_Modelo de datos_template.xlsx"
            writable_file = model.create_writable_file(parsed_file)
            writer = ExcelWriter()
            
        else:
            return ParsingError("Model not implemented.")

        # parsed_file.title = self._parse_main_title(sheet)
        # parsed_file.sections = self._parse_sections(workbook)

        result = ParseResult(writable_file, model, writer, template)

        return result

        
        
        result = ParseResult(parsed_file)

        return result
