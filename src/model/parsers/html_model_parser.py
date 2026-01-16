from io import BytesIO
from src.errors import Error
from src.model.model import ParsedFile, Section
from bs4 import BeautifulSoup

class HTMLModelParser:


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
        

    def parse(self, data: bytes) -> ParsedFile | Error:
        """Parse the excel data and return a ParsedFile or an Error."""

        soup: BeautifulSoup = BeautifulSoup(data, 'html.parser')

        

        parsed_file = ParsedFile()

        parsed_file.title = self._parse_main_title(soup)
        parsed_file.sections = self._parse_sections(soup)

        return parsed_file
