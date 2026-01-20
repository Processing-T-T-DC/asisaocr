from collections.abc import Mapping
from io import BytesIO
from typing import Any, TypedDict, cast
import pdfplumber
from src.errors import Error, ParsingError
from src.model.model import HeaderLevel, Model, ParseResult, ParsedFile, Parser, Section, WritableFile, Writer
import re
import openpyxl

from src.model.models.AARR_model import AARR_Model
from src.model.writers.excel_writer import ExcelWriter

class ExcelModelParser(Parser):


    def _parse_main_title(self, workbook) -> str:
        """Extract the main title from the text."""

        return ""
        pass

    def _parse_subsection(self, workbook) -> Section:
        """Extract the main title from the text."""

        return Section()
        pass

    def _parse_sections(self, workbook) -> list[Section]:
        """Extract sections from the textlines."""

        sections: list[Section] = []

        for sheet in workbook.worksheets:
            section = Section()
            section.title = sheet.title
            section.content = ""
            sections.append(section)
            section.subsections = []

        return sections
        

    def parse(self, data: bytes) -> ParseResult | ParsingError:
        """Parse the excel data and return a ParsedFile or an Error."""
        
        workbook = openpyxl.load_workbook(filename=BytesIO(data))
        sheet = workbook.active

        if sheet is None:
            return ParsingError("Sheet is None")
        
        writable_file: WritableFile | None
        model: Model | None = None
        writer: Writer | None = None
        
        if sheet.title == "Evaluacion objetiva":
            name = "REPLACE_WITH_NAME"
            model = AARR_Model()
            model.process(workbook,  f"output/{name}_evaluacion_objetiva.xlsx")
            writable_file = model.create_writable_file(ParsedFile())
            writer = ExcelWriter()
        else:
            raise RuntimeError("Not Implemented")

        # parsed_file.title = self._parse_main_title(sheet)
        # parsed_file.sections = self._parse_sections(workbook)

        result = ParseResult(writable_file, model, writer)

        return result
