from collections.abc import Mapping
from io import BytesIO
from typing import Any, TypedDict, cast
import pdfplumber
from src.errors import Error, ParsingError
from src.model.model import File, HeaderLevel, Model, ParseResult, ParsedFile, Parser, Section, WritableFile, Writer
import re
import openpyxl
from openpyxl import Workbook

from src.model.models.AARR_model import AARR_Model
from src.model.models.RAT_model import RAT_Model
from src.model.writers.excel_writer import ExcelWriter

class ExcelModelParser(Parser):


    def _parse_main_title(self, workbook: Workbook) -> str:
        """Extract the main title from the text."""

        return ""
        pass

    def _parse_subsection(self, workbook: Workbook) -> Section:
        """Extract the main title from the text."""

        return Section()
        pass

    def _parse_sections(self, workbook: Workbook) -> list[Section]:
        """Extract sections from the textlines."""

        sections: list[Section] = []

        for sheet in workbook.worksheets:
            section = Section()
            section.title = sheet.title
            section.content = ""
            sections.append(section)
            section.subsections = []

        return sections
        

    def parse(self, file: File) -> ParseResult | ParsingError:
        """Parse the excel data and return a ParsedFile or an Error."""
        
        workbook = openpyxl.load_workbook(BytesIO(file.data))
        sheet = workbook.worksheets[0]

        if sheet is None:
            return ParsingError("Sheet is None")
        
        writable_file: WritableFile | None
        model: Model | None = None
        writer: Writer | None = None
        template: str | None = None

        if sheet.title == "Evaluacion objetiva":
            model = AARR_Model()
            model.process(workbook,  "output/evaluaciones_objetivas_master.xlsx")
            template = "templates/AARR_Modelo de datos_template.xlsx"
            writable_file = model.create_writable_file(ParsedFile())
            writer = ExcelWriter()
        
        elif sheet["A1"].value == "Informe RAT":
            model = RAT_Model()
            model.process(workbook, "output/RAT_master.xlsx")
            template = "templates/RAT_Modelo de datos_template.xlsx"
            writable_file = model.create_writable_file(ParsedFile())
            writer = ExcelWriter()
            
        else:
            return ParsingError("Model not implemented.")

        # parsed_file.title = self._parse_main_title(sheet)
        # parsed_file.sections = self._parse_sections(workbook)

        result = ParseResult(writable_file, model, writer, template)

        return result
