from collections.abc import Mapping
from io import BytesIO
from typing import Any, TypedDict, cast
import pdfplumber
from src.errors import Error, ParsingError
from src.model.model import HeaderLevel, ParsedFile, Section, WritableFile
import re
import openpyxl


class ExcelModelParser:


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
        

    def parse(self, data: bytes) -> WritableFile | Error:
        """Parse the excel data and return a ParsedFile or an Error."""
        
        workbook = openpyxl.load_workbook(filename=BytesIO(data))
        sheet = workbook.active


        # copy first column third row to the end of row and paste it two rows below.
        value_list = []
        range_start = 3

        if sheet is None:
            return ParsingError("Sheet is None")

        for row in sheet.iter_rows(min_row=range_start, max_row=sheet.max_row, min_col=1, max_col=1):
            cell_value = row[0].value
            
            if cell_value is not None:
                value_list.append(cell_value)

            if row[0].row is not None:
                sheet.cell(row=row[0].row, column=sheet.max_column + 1, value=cell_value)

        writable_file = WritableFile()

        # parsed_file.title = self._parse_main_title(sheet)
        # parsed_file.sections = self._parse_sections(workbook)

        return writable_file
