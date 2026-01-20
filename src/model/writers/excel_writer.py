from typing import TypeVar
from src.model.model import Model, WritableEntry, WritableFile, WritableParameters, Writer, WritingTarget
import openpyxl
from openpyxl.cell import Cell
from src.errors import Error, FileWriteError


class ExcelWriter(Writer):

    class WritableExcelFile(WritableFile):

        class WritableExcelParameters(WritableParameters):
            cell: Cell

        class WritableExcelEntry(WritableEntry):
            content: Cell
            parameters: "ExcelWriter.WritableExcelFile.WritableExcelParameters"

            def __init__(self, content: Cell, parameters: "ExcelWriter.WritableExcelFile.WritableExcelParameters"):
                self.content = content
                self.parameters = parameters

        entries: list[WritableExcelEntry] = []

        def __init__(self, target: str, entries: list[WritableExcelEntry]):
            self.entries = entries
            self.target = target

    def write(self, writable_file: WritableExcelFile) -> None | Error:
        """Write the parsed file to excel format and return as bytes."""
        # Dummy implementation for demonstration

        workbook = openpyxl.Workbook()
        sheet = workbook.active

        if sheet is None:
            raise FileWriteError("Sheet is empty right after creation.")
        
        for entry in writable_file.entries:
            cell = entry.content
            sheet[cell.coordinate] = cell.value
        

        if self.target is None:
            raise FileWriteError("Writing target is None.")
            
        workbook.save(self.target)
