from src.model.model import WritableFile
import openpyxl

class ExcelWriter:

    class WritableExcelFile(WritableFile):

        class WritableExcelParameters(WritableFile.WritableParameters):
            excel_specific_param: str = "default_value"

        class WritableEntry(WritableFile.WritableEntry):
            # content: str
            # parameters: "ExcelWriter.WritableExcelFile.WritableExcelParameters"
            pass

        # entries: list[WritableEntry]


    def write(self, parsed_file: WritableExcelFile) -> bytes:
        """Write the parsed file to excel format and return as bytes."""
        # Dummy implementation for demonstration

        workbook = openpyxl.Workbook()
        sheet = workbook.active

        # for entry in parsed_file.entries:
        #     sheet.append([entry.content])

        return b"Excel data"