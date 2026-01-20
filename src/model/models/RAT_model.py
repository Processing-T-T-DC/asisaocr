
from src.model.model import Model, ParsedFile, WritableFile
from src.model.writers.excel_writer import ExcelWriter

class RAT_Model(Model):

    FONT_SIZES_MAPPING = {
        18: "title", # Only title for both pdfs
        14: "H1", # 
        12: "H2", # 
        11: "H4", # 
    }


    def create_excel_writable_file(self, parsed_file: ParsedFile) -> WritableFile:
        """Create a writable file from the parsed file."""
        entries = []
        target = ""
        writable_file = ExcelWriter.WritableExcelFile(target, entries)
        # for section in parsed_file.sections:
        #     entry = ExcelWriter.WritableExcelFile.WritableExcelEntry()
        #     entry.content = section.content if section.content else ""
        #     # entry.parameters = {}
        #     writable_file.entries.append(entry)
        return writable_file

