"""
Base Model for data parsing and representation.
"""
class Model:
    
    raw: str
    
    def parse(self, file_text: str) -> dict:
        # Dummy parse method for demonstration
        return {"parsed_data": file_text.upper()}
    
    def _parse_inner_table(self, table: list[list[str]]) -> list[dict[str, str]]:
        # Dummy inner table parsing for demonstration
        parsed_table = []
        for row in table:
            parsed_row = {f"col_{i}": cell for i, cell in enumerate(row)}
            parsed_table.append(parsed_row)
        return parsed_table



class Table:
    
    def get_data( self) -> list[object] :
        pass

    def get_headers(self ) -> list[str]:
        pass

    def get_with_headers(self) -> list[dict[str, str]]:
        pass


class Section:
    title: str
    content: str

    subsections: list['Section']

    table: Table


class ParsedFile:
    title: str
    sections: list[Section]

    raw: str

