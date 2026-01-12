"""
Base Model for data parsing and representation.
"""


class Model:
    raw: str


class Table:
    headers: list[str]
    content: list[list[object]]

    def get_data(self) -> list[list[object]]:
        return self.content

    def get_headers(self) -> list[str]:
        return self.headers

    def get_with_headers(self) -> list[dict[str, object]]:
        data_with_headers: list[dict[str, object]] = []
        for row in self.content:
            row_dict: dict[str, object] = {}
            for idx, header in enumerate(self.headers):
                row_dict[header] = row[idx] if idx < len(row) else ""
            data_with_headers.append(row_dict)
        return data_with_headers


class Section:

    title: str
    content: str | None
    table: Table | None

    subsections: list["Section"]



class ParsedFile:
    title: str
    sections: list[Section]

    raw: str
