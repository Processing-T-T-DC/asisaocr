"""
Base Model for data parsing and representation.
"""


from collections.abc import Mapping
from typing import Literal
from abc import ABC



type HeaderLevel = Literal["title", "H1", "H2", "H3", "H4", "p"]

class Model(ABC):
    """ Abstract base class for models """
 
    raw: str

    FONT_SIZES_MAPPING: Mapping[int, HeaderLevel]


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
    table: list[Table]

    subsections: list["Section"]



class ParsedFile:
    title: str
    sections: list[Section]

    raw: str




class WritableFile:
    class WritableParameters:
        pass

    class WritableEntry:
        content: str
        parameters: "WritableFile.WritableParameters"


    entries: list[WritableEntry]