"""
Base Model for data parsing and representation.
"""


from abc import ABC, abstractmethod
from collections.abc import Mapping
from typing import Literal

from src.errors import Error, ParsingError

type HeaderLevel = Literal["title", "H1", "H2", "H3", "H4", "p"]

class WritableParameters:
    pass

class WritableContent:
    pass

class WritingTarget:
    pass

class WritableEntry[P: WritableParameters, C: WritableContent]:
    content: C
    parameters: P

class WritableFile[T: WritableFile, E: WritableEntry](ABC):
    target: T
    entries: list[E]

    @abstractmethod
    def __init__(self, target: T, entries: list[E]):
        self.target = target
        self.entries = entries

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

class File:
    data: bytes
    filename: str

    def __init__(self, data: bytes, filename: str):
        self.data = data
        self.filename = filename

type FieldType = Literal['text', 'yes_no', 'fixed_number_range', 'treatment_duration', 'geographic_extension']
class Field:
    name: str
    coords: str
    data_type: FieldType
    
    is_required: bool

    def __init__(self, name: str, coords: str, data_type: FieldType, is_required: bool):
        self.name = name
        self.coords = coords
        self.data_type = data_type
        self.is_required = is_required
class Model(ABC):
    """ Abstract base class for models """
 
    raw: str

    FONT_SIZES_MAPPING: Mapping[int, HeaderLevel]    
    
    @property
    @abstractmethod
    def fields(self) -> list[Field]:
        pass

    @abstractmethod
    def create_writable_file(self, parsed_file: ParsedFile) -> WritableFile:
        """ Create a writable file from the parsed file. """

        pass


class Writer[TARGET: WritingTarget, WRITABLE_FILE: WritableFile](ABC):

    def set_target(self, value: TARGET):
        self.target = value
    
    @abstractmethod
    def write(self, writable_file: WRITABLE_FILE, template: str | None = None) -> None | Error:
        pass

class ParseResult:

    file: ParsedFile | WritableFile
    
    model: Model | None
    writer: Writer | None 
    """It can be None, because we might know what writer to use after parsing if the model already processed it (WritableFile is returned), 
    but it can also be unknown until the model processes it.
    This way, we can skip processing the data into a ParsedFile if the known writing format is already the same as the parsing one.
    """

    template: str | None

    def __init__(self, file, model: Model | None = None, writer: Writer | None = None, template: str | None = None):
        self.file = file
        self.model = model
        self.writer = writer
        self.template = template

class Parser(ABC):
    
    @abstractmethod
    def parse(self, file: File) -> ParseResult | ParsingError:
        pass

