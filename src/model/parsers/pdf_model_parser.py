from collections.abc import Mapping
from io import BytesIO
from typing import Any, TypedDict, cast
import pdfplumber
from src.errors import Error
from src.model.model import HeaderLevel, ParsedFile, Section
import re


class PDFModelParser:

    font_sizes_to_header_type: Mapping[int, HeaderLevel] # H1, H2, H3, H4, p

    def __init__(self, font_sizes: Mapping[int, HeaderLevel]) -> None:
        self.font_sizes_to_header_type = font_sizes

    class TextLineObject(TypedDict):
        text: str
        x0: float
        top: float
        x1: float
        bottom: float
        chars: list[dict[str, Any]]
        page_number: int

    def _text_size_to_header(self, size: float) -> str | ValueError:
        """Map a font size to a heading level."""

        header_size = self.font_sizes_to_header_type.get(int(size))

        if not header_size:
            return 'p'
            # raise ValueError(f"Unsupported font size for header: {size}")

        return header_size

    def _parse_main_title(self, textlines: list[TextLineObject]) -> str:
        """Extract the main title from the text."""
        # Dummy implementation for demonstration
        if textlines and len(textlines) > 1:
            first_line = textlines[0]
            first_line_char_size = first_line["chars"][0]["size"]

            second_line = textlines[1]
            second_line_char_size = second_line["chars"][0]["size"]

            if (
                first_line_char_size
                and second_line_char_size
                and first_line_char_size == second_line_char_size
            ):
                return first_line.get("text") + " " + second_line.get("text")
            else:
                return first_line.get("text")

        return "Untitled"
    

    def _parse_subsection(self, textlines: list[TextLineObject]) -> Section:
        """Extract the main title from the text."""
        # Dummy implementation for demonstration
        section = Section()
        section.title = "Subsection"
        section.content = "\n".join([line["text"] for line in textlines])
        section.table = []
        section.subsections = []
        return section
    

    def _concatenate_close_lines_inplace(self, lines: list[TextLineObject]) -> None:
        """
        If textlineobject has a separation of less than 15 points of height,
        concat it to the end of the previous one.
        It must take into account the text size.
        """

        # iterating backwards to take into account element deletion 
        for index in range(len(lines) - 1, 0, -1):
            line = lines[index]
            prev = lines[index - 1]

            gap = (prev["bottom"] - line["top"]) - line["chars"][0]["height"]

            if line["page_number"] == prev["page_number"] and abs(gap) < 15:
                prev["text"] += " " + line["text"]
                lines.pop(index)

    HEADER_RE = re.compile(r"^(\d+(?:\.\d+)*)\s+.+")
    FIRST_HEADER_RE = re.compile(r"^(1)\s+.+")

    def _get_lines_per_header_section(
        self,
        textlines: list[TextLineObject],  # TextLineObject is a dict from pdfplumber
    ) -> dict[str, list[TextLineObject]]:

        sections: dict[str, list[PDFModelParser.TextLineObject]] = {}

        current_header = None
        current_level = None

        first_header_found = False

        for line in textlines:
            text = line.get("text", "").strip()
            if not text:
                # Empty lines just get appended if we're inside a section
                if current_header is not None:
                    sections[current_header].append(line)
                continue

            match = self.HEADER_RE.match(text) if first_header_found else self.FIRST_HEADER_RE.match(text)

            if first_header_found is False and match:
                first_header_found = True

            if match:
                header_number = match.group(1)
                level = (header_number.count(".") or 0) + 1
                current_header = text

                # Start a new section if:
                # - we have no current section
                # - or this header is same or higher level than the current one
                if current_level is None or level <= current_level:
                    current_header = text
                    current_level = level
                    sections[current_header] = []
                    continue

            # Normal content line
            if current_header is not None:
                sections[current_header].append(line)

        return sections


    def _parse_sections(self, textlines: list[TextLineObject]) -> list[Section]:
        """Extract the main title from the text."""
        # Dummy implementation for demonstration

        # map each textline to its size header
        self._concatenate_close_lines_inplace(textlines)
        lines_per_header_section = self._get_lines_per_header_section(textlines)

        


        grouped_textlines: dict[str, list[PDFModelParser.TextLineObject]] = {}

        # for line in textlines:
        #     header_level = self._text_size_to_header(line["chars"][0]["size"])
        #     if isinstance(header_level, ValueError):
        #         raise header_level

        #     if header_level not in grouped_textlines:
        #         grouped_textlines[header_level] = []
        #     grouped_textlines[header_level].append(line)


        # starting from the largest header to the smallest,
        # check if there are textlines that contain the 
        # corresponding header level (e.g., "1 - ..." for H1. "1.1.1.2 - ..." for H4)
        # for line in textlines:
            
        sections: list[Section] = []


        return sections
        

    def parse(self, data: bytes) -> ParsedFile | Error:
        """Parse the PDF data and return a ParsedFile or an Error."""

        parsed_file = ParsedFile()

        with pdfplumber.open(BytesIO(data)) as pdf:
            text_info: list[PDFModelParser.TextLineObject] = []
            # Implement parsing logic here
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"

                text_lines = page.extract_text_lines()
                for line in text_lines:
                    line["page_number"] = page.page_number

                text_info += cast(
                    list[PDFModelParser.TextLineObject], text_lines
                )


            parsed_file.raw = text  # Example: extract text from all pages

            # Get unique font sizes, sorted largest first, saving the text as well
            sizes = sorted(
                {(line["text"], round(line["chars"][0]["size"], 3)) for line in text_info},
                reverse=True,
                key=lambda x: x[1],
            )

            size_to_level = {
                size: f"H{i + 1}"
                for i, size in enumerate(sizes[:4])  # limit depth
            }

            print(size_to_level)

            parsed_file.title = self._parse_main_title(text_info)
            # parsed_file.title = text.index()
            parsed_file.sections = self._parse_sections(text_info)

            return parsed_file
