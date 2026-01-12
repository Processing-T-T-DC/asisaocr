from io import BytesIO
from typing import Any, TypedDict, cast
import pdfplumber
from src.errors import Error
from src.model.model import ParsedFile, Section


class PDFModelParser:
    class TextLineObject(TypedDict):
        text: str
        x0: float
        top: float
        x1: float
        bottom: float
        chars: list[dict[str, Any]]

    def _text_size_to_header(self, size: float) -> str | ValueError:
        """Map a font size to a heading level."""
        size_to_level = {
            18: "H1",
            14: "H2",
            12: "H3",
            11: "H4",
        }

        header_size = size_to_level.get(int(size))

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
        section.table = None
        section.subsections = []
        return section
    

    def _concatenate_close_lines_inplace(self, grouped_textlines: dict[str, list[TextLineObject]]) -> None:
        """
        If textlineobject has a separation of less than 5 points of height,
        concat it to the end of the previous one.
        It must take into account the text size.
        """

        for _, lines in grouped_textlines.items():

            # iterating backwards to take into account element deletion 
            for index in range(len(lines) - 1, 0, -1):
                line = lines[index]
                prev = lines[index - 1]

                gap = (prev["bottom"] - line["top"]) - line["chars"][0]["height"]

                if gap < 5:
                    prev["text"] += " " + line["text"]
                    lines.pop(index)



    def _parse_sections(self, textlines: list[TextLineObject]) -> list[Section]:
        """Extract the main title from the text."""
        # Dummy implementation for demonstration

        # map each textline to its size header

        grouped_textlines: dict[str, list[PDFModelParser.TextLineObject]] = {}

        for line in textlines:
            header_level = self._text_size_to_header(line["chars"][0]["size"])
            if isinstance(header_level, ValueError):
                raise header_level

            if header_level not in grouped_textlines:
                grouped_textlines[header_level] = []
            grouped_textlines[header_level].append(line)


        self._concatenate_close_lines_inplace(grouped_textlines)


        sections: list[Section] = []

        # for extracted_lines in grouped_textlines.get('H2', []):
        #     extracted_lines
        #     section = Section()
        #     # section.level = level
        #     section.content = "\n".join([line["text"] for line in extracted_lines])
        #     sections.append(section)

        return sections
            


            # if not current_section or current_section["level"] != header_level:
            #     if current_section:
            #         sections.append(current_section)
            #     current_section = {
            #         "level": header_level,
            #         "content": line["text"],
            #     }
            # else:
                # current_section["content"] += "\n" + line["text"]

        # if current_section:
        #     sections.append(current_section)

        # return sections

    def parse(self, data: bytes) -> ParsedFile | Error:
        """Parse the PDF data and return a ParsedFile or an Error."""

        parsed_file = ParsedFile()

        with pdfplumber.open(BytesIO(data)) as pdf:
            text_info: list[PDFModelParser.TextLineObject] = []
            # Implement parsing logic here
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
                text_info += cast(
                    list[PDFModelParser.TextLineObject], page.extract_text_lines()
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
