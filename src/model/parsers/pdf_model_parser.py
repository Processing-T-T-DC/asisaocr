

from io import BytesIO
from typing import Any, TypedDict, cast
import pdfplumber
from src.errors import Error
from src.model.model import ParsedFile

class PDFModelParser:

    class TextLineObject(TypedDict):
        text: str
        x0: float
        top: float
        x1: float
        bottom: float
        chars: list[dict[str, Any]]


    def _parse_main_title(self, textlines: list[TextLineObject]) -> str:
        """Extract the main title from the text."""
        # Dummy implementation for demonstration
        if textlines and len(textlines) > 1:

            first_line = textlines[0]
            first_line_char_size = first_line["chars"][0]["size"]

            second_line = textlines[1]
            second_line_char_size = second_line["chars"][0]["size"]

            
            if first_line_char_size and second_line_char_size \
              and first_line_char_size == second_line_char_size:
                
                return first_line.get("text") + " " + second_line.get("text")
            else:
                return first_line.get("text")
            
        return "Untitled"

    def parse(self, data: bytes) -> ParsedFile | Error:
        """Parse the PDF data and return a ParsedFile or an Error."""

        parsed_file = ParsedFile()

        with pdfplumber.open(BytesIO(data)) as pdf:

            text_info: list[PDFModelParser.TextLineObject] = []
            # Implement parsing logic here
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n"
                text_info += cast(list[PDFModelParser.TextLineObject], page.extract_text_lines())
                

            parsed_file.raw = text  # Example: extract text from all pages
            print(text[-100:])

            
            # Get unique font sizes, sorted largest first, saving the text as well
            sizes = sorted(
                {
                    (line["text"], line["chars"][0]["size"])
                    for line in text_info
                },
                reverse=True,
                key=lambda x: x[1]
            )

            size_to_level = {
                size: f"H{i+1}"
                for i, size in enumerate(sizes[:4])  # limit depth
            }

            print(size_to_level)


            parsed_file.title = self._parse_main_title(text_info)
            # parsed_file.title = text.index()
            parsed_file.sections = []  # Populate sections as needed

            return parsed_file
        

