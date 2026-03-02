import sys

from src.errors import Error, FileReadError, ParsingError
from src.model.model import WritableFile
from src.model.readers.file_model_reader import FileReader
from src.model.validators.validator import Validator
from src.utils import clean_folder, get_all_file_paths_in_dir

sys.stdout.reconfigure(encoding="utf-8")# type: ignore

if __name__ == "__main__":
    reader = FileReader()

    if isinstance(clean_folder("output"), Error):
        print("Error cleaning output folder. Check if the folder exists and is not being used by another process.")
        exit(1)

    for file_name in get_all_file_paths_in_dir("input"): #\\extras
        print(file_name)
        reader.set_target(file_name)
        result = reader.read()

        if isinstance(result, FileReadError):
            print(f"Error reading file: {result.message}")
            continue
        
        # parsing according to file received
        parser, data = result # Now it's safe to assume it didn't error out, so we can get the items out of the tuple
        
        parse_result = parser.parse(data)

        if isinstance(parse_result, ParsingError):
            print(f"Error parsing file: {parse_result.message}")
            continue
        
        if parse_result.writer and parse_result is not None:
            
            assert isinstance(parse_result.file, WritableFile)
            
            parse_result.writer.set_target(parse_result.file.target)
            parse_result.writer.write(parse_result.file, parse_result.template)
    
    validator = Validator()
    validator.validate_output()
    
            

