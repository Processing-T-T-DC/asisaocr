import os
import shutil
from typing import Literal

from src.errors import Error


def get_all_file_paths_in_dir(dir: str) -> list[str]:
    file_paths = []
    for root, _, files in os.walk(dir):
        for file in files:
            file_paths.append(os.path.join(root, file))
    return file_paths

def clean_folder(dir: str) -> Error | None:
    """Remove one directory without iterating through the files and create it again."""
    if os.path.exists(dir):
        try:
            shutil.rmtree(dir)
        except Exception as e:
            return Error(f"Error cleaning folder '{dir}': {str(e)}")
    os.makedirs(dir)

type FileType = Literal["ExcelFile", "PDFFile", "HTMLFile", "ExcelLockFile", "Unknown"]

def __get_file_suffix(file: str) -> str:
    i = file.rfind('.')
    if 0 < i < len(file) - 1:
        return file[i:]
    else:
        return ''

def get_file_type_from_file_path(file: str) -> FileType:

    extension = __get_file_suffix(file).lower()

    if extension in {".xls", ".xlsx", ".xlsm", ".xlsb"}:
        if file.startswith("~$"):
            return "ExcelLockFile"
        
        return "ExcelFile"
    elif extension == ".pdf":
        return "PDFFile"
    elif extension in {".html", ".htm"}:
        return "HTMLFile"
    else:
        return "Unknown"
