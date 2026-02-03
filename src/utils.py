import os
import shutil


def get_all_file_paths_in_dir(dir: str) -> list[str]:
    file_paths = []
    for root, _, files in os.walk(dir):
        for file in files:
            file_paths.append(os.path.join(root, file))
    return file_paths

def clean_folder(dir: str) -> None:
    """Remove one directory without iterating through the files and create it again."""
    if os.path.exists(dir):
        shutil.rmtree(dir)
    os.makedirs(dir)