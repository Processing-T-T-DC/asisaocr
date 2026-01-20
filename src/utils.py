import os


def get_all_file_paths_in_dir(dir: str) -> list[str]:
    file_paths = []
    for root, _, files in os.walk(dir):
        for file in files:
            file_paths.append(os.path.join(root, file))
    return file_paths
