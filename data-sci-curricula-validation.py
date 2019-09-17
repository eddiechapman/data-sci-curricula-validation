import pathlib
import re

import docx
import pandas as pd

DOCS_DIR = pathlib.Path('input/Documents')
DEGREE_XLSX = pathlib.Path('input/Data science programs.xlsx')


def normalize_filesnames(directory_path):
    """Rename degree documents for filename consistency.

    Ex: '630-skills.docx'

    Args:
        directory_path (pathlib.Path): a path object to a directory of
            degree documents

    """
    for child in directory_path.iterdir():
        name_normal = child.name.replace(' ', '').lower()
        path_normal = child.with_name(name_normal)
        child.rename(path_normal)


def filter_filenames(path):
    """Move and log any inconsistent filenames.

    Args:
        path (pathlib.Path): a path object to a directory of
            degree documents

    """
    pattern = r'\d{3}-(courses|skills|mission)\.docx'
    bad_docs = [f for f in path.iterdir() if not re.match(pattern, f.name)]
    pathlib.Path('output/bad_docs').mkdir(exist_ok=True)
    for f in bad_docs:
        f.rename(pathlib.Path('output/bad_docs') / f.name)


def load_excel():
    with pd.ExcelFile(DEGREE_XLSX) as xlsx:
        degrees_us = pd.read_excel(xlsx, 'US Only', true_values='T', false_values='F')
        degrees_non_us = pd.read_excel(xlsx, 'Non-US', true_values='T', false_values='F')
        codes = pd.read_excel(xlsx, 'Codes', header=None)

    print(len(degrees_us))


def main():
    filter_filenames(DOCS_DIR)


if __name__ == '__main__':
    main()
