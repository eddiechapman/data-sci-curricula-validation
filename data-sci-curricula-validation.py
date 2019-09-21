import pathlib
import re

import docx
import pandas as pd

DOCS_DIR = pathlib.Path('input/Documents')
DEGREE_XLSX = pathlib.Path('input/Data science programs.xlsx')


def normalize_filesnames(path):
    """
    Rename degree documents for filename consistency.
    """
    for child in path.iterdir():
        name_normal = child.name.replace(' ', '').lower()
        path_normal = child.with_name(name_normal)
        child.rename(path_normal)


def filter_filenames(path):
    """
    Move and log any inconsistent filenames.
    """
    pattern = r'\d{3}-(courses|skills|mission)\.docx'
    bad_docs = [f for f in path.iterdir() if not re.match(pattern, f.name)]
    pathlib.Path('output/bad_docs').mkdir(exist_ok=True)
    for f in bad_docs:
        f.rename(pathlib.Path('output/bad_docs') / f.name)


def find_null_docs(path):
    """
    Check if any documents in the path say 'no skills' etc.
    """
    null_docs = pathlib.Path('output/null_docs')
    null_docs.mkdir(exist_ok=True)

    null_docs_log = dict()
    for child in path.iterdir():
        doc = docx.Document(child)
        text = '\n'.join([p.text for p in doc.paragraphs])
        if not text or len(text) < 15:
            null_docs_log[child.name] = text
            child.rename(null_docs / child.name)

    with open(pathlib.Path(null_docs / 'null_docs.log'), 'w') as f:
        for filename, text in null_docs_log.items():
            f.write(f'{filename}\n')
            f.write('---------\n')
            f.writelines(f'{text}\n')
            f.write('\n')


def load_excel():
    with pd.ExcelFile(DEGREE_XLSX) as xlsx:
        degrees_us = pd.read_excel(
            io=xlsx,
            sheet_name='US Only',
            true_values='T',
            false_values='F'
        )
        degrees_non_us = pd.read_excel(
            io=xlsx,
            sheet_name='Non-US',
            true_values='T',
            false_values='F'
        )
        codes = pd.read_excel(
            io=xlsx,
            sheet_name='Codes',
            header=None
        )

    print(len(degrees_us))


def main():
    find_null_docs(DOCS_DIR)


if __name__ == '__main__':
    main()
