"""
record_document_status.py

Update an Excel sheet with the document statuses of each degree.
"""
import pathlib
import re
import pandas as pd

PATH_VALIDATED = pathlib.Path('output/validated')
PATH_XLSX = pathlib.Path('input/Data science programs.xlsx')
PATTERN = r'(\d{3})-(courses|skills|mission)\.docx'


def iter_file_matches(path, pattern):
    return (re.match(pattern, f.name) for f in path.iterdir())


def load_xlsx(path):
    with pd.ExcelFile(PATH_XLSX) as xlsx:
        degrees_us = pd.read_excel(
            io=xlsx,
            sheet_name='US Only',
            true_values='T',
            false_values='F'
        )
        degrees_intl = pd.read_excel(
            io=xlsx,
            sheet_name='Non-US',
            true_values='T',
            false_values='F'
        )
        return degrees_us, degrees_intl


def generate_blank_degree_dict(size=1000):
    degrees = {}
    for i in range(1, size):
        degree_id = str(i).zfill(3)
        degrees[degree_id] = {
            'skills': False,
            'courses': False,
            'mission': False
        }
    return degrees


def generate_lookup():
    return {
        'skills': 'skills',
        'courses': 'courses',
        'mission': 'mission'
    }


def degree_documents(path, pattern):
    lookup = generate_lookup()
    degrees = generate_blank_degree_dict()

    for match in iter_file_matches(path, pattern):
        if not match:
            continue

        degree = degrees.get(match.group(1))
        component = lookup.get(match.group(2))

        # TODO: Log this instead
        if not match or not component:
            print(match.group(0))

        degree[component] = True

    return degrees


def generate_row_update(row):
    return {
        'ID': getattr(row, '_1', None),
        'SKILLS': getattr(row, 'SKILLS', False),
        'COURSES': getattr(row, 'COURSES', False),
        'MISSION': getattr(row, 'MISSION', False),
        'ERROR': False
    }


def main():
    degrees_us, degrees_intl = load_xlsx(PATH_XLSX)
    document_progress = degree_documents(PATH_VALIDATED, PATTERN)
    degree_updates = []

    df = degrees_us
    for row in df.itertuples():

        if not getattr(row, '_1', None):
            print(f'ERROR: No degree ID found in {getattr(row, "Index")}')
            row_update = generate_row_update(row)
            row_update['ERROR'] = 'No degree ID found in row'
            degree_updates.append(row_update)
            continue

        degree_id = str(getattr(row, '_1')).zfill(3)
        row_update = generate_row_update(row)
        row_update['ID'] = degree_id
        degree_docs = document_progress.get(degree_id)

        for name in ['skills', 'courses', 'mission']:
            # When the Excel sheet thinks a doc exists but it can't be found
            if row_update[name.upper()] and not degree_docs.get(name):
                row_update['ERROR'] = f'Not found: {degree_id}, {name}'
            else:
                row_update[name.upper()] = degree_docs[name]

        row_update = {k: str(v) for k, v in row_update.items()}
        degree_updates.append(row_update)

    update = pd.DataFrame(degree_updates)
    update.to_csv(pathlib.Path('output/document_progress-us.csv'))


if __name__ == '__main__':
    main()
