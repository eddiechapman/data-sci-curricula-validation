import pathlib
import re
import docx

DOCS_DIR = pathlib.Path('input/Documents')
PATH_VALIDATED = pathlib.Path('output/validated')
PATH_XLSX = pathlib.Path('input/Data science programs.xlsx')
PATTERN = r'(\d{3})-(courses|skills|mission)\.docx'


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
    bad_docs = [f for f in path.iterdir() if not re.match(PATTERN, f.name)]
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


def find_incomplete_sets(path):
    incomplete_sets_dir = pathlib.Path('output/incomplete_sets')
    incomplete_sets_dir.mkdir(exist_ok=True)
    incomplete_set_log = pathlib.Path(
        incomplete_sets_dir / 'incomplete_sets.log'
    )

    sets = {}
    for child in path.iterdir():
        match = re.match(PATTERN, child.name)
        if not match:
            continue
        if not match.group(1) in sets:
            sets[match.group(1)] = []
        sets[match.group(1)].append(match.group(2))

    incomplete_sets = []
    with open(incomplete_set_log, 'w') as f:
        for degree, docs in sorted(sets.items()):
            if not len(docs) == 3:
                incomplete_sets.append(degree)
                f.write(f'{degree}\n')
                f.write(f'---\n')
                f.write(f'skills:\t\t{"skills" in docs}\n')
                f.write(f'mission:\t{"mission" in docs}\n')
                f.write(f'courses:\t{"courses" in docs}\n')
                f.write(f'\n')

    for child in path.iterdir():
        match = re.match(PATTERN, child.name)
        if not match:
            continue
        if match.group(1) in incomplete_sets:
            child.rename(incomplete_sets_dir / child.name)


def name_complete_sets(path):
    complete_sets = set()
    for child in path.iterdir():
        match = re.match(PATTERN, child.name)
        if match:
            complete_sets.add(match.group(1))
    with open('complete_sets.log', 'w') as f:
        for degree in sorted(complete_sets):
            f.write(f'{degree}\n')


def main():
    pass


if __name__ == '__main__':
    main()
