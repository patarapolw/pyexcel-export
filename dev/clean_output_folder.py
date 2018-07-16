from send2trash import send2trash
from pathlib import Path


def clean_output_folder(output_folder_path=Path('../tests/output'), ignore=('README.md', '.gitignore')):
    for filename in output_folder_path.glob('*'):
        if filename.name not in ignore:
            send2trash(str(filename.absolute()))


if __name__ == '__main__':
    clean_output_folder()
