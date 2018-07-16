import os
from send2trash import send2trash

from pyexcel_export.dir import root_path


def clean_output_folder(output_folder_path='tests/output', ignore=('README.md', '.gitignore')):
    for abs_path in [root_path(os.path.join(output_folder_path, filename))
                     for filename in os.listdir(root_path(output_folder_path))
                     if filename not in ignore]:
        send2trash(abs_path)


if __name__ == '__main__':
    clean_output_folder()
