import glob
import re

import pandas as pd


def copy_and_convert_all_excel(base_path, converted_path):
    """Excel files under the base_path folder are converted to csv
    and placed under the converted_path folder.
    The folder structure in base_path is maintained in converted_path.

    Args:
        base_path (str): Path of the folder to be converted
        converted_path (str): Path of the folder after conversion
    """

    for file_path in glob.glob("**", recursive=True, root_dir=base_path):
        if re.search("\.(xls|xlsx)$", file_path):
            print(file_path)
