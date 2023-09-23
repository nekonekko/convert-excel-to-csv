import glob
import re
import os

import pandas as pd
import PySimpleGUI as sg


def copy_and_convert_all_excel(base_path, converted_path):
    """Excel files under the base_path folder are converted to csv
    and placed under the converted_path folder.
    The folder structure in base_path is maintained in converted_path.

    Args:
        base_path (str): Path of the folder to be converted
        converted_path (str): Path of the folder after conversion

    Return:
        (boolean): Whether the conversion was completed to the end or not.
    """

    excel_path_list = []
    for file_path in glob.glob("**", recursive=True, root_dir=base_path):
        if re.search("\.(xls|xlsx)$", file_path):
            excel_path_list.append(file_path)

    layout = [
        [
            sg.ProgressBar(
                len(excel_path_list), orientation="h", size=(30, 10), key="-PROG-"
            )
        ],
        [sg.Text(key="-CURRENT-FILE-NAME-")],
    ]
    window = sg.Window("Progress", layout)

    for i, excel_path in enumerate(excel_path_list):
        event, _ = window.read(timeout=10, timeout_key="-READ-NEXT-FILE-")

        if event == sg.WIN_CLOSED:
            window.close()
            return False
        elif event == "-READ-NEXT-FILE-":
            dirname = os.path.dirname(excel_path)
            filename_without_ext = os.path.splitext(os.path.basename(excel_path))[0]

            try:
                excel_file = pd.read_excel(os.path.join(base_path, excel_path))
            except:
                sg.popup(f"{excel_path}を開けませんでした！")  # e.g. temporary file
                continue

            window["-CURRENT-FILE-NAME-"].update(excel_path)
            os.makedirs(os.path.join(converted_path, dirname), exist_ok=True)
            excel_file.to_csv(
                os.path.join(converted_path, dirname, filename_without_ext + ".csv"),
                index=None,
                header=False,
            )

            window["-PROG-"].update(i + 1)

    window.close()

    return True
