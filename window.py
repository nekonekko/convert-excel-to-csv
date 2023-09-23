import PySimpleGUI as sg

from convert import copy_and_convert_all_excel


def main():
    sg.theme("BlueMono")

    layout = [
        [
            sg.Text(
                "Converts all xlsx and xls files under the selected conversion source folder to csv."
            )
        ],
        [
            sg.Text(
                "The converted csv file will be placed under the target folder, keeping the folder structure."
            )
        ],
        [
            sg.Text('Source folder', size=(15, 1)),
            sg.InputText(
                key="-BASE-FOLDER-INPUT-",
                enable_events=True,
            ),
            sg.FolderBrowse(button_text="Select", key="-BASE-FOLDER-BROWSE-"),
        ],
        [
            sg.Text('Target folder', size=(15, 1)),
            sg.InputText(
                key="-CONVERTED-FOLDER-INPUT-",
                enable_events=True,
            ),
            sg.FolderBrowse(button_text="Select", key="-CONVERTED-FOLDER-BROWSE-"),
        ],
        [sg.Submit()],
    ]

    window = sg.Window("ExcelToCSV", layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == "Submit":
            copy_and_convert_all_excel(
                values["-BASE-FOLDER-INPUT-"], values["-CONVERTED-FOLDER-INPUT-"]
            )
            sg.popup("Conversion finished.")

    window.close()


if __name__ == "__main__":
    main()
