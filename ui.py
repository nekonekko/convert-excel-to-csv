import PySimpleGUI as sg

from convert import copy_and_convert_all_excel


def main():
    sg.theme("DarkAmber")

    layout = [
        [
            sg.Text(
                "Converts all xlsx and xls files under the selected conversion source folder to csv."
            )
        ],
        [
            sg.Text(
                "The converted csv file will be placed under the converted folder, keeping the folder structure."
            )
        ],
        [
            sg.InputText(
                "Select source folder",
                key="-BASE-FOLDER-INPUT-",
                enable_events=True,
            ),
            sg.FolderBrowse(button_text="Select", key="-BASE-FOLDER-BROWSE-"),
        ],
        [
            sg.InputText(
                "Select a folder after conversion",
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
