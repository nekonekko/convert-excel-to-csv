import PySimpleGUI as sg

from convert import copy_and_convert_all_excel


def main():
    sg.theme("BlueMono")

    layout = [
        [sg.Text("選択した変換元フォルダの下にあるすべてのxlsx/xlsファイルをcsvに変換します。")],
        [sg.Text("変換されたcsvファイルは、フォルダ構造を維持したまま、変換後フォルダの下に配置されます。")],
        [
            sg.Text("変換元フォルダ", size=(15, 1)),
            sg.InputText(
                key="-SOURCE-FOLDER-INPUT-",
                enable_events=True,
            ),
            sg.FolderBrowse(button_text="選択", key="-SOURCE-FOLDER-BROWSE-"),
        ],
        [
            sg.Text("変換後フォルダ", size=(15, 1)),
            sg.InputText(
                key="-TARGET-FOLDER-INPUT-",
                enable_events=True,
            ),
            sg.FolderBrowse(button_text="選択", key="-TARGET-FOLDER-BROWSE-"),
        ],
        [sg.Push(), sg.Submit("決定"), sg.Push()],
    ]

    window = sg.Window("ExcelToCSV", layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == "Submit":
            is_completed = copy_and_convert_all_excel(
                values["-SOURCE-FOLDER-INPUT-"], values["-TARGET-FOLDER-INPUT-"]
            )
            if is_completed:
                sg.popup("変換が完了しました！")
            else:
                sg.popup("変換を完了できませんでした！")
    window.close()


if __name__ == "__main__":
    main()
