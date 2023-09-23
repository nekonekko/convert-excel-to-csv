import PySimpleGUI as sg

from convert import copy_and_convert_all_excel

sg.theme("DarkAmber")

layout = [
    [sg.Text("選択された変換元フォルダ下にあるすべてのxlsxファイルとxlsファイルをcsvに変換します。")],
    [sg.Text("変換後のcsvファイルは、変換後フォルダ下にフォルダ構造を保った状態で入ります。")],
    [
        sg.InputText(
            "変換元フォルダを選択",
            key="-BASE-FOLDER-INPUT-",
            enable_events=True,
        ),
        sg.FolderBrowse(button_text="選択", font=("メイリオ"), key="-BASE-FOLDER=BROWSE-"),
    ],
    [
        sg.InputText(
            "変換後フォルダを選択",
            key="-CONVERTED-FOLDER-INPUT-",
            enable_events=True,
        ),
        sg.FolderBrowse(
            button_text="選択", font=("メイリオ"), key="-CONVERTED-FOLDER=BROWSE-"
        ),
    ],
    [sg.Submit()],
]

window = sg.Window("ExcelToCSV", layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED:  # if user closes window or clicks cancel
        break
    if event == "Submit":
        # csvへの変換を行う
        copy_and_convert_all_excel(
            values["-BASE-FOLDER-INPUT-"], values["-CONVERTED-FOLDER-INPUT-"]
        )

window.close()
