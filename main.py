import PySimpleGUI as sg
import pandas as pd
from pathlib import Path


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct")
    return False


def convert_worksheets(filepath):
    df = pd.concat(pd.read_excel(filepath, sheet_name=None), ignore_index=True)
    df.to_csv("Test-new.csv", index=False, header=True)
    sg.popup_no_titlebar("Done! :)")


def convert_to_csv(excel_file_path, output_folder, sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).stem
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, index=False)
    sg.popup_no_titlebar("Done! :)")


def main_window():
    layout = [[sg.T("Input File:", s=15, justification="r"), sg.I(key="-IN-"),
               sg.FilesBrowse(file_types=(("Excel Files", "*.xls*"), ("All Files", "*.*")))],
              [sg.T("Output Folder:", s=15, justification="r"), sg.I(key="-OUT-"), sg.FolderBrowse()],
              [sg.T("Input File Type:", s=15, justification="r"), sg.Radio("Worksheet", "dType", default=True), sg.Radio("Workbook", "dType")],
              [sg.T("Output File Type:", s=15, justification="r"), sg.Checkbox(".csv", default=True), sg.Checkbox(".xlsx")],
              [sg.B("Combine", s=16), sg.B("Convert", s=16), sg.Exit(button_color="tomato", s=16)]]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Combine":
            convert_worksheets(values["-IN-"])


if __name__ == "__main__":
    SETTINGS_PATH = Path.cwd()
    # create the settings object and use ini format
    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename="config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"]
    font_family = settings["GUI"]["font_family"]
    font_size = int(settings["GUI"]["font_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()
