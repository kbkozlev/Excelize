import PySimpleGUI as sg
import pandas as pd
from pathlib import Path


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_no_titlebar("File not found")
    return False


def combine_worksheets(excel_file_path):
    df = pd.concat(pd.read_excel(excel_file_path, sheet_name=None), ignore_index=True)
    filename = Path(excel_file_path).stem + "_combined"
    return df, filename


def combine_workbooks():
    pass


def convert_to_csv(df, filename, output_folder):
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, index=False)


def convert_to_excel(df, filename, output_folder):
    outputfile = Path(output_folder) / f"{filename}.xlsx"
    df.to_excel(outputfile, index=False)


def main_window():
    layout = [[sg.T("Input File:", s=15, justification="r"), sg.I(key="-IN-"),
               sg.FilesBrowse(file_types=(("Excel Files", "*.xls*"), ("All Files", "*.*")))],
              [sg.T("Output Folder:", s=15, justification="r"), sg.I(key="-OUT-"), sg.FolderBrowse()],
              [sg.T("Input File Type:", s=15, justification="r"), sg.Radio("Worksheet", "dType", default=True, key="-WS-"), sg.Radio("Workbook", "dType", key="-WB-")],
              [sg.T("Output File Type:", s=15, justification="r"), sg.Checkbox(".csv", default=True, key="-CSV-"), sg.Checkbox(".xlsx", key="-XLS-")],
              [sg.B("Execute", s=16), sg.Exit(button_color="tomato", s=16)]]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        df, filename = combine_worksheets(values["-IN-"])

        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Execute":
            if values["-WS-"]:
                if is_valid_path(values["-IN-"]):
                    if values["-CSV-"]:
                        convert_to_csv(df, filename, values["-OUT-"])
                    if values["-XLS-"]:
                        convert_to_excel(df, filename, values["-OUT-"])
                    sg.popup_quick_message("Done! :)")
            elif values["-WB-"]:
                print("fail")


if __name__ == "__main__":
    SETTINGS_PATH = Path.cwd()
    # create the settings object and use ini format
    settings = sg.UserSettings(
        path=str(SETTINGS_PATH), filename="config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"]
    font_family = settings["GUI"]["font_family"]
    font_size = int(settings["GUI"]["font_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()
