import PySimpleGUI

from functions import *
import time


def main_window():
    layout = [[sg.T("Input File(s):", s=l_side_t_size, justification="r"), sg.I(key="-IN-"),
               sg.FilesBrowse(file_types=(("Excel Files", "*.xls*"), ("All Files", "*.*")), s=r_side_b_size)],
              [sg.T("Output Folder:", s=l_side_t_size, justification="r"), sg.I(key="-OUT-"),
               sg.FolderBrowse(s=r_side_b_size)],
              [sg.T("Input File Type:", s=l_side_t_size, justification="r"),
               sg.Radio("Worksheet", "dType", default=True, key="-WS-"),
               sg.Radio("Workbook", "dType", key="-WB-"),
               sg.Radio("Combined", "dType", key="-CB-")],
              [sg.T("Output File Type(s):", s=l_side_t_size, justification="r"),
               sg.Checkbox(".csv", default=True, key="-CSV-"),
               sg.Checkbox(".xlsx", key="-XLS-")],
              [sg.T("Exec. Status:", s=l_side_t_size, justification="r", font=(font_family, font_size, "bold")),
               sg.T(s=56, justification="l", key="-OUTPUT-")],
              center([sg.B("Execute", s=b_side_b_size), sg.Exit(button_color="tomato", s=b_side_b_size)])]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True, keep_on_top=False)

    while True:
        event, values = window.read()
        input_path = values["-IN-"]
        output_path = values["-OUT-"]
        csv = values['-CSV-']
        xls = values['-XLS-']

        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Execute":
            if is_valid_path(input_path, window) and is_valid_path(output_path, window):

                if values["-WS-"]:
                    combine_and_convert_ws(input_path, csv, xls, output_path, window)

                elif values["-WB-"]:
                    combine_workbooks(input_path, csv, xls, output_path, window)

                elif values["-WB-"]:
                    window["-OUTPUT-"].update("*** Function not yet implemented ***")
                    window.refresh()

        time.sleep(1)
        window["-OUTPUT-"].update(" ")


if __name__ == "__main__":
    SETTINGS_PATH = Path.cwd()
    # create the settings object and use ini format
    settings = sg.UserSettings(
        path=str(SETTINGS_PATH), filename="config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"]
    font_family = settings["GUI"]["font_family"]
    font_size = int(settings["GUI"]["font_size"])
    r_side_b_size = int(settings["ELMS"]["r_side_b_size"])
    b_side_b_size = int(settings["ELMS"]["b_side_b_size"])
    l_side_t_size = int(settings["ELMS"]["l_side_t_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))

    main_window()
