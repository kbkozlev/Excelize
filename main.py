from functions import *
import time


def main_window():
    layout = [[sg.T("Input File(s):", s=l_side_t_size, justification="r"), sg.I(key="-IN-"),
               sg.FilesBrowse(file_types=(("Excel Files", "*.xls*"), ("All Files", "*.*")), s=r_side_b_size, button_color=b_colour)],
              [sg.T("Output Folder:", s=l_side_t_size, justification="r"), sg.I(key="-OUT-"),
               sg.FolderBrowse(s=r_side_b_size, button_color=b_colour)],
              [sg.T("Input File Type:", s=l_side_t_size, justification="r"),
               sg.Radio("Worksheet", "dType", default=True, key="-WS-"),
               sg.Radio("Workbook", "dType", key="-WB-"),
               sg.Radio("Combined", "dType", key="-CB-")],
              [sg.T("Output File Type(s):", s=l_side_t_size, justification="r"),
               sg.Checkbox(".csv", default=True, key="-CSV-"),
               sg.Checkbox(".xlsx", key="-XLS-")],
              [sg.T("Exec. Status:", s=l_side_t_size, justification="r", font=(font_family, font_size, "bold")),
               sg.T(s=56, justification="l", key="-OUTPUT-")],
              [sg.T(s=l_side_t_size), sg.B("Combine", s=b_side_b_size, button_color=b_colour), sg.B("Split", s=b_side_b_size, button_color=b_colour),
                    sg.Push(), sg.Exit(button_color=exit_b_colour, s=15)]]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=False, keep_on_top=False)

    while True:
        event, values = window.read()
        in_list = values["-IN-"].split(";")
        output_path = values["-OUT-"]
        csv = values['-CSV-']
        xls = values['-XLS-']

        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Combine":

            if is_valid_path(in_list, window) and is_valid_path(in_list, window):
                if csv is not False or xls is not False:

                    if values["-WS-"]:
                        combine_and_convert_ws(in_list, csv, xls, output_path, window)

                    elif values["-WB-"]:
                        name = sg.popup_get_text("New Workbook Name:", default_text="Workbook-Combined",
                                                 no_titlebar=False, grab_anywhere=True,
                                                 font=(font_family, font_size), size=(30, 5), button_color=b_colour)

                        if name is not None:
                            combine_and_convert_wb(in_list, csv, xls, output_path, window, name)

                        else:
                            window["-OUTPUT-"].update("*** Missing Workbook Name ***")
                            window.refresh()

                    elif values["-CB-"]:
                        window["-OUTPUT-"].update("*** Function Not Yet Implemented ***")
                        window.refresh()
                else:
                    window["-OUTPUT-"].update("*** No Output File Type Selected ***")
                    window.refresh()

        elif event == "Split":
            if is_valid_path(in_list, window) and is_valid_path(in_list, window):
                if csv is not False or xls is not False:
                    pass

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
    b_colour = settings["GUI"]["b_colour"]
    exit_b_colour = settings["GUI"]["exit_b_colour"]

    r_side_b_size = int(settings["ELMS"]["r_side_b_size"])
    b_side_b_size = int(settings["ELMS"]["b_side_b_size"])
    l_side_t_size = int(settings["ELMS"]["l_side_t_size"])

    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))

    main_window()

