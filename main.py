from functions import *

def main_window():

    layout = [  [sg.T("Input File(s):", s=15, justification="r"), sg.I(key="-IN-"),
                sg.FilesBrowse(file_types=(("Excel Files", "*.xls*"), ("All Files", "*.*")), s=15)],
                [sg.T("Output Folder:", s=15, justification="r"), sg.I(key="-OUT-"), sg.FolderBrowse(s=15)],
                [sg.T("Input File Type:", s=15, justification="r"), sg.Radio("Worksheet", "dType", default=True, key="-WS-"),
                sg.Radio("Workbook", "dType", key="-WB-")],
                [sg.T("Output File Type(s):", s=15, justification="r"), sg.Checkbox(".csv", default=True, key="-CSV-"),
                sg.Checkbox(".xlsx", key="-XLS-")],
                [sg.T("Exec. Status:", s=15, justification="r", font=(font_family,font_size, "bold")), sg.T(s=38, justification="l", key="-OUTPUT-")],
                c([sg.B("Execute", s=16), sg.Exit(button_color="tomato", s=16)])    ]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True, keep_on_top=True)

    def is_valid_path(filepath):
        if filepath and Path(filepath).exists():
            return True
        window["-OUTPUT-"].update("***Filepath not valid***")
        return False

    while True:
        event, values = window.read()
        input_path = values["-IN-"]
        output_path = values["-OUT-"]

        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        if event == "Execute":
            if is_valid_path(input_path) and is_valid_path(output_path):

                if values["-WS-"]:
                    window["-OUTPUT-"].update("Combining Worksheets")
                    df, filename = combine_worksheets(input_path)

                    if values["-CSV-"]:
                        window["-OUTPUT-"].update("Starting conversion to CSV")
                        convert_to_csv(df, filename, output_path)
                        window["-OUTPUT-"].update("Conversion finished")

                    if values["-XLS-"]:
                        convert_to_excel(df, filename, output_path)

                elif values["-WB-"]:
                    print("fail")

        elif event == "Test":
            window["-OUTPUT-"].update("works")


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
