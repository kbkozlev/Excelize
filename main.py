import PySimpleGUI as sg
import functions as fn
import re
import time
import webbrowser


def about_window():
    layout = [[sg.Push(), sg.T(str(window_title), font=(font_family, 12, "bold")), sg.Push()],
              [sg.T(s=40)],
              [sg.Push(), sg.T(github_url['name'], enable_events=True, font=(font_family, font_size, "underline"),
                               justification='l', text_color='#0066CC',
                               auto_size_text=True, key='download'), sg.Push()],
              [sg.Push(), sg.T("License: Apache-2.0", justification='c'), sg.Push()],
              [sg.T()],
              [sg.Push(), sg.T("Copyright © 2023 Kaloian Kozlev", text_color='light grey'), sg.Push()]]

    window = sg.Window("About", layout, icon=icon)

    while True:
        event, values = window.read()

        match event:

            case sg.WIN_CLOSED:
                break

            case 'download':
                webbrowser.open(github_url['url'])
                window.close()


def updates_window(current_release):
    latest_release, download_url = fn.get_latest_version()
    layout = [[sg.Push(), sg.T('Version Info', font=(font_family, 12, 'bold')), sg.Push()],
              [sg.T()],
              [sg.T('Current Version:', s=13), sg.T(f'{current_release}', font=(font_family, 10, 'bold'))],
              [sg.T(f'Latest Version:', s=13), sg.T(f'{latest_release}', font=(font_family, 10, 'bold'))],
              [sg.T(s=40, justification="c", key="-INFO-")],
              [sg.Push(), sg.B('Download', key='download', button_color=b_colour, s=16), sg.Push()]]

    window = sg.Window("Check for Updates", layout, icon=icon)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED:
            break

        match event:

            case 'download':
                if latest_release is not None:
                    current_release = re.sub(r'[^0-9]', '', current_release)
                    latest_release = re.sub(r'[^0-9]', '', latest_release)

                    if int(latest_release) > int(current_release):
                        webbrowser.open(download_url)
                        window.close()

                    else:
                        window['-INFO-'].update("You have the latest version", text_color='green')

                else:
                    window['-INFO-'].update("Cannot fetch version data", text_color='red')

        window.refresh()
        time.sleep(1)
        window["-INFO-"].update(" ")


def main_window():
    menu_bar = [['Help', ['About', 'Check for Updates']]]

    layout = [[sg.Menubar(menu_bar)],
              [sg.T("Input File(s):", s=l_side_t_size, justification="r"),
               sg.I(key="-IN-", default_text="Select Input File(s)", text_color="grey"),
               sg.FilesBrowse(file_types=(("Excel Files", "*.xls*"), ("All Files", "*.*")), s=r_side_b_size,
                              button_color=b_colour)],
              [sg.T("Output Folder:", s=l_side_t_size, justification="r"),
               sg.I(key="-OUT-", default_text="Select Output Folder", text_color="grey"),
               sg.FolderBrowse(s=r_side_b_size, button_color=b_colour)],
              [sg.T("Input File Type:", s=l_side_t_size, justification="r"),
               sg.Radio("Worksheet", "dType", default=True, key="-WS-"),
               sg.Radio("Workbook", "dType", key="-WB-")],
              [sg.T("Output File Type(s):", s=l_side_t_size, justification="r"),
               sg.Checkbox(".csv", default=True, key="-CSV-"),
               sg.Checkbox(".xlsx", key="-XLS-")],
              [sg.T("Exec. Status:", s=l_side_t_size, justification="r", font=(font_family, font_size)),
               sg.T(s=58, justification="l", key="-OUTPUT-", auto_size_text=True)],
              [sg.T(s=l_side_t_size), sg.B("Combine", s=b_side_b_size, button_color=b_colour),
               sg.B("Split", s=b_side_b_size, button_color=b_colour),
               sg.Push(), sg.Exit(button_color=exit_b_colour, s=16)]]

    window = sg.Window(window_title, layout, icon=icon)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        in_list = values["-IN-"].split(";") if values["-IN-"] is not None else ''
        output_path = values["-OUT-"]
        csv = values['-CSV-']
        xls = values['-XLS-']

        match event:

            case "About":
                about_window()

            case "Check for Updates":
                updates_window(release)

            case "Combine":
                if fn.is_valid_path(in_list, window) and fn.is_valid_path(output_path, window):
                    if csv or xls:
                        if values["-WS-"]:
                            fn.combine_and_convert_ws(in_list, csv, xls, output_path, window)

                        elif values["-WB-"]:
                            name = sg.popup_get_text("New Workbook Name:", default_text="Workbook-Combined",
                                                     no_titlebar=False, grab_anywhere=True,
                                                     font=(font_family, font_size), size=(30, 5), button_color=b_colour,
                                                     icon=icon)

                            if name is not None:
                                fn.combine_and_convert_wb(in_list, csv, xls, output_path, window, name)

                            else:
                                window["-OUTPUT-"].update("*** Missing Workbook Name ***", text_color='red')
                                window.refresh()

                    else:
                        window["-OUTPUT-"].update("*** No Output File Type Selected ***", text_color='red')
                        window.refresh()

            case "Split":
                if fn.is_valid_path(in_list, window) and fn.is_valid_path(output_path, window):
                    if csv or xls:

                        fn.split_wb(in_list, csv, xls, output_path, window)

                    else:
                        window["-OUTPUT-"].update("*** No Output File Type Selected ***", text_color='red')
                        window.refresh()

        time.sleep(1)
        window["-OUTPUT-"].update(" ")


if __name__ == "__main__":
    release = '1.2.2-beta'
    window_title = f"Excelize v{release}"
    font_family = "Arial"
    font_size = 10
    b_colour = "#015FB8"
    exit_b_colour = "#D44A5A"
    icon = "icon.ico"

    b_side_b_size = 16
    r_side_b_size = 16
    l_side_t_size = 16

    sg.theme("Reddit")
    sg.set_options(font=(font_family, font_size), force_modal_windows=True, dpi_awareness=True, auto_size_buttons=True,
                   auto_size_text=True)

    github_url = {'name': 'Official GitHub Page',
                  'url': 'https://github.com/kbkozlev/Excelize'}

    main_window()
