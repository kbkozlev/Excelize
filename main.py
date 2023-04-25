# Created by Kaloian Kozlev

import PySimpleGUI as sg
from functions import *
import time
import ctypes
import platform


def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)


def about_window():
    layout = [[sg.T(str(window_title), font=(font_family, 12, "bold"))],
              [sg.T("GitHub: https://github.com/kbkozlev/Excelize")],
              [sg.T("License: Apache-2.0")],
              [sg.T("Copyright © 2023 Kaloian Kozlev")]]
    window = sg.Window("About", layout, modal=True, icon=icon, size=(400, 150))
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break

    window.close()


def main_window():
    menu_bar = [['Help', 'About']]

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
               sg.T(s=56, justification="l", key="-OUTPUT-")],
              [sg.T(s=l_side_t_size), sg.B("Combine", s=b_side_b_size, button_color=b_colour),
               sg.B("Split", s=b_side_b_size, button_color=b_colour),
               sg.Push(), sg.Exit(button_color=exit_b_colour, s=15)]]

    window = sg.Window(window_title, layout, use_custom_titlebar=False, keep_on_top=False, icon=icon)

    while True:
        event, values = window.read()

        if event in (sg.WINDOW_CLOSED, "Exit"):
            break

        in_list = values["-IN-"].split(";")
        output_path = values["-OUT-"]
        csv = values['-CSV-']
        xls = values['-XLS-']

        if event == "About":
            about_window()

        if event == "Combine":

            if is_valid_path(in_list, window) and is_valid_path(output_path, window):
                if csv is not False or xls is not False:

                    if values["-WS-"]:
                        combine_and_convert_ws(in_list, csv, xls, output_path, window)

                    elif values["-WB-"]:
                        name = sg.popup_get_text("New Workbook Name:", default_text="Workbook-Combined",
                                                 no_titlebar=False, grab_anywhere=True,
                                                 font=(font_family, font_size), size=(30, 5), button_color=b_colour,
                                                 icon=icon)

                        if name is not None:
                            combine_and_convert_wb(in_list, csv, xls, output_path, window, name)

                        else:
                            window["-OUTPUT-"].update("*** Missing Workbook Name ***")
                            window.refresh()

                else:
                    window["-OUTPUT-"].update("*** No Output File Type Selected ***")
                    window.refresh()

        elif event == "Split":
            if is_valid_path(in_list, window) and is_valid_path(output_path, window):
                if csv is not False or xls is not False:

                    split_wb(in_list, csv, xls, output_path, window)

                else:
                    window["-OUTPUT-"].update("*** No Output File Type Selected ***")
                    window.refresh()

        time.sleep(1)
        window["-OUTPUT-"].update(" ")


if __name__ == "__main__":
    window_title = "Excelize v.1.0"
    font_family = "Arial"
    font_size = 10
    b_colour = "#015FB8"
    exit_b_colour = "#D44A5A"
    icon = "icon.ico"

    b_side_b_size = 16
    r_side_b_size = 15
    l_side_t_size = 16

    sg.theme("Reddit")
    sg.set_options(font=(font_family, font_size))

    main_window()
    make_dpi_aware()
