import PySimpleGUI as sg
import pandas as pd
from pathlib import Path


def center(elms):
    return [sg.Stretch(), *elms, sg.Stretch()]


def border(elms):
    return sg.Frame('', [[elms]])


def is_valid_path(filepath, window):
    wb_list = filepath.split(';')
    for item in wb_list:
        if item and Path(item).exists():
            return True
        window["-OUTPUT-"].update("***Filepath not valid***")
        window.refresh()


def combine_and_convert_ws(excel_file_path, csv, xls, output_folder, window):
    ws_list = excel_file_path.split(';')
    for item in ws_list:
        window["-OUTPUT-"].update(f"*** Combining Worksheet from {Path(item).stem} ***")
        window.refresh()
        df = pd.concat(pd.read_excel(item, sheet_name=None), ignore_index=True)
        filename = Path(item).stem + "_combined"
        if csv:
            outfile = Path(output_folder) / f"{filename}.csv"
            window["-OUTPUT-"].update(f"*** Converting {Path(item).stem} to CSV ***")
            window.refresh()
            df.to_csv(outfile, index=False)
        if xls:
            outfile = Path(output_folder) / f"{filename}.xlsx"
            window["-OUTPUT-"].update(f"*** Converting {Path(item).stem} to XLSX ***")
            window.refresh()
            df.to_excel(outfile, index=False)

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()


def combine_and_convert_wb(excel_file_path, csv, xls, output_folder, window, name):
    final_df = pd.DataFrame()
    wb_list = excel_file_path.split(';')

    for item in wb_list:
        window["-OUTPUT-"].update(f"*** Loading File {Path(item).stem} ***")
        window.refresh()
        df = pd.read_excel(item)
        final_df = final_df._append(df, ignore_index=True)

    if csv:
        outfile = Path(output_folder) / f"{name}.csv"
        window["-OUTPUT-"].update(f"*** Converting {name} to CSV ***")
        window.refresh()
        final_df.to_csv(outfile, index=False)

    if xls:
        outfile = Path(output_folder) / f"{name}.xlsx"
        window["-OUTPUT-"].update(f"*** Converting {name} to XLSX ***")
        window.refresh()
        final_df.to_excel(outfile, index=False)

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()

