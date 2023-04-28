import pandas as pd
from pathlib import Path
import ctypes
import platform
import requests



def is_valid_path(in_list, window):
    for item in in_list:
        if item and Path(item).exists():
            return True
        else:
            window["-OUTPUT-"].update("*** Filepath not valid ***")
            window.refresh()


def get_latest_version():
    try:
        response = requests.get("https://api.github.com/repos/kbkozlev/Excelize/releases/latest")
        latest_release = response.json()['tag_name']
        download_url = response.json()['html_url']

    except:
        latest_release = None
        download_url = None

    return latest_release, download_url


def convert_to_csv(output_folder, name, window, df):
    outfile = Path(output_folder) / f"{name}.csv"
    window["-OUTPUT-"].update(f"*** Converting {name} to CSV ***")
    window.refresh()
    try:
        df.to_csv(outfile, index=False)

    except:
        window["-OUTPUT-"].update(f"*** Error converting {name} to CSV ***")


def convert_to_xls(output_folder, name, window, df):
    outfile = Path(output_folder) / f"{name}.xlsx"
    window["-OUTPUT-"].update(f"*** Converting {name} to XLSX ***")
    window.refresh()
    try:
        df.to_excel(outfile, index=False)

    except:
        window["-OUTPUT-"].update(f"*** Error converting {name} to XLSX ***")


def split_wb(in_list, csv, xls, output_folder, window):
    for item in in_list:
        window["-OUTPUT-"].update(f"*** Splitting {Path(item).stem} into Worksheets ***")
        window.refresh()
        try:
            xl = pd.ExcelFile(item)
            for name in xl.sheet_names:
                df = pd.read_excel(xl, sheet_name=name)

                if csv:
                    convert_to_csv(output_folder, name, window, df)
                if xls:
                    convert_to_xls(output_folder, name, window, df)

        except:
            window["-OUTPUT-"].update(f"*** Error splitting {Path(item).stem} ***")

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()


def combine_and_convert_ws(in_list, csv, xls, output_folder, window):
    for item in in_list:
        window["-OUTPUT-"].update(f"*** Combining Worksheets from {Path(item).stem} ***")
        window.refresh()
        try:
            df = pd.concat(pd.read_excel(item, sheet_name=None), ignore_index=True)
            name = Path(item).stem + "_combined"

            if csv:
                convert_to_csv(output_folder, name, window, df)
            if xls:
                convert_to_xls(output_folder, name, window, df)

        except:
            window["-OUTPUT-"].update(f"*** Error combining Worksheets from {Path(item).stem} ***")

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()


def combine_and_convert_wb(in_list, csv, xls, output_folder, window, name):
    final_df = pd.DataFrame()

    for item in in_list:
        window["-OUTPUT-"].update(f"*** Loading File {Path(item).stem} ***")
        window.refresh()
        try:
            df = pd.read_excel(item)
            final_df = final_df._append(df, ignore_index=True)

            if csv:
                convert_to_csv(output_folder, name, window, final_df)
            if xls:
                convert_to_xls(output_folder, name, window, final_df)

        except:
            window["-OUTPUT-"].update(f"*** Error loading File {Path(item).stem} ***")

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()
