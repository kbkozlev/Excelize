import pandas as pd
from pathlib import Path


def is_valid_path(in_list, window):
    for item in in_list:
        if item and Path(item).exists():
            return True
        window["-OUTPUT-"].update("***Filepath not valid***")
        window.refresh()


def split_wb(in_list, csv, xls, output_folder, window):
    for item in in_list:
        window["-OUTPUT-"].update(f"*** Splitting {Path(item).stem} into Worksheets ***")
        window.refresh()
        xl = pd.ExcelFile(item)
        for filename in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=filename)
            if csv:
                outfile = Path(output_folder) / f"{filename}.csv"
                window["-OUTPUT-"].update(f"*** Converting {filename} to CSV ***")
                window.refresh()
                df.to_csv(outfile, index=False)
            if xls:
                outfile = Path(output_folder) / f"{filename}.xlsx"
                window["-OUTPUT-"].update(f"*** Converting {filename} to XLSX ***")
                window.refresh()
                df.to_excel(outfile, index=False)

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()


def combine_and_convert_ws(in_list, csv, xls, output_folder, window):
    for item in in_list:
        window["-OUTPUT-"].update(f"*** Combining Worksheets from {Path(item).stem} ***")
        window.refresh()
        df = pd.concat(pd.read_excel(item, sheet_name=None), ignore_index=True)
        filename = Path(item).stem + "_combined"
        if csv:
            outfile = Path(output_folder) / f"{filename}.csv"
            window["-OUTPUT-"].update(f"*** Converting {filename} to CSV ***")
            window.refresh()
            df.to_csv(outfile, index=False)
        if xls:
            outfile = Path(output_folder) / f"{filename}.xlsx"
            window["-OUTPUT-"].update(f"*** Converting {filename} to XLSX ***")
            window.refresh()
            df.to_excel(outfile, index=False)

    window["-OUTPUT-"].update("*** Done ***")
    window.refresh()


def combine_and_convert_wb(in_list, csv, xls, output_folder, window, name):
    final_df = pd.DataFrame()

    for item in in_list:
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

