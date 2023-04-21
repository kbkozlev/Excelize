import PySimpleGUI as sg
import pandas as pd
from pathlib import Path


def c(elems):
    return [sg.Stretch(), *elems, sg.Stretch()]


def border(elems):
    return sg.Frame('', [[elems]])


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
