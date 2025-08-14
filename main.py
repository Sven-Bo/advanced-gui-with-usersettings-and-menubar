from pathlib import Path  # core python module

import pandas as pd  # pip install pandas openpyxl
import FreeSimpleGUI as sg  # pip install FreeSimpleGUI


def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("Filepath not correct")
    return False


def display_excel_file(excel_file_path, sheet_name):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).name
    sg.popup_scrolled(df.dtypes, "=" * 50, df, title=filename)


def convert_to_csv(excel_file_path, output_folder, sheet_name, separator, decimal):
    df = pd.read_excel(excel_file_path, sheet_name)
    filename = Path(excel_file_path).stem
    outputfile = Path(output_folder) / f"{filename}.csv"
    df.to_csv(outputfile, sep=separator, decimal=decimal, index=False)
    sg.popup_no_titlebar("Done! :)")


def settings_window(settings):
    # ------ GUI Definition ------ #
    layout = [[sg.T("SETTINGS")],
              [sg.T("Separator"), sg.I(settings["CSV"]["separator"], s=1, key="-SEPARATOR-"),
               sg.T("Decimal"), sg.Combo(settings["CSV"]["decimal"].split("|"),
                                   default_value=settings["CSV"]["decimal_default"],
                                   s=1, key="-DECIMAL-"),
               sg.T("Sheet Name:"), sg.I(settings["EXCEL"]["sheet_name"], s=20, key="-SHEET_NAME-")],
              [sg.B("Save Current Settings", s=20)]]

    window = sg.Window("Settings Window", layout, modal=True, use_custom_titlebar=True)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        if event == "Save Current Settings":
            # Write to ini file
            settings["CSV"]["separator"] = values["-SEPARATOR-"]
            settings["CSV"]["decimal_default"] = values["-DECIMAL-"]
            settings["EXCEL"]["sheet_name"] = values["-SHEET_NAME-"]

            # Display success message & close window
            sg.popup_no_titlebar("Settings saved!")
            break
    window.close()


def main_window():
    # ------ Menu Definition ------ #
    menu_def = [["Toolbar", ["Command 1", "Command 2", "---", "Command 3", "Command 4"]],
                ["Help", ["Settings", "About", "Exit"]]]


    # ------ GUI Definition ------ #
    layout = [[sg.MenubarCustom(menu_def, tearoff=False)],
              [sg.T("Input File:", s=15, justification="r"), sg.I(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),))],
              [sg.T("Output Folder:", s=15, justification="r"), sg.I(key="-OUT-"), sg.FolderBrowse()],
              [sg.Exit(s=16, button_color="tomato"),sg.B("Settings", s=16), sg.B("Display Excel File", s=16), sg.B("Convert To CSV", s=16)],]

    window_title = settings["GUI"]["title"]
    window = sg.Window(window_title, layout, use_custom_titlebar=True)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "Exit"):
            break
        if event == "About":
            window.disappear()
            sg.popup(window_title, "Version 1.0", "Convert Excel files to CSV", grab_anywhere=True)
            window.reappear()
        if event in ("Command 1", "Command 2", "Command 3", "Command 4"):
            sg.popup_error("Not yet implemented")
        if event == "Display Excel File":
            if is_valid_path(values["-IN-"]):
                display_excel_file(values["-IN-"], settings["EXCEL"]["sheet_name"])
        if event == "Settings":
            settings_window(settings)
        if event == "Convert To CSV":
            if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):
                convert_to_csv(
                    excel_file_path=values["-IN-"],
                    output_folder=values["-OUT-"],
                    sheet_name=settings["EXCEL"]["sheet_name"],
                    separator=settings["CSV"]["separator"],
                    decimal=settings["CSV"]["decimal"],
                )
    window.close()


if __name__ == "__main__":
    SETTINGS_PATH = Path.cwd()
    # create the settings object and use ini format
    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename="config.ini", use_config_file=True, convert_bools_and_none=True
    )
    theme = settings["GUI"]["theme"]
    font_family = settings["GUI"]["font_family"]
    font_size = int(settings["GUI"]["font_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()
