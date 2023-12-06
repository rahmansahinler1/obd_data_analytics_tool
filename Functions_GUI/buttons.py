import pandas
import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import *
import tkinter.messagebox
import pathlib
import openpyxl

# Global Path Information
default_path = pathlib.Path().resolve()  # Path for default output.
excel_path = ""  # Configuration Excel Path
save_path = ""  # Save path for output
diagra_path = pd.DataFrame(columns=['Path'])  # Stores every diagrashot path in it. Updates itself while buttons activated.

# Global Data Storing Variables
shot_names = []  # Stores diagrashot data name within.
diagra_list = []  # Stores every readed diagrashot information in every index value parallel to the listbox1 and diagra_path.
config_excel_list = []  # Stores Configuration Excel sheets as dataframe input.
basic_settings = []  # Stores basic settings selections of user
evaluation_settings = []  # Stores evaluation settings of user (which function will be used for evaluation.)
slide_positions = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]  # Stores slide positions for sections. Updates within functions after appending ppt.

"""
Every index of slide_positions represents a section for output ppt.
Output sections as follows:
 - Laufleistung über Zeit / Fahrzeugübersicht // starts at 3  // index 0
 - Mode 9 - Aktuellste // starts at 4  // index 1
 - Mode 9 - Grafiken // starts at 5  // index 2
 - IUMPR - Aktuellste // starts at 6   // index 3
 - IUMPR – Grafiken // starts at 7  // index 4
 - RAM Zellen // starts at 8  // index 5
 - OBD Radar Statistik // starts at 9  // index 6
 - Tabellen // starts at 10  // index 7
 - Verwendete DiagRA-Dateien // starts at 11  // index 8
 - P-Code - Fehlerspeicher // starts at 12  // index 9
 - Mode 9 // starts at 13  // index 10
 - IUMPR // starts at 14  // index 11
"""
ppt_path = None  # ppt report path. Updates within ppt functions.


def update_user(list2: tk.Listbox, message:str):
    """
    This function is used in other functions to make update on GUI.
    :param list2: Widget for user update on GUI
    :param message: Update message
    :return: Update the user.
    """
    list2.see("end")
    list2.insert(END, message)
    list2.update()
    list2.see("end")


def select_excel(text1: tk.Text):
    """
    Selects the configuration excel, reads the data inside.
    :return:
    Updates the global excel configuration dataframe.
    """
    global excel_path  # Configuration excel file path
    # Setting global excel path value
    excel_path = ""
    excel_path = filedialog.askopenfilename(title='Please select Excel File to read',
                                            filetypes=[("Excel files", ".xlsx .xls .xlsm")],
                                            initialdir=default_path)
    last_index = excel_path.rfind("/")

    # Writing file names into text boxes
    if excel_path:
        text1.configure(state="normal")
        text1.delete('1.0', END)
        text1.insert("end-1c", excel_path[last_index + 1:len(excel_path)])
        text1.configure(state="disabled")


def select_save_folder(text2: tk.Text):
    """
    Selects the save path for output.

    :return:
    Updates the global save path value.
    """
    global save_path  # Output save path.

    # Setting global excel path value
    save_path = ""
    save_path = filedialog.askdirectory(title="Select the save path",
                                        initialdir=default_path)
    # Writing folder names into text boxes
    if save_path:
        text2.configure(state="normal")
        text2.delete('1.0', END)
        text2.insert("end-1c", save_path)
        text2.configure(state="disabled")


def set_save_default(text2: tk.Text):
    """
    Checks the main folder, if Output folder exists selects is as default save path, if not creates one.

    :return:
    Output folder path.
    """
    global save_path  # Output save path.

    # Checking Output path exists
    default_path_tmp = default_path.joinpath('Output')
    if not default_path_tmp.exists():
        default_path_tmp.mkdir()
        tkinter.messagebox.showinfo("Dauerlauf Tool", "Output directory created in main folder")

    save_path = default_path_tmp
    # Writing folder names into text boxes
    if save_path:
        text2.configure(state="normal")
        text2.delete('1.0', END)
        text2.insert("end-1c", str(save_path))
        text2.configure(state="disabled")


def import_diagra_shots(root, list1: tk.Listbox, label2: tk.Label):
    """
    Imports diagrashots. Diagrashot data stores into dataframes.

    :return:
    dataframes for each imported diagrashot.txt file
    """
    global diagra_path
    files = filedialog.askopenfilename(parent=root,
                                       title='Select DiagRA Shots',
                                       filetypes=[('text files', 'txt')],
                                       multiple=True,
                                       initialdir=default_path)
    # Reading diagra file information.
    temp_path = pd.DataFrame(files, columns=['Path'])
    # Inserting diagra names into listbox
    for i in temp_path.index:
        path = temp_path.loc[i, 'Path']
        last_index = path.rfind("/")
        shot_name = path[last_index + 1:len(path)]
        list1.insert(END, shot_name)
        shot_names.append(shot_name)

    diagra_path = diagra_path.append(temp_path, ignore_index=True)
    # Update the amount of diagrashot label.
    label2.config(text="Amount of DiagRA Shots: " + str(list1.size()))


def delete_diagra_shot(list1: tk.Listbox, label2: tk.Label):
    """
    Delete imported diagrashots. Enables to edit the diagrashots before reading process.

    :return:
    None. Updates the diagrashot list.
    """
    global diagra_path
    selected_indices = list1.curselection()  # Holds the selected values inside listbox
    for selection in sorted(selected_indices, reverse=True):
        # This loop updates the corresponding listbox and global diagra_path dataframe at the same time.
        diagra_path = diagra_path[~diagra_path['Path'].str.contains(list1.get(selection))]
        diagra_path = diagra_path.reset_index(drop=True)
        list1.delete(selection)

    # Update the amount of diagrashot label.
    label2.config(text="Amount of DiagRA Shots: " + str(list1.size()))


def delete_all_shots(list1: tk.Listbox, label2: tk.Label):
    """
    Delete all imported diagrashots. Enables to edit the diagrashots before reading process.

    :return:
    None. Updates the diagrashot list.
    """
    global diagra_path
    # Delete whole diagrashot labels inside the list and diagra_path values.
    if not diagra_path.empty:
        # if the dataframe is not empty, delete all values
        diagra_path.drop(diagra_path.index, inplace=True)
        # reset the index to start from 0
        diagra_path.reset_index(drop=True, inplace=True)
    list1.delete(0, tk.END)
    # Update the amount of diagrashot label.
    label2.config(text="Amount of DiagRA Shots: " + str(list1.size()))


def button_start(list2: tk.Listbox,
                 button_start: tk.Button,button_delete: tk.Button, button_delete_all: tk.Button, button_import: tk.Button,
                 button_excel: tk.Button, button_save: tk.Button, button_save_default: tk.Button,
                 button_reset: tk.Button, button_stop: tk.Button,
                 option_var1: tk.IntVar, option_var2: tk.IntVar, option_var3: tk.IntVar, option_var4: tk.IntVar,
                 checkbox_vars: list):
    """
    Reads out the all necessary information which user is selected on to the GUI. Stores them into global variables for later use.

    :return:
    Updates the diagra_list
    Updates the config_excel_list
    Updates the basic_settings
    Updates the evaluation settings
    """
    global diagra_list
    global diagra_path
    global excel_path
    global config_excel_list
    global basic_settings
    global evaluation_settings

    # Initializing values for basic and evaluation settings.
    basic_settings = [0, 0, 0, 0]  # Initialize values for every selection. 0 means that selection haven't been made. 1 for first, 2 for second selection.
    evaluation_settings = [0, 0, 0, 0, 0, 0, 0, 0]  # Initialize values for function selection. 0 Means not checked. 1 for selected 0 for not selected.

    # Take and store settings if user selected correctly.
    basic_settings[0] = option_var1.get()
    basic_settings[1] = option_var2.get()
    basic_settings[2] = option_var3.get()
    basic_settings[3] = option_var4.get()
    for i, var in enumerate(checkbox_vars):
        evaluation_settings[i] = var.get()

    # Disable and update buttons which won't be using in the evaluation.
    # It prevents bugs and freezing condition of tool.
    button_start.config(state="disabled")
    button_start.update()
    button_delete.config(state="disabled")
    button_delete.update()
    button_delete_all.config(state="disabled")
    button_delete_all.update()
    button_import.config(state="disabled")
    button_import.update()
    button_save.config(state="disabled")
    button_save.update()
    button_excel.config(state="disabled")
    button_excel.update()
    button_save_default.config(state="disabled")
    button_save_default.update()

    # Enable Stop and Reset buttons. They will be activated when the run starts.
    button_stop.config(state="normal")
    button_stop.update()
    button_reset.config(state="normal")
    button_reset.update()

    # Check if dataframe paths is empty or not.
    if not diagra_path.empty:
        update_user(list2, 'DiagRA Shots reading process started')
        progress_update = 10
        for idx, row in enumerate(diagra_path['Path']):
            df = pd.read_fwf(row, infer_nrows=20000,  header=None)
            diagra_list.append(df.iloc[:,0].to_frame())
            progress = idx / len(diagra_path) * 100
            if str(progress)[:2] == str(progress_update):
                update_user(list2, 'DiagRA Shots reading... %' + str(progress_update))
                progress_update += 10
        update_user(list2, 'DiagRA Shots reading... %100\u2713')
    else:
        tkinter.messagebox.showerror('Error!', 'Please import the DiagRA Shots first.')

    #  Deleting dataframe for further use of variable
    del df
    # Check if excel path is empty or not.
    if excel_path:
        # Load the configuration workbook
        workbook = openpyxl.load_workbook(excel_path)
        # Get the number of sheets within the workbook
        sheet_names = workbook.sheetnames
        num_sheets = len(sheet_names)
        for i in range(num_sheets):
            df = pandas.read_excel(io=excel_path, sheet_name=i)
            config_excel_list.append(df.copy())
        update_user(list2, "Configuration Excel reading process has finished successfully\u2713")
    else:
        tkinter.messagebox.showerror('Error!', 'Please select the Configuration Excel first.')

    import Functions_Basic.main_run as main_run
    main_run.main(list2)
