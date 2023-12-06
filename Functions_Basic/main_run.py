import tkinter as tk
import tkinter.messagebox
import Functions_GUI.buttons as bt
import time
from datetime import datetime


def main(list2: tk.Listbox):
    """
    Main run order of the tool. Calls the functions within.

    :return:
    None.
    """
    # Creating output folder for evaluation.
    now = str(datetime.now())  # Taking current time.
    now = now[:-7]

    # Update user for current evaluation process
    time.sleep(0.3)
    bt.update_user(list2, "Evaluation Started: " + now)

    # Create result folder for output
    result_folder_name = 'Evaluation ' + now
    result_folder_name = result_folder_name.replace(':', '.')
    if bt.save_path:
        bt.save_path = bt.save_path.joinpath(result_folder_name)
        bt.save_path.mkdir()
    else:
        tk.messagebox.showerror('Error on save directory', "No save path selected or It couldn't be created")
        bt.update_user(list2, 'Error on run!, Please select correct save path')
        return None

    # Running read_measurement_data.py function. // First function of tool.
    bt.update_user(list2, " - Function on run: read_measurement_data.py")
    import Functions_Basic.read_measurement_data as rmd
    data_id, unique_vehicle = rmd.read_measurement_data(list2)
    bt.update_user(list2, " - Function run successfully completed: read_measurement_data.py \u2713")

    if bt.evaluation_settings[0]:
        # Checking Evaluation Settings for Function: date_kilometer.py
        bt.update_user(list2, " - Function on run: date_kilometer.py")  # Running date_kilometer.py function. // Second function of tool.
        import Functions_Basic.date_kilometer as dk
        dk.date_kilometer(data_id, unique_vehicle)
        bt.update_user(list2, " - Function run successfully completed: date_kilometer.py \u2713")

    # Checking Evaluation Settings for Function: used_diagra_files.py
    if bt.evaluation_settings[1]:
        bt.update_user(list2, " - Function on run: used_diagra_files.py")
        import Functions_Basic.used_diagra_files as udf
        udf.used_diagra_files(data_id.copy(), unique_vehicle)
        bt.update_user(list2, " - Function run successfully completed: used_diagra_files.py \u2713")

    # Checking Evaluation Settings for Function: p_code.py
    if bt.evaluation_settings[2]:
        bt.update_user(list2, " - Function on run: p_code.py")
        import Functions_Basic.p_code as p_code
        p_code.p_code(data_id)
        bt.update_user(list2, " - Function run successfully completed: p_code.py \u2713")

    # Checking Evaluation Settings for Function: mode9.py
    if bt.evaluation_settings[3]:
        bt.update_user(list2, " - Function on run: mode9.py")
        import Functions_Basic.mode9 as mode9
        mode9.mode9(data_id)
        bt.update_user(list2, " - Function run successfully completed: mode9.py \u2713")

    # Checking Evaluation Settings for Function: iumpr.py
    if bt.evaluation_settings[4]:
        bt.update_user(list2, " - Function on run: iumpr.py")
        import Functions_Basic.iumpr as iumpr
        iumpr.iumpr(data_id.copy())
        bt.update_user(list2, " - Function run successfully completed: iumpr.py \u2713")

    # Checking Evaluation Settings for Function: iumpr.py
    if bt.evaluation_settings[5]:
        bt.update_user(list2, " - Function on run: ramzellen.py")
        import Functions_Basic.ramzellen as ramzellen
        ramzellen.ramzellen(data_id.copy())
        bt.update_user(list2, " - Function run successfully completed: ramzellen.py \u2713")

    # Checking Evaluation Settings for Function: obd_radar.py
    if bt.evaluation_settings[6]:
        bt.update_user(list2, " - Function on run: obd_radar_statistics.py")
        import Functions_Basic.obd_radar_statistics as obd_radar
        obd_radar.obd_radar(data_id.copy())
        bt.update_user(list2, " - Function run successfully completed: obd_radar.py \u2713")

    # Run finished!
    bt.update_user(list2, 'Run successfully finished!')
