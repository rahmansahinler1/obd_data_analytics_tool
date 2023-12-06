import pandas as pd
import numpy as np
import re
import tkinter as tk
import Functions_GUI.buttons as bt


def read_measurement_data(list2: tk.Listbox):
    """
    Reads through all diagrashots, find and returns necessary data within.
    :param list2: User update listbox.
    :return: data_id. Holds 'column_names' variables inside.
    """
    # Variable to hold basic vehicle information inside.
    column_names = ['file_name', 'vehicle_name', 'software/data_status', 'date', 'km-state', 'time', 'vehicle_id']
    data_id = pd.DataFrame(columns=column_names)
    unique_vehicle = pd.DataFrame(columns=['name'])  # Stores unique vehicle names within.

    ##### VEHICLE NAME #####
    if bt.basic_settings[1] == 1:  # If run type is 'DataType-1'
        vehicle_name = {'vehicle_name': []}  # Creating empty dict. for storing vehicle name within every shot. Only for DataType-1.
        regex_vehicle_name_dtype1 = r'([A-Za-z]{2}\d{4}\-\d{1}\-\d{4})'
        for shot in bt.diagra_list:
            for i, var in enumerate(shot.loc[:, 0]):
                if 'Fahrzeug' in var:
                    index = i + 2
                    match = re.search(regex_vehicle_name_dtype1, shot.loc[index, 0])
                    if match:
                        vehicle_name['vehicle_name'].append(match.group(0))
                        break
                    else:
                        vehicle_name['vehicle_name'].append('-')
                        bt.basic_settings[0] = 2
                        warning = 'Shot number: ' + str(i) + 'does not have vehicle name. Process will continue with VIN.'
                        bt.update_user(list2, warning)
                        break
        data_id['vehicle_name'] = pd.DataFrame(vehicle_name)

    ##### VEHICLE ID #####
    vehicle_id = {'vehicle_id': []}  # Creating empty dict. for storing vehicle id within every shot.
    regex_vehicle_id = r'(?<= \t)[^ ]+(?= )'
    for shot in bt.diagra_list:
        for i, var in enumerate(shot.loc[:, 0]):
            if 'Vehicle Identification Number' in var:
                match = re.search(regex_vehicle_id, shot.loc[i, 0])
                if match:
                    vehicle_id['vehicle_id'].append(match.group(0))
                    break
                else:
                    vehicle_id['vehicle_id'].append('-')
                    break
    data_id['vehicle_id'] = pd.DataFrame(vehicle_id)

    if bt.basic_settings[1] == 2:  # If run type is 'DataType-2' then vehicle_name == vehicle_id
        data_id['vehicle_name'] = data_id['vehicle_id']

    ##### FILE NAME #####
    data_id['file_name'] = pd.DataFrame(bt.shot_names)

    ##### DATE #####
    shot_date = {'date': []}  # Creating empty dict. for storing vehicle id within every shot.
    date_regex_dot = r'(\d{2}\.\d{2}\.\d{4})'
    date_regex_slash = r'(\d{2}\/\d{2}\/\d{4})'
    for shot in bt.diagra_list:
        if bt.basic_settings[1] == 1:  # If run type is 'DataType-1'
            for i, var in enumerate(shot.loc[0:100, 0]):
                if 'Datum' in var:
                    match_dot = re.search(date_regex_dot, shot.loc[i, 0])
                    match_slash = re.search(date_regex_slash, shot.loc[i, 0])
                    if match_dot:
                        date = match_dot.group(0)
                        shot_date['date'].append(date)
                        break
                    elif match_slash:
                        date = match_slash.group(0)
                        date = date.replace('/', '.', 2)
                        date = date[3:5] + '.' + date[0:2] + '.' + date[-4:]
                        shot_date['date'].append(date)
                        break
                    else:
                        shot_date['date'].append('-')
                        break
        elif bt.basic_settings[1] == 2:  # If run type is 'DataType-2'
            for i, var in enumerate(shot.loc[0:100, 0]):
                if 'Date' in var:
                    match_dot = re.search(date_regex_dot, shot.loc[i, 0])
                    match_slash = re.search(date_regex_slash, shot.loc[i, 0])
                    if match_dot:
                        date = match_dot.group(0)
                        shot_date['date'].append(date)
                        break
                    elif match_slash:
                        date = match_slash.group(0)
                        date = date.replace('/', '.', 2)
                        date = date[3:5] + '.' + date[0:2] + '.' + date[-4:]
                        shot_date['date'].append(date)
                        break
                    else:
                        shot_date['date'].append('-')
                        break
        else:  # Run type is not selected.
            ### RAISE ERROR ###
            pass
    data_id['date'] = pd.DataFrame(shot_date)

    ##### TIME #####
    time = {'time': []}  #  Creating empty dict. for storing every time for every shot
    for shot in bt.diagra_list:
        #  Loop through shots for time input. If not find, dict will be appended with nan value.
        date = shot[0].str.extractall(r'(\d{2}\:\d{2}\:\d{2})')
        if not date.empty:
            time['time'].append(date.iloc[0, 0])
        else:
            time['time'].append(np.nan)
    data_id['time'] = pd.DataFrame(time)

    ##### SOFTWARE / DATA STATUS #####
    software = {'software/data_status': []}  #  Empty dict for storing software and data status within every shot.
    if bt.basic_settings[1] == 1:  #  If user selected DataType-1.
        for shot in bt.diagra_list:
            #  Loop through shots for time input. If not find, dict will be appended with nan value.
            temp_str = ''  #  Empty string will have software and data status inside.
            regex_engine = '(?<=\\t)\S+(?=\s)'  #  Regex formula for System Name Or Engine Type search
            regex_spare = r'(?<= \t)[^ ]+(?= )'  #  Regex formula for VW Spare Part Number search
            for i, var in enumerate(shot.loc[:, 0]):
                if (i < 100) and ('VW System Name Or Engine Type' in var):
                    match = re.search(regex_engine, shot.loc[i, 0])
                    if match:
                        temp_str = temp_str + match.group(0)
                        break
            if not temp_str:
                temp_str = temp_str + '-'
            for i, var in enumerate(shot.loc[:, 0]):
                if (i < 100) and ('VW Application Software Version Number' in var):
                    match = re.search(regex_spare, shot.loc[i, 0])
                    if match:
                        temp_str = temp_str + ' ' + match.group(0)
                        break
            if not temp_str:
                temp_str = temp_str + ' ' + '-'
            software['software/data_status'].append(temp_str)
        data_id['software/data_status'] = pd.DataFrame(software)
    elif bt.basic_settings[1] == 2:  # 'DataType-2' selected
        for shot in bt.diagra_list:
            #  Loop through shots for time input. If not find, dict will be appended with nan value.
            temp_str = ''  #  Empty string will have software and data status inside.
            regex_engine = '(?<=\\t)\S+(?=\s)'  #  Regex formula for System Name Or Engine Type search
            regex_spare = '(?<=\\t)\S+(?=\s)'  #  Regex formula for VW Spare Part Number search
            for i, var in enumerate(shot.loc[:, 0]):
                if (i < 100) and ('VW Application Software Version Number' in var):
                    match = re.search(regex_engine, shot.loc[i, 0])
                    if match:
                        temp_str = temp_str + match.group(0)
                        break
            if not temp_str:
                temp_str = temp_str + '-'
            for i, var in enumerate(shot.loc[:, 0]):
                if (i < 100) and ('VW Spare Part Number' in var):
                    match = re.search(regex_spare, shot.loc[i, 0])
                    if match:
                        temp_str = temp_str + ' ' + match.group(0)
                        break
            if not temp_str:
                temp_str = temp_str + ' ' + '-'
            software['software/data_status'].append(temp_str)
        data_id['software/data_status'] = pd.DataFrame(software)
    else:
        data_id['software/data_status'] = 'could not be acquired due to run type selection (DataType-1 / DataType-2)'

    ##### KM STATE #####
    km_state = {'km-state': []}  # Creating empty dict. for storing vehicle name within every shot.
    regex_km = r'(\d*\.\d*)'
    km_type = ['PID A6', 'Dspl_lMlg', 'DIST_CAN', 'VehicOdome.VehicOdomeReadi']
    for shot in bt.diagra_list:
        status = 0
        for km_index in km_type:
            for i, var in enumerate(shot.loc[:, 0]):
                if km_index in var:
                    match = re.search(regex_km, shot.loc[i, 0])
                    if match:
                        km_state['km-state'].append(match.group(0))
                        status = 1
                        break
            if status:
                break
            if km_index == 'VehicOdome.VehicOdomeReadi' and status == 0:
                km_state['km-state'].append('-')
                break
    data_id['km-state'] = pd.DataFrame(km_state)

    ##### UNIQUE VEHICLE NAME #####
    unique_vehicle['name'] = pd.DataFrame(data_id['vehicle_name'].unique())

    ##### Vehicle Name Type Check ####
    if bt.basic_settings[0] == 2:  # 'VIN' selected. DiagRA order should be based on vehicle name.
        """
        Takes every unique vehicle index and reorder the data_id correspond to this order.
        """
        unique_vehicle_index = {key: [] for key in unique_vehicle.iloc[:, 0]}
        for i, value in enumerate(data_id.loc[:, 'vehicle_name']):
            unique_vehicle_index[value].append(i)
        index_order = []
        for value in unique_vehicle_index.values():
            for idx in value:
                index_order.append(idx)
        data_id = data_id.iloc[index_order]
        data_id = data_id.reset_index(drop=True)
    return data_id, unique_vehicle
