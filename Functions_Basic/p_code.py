import pandas as pd
import re
import Functions_GUI.buttons as bt


def p_code(data_id: pd.DataFrame):
    """
    This function creates PCode graphics, excel table for given data_id dataframe.

    :param data_id: Dataframe which contains vehicle data for corresponding diagrashot.
    :return: None
    """
    ##### P_CODE DATAFRAME, DATA ACQUISITION ####

    # Deciding failure detection string for run type.
    if bt.basic_settings[1] == 1:  # Bosch selected
        failure_detection_str = 'Fehler gespeichert'
    elif bt.basic_settings[1] == 2:  # Conti selected
        failure_detection_str = 'fault code entries'
    else:
        # ERROR FUNCTION CALL
        return 0

    # Empty dataframe to store failure entry data
    columns = ['Dateiname', 'PCode', 'DFCC', 'Diagnosename', 'Beschreibung', 'Fehlerstatus',
               'MIL-Status', 'Fahrzeug', 'Datum', 'Km', 'Occ', 'Datenstand']
    df_failure = pd.DataFrame(columns=columns)
    failure_amount_regex = r'(\d+)'  # Regex formula to take amount of failure entry within diagrashot
    failure_detection_regex = '(\d+\\t[A-Z]\d+\w+\d+)'  # Regex formula for detection different failure entries

    # Main loop to fill failure dataframe
    for i, shot in enumerate(bt.diagra_list):
        # Loop through shots for time input. If not find, dict will be appended with nan value.
        for j, var in enumerate(shot.loc[0:500, 0]):
            # Loop through within shot, from start first index
            if failure_detection_str in var:  # Search for failure Ex: '5 Fehler gespeichert'
                failure_index = j + 1  # Failure index will be aligned with P_Code starting index.
                match_failure_amount = re.search(failure_amount_regex, var)
                if match_failure_amount:
                    failure_amount = int(match_failure_amount.group(0))
                    for k in range(failure_amount):
                        #### Data Extraction For Stationary Data ####
                        failure_entry = {key: None for key in columns}  # Temporary failure entry, it will be appended the failure_regex

                        failure_entry['Dateiname'] = data_id.loc[i, 'file_name']
                        failure_entry['Fahrzeug'] = data_id.loc[i, 'vehicle_name']
                        failure_entry['Datenstand'] = data_id.loc[i, 'software/data_status']

                        # DFCC Extraction
                        temp_string = shot.loc[failure_index, 0]
                        regex_first_number = r'(\d+)'
                        match_dfcc = re.search(regex_first_number, temp_string)
                        if match_dfcc:
                            failure_entry['DFCC'] = match_dfcc.group(0)
                        else:
                            failure_entry['DFCC'] = 'nan'

                        # PCode Extraction
                        temp_string = shot.loc[failure_index, 0]
                        regex_pcode = r'([A-Z]\d+\w+\d+)'
                        match_pcode = re.search(regex_pcode, temp_string)
                        if match_pcode:
                            if bt.basic_settings[1] == 1:  # 'Bosch' selected
                                failure_entry['PCode'] = match_pcode.group(0)
                            else:
                                failure_entry['PCode'] = match_pcode.group(0)[:-2]
                        else:
                            failure_entry['PCode'] = 'nan'

                        # Diagnosename Extraction
                        temp_string = shot.loc[failure_index + 1, 0]
                        regex_diagnosename = '^([^\\t]*)'
                        match_diagnosename = re.search(regex_diagnosename, temp_string)
                        if match_diagnosename:
                            failure_entry['Diagnosename'] = match_diagnosename.group(0)
                        else:
                            failure_entry['Diagnosename'] = 'nan'

                        # Beschreibung Extraction
                        temp_string = shot.loc[failure_index + 1, 0]
                        regex_beschreibung = '(?<=\\t)(.*)'
                        match_beschreibung = re.search(regex_beschreibung, temp_string)
                        if match_beschreibung:
                            failure_entry['Beschreibung'] = match_beschreibung.group(0)
                        else:
                            failure_entry['Beschreibung'] = 'nan'

                        # Fehlerstatus Extraction
                        temp_string = shot.loc[failure_index + 2, 0]
                        regex_fehlerstatus = '(?<=\\t)(.*)'
                        match_fehlerstatus = re.search(regex_fehlerstatus, temp_string)
                        if match_fehlerstatus:
                            failure_entry['Fehlerstatus'] = match_fehlerstatus.group(0)
                        else:
                            failure_entry['Fehlerstatus'] = 'nan'

                        # MIL-Status Extraction
                        if bt.basic_settings[1] == 1:  # 'Bosch' selected
                            mil_status_str = 'Warnlampe'
                        elif bt.basic_settings[1] == 2:  # 'Conti' selected
                            mil_status_str = 'Warning lamp'
                        else:
                            # ERROR FUNCTION CALL
                            return 0
                        if mil_status_str in shot.loc[failure_index + 4, 0]:
                            failure_entry['MIL-Status'] = shot.loc[failure_index + 4, 0]
                        elif mil_status_str in shot.loc[failure_index + 5, 0]:
                            failure_entry['MIL-Status'] = shot.loc[failure_index + 5, 0]
                        else:
                            failure_entry['MIL-Status'] = 'nan'

                        # DATE - KM - OCCURRENCE EXTRACTION
                        day, month, year = '-', '-', '-'  # Date components initialization for failure entry
                        km, occurrence = 'nan', 'nan'
                        if bt.basic_settings[1] == 1:  # 'Bosch' selected
                            failure_end_str = 'Messwerte'
                        elif bt.basic_settings[1] == 2:  # 'Conti' selected
                            failure_end_str = 'Measured values'
                        else:
                            # ERROR FUNCTION CALL
                            return 0
                        for f_index, f_row in enumerate(shot.loc[failure_index + 5:, 0]):
                            # Loop to extract non-stationary data (date, km-mileage, occurrence)
                            """
                            Breaking conditions:
                            If next failure DFCC entry found --> Go to the next failure search
                            If 'Messwerte' section is found within f_row --> End of failure entry search
                            """
                            match_next_failure = re.search(failure_detection_regex, f_row)
                            if match_next_failure:
                                failure_index = failure_index + 5 + f_index
                                break
                            if failure_end_str in f_row:
                                break
                            # Occurrence Extraction
                            if 'Occurence' in f_row:
                                match_occ = re.search(regex_first_number, f_row)
                                if match_occ:
                                    occurrence = match_occ.group(0)
                            # Km Extraction
                            if 'km-Mileage' in f_row:
                                match_km = re.search(regex_first_number, f_row)
                                if match_km:
                                    km = match_km.group(0)
                            # Day Extraction
                            if 'Day' in f_row:
                                match_day = re.search(regex_first_number, f_row)
                                if match_day:
                                    day = match_day.group(0)
                                    if len(day) < 2 and bt.basic_settings[1] == 1:  # Only for Bosch
                                        day = '0' + day
                            # Month Extraction
                            if 'Month' in f_row:
                                match_month = re.search(regex_first_number, f_row)
                                if match_month:
                                    month = match_month.group(0)
                                    if len(month) < 2 and bt.basic_settings[1] == 1:  # Only for Bosch
                                        month = '0' + month
                            # Year Extraction
                            if 'Year' in f_row:
                                match_year = re.search(regex_first_number, f_row)
                                if match_year:
                                    year = match_year.group(0)

                        if bt.basic_settings[1] == 1:  # 'Bosch' selected
                            failure_entry['Datum'] = day + '.' + month + '.' + year
                        elif bt.basic_settings[1] == 2:  # 'Conti' selected
                            failure_entry['Datum'] = month + '/' + day + '/' + year
                        else:
                            failure_entry['Datum'] = 'nan'

                        failure_entry['Occ'] = occurrence
                        failure_entry['Km'] = km
                        df_failure = df_failure.append(failure_entry, ignore_index=True)
                break

    ##### EXCEL OUTPUT CREATION ####
    # Create file save paths
    path = bt.save_path.joinpath('Excel Table')
    if not path.exists():
        path.mkdir()
    excel_path = path.joinpath('Auswertung_Fehlerspeicher.xlsx')

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    df_failure.iloc[:, 0:11].to_excel(writer, sheet_name="Auswertung_Fehlerspeicher", index=False, na_rep='NaN')

    # Auto-adjust columns' width
    for column in df_failure:
        column_width = max(df_failure[column].astype(str).map(len).max(), len(column))
        col_idx = df_failure.columns.get_loc(column)
        writer.sheets['Auswertung_Fehlerspeicher'].set_column(col_idx, col_idx, column_width)
    writer.save()

    # P_Code PPT Table Output
    import Functions_Basic.create_ppt_table as ppt_table
    ppt_table.create_ppt_table(9, df_failure.iloc[:, 0:11])

    ##### SWP EXCEL UPDATE FOR CONTI ####
    if bt.evaluation_settings[7]:  # SWP Excel Update Option Selected
        import Functions_Basic.p_code_swp_update as swp_update
        swp_update.p_code_swp_update(df_failure.copy())
