import pandas as pd
import re
import Functions_GUI.buttons as bt
import re
import numpy as np
import datetime

def change_date_format(column: pd.DataFrame):
    """
    This function changes the date format from 'yyyy-mm-dd' to 'dd.mm.yyyy'.
    
    :column: Column of a dataframe
    :return: Column of a dataframe
    """
    column = column.dt.strftime('%d.%m.%Y')
    column = column.astype(str)
    return column

def validate_data(column: pd.DataFrame, name: str):
    """
    This function replaces the string values wich are not in format of 'dd.mm.yyyy' with 'NaN'.
    
    :column: Column of a dataframe
    :return: Column of a dataframe
    """
    def check_date(date_str: str):
        if not re.match(r'\d{2}\.\d{2}\.\d{4}', date_str):
            return '-'
        try:
            datetime.datetime.strptime(date_str, '%d.%m.%Y')
            return date_str
        except ValueError:
            return '-'
    
    if name == 'Datum':
        column = column.apply(check_date)
    else:
        column = column.replace({None: '-', np.nan: '-', '': '-', 'nan': '-'})
    
    return column.astype(str)

def generate_hpyerlink(column: pd.DataFrame):
    """
    This function generates hyperlinks for given column.
    
    :column: Column of a dataframe
    :return: New pandas dataframe column with hyperlinks
    """
    def add_hyperlink(name):
        url = base_folder.joinpath(name)
        return f'=HYPERLINK("{url}","Link")'
    
    base_folder = bt.default_path.joinpath('Measurements/test-error_db/')
    link_column = column.apply(add_hyperlink)
    return link_column
    

def p_code_swp_update(df_failure: pd.DataFrame):
    """
    This function updates the Fault Memory Database Excel file with readed failure memory data.
    
    :df_failure: Dataframe which contains failure memory data
    :return: None
    """
    # Fault Memory Database #
    database_excel_path = bt.default_path.joinpath('data/FaultMemoryDatabase2.xlsx')
    
    # Reading database excel
    df_failure_database = pd.read_excel(database_excel_path, sheet_name='Fehlerspeichereinträge')
      
    # Filtering df_failure with given list
    columns = ['Datum', 'Fahrzeug', 'PCode', 'Diagnosename', 'Km', 'MIL-Status', 'Occ' , 'Datenstand', 'Dateiname']
    df_failure = df_failure[columns]
    df_failure.rename(columns={'Occ': 'Occurrence'}, inplace=True)
    columns[columns.index('Occ')] = 'Occurrence'
    
    # Validate data for both dataframes
    for column in columns:
        df_failure[column] = validate_data(df_failure[column], column)
        df_failure_database[column] = validate_data(df_failure_database[column], column)           
    
    # Select 'df_failure' the rows which has not occurred in 'df_failure_database'
    # Append 'df_failure_database' with unique rows of 'df_failure'
    unique_rows = []  # List of unique rows' indexes
    unique_rows_values = []  # List of unique rows' values
    for failure_row_number in range(len(df_failure)):
        matched = False
        for database_row_number in range(len(df_failure_database)):
            if (df_failure.loc[failure_row_number, columns].values == df_failure_database.loc[database_row_number, columns].values).all():
                matched = True
                break
        if not matched:
            unique_rows.append(failure_row_number)
    
    # Append 'df_failure_database' with unique rows of 'df_failure'
    if unique_rows:
        unique_rows_values = df_failure.loc[unique_rows, columns].values
        df_failure_database = df_failure_database.append(df_failure.loc[unique_rows, :], ignore_index=True)      
            
    # Re-Sort 'df_failure_database' by given levels: date, vehicle, pcdoe
    # First change date column to datetime type for accurate sorting
    # After sorting re-validate date column
    df_failure_database['Datum'] = pd.to_datetime(df_failure_database['Datum'], format='%d.%m.%Y', errors='coerce')
    df_failure_database = df_failure_database.sort_values(by=['Datum', 'Fahrzeug', 'PCode'], ignore_index=True)
    df_failure_database['Datum'] = change_date_format(df_failure_database['Datum'])
    df_failure_database['Datum'] = validate_data(df_failure_database['Datum'], 'Datum')
    
    # Add hyperlink column to the database dataframe
    df_failure_database['Hyperlink'] = generate_hpyerlink(df_failure_database['Dateiname'])
    
    # Re'write database excel sheet with new data
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(database_excel_path, engine='xlsxwriter')
    df_failure_database.to_excel(writer, sheet_name="Fehlerspeichereinträge", index=False)

    workbook = writer.book
    worksheet = writer.sheets['Fehlerspeichereinträge']
    
    # Create a format object with orange background color
    orange_format = workbook.add_format({
        'bg_color': '#FFA500',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    light_blue_format = workbook.add_format({
        'bg_color': '#B0E0E6',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    white_format = workbook.add_format({
        'bg_color': '#FFFFFF',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})
    
    # Change background color for rows which has '-' value inside
    for row_number in range(len(df_failure_database)):
        values = df_failure_database.loc[row_number, columns].values
        matched = False
        for unique_row_value in unique_rows_values:
            if (values == unique_row_value).all():
                matched = True
                break
        if matched:
            for col_number, value in enumerate(df_failure_database.loc[row_number, :].values):
                if value == '-':
                    worksheet.write(row_number + 1, col_number, value, orange_format)
                else:
                    worksheet.write(row_number + 1, col_number, value, light_blue_format)
        else:
            for col_number, value in enumerate(df_failure_database.loc[row_number, :].values):
                if value == '-':
                    worksheet.write(row_number + 1, col_number, value, orange_format)
                else:
                    worksheet.write(row_number + 1, col_number, value, white_format)

            
    # Auto-adjust columns' width
    for column in df_failure_database:
        if column == 'Hyperlink':
            continue
        column_width = max(df_failure_database[column].astype(str).map(len).max(), len(column))
        col_idx = df_failure_database.columns.get_loc(column)
        writer.sheets['Fehlerspeichereinträge'].set_column(col_idx, col_idx, column_width)
    
    # Freeze the header row
    worksheet.freeze_panes(1, 0)
    
    # Save the workbook
    writer.save()
