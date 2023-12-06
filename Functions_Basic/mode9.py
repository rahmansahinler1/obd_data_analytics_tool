import pandas as pd
import numpy as np
import re
import itertools
import Functions_GUI.buttons as bt
import matplotlib.pyplot as plt


def mode9(data_id: pd.DataFrame):
    """
    This function creates mode9 graphics / tables for given data_id dataframe.

    :param data_id: Dataframe which contains vehicle data for corresponding diagrashot.
    :return: None
    """
    # Path generation for results
    mode9_path = bt.save_path.joinpath('Mode9')
    excel_path = bt.save_path.joinpath('Excel Table')
    if not mode9_path.exists():
        mode9_path.mkdir()
    if not excel_path.exists():
        excel_path.mkdir()

    # Excel list extraction
    config_label_list = bt.config_excel_list[5].loc[:, 'Label'].to_list()

    # Mode9 Data Extraction Loop
    columns = ['Fahrzeugname', 'Software/Datenstand', 'Spare Part Number', 'Datum', 'Km-Stand']
    mode9 = pd.DataFrame(columns=columns.__add__(config_label_list))  # Stores mode9 data inside for every shot
    for si, shot in enumerate(bt.diagra_list):
        check_index = 0
        start_index = 0
        end_index = 0
        mode9_record = {key: [] for key in columns}  # Empty dataframe for every iumpr record within every shot
        for i, var in enumerate(shot.loc[:, 0]):
            if ('Scan-Tool Mode 9 - Fahrzeuginformation' == var) or ('Scan-Tool Mode 9 - Vehicle information' == var):
                check_index = i
                break
        if check_index:
            for i, var in enumerate(shot.loc[check_index:, 0]):
                if 'INFOTYPE 08' in var:
                    start_index = check_index + i
                elif 'INFOTYPE 0A' in var:
                    end_index = check_index + i
            if not start_index or not end_index:  # NO MODE9 ENTRY WITHIN SHOT
                # Update user, no mode9 entry
                continue

        else:  # NO MODE9 ENTRY WITHIN SHOT
            # Update user, no mode9 entry. What will we do?   ##### OPEN POINT #####
            continue

        # Mode9 Data gathering for corresponding shot
        temp_record = shot.loc[start_index + 1:end_index - 1, 0].reset_index(drop=True)

        # Data From data_id
        mode9_record['Fahrzeugname'] = data_id.loc[si, 'vehicle_name']  # Vehicle name acquisition
        mode9_record['Software/Datenstand'] = data_id.loc[si, 'software/data_status']  # File name acquisition
        mode9_record['Datum'] = data_id.loc[si, 'date']  # File name acquisition
        mode9_record['Km-Stand'] = data_id.loc[si, 'km-state']  # Km-Stand acquisition

        # Spare Part Number Acquisition
        spare_part_number = 'NaN'
        for i, var in enumerate(shot.loc[0:1000, 0]):  # Loop for spare part number acquisition
            if 'F187' in var:
                spare_part_number = var.split()[1]
                break
        if len(spare_part_number) > 1:
            mode9_record['Spare Part Number'] = spare_part_number

        # Mode9 data acquisition
        comp = ''
        cond = ''
        ratio = ''
        label = ''
        label_list = []
        for record in temp_record:
            record_list = record.split()
            if len(record_list) >= 4:
                if 'COMP' in record_list[0]:
                    mode9_record[record_list[0]] = record_list[1]
                    comp = record_list[1]
                    label = record_list[0]
                elif 'COND' in record_list[0]:
                    cond = record_list[1]
                elif 'Ratio' in record_list[3]:
                    ratio = record_list[0]
                if ratio:
                    mode9_record[label] = ratio + ' (' + comp + '/' + cond + ')'
                    label_list.append(label)
                    ratio = ''
            else:
                pass
        # Delete the labels which are not in the configuration excel
        for check_label in label_list:
            if check_label not in config_label_list:
                mode9_record.pop(check_label)

        mode9 = mode9.append(pd.DataFrame(mode9_record, index=[si]))

    # Deleting Empty Rows -- Without label value --
    mode9 = mode9.dropna(axis='rows')
    mode9 = mode9.reset_index(drop=True)
    mode9.insert(loc=0, column='Dateiname', value=(mode9.index + 1).astype(str))

    ##### Graph Output #####
    unique_vehicle = mode9.loc[:, 'Fahrzeugname'].unique()
    for vn, vehicle in enumerate(unique_vehicle):  # Main loop for unique vehicles within dataset
        filter_by_vehicle = mode9['Fahrzeugname'] == vehicle  # Making filter for unique corresponding vehicle
        temp_km_stand = mode9.where(filter_by_vehicle, inplace=False).dropna(axis='rows').reset_index(drop=True)['Km-Stand'].astype(float)  # Applying filter and extracting km-stand values for corresponding vehicle

        # Graphic generation
        color_cycle = itertools.cycle(['k', 'b', 'r', 'g', 'y', 'm', 'c', 'darkorange', 'brown', 'olive', 'pink', 'limegreen', 'lightyellow', 'cyan', 'purple', 'gold'])
        markers = itertools.cycle(['o', 's', 'D', 'v', '^', '>', '<', 'p', '*', 'x', 'h', 'H', '+', '|', '_', 'd'])

        evaluation_graphic = plt.figure(figsize=(15, 5), facecolor='w')
        plt.grid(True)
        color_hold = []
        marker_hold = []
        latest_ratio_list = []

        for label in config_label_list:
            color = next(color_cycle)
            marker = next(markers)
            plot_temp_ratio = []
            temp_ratio = mode9.where(filter_by_vehicle, inplace=False).dropna(axis='rows')[label]
            for ratio in temp_ratio:
                ratio_value = float(ratio.split()[0])
                if ratio_value == 7.995:
                    plot_temp_ratio.append(1.20)
                elif ratio_value > 1:
                    plot_temp_ratio.append(1)
                else:
                    plot_temp_ratio.append(ratio_value)

            plt.plot(temp_km_stand, plot_temp_ratio, marker, color=color, markerfacecolor=color, linewidth=1)
            for j in range(len(temp_km_stand) - 1):  # connect dots with lines when second dot created
                plt.plot([temp_km_stand[j], temp_km_stand[j + 1]],
                         [plot_temp_ratio[j], plot_temp_ratio[j + 1]],
                         color=color, markerfacecolor=color, linewidth=1)

            try:
                latest_ratio_value = float(temp_ratio.iloc[-1].split()[0])
            except ValueError:
                latest_ratio_value = 'n/a'
            except KeyError:
                latest_ratio_value = 'n/a'
            color_hold.append(color)
            marker_hold.append(marker)
            latest_ratio_list.append(latest_ratio_value)

        plt.axhline(y=0.5, color='#964B00', linestyle='--')
        plt.axhline(y=0.366, color='#FFB90F', linestyle='--')

        y_ticks = [0, 0.1, 0.33, 0.52, 0.75, 1, 1.09, 1.21]
        plt.gca().set_yticks(y_ticks)
        plt.gca().set_yticklabels(['{:.2f}'.format(y) for y in y_ticks])
        plt.gcf().set_facecolor('white')
        plt.ylim([0, 1.21])
        plt.yticks([0, 0.10, 0.33, 0.52, 0.75, 1.00, 1.09, 1.21],
                   ['0', '0.10', '0.33', '0.52', '0.75', '1.00', '4.00', '8.00'])

        plt.title('Fahrzeug: ' + vehicle + ' // Mode09 Graphic', fontsize=16, color='blue')
        plt.ylabel('Ratio')
        plt.xlabel('Kilometer [KM]')

        # Legends
        legend_list = []
        for i, label in enumerate(config_label_list):
            legend_list.append(label + ' | Letzter Wert:' + str(latest_ratio_list[i]))
        legend_list.append('0.5')
        legend_list.append('0.366')
        leg = plt.legend(legend_list, loc='center left', bbox_to_anchor=(1, 0.5))
        for i, text in enumerate(leg.get_texts()):
            if i == 6:
                leg.legendHandles[i].set_marker('.')
                leg.legendHandles[i].set_color('#964B00')
                leg.legendHandles[i].set_markerfacecolor('#964B00')
                leg.legendHandles[i].set_linewidth(0)
                text.set_text(text.get_text())
                text.set_color('#000000')
                continue
            elif i == 7:
                leg.legendHandles[i].set_marker('.')
                leg.legendHandles[i].set_color('#FFB90F')
                leg.legendHandles[i].set_markerfacecolor('#FFB90F')
                leg.legendHandles[i].set_linewidth(0)
                text.set_text(text.get_text())
                text.set_color('#000000')
                continue
            leg.legendHandles[i].set_marker(marker_hold[i])
            leg.legendHandles[i].set_color(color_hold[i])
            leg.legendHandles[i].set_markerfacecolor(color_hold[i])
            leg.legendHandles[i].set_linewidth(0)
            text.set_text(text.get_text())
            text.set_color('#000000')

        # Saving corresponding plot
        if vn < 10:
            graph_path = mode9_path.joinpath('00' + str(vn) + '_' + 'Auswertung_Mode9_' + vehicle + '.png')
        elif vn < 100:
            graph_path = mode9_path.joinpath('0' + str(vn) + '_' + 'Auswertung_Mode9_' + vehicle + '.png')
        else:
            graph_path = mode9_path.joinpath(str(vn) + '_' + 'Auswertung_Mode9_' + vehicle + '.png')
        plt.savefig(graph_path, bbox_inches='tight', dpi=100)
        plt.close(evaluation_graphic)

    ##### Excel Output #####
    path = excel_path.joinpath('Auswertung_Mode9.xlsx')

    def color_rule(val):  # Color rule logic for mode9
        styles = []  # Holds the styles
        for x in val:  # val represents the row, x represents the row value
            try:
                group = x.split()  # we take first part of every value
            except AttributeError:
                styles.append('')
                continue
            try:
                value = float(group[0])
                nom_denom_part = group[1]
                if '0/0' in nom_denom_part:
                    styles.append('')
                elif float(value) >= 0.5:
                    styles.append('background-color: green')
                elif float(value) >= 0.334:
                    styles.append('background-color: yellow')
                else:
                    styles.append('background-color: red')
            except ValueError:
                styles.append('')
        return styles

    mode9_column = mode9.style.apply(color_rule, axis=1, subset=config_label_list)

    # Convert to Excel and auto-adjust column widths
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        mode9_column.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for i, column in enumerate(mode9.columns):
            column_width = max(mode9[column].astype(str).str.len().max(), len(column))
            worksheet.column_dimensions[worksheet.cell(row=1, column=i + 1).column_letter].width = column_width + 2

    ##### Table Output #####
    import Functions_Basic.create_ppt_table as ppt_table
    ppt_table.create_ppt_table(10, mode9, True)

    # Mode9 Aktuellste Table
    mode9.pop('Dateiname')
    index_list = []
    sorted_index = pd.to_datetime(mode9['Datum'].copy(), format='%d.%m.%Y').sort_values().index.to_list()
    mode9_aktuellste = mode9.reindex(sorted_index)
    mode9_aktuellste = mode9_aktuellste.reset_index(drop=True)
    for vehicle in unique_vehicle:
        last_index = mode9_aktuellste.where(mode9_aktuellste['Fahrzeugname'] == vehicle).last_valid_index()
        index_list.append(last_index)
    mode9_aktuellste = mode9_aktuellste.iloc[index_list].reset_index(drop=True)

    # Durchschnitt(Average) and Flotte > 50%(green ratio to all) calculation
    durchschnitt_dict = {key: None for key in config_label_list}
    flotte_dict = {key: None for key in config_label_list}
    for label in config_label_list:
        nominator = 0
        denominator = 0
        durchschnitt = '0'
        flotte = 0
        for value in mode9_aktuellste.loc[:, label]:
            if not value:
                nominator += 0
                denominator += 0
                ratio = 0
            else:
                temp_nominator, temp_denominator = re.findall(r'\d+', value.split()[1])
                if temp_nominator == '0':
                    continue
                ratio = value.split()[0]
                nominator += float(temp_nominator)
                denominator += float(temp_denominator)
            if float(ratio) >= 0.5:
                flotte += 1

        if denominator:
            durchschnitt = str(round(nominator / denominator, 3))
        else:
            durchschnitt = '7.995'

        if '7.995' in durchschnitt:
            flotte = 100
        else:
            flotte = round(flotte / len(mode9_aktuellste.index) * 100, 3)
        flotte_dict[label] = str(flotte) + ' %'
        durchschnitt_dict[label] = durchschnitt + ' (' + str(round(nominator, 3)) + '/' + str(round(denominator, 3)) + ')'

    # Durchschnitt to mode9_aktuellste
    mode9_aktuellste = mode9_aktuellste.transpose()
    mode9_aktuellste.columns = mode9_aktuellste.columns + 1
    mode9_aktuellste[''] = ' '
    for index in mode9_aktuellste.index:
        if index in config_label_list:
            mode9_aktuellste.loc[index, ''] = durchschnitt_dict[index]
        elif index == 'Km-Stand':
            mode9_aktuellste.loc[index, ''] = 'Durchschnitt'
        else:
            pass

    # Flotte to mode9_aktuellste
    mode9_aktuellste[' '] = ' '
    for index in mode9_aktuellste.index:
        if index in config_label_list:
            mode9_aktuellste.loc[index, ' '] = flotte_dict[index]
        elif index == 'Km-Stand':
            mode9_aktuellste.loc[index, ' '] = 'Flotte > 50%'
        else:
            pass
    # Extra column for index names
    # mode9_aktuellste.insert(loc=0, column='Aktuellste Information', value=list(mode9_aktuellste.index))

    mode9_aktuellste_tables = []  # Stores Mode9 Table Page values inside
    cut_counter_index = 0  # cut starts from
    cut_counter = 0
    cut_column_amount = 6
    column_amount = len(mode9_aktuellste.columns) - 2
    if column_amount <= cut_column_amount:
        mode9_aktuellste_tables.append(mode9_aktuellste.iloc[:, 0:column_amount])
    else:
        for i in range(1, column_amount + 1):
            if i % cut_column_amount == 0:  # --> Every cut_column_amount row will be count as one page
                mode9_aktuellste_tables.append(mode9_aktuellste.iloc[:, cut_counter_index:i])
                cut_counter_index = i
                cut_counter += 1
            elif (cut_counter * cut_column_amount) + (i % cut_column_amount) == column_amount:
                mode9_aktuellste_tables.append(mode9_aktuellste.iloc[:, cut_counter_index:column_amount])

    # IUMPR Aktuellste to PPT
    for i, table in enumerate(mode9_aktuellste_tables):
        if i + 1 == len(mode9_aktuellste_tables):
            table[''] = mode9_aktuellste['']
            table[' '] = mode9_aktuellste[' ']
        table.insert(loc=0, column='Aktuellste Information', value=list(mode9_aktuellste.index))
        ppt_table.create_ppt_table(1, table, True, config_label_list)

    # Table of Contents for Mode9 Graphics
    label_names = unique_vehicle  # Names of graphic labels
    number_of_toc_pages = int(len(label_names) / 20) + 1
    page_numbers = np.arange(len(label_names)) + bt.slide_positions[2] + number_of_toc_pages + 1
    ### page_numbers = how much label as numeric array + where is the start slide + how many toc slides will be created
    toc = pd.DataFrame(data=label_names, columns=['Fenster Name'])
    toc['Seite'] = page_numbers.astype(str)
    toc_tables = []  # Stores TOC tables inside with 20 labels each (10 row - 10 row)
    cut_index = 0  # Counter for cut index
    temp_df = pd.DataFrame()
    for j in range(1, len(label_names) + 1):
        if j % 20 == 0:  # Completed page
            temp_df['Fenster Name '] = toc.iloc[cut_index:j, 0].values
            temp_df['Seite '] = toc.iloc[cut_index:j, 1].values
            cut_index = j
            toc_tables.append(temp_df)  # Append the toc_tables
            temp_df = pd.DataFrame()  # Initialize the temporary dataframe again
        elif j % 10 == 0:  # First part of the page
            temp_df['Fenster Name'] = toc.iloc[cut_index:j, 0].values
            temp_df['Seite'] = toc.iloc[cut_index:j, 1].values
            cut_index = j
        elif j == len(label_names):
            row_amount = j % 10
            start_row = j - row_amount
            if row_amount:
                if start_row % 20 == 0:
                    temp_df['Fenster Name'] = np.append(toc.iloc[start_row:j, 0].values, [''] * (10 - row_amount))
                    temp_df['Seite'] = np.append(toc.iloc[start_row:j, 1].values, [''] * (10 - row_amount))
                else:
                    temp_df['Fenster Name '] = np.append(toc.iloc[start_row:j, 0].values, [''] * (10 - row_amount))
                    temp_df['Seite '] = np.append(toc.iloc[start_row:j, 1].values, [''] * (10 - row_amount))
            toc_tables.append(temp_df)

    # TOC to PPT
    import Functions_Basic.create_ppt_table as ppt_table
    for table in toc_tables:
        ppt_table.create_ppt_table(2, table)

    # Graphs to PPT
    import Functions_Basic.create_ppt_graph as ppt_graph
    ppt_graph.create_ppt_graph(mode9_path, 2)
