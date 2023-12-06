import numpy as np
import pandas as pd
import re
import itertools
import matplotlib.pyplot as plt
import Functions_GUI.buttons as bt


def iumpr(data_id: pd.DataFrame):
    """
    This function creates iumpr graphics / tables for given data_id dataframe.

    :param data_id: Dataframe which contains vehicle data for corresponding diagrashot.
    :return: None
    """
    # Path generation for results
    iumpr_path = bt.save_path.joinpath('IUMPR')
    excel_path = bt.save_path.joinpath('Excel Table')
    if not iumpr_path.exists():
        iumpr_path.mkdir()
    if not excel_path.exists():
        excel_path.mkdir()

    # Excel list extraction
    config_label_list = bt.config_excel_list[1].loc[:, 'Label'].to_list()

    # IUMPR Data Extraction Loop
    iumpr_record_columns = ['Index', 'Group', 'FID', 'DFC', 'Numerator', 'Denominator', 'Ratio', 'Status']
    iumpr_table_columns = ['Dateiname', 'Fahrzeugname', 'Software/Datenstand', 'Spare Part Number', 'Datum', 'Km-Stand']
    iumpr_list = []
    spare_part_list = []
    iumpr_labels = []  # Stores used IUMPR labels inside
    iumpr_label_check = []  # Checks the IUMPR labels are added correctly
    for si, shot in enumerate(bt.diagra_list):  # iumpr type and start index detection
        start_index = 0
        end_index = 0
        iumpr_record = {key: [] for key in iumpr_record_columns}  # Empty dataframe for every iumpr record within every shot
        for i, var in enumerate(shot.loc[:, 0]):
            if ('IUMPR table' == var) or ('IUMPR-Tabelle' == var):
                start_index = i
                break

        if start_index:  # table or tabelle
            for i, var in enumerate(shot.loc[start_index:, 0]):
                if ('Conditions for general denominator' == var) or ('Bedingungen General Denominator' == var):
                    end_index = start_index + i
                    break
        if not start_index or not end_index:
            data_id.drop(si, inplace=True)
            continue  # No IUMPR data for corresponding shot

        # Spare Part Number Acquisition
        spare_part_number = None
        for i, var in enumerate(shot.loc[0:1000, 0]):  # Loop for spare part number acquisition
            if 'F187' in var:
                spare_part_number = var.split()[1]
                break
        if len(spare_part_number) > 1:
            spare_part_list.append(spare_part_number)

        # IUMPR Data gathering for corresponding shot
        temp_record = shot.loc[start_index + 3:end_index - 1, 0].reset_index(drop=True)

        for i in range(len(temp_record.index)):
            iumpr_row = temp_record[i].split()  # IUMPR row values for corresponding shot

            # Take checking label from iumpr row
            try:
                check_label_group = iumpr_row[1]
                check_label_fid = iumpr_row[2]
                check_label_dfc = iumpr_row[3]
            except IndexError:
                check_label_group = None
                check_label_fid = None
                check_label_dfc = None

            # Separation of FLD-Unused from iumpr_list
            if (check_label_group == 'Unused') \
                    or (check_label_group.__contains__('n/a')) \
                    or (check_label_fid.__contains__('n/a')) \
                    or (check_label_dfc.__contains__('n/a')) \
                    or (check_label_dfc == 'Unused') \
                    or (check_label_group == 'Private' and check_label_dfc == 'Unused') \
                    or (check_label_fid == 'FId_Unused') \
                    or len(iumpr_row) < 3:
                continue

            # Separate the data if 'FID' is not in config_label_list
            if iumpr_row[2] not in config_label_list:
                continue

            # If label is not added to the label list, add it with adjustment
            if iumpr_row[2] not in iumpr_label_check:
                if bt.basic_settings[1] == 1:  # If DataType-1 selected
                    iumpr_labels.append(iumpr_row[3] + ' // ' + iumpr_row[2])  # DataType-1 adjustment - DFC // FID
                    iumpr_label_check.append(iumpr_row[2])
                else:  # DataType-2 selected. #### OPEN POINT -- > What to do with DataType-2 Labels? ####
                    iumpr_labels.append(iumpr_row[2])
                    iumpr_label_check.append(iumpr_row[2])

            # Assign values to iumpr_record dictionary
            for j in range(len(iumpr_record)):
                if iumpr_record_columns[j] == 'Status':
                    status = iumpr_row[-1] if len(iumpr_row) == 7 else iumpr_row[-2] + ' ' + iumpr_row[-1]
                    iumpr_record[iumpr_record_columns[j]].append(status)
                    break
                else:
                    iumpr_record[iumpr_record_columns[j]].append(iumpr_row[j])

        # IUMPR Data list
        iumpr_list.append(pd.DataFrame(iumpr_record))

    # Initialize IUMPR Table
    iumpr_table_columns.extend(iumpr_labels)
    iumpr_table = pd.DataFrame(columns=iumpr_table_columns)
    data_id.reset_index(drop=True, inplace=True)

    # IUMPR Table Filling Main Loop
    for i, iumpr_entry in enumerate(iumpr_list):
        temp_iumpr = {key: [None] for key in iumpr_table_columns}
        temp_iumpr['Dateiname'] = data_id.loc[i, 'file_name']
        temp_iumpr['Fahrzeugname'] = data_id.loc[i, 'vehicle_name']
        temp_iumpr['Software/Datenstand'] = data_id.loc[i, 'software/data_status']
        temp_iumpr['Datum'] = data_id.loc[i, 'date']
        temp_iumpr['Km-Stand'] = data_id.loc[i, 'km-state']
        temp_iumpr['Spare Part Number'] = spare_part_list[i]

        for j in range(len(iumpr_entry)):
            numerator = iumpr_entry.loc[j, 'Numerator']
            denominator = iumpr_entry.loc[j, 'Denominator']
            ratio = iumpr_entry.loc[j, 'Ratio']
            fid = iumpr_entry.loc[j, 'FID']
            dfc = iumpr_entry.loc[j, 'DFC']
            label = dfc + ' // ' + fid
            if label in iumpr_labels:
                temp_iumpr[label] = ratio + ' (' + numerator + '/' + denominator + ')'

        iumpr_table = iumpr_table.append(pd.DataFrame(temp_iumpr, index=[i]))

    iumpr_table.dropna(axis='rows', subset=iumpr_labels, inplace=True, how='all')
    iumpr_table.reset_index(drop=True, inplace=True)
    iumpr_table_t = iumpr_table.transpose()  # Transposed IUMPR Table

    ##### Graph Output #####
    unique_vehicle = iumpr_table.loc[:, 'Fahrzeugname'].unique()
    for ln, label in enumerate(iumpr_labels):  # Main loop for unique vehicles within dataset
        # Graphic generation
        color_cycle = itertools.cycle(['k', 'b', 'r', 'g', 'y', 'm', 'c', 'darkorange', 'brown', 'olive', 'pink', 'limegreen', 'lightyellow', 'cyan', 'purple', 'gold'])
        markers = itertools.cycle(['o', 's', 'D', 'v', '^', '>', '<', 'p', '*', 'x', 'h', 'H', '+', '|', '_', 'd'])

        evaluation_graphic = plt.figure(figsize=(15, 5), facecolor='w')
        plt.grid(True)
        color_hold = []
        marker_hold = []
        latest_ratio_list = []

        for vehicle in unique_vehicle:
            filter_by_vehicle = iumpr_table['Fahrzeugname'] == vehicle  # Making filter for unique corresponding vehicle
            color = next(color_cycle)
            marker = next(markers)
            plot_temp_ratio = []

            temp_label_graph = iumpr_table.where(filter_by_vehicle, inplace=False)[['Km-Stand', label]].dropna(axis='rows')
            temp_label_graph.reset_index(drop=True, inplace=True)

            for ri, ratio in enumerate(temp_label_graph.loc[:, label]):
                try:
                    ratio_value = float(ratio.split()[0])
                except ValueError:
                    ratio_value = None

                if not ratio_value:
                    temp_label_graph.drop(ri, inplace=True)
                elif ratio_value == 7.995:
                    plot_temp_ratio.append(1.20)
                elif ratio_value > 1:
                    plot_temp_ratio.append(1)
                else:
                    plot_temp_ratio.append(ratio_value)

            temp_km_graph = temp_label_graph['Km-Stand'].astype(float)
            temp_km_graph.reset_index(drop=True, inplace=True)
            plt.plot(temp_km_graph, plot_temp_ratio, marker, color=color, markerfacecolor=color, linewidth=1)
            for j in range(len(temp_km_graph) - 1):  # connect dots with lines when second dot created
                plt.plot([temp_km_graph[j], temp_km_graph[j + 1]],
                         [plot_temp_ratio[j], plot_temp_ratio[j + 1]],
                         color=color, markerfacecolor=color, linewidth=1)

            try:
                latest_ratio_value = temp_label_graph[label].loc[temp_label_graph[label].last_valid_index()]
            except ValueError:
                latest_ratio_value = 'n/a'
            except KeyError:
                latest_ratio_value = 'n/a'
            color_hold.append(color)
            marker_hold.append(marker)
            latest_ratio_list.append(str(latest_ratio_value))

        plt.axhline(y=0.5, color='#964B00', linestyle='--')
        plt.axhline(y=0.366, color='#FFB90F', linestyle='--')

        y_ticks = [0, 0.1, 0.33, 0.52, 0.75, 1, 1.09, 1.21]
        plt.gca().set_yticks(y_ticks)
        plt.gca().set_yticklabels(['{:.2f}'.format(y) for y in y_ticks])
        plt.gcf().set_facecolor('white')
        plt.ylim([0, 1.21])
        plt.yticks([0, 0.10, 0.33, 0.52, 0.75, 1.00, 1.09, 1.21],
                   ['0', '0.10', '0.33', '0.52', '0.75', '1.00', '4.00', '8.00'])

        plt.title('IUMPR Diagram', fontsize=16, color='blue')
        plt.ylabel(label)
        plt.xlabel('Kilometer [KM]')

        # Legends
        legend_list = []
        for i, vehicle in enumerate(unique_vehicle):
            legend_list.append(vehicle + ' | Letzter Wert:' + latest_ratio_list[i])
        legend_list.append('0.5')
        legend_list.append('0.366')
        leg = plt.legend(legend_list, loc='center left', bbox_to_anchor=(1, 0.5))
        for i, text in enumerate(leg.get_texts()):
            if i == len(legend_list) - 2:
                leg.legendHandles[i].set_marker('.')
                leg.legendHandles[i].set_color('#964B00')
                leg.legendHandles[i].set_markerfacecolor('#964B00')
                leg.legendHandles[i].set_linewidth(0)
                text.set_text(text.get_text())
                text.set_color('#000000')
                continue
            elif i == len(legend_list) - 1:
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
        if ln < 10:
            graph_path = iumpr_path.joinpath('00' + str(ln) + '_' + 'Auswertung_IUMPR_' + label.replace('//', '--') + '.png')
        elif ln < 100:
            graph_path = iumpr_path.joinpath('0' + str(ln) + '_' + 'Auswertung_IUMPR_' + label.replace('//', '--') + '.png')
        else:
            graph_path = iumpr_path.joinpath(str(ln) + '_' + 'Auswertung_IUMPR_' + label.replace('//', '--') + '.png')
        plt.savefig(graph_path, bbox_inches='tight', dpi=100)
        plt.close(evaluation_graphic)

    ##### Excel Output ####
    path = excel_path.joinpath('Auswertung_IUMPR.xlsx')

    def color_rule(val):  # Color rule logic for iumpr
        styles = []  # Holds the styles
        for i, x in enumerate(val):  # val represents the row, x represents the row value
            if i >= 6:
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
                    elif value >= 0.5:
                        styles.append('background-color: green')
                    elif value >= 0.334:
                        styles.append('background-color: yellow')
                    else:
                        styles.append('background-color: red')
                except ValueError:
                    styles.append('')
            else:
                styles.append('')
        return styles

    iumpr_column = iumpr_table_t.style.apply(color_rule, axis=0, subset=iumpr_table_t.columns)

    # Convert to Excel and auto-adjust column widths
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        iumpr_column.to_excel(writer, index=True, sheet_name='IUMPR_Auswertung', header=False)
        worksheet = writer.sheets['IUMPR_Auswertung']
        for i, column in enumerate(iumpr_table.columns):
            column_width = max(iumpr_table[column].astype(str).str.len().max(), len(column))
            worksheet.column_dimensions[worksheet.cell(row=1, column=i + 1).column_letter].width = column_width + 2

    ##### Table Output #####
    iumpr_tables = []  # Stores IUMPR Table Page values inside
    iumpr_table_t.columns = iumpr_table_t.columns + 1
    iumpr_table_t.drop(labels='Dateiname', inplace=True)
    cut_index = 0  # cut starts from
    cut_counter = 0
    column_amount = len(iumpr_table_t.columns)
    for i in range(1, column_amount):
        if i % 6 == 0:  # --> Every 6 row will be count as one page
            iumpr_tables.append(iumpr_table_t.iloc[:, cut_index:i])
            cut_index = i
            cut_counter += 1
        elif (cut_counter * 6) + (i % 6) == column_amount:
            iumpr_tables.append(iumpr_table_t.iloc[:, cut_index:column_amount])
    import Functions_Basic.create_ppt_table as ppt_table
    for table in iumpr_tables:
        table.insert(0, '', table.index)
        ppt_table.create_ppt_table(11, table, True, iumpr_labels)

    # IUMPR Aktuellste Table
    index_list = []
    sorted_index = pd.to_datetime(iumpr_table['Datum'].copy(), format='%d.%m.%Y').sort_values().index.to_list()
    iumpr_aktuellste = iumpr_table.reindex(sorted_index)
    iumpr_aktuellste = iumpr_aktuellste.reset_index(drop=True)
    for vehicle in unique_vehicle:
        last_index = iumpr_aktuellste.where(iumpr_aktuellste['Fahrzeugname'] == vehicle).last_valid_index()
        index_list.append(last_index)
    iumpr_aktuellste = iumpr_aktuellste.iloc[index_list].reset_index(drop=True)
    iumpr_aktuellste_t = iumpr_aktuellste.transpose()
    iumpr_aktuellste_t.columns = iumpr_aktuellste_t.columns + 1

    # Durchschnitt(Average) and Flotte > 50%(green ratio to all) calculation
    durchschnitt_dict = {key: None for key in iumpr_labels}
    flotte_dict = {key: None for key in iumpr_labels}
    for label in iumpr_labels:
        nominator = 0
        denominator = 0
        durchschnitt = '0'
        flotte = 0
        for value in iumpr_aktuellste_t.loc[label, :]:
            if not value:
                nominator += 0
                denominator += 0
                ratio = 0
            else:
                nominator_denominator_part = re.findall(r'\d+', value.split()[1])
                if nominator_denominator_part[0] == '0':
                    continue
                nominator += float(nominator_denominator_part[0] + '.' + nominator_denominator_part[1])
                denominator += float(nominator_denominator_part[2] + '.' + nominator_denominator_part[3])
                ratio = value.split()[0]
            if float(ratio) >= 0.5:
                flotte += 1

        if denominator:
            durchschnitt = str(round(nominator / denominator, 3))
        else:
            durchschnitt = '7.995'

        if '7.995' in durchschnitt:
            flotte = 100
        else:
            flotte = round(flotte / len(iumpr_aktuellste.index) * 100, 3)
        flotte_dict[label] = str(flotte) + ' %'
        durchschnitt_dict[label] = durchschnitt + ' (' + str(round(nominator, 3)) + '/' + str(round(denominator, 3)) + ')'

    # Durchschnitt to iumpr_aktuellste
    for index in iumpr_aktuellste_t.index:
        if index in iumpr_labels:
            iumpr_aktuellste_t.loc[index, ''] = durchschnitt_dict[index]
        elif index == 'Km-Stand':
            iumpr_aktuellste_t.loc[index, ''] = 'Durchschnitt'
        else:
            pass

    # Flotte to mode9_aktuellste
    for index in iumpr_aktuellste_t.index:
        if index in iumpr_labels:
            iumpr_aktuellste_t.loc[index, ' '] = flotte_dict[index]
        elif index == 'Km-Stand':
            iumpr_aktuellste_t.loc[index, ' '] = 'Flotte > 50%'
        else:
            pass

    # Clean na
    iumpr_aktuellste_t[['', ' ']] = iumpr_aktuellste_t[['', ' ']].fillna(' ')

    iumpr_aktuellste_tables = []  # Stores IUMPR Table Page values inside
    iumpr_aktuellste_t.drop(labels='Dateiname', inplace=True)
    cut_counter_index = 0  # cut starts from
    cut_counter = 0
    cut_column_amount = 4
    column_amount = len(iumpr_aktuellste_t.columns) - 2
    if column_amount < cut_column_amount:
        iumpr_aktuellste_tables.append(iumpr_aktuellste_t.iloc[:, 0:column_amount])
    else:
        for i in range(1, column_amount + 1):
            if i % cut_column_amount == 0:  # --> Every cut_column_amount row will be count as one page
                iumpr_aktuellste_tables.append(iumpr_aktuellste_t.iloc[:, cut_counter_index:i])
                cut_counter_index = i
                cut_counter += 1
            elif (cut_counter * cut_column_amount) + (i % cut_column_amount) == column_amount:
                iumpr_aktuellste_tables.append(iumpr_aktuellste_t.iloc[:, cut_counter_index:column_amount])

    # IUMPR Aktuellste to PPT
    for i, table in enumerate(iumpr_aktuellste_tables):
        if i + 1 == len(iumpr_aktuellste_tables):
            table[''] = iumpr_aktuellste_t['']
            table[' '] = iumpr_aktuellste_t[' ']
        table.insert(loc=0, column='Aktuellste Information', value=list(iumpr_aktuellste_t.index))
        ppt_table.create_ppt_table(3, table, True, iumpr_labels)

    # Table of Contents for Mode9 Graphics
    label_names = iumpr_labels  # Names of graphic labels
    number_of_toc_pages = int(len(label_names) / 20) + 1
    page_numbers = np.arange(len(label_names)) + bt.slide_positions[4] + number_of_toc_pages + 1
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
        ppt_table.create_ppt_table(4, table)

    # Graphs to PPT
    import Functions_Basic.create_ppt_graph as ppt_graph
    ppt_graph.create_ppt_graph(iumpr_path, 4)
