import numpy as np
import pandas as pd
import itertools
import matplotlib.pyplot as plt
import Functions_GUI.buttons as bt


def ramzellen(data_id: pd.DataFrame):
    """
    This function creates Ramzellen Graphics for given data_id dataframe.

    :param data_id: Dataframe which contains vehicle data for corresponding diagrashot.
    :return: None
    """
    # Path generation for results
    ramzellen_path = bt.save_path.joinpath('Ramzellen')
    kilometer_path = bt.save_path.joinpath('Kilometer')
    if not ramzellen_path.exists():
        ramzellen_path.mkdir()
    if not kilometer_path.exists():
        kilometer_path.mkdir()

    # Excel list extraction
    config_label_list = bt.config_excel_list[3].loc[:, 'Label'].to_list()
    additional_graph_labels = ['GEMInMnf_arThrValAdpObs_VW', 'GEMInVlvSoot_facAdpInt_VW']  # Additional Scatter Graphs to the Ram-zellen function
    config_darlegung_list = bt.config_excel_list[3].loc[:, 'Darlegung'].to_list()
    config_darlegung_dict = {key: '' for key in config_label_list}
    for i, darlegung in enumerate(config_darlegung_list):
        config_darlegung_dict[config_label_list[i]] = darlegung

    ramzellen_columns = ['Dateiname', 'Fahrzeugname', 'Software/Datenstand', 'Datum', 'Km-Stand']
    ramzellen_table = pd.DataFrame(columns=ramzellen_columns.__add__(config_label_list))  # Ramzellen data for every diagrashot
    
    # Ramzellen data extraction loop
    for i, shot in enumerate(bt.diagra_list):  # iumpr type and start index detection
        ramzellen_record = {key: '' for key in ramzellen_columns.__add__(config_label_list)}
        start_index = shot.where(shot[0] == 'RAM cells').last_valid_index()
        if not start_index:
            start_index = shot.where(shot[0] == 'Ramzellen').last_valid_index()
        if not start_index:  # No Ramzellen Entry found
            data_id.drop(index=i)
            continue
        start_index += 2  # Index where Ramzellen data starts

        # Ramzellen entry from corresponding shot
        ramzellen_entry = shot.loc[start_index:, 0]

        # Data Acquisition from data_id
        ramzellen_record['Dateiname'] = data_id.loc[i, 'file_name']
        ramzellen_record['Fahrzeugname'] = data_id.loc[i, 'vehicle_name']
        ramzellen_record['Software/Datenstand'] = data_id.loc[i, 'software/data_status']
        ramzellen_record['Datum'] = data_id.loc[i, 'date']
        ramzellen_record['Km-Stand'] = data_id.loc[i, 'km-state']

        # Labeled data value acquisition
        for label in config_label_list:
            temp_label_value = None
            for row in ramzellen_entry:
                if label in row:
                    split_row = row.split()
                    if split_row[0] == label:
                        try:
                            value_index = split_row.index('=') + 1
                        except ValueError:
                            value_index = None

                        if value_index:
                            try:
                                temp_label_value = float(split_row[value_index])
                            except ValueError:
                                temp_label_value = None
                        else:
                            temp_label_value = None
                        break
            ramzellen_record[label] = temp_label_value
        ramzellen_table = ramzellen_table.append(pd.DataFrame(ramzellen_record, index=[i]))

    ramzellen_table = ramzellen_table.dropna(axis='columns', how='all')
    ramzellen_table.reset_index(drop=True, inplace=True)
    ramzellen_label = ramzellen_table.columns.to_list()

    ##### Graph Output #####
    unique_vehicle = ramzellen_table.loc[:, 'Fahrzeugname'].unique()
    for ln, label in enumerate(ramzellen_label[5:]):  # Main loop for unique vehicles within dataset
        # Graphic generation
        color_cycle = itertools.cycle(['k', 'b', 'r', 'g', 'y', 'm', 'c', 'darkorange', 'brown', 'olive', 'pink', 'limegreen', 'lightyellow', 'cyan', 'purple', 'gold'])
        markers = itertools.cycle(['o', 's', 'D', 'v', '^', '>', '<', 'p', '*', 'x', 'h', 'H', '+', '|', '_', 'd'])
        color_hold = []
        marker_hold = []

        # Regular graphic generation
        evaluation_graphic = plt.figure(figsize=(15, 5), facecolor='w')
        if label not in additional_graph_labels:
            plt.grid(True)
            plt.xlabel('Kilometer [KM]')
            plt.ylabel(label)
            for vehicle in unique_vehicle:
                filter_by_vehicle = ramzellen_table['Fahrzeugname'] == vehicle  # Making filter for unique corresponding vehicle
                color = next(color_cycle)
                marker = next(markers)
                temp_label_graph = ramzellen_table.where(filter_by_vehicle, inplace=False)[['Km-Stand', label]].dropna(axis='rows').astype(float)
                temp_label_graph.reset_index(drop=True, inplace=True)

                plt.plot(temp_label_graph['Km-Stand'], temp_label_graph[label], marker, color=color, markerfacecolor=color, linewidth=1)
                for j in range(len(temp_label_graph) - 1):  # connect dots with lines when second dot created
                    plt.plot([temp_label_graph.loc[j, 'Km-Stand'], temp_label_graph.loc[j + 1, 'Km-Stand']],
                             [temp_label_graph.loc[j, label], temp_label_graph.loc[j + 1, label]],
                             color=color, markerfacecolor=color, linewidth=1)

                color_hold.append(color)
                marker_hold.append(marker)
                plt.title('RAM-Zellen: ' + config_darlegung_dict[label], fontsize=16, color='blue')

            # Legends
            legend_list = []
            for i, vehicle_name in enumerate(unique_vehicle):
                legend_list.append(vehicle_name)
            leg = plt.legend(legend_list, loc='center left', bbox_to_anchor=(1, 0.5))
            for i, text in enumerate(leg.get_texts()):
                leg.legendHandles[i].set_marker(marker_hold[i])
                leg.legendHandles[i].set_markerfacecolor(color_hold[i])
                leg.legendHandles[i].set_color(color_hold[i])
                leg.legendHandles[i].set_linewidth(0)
                text.set_text(text.get_text())
                text.set_color('#000000')

        # Scatter graphic generation
        else:
            # Subplot 1
            plt.subplot(2, 1, 1)
            plt.grid(True)
            plt.xlabel('Kilometer [KM]')
            plt.ylabel(label)

            # Subplot 2
            plt.subplot(2, 1, 2)
            plt.grid(True)
            plt.xlabel('Fahrzeugname')
            plt.ylabel(label)

            for vehicle in unique_vehicle:
                filter_by_vehicle = ramzellen_table['Fahrzeugname'] == vehicle  # Making filter for unique corresponding vehicle
                color = next(color_cycle)
                marker = next(markers)
                temp_label_graph = ramzellen_table.where(filter_by_vehicle, inplace=False)[['Km-Stand', label]].dropna(axis='rows').astype(float)
                temp_label_graph.reset_index(drop=True, inplace=True)

                plt.subplot(2, 1, 1)
                plt.plot(temp_label_graph['Km-Stand'], temp_label_graph[label], marker, color=color, markerfacecolor=color, linewidth=1)
                for j in range(len(temp_label_graph) - 1):  # connect dots with lines when second dot created
                    plt.plot([temp_label_graph.loc[j, 'Km-Stand'], temp_label_graph.loc[j + 1, 'Km-Stand']],
                             [temp_label_graph.loc[j, label], temp_label_graph.loc[j + 1, label]],
                             color=color, markerfacecolor=color, linewidth=1)

                plt.subplot(2, 1, 2)
                plt.scatter([vehicle] * len(temp_label_graph.index), temp_label_graph[label], color=color, linewidth=1)

                color_hold.append(color)
                marker_hold.append(marker)

            plt.subplots_adjust(hspace=0.3)  # Space between subplots
            plt.suptitle('RAM-Zellen: ' + config_darlegung_dict[label], fontsize=16, color='blue')  # Main title

            # Legends
            legend_list = []
            for i, vehicle in enumerate(unique_vehicle):
                legend_list.append(vehicle)
            leg = plt.legend(legend_list, loc='center left', bbox_to_anchor=(1, 0.5))
            for i, text in enumerate(leg.get_texts()):
                leg.legendHandles[i].set_color(color_hold[i])
                leg.legendHandles[i].set_linewidth(0)
                text.set_text(text.get_text())
                text.set_color('#000000')

        # Saving corresponding plot
        if ln < 10:
            graph_path = ramzellen_path.joinpath('00' + str(ln) + '_' + 'Auswertung_Ramzellen_' + label + '.png')
        elif ln < 100:
            graph_path = ramzellen_path.joinpath('0' + str(ln) + '_' + 'Auswertung_Ramzellen_' + label + '.png')
        else:
            graph_path = ramzellen_path.joinpath(str(ln) + '_' + 'Auswertung_Ramzellen_' + label + '.png')

        plt.savefig(graph_path, bbox_inches='tight', dpi=100)
        plt.close(evaluation_graphic)

    # Table of Contents for Ramzellen Graph
    label_names = ramzellen_label[5:]  # Names of graphic labels
    number_of_toc_pages = int(len(label_names) / 20) + 1
    page_numbers = np.arange(len(label_names)) + bt.slide_positions[5] + number_of_toc_pages + 1
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
            toc_tables.append(temp_df)
            temp_df = pd.DataFrame()
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
        ppt_table.create_ppt_table(5, table)

    # Graphs to PPT
    import Functions_Basic.create_ppt_graph as ppt_graph
    ppt_graph.create_ppt_graph(ramzellen_path, 5)
