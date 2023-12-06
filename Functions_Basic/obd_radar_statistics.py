import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import Functions_GUI.buttons as bt


def obd_radar(data_id: pd.DataFrame):
    """
    This function creates OBD-Radar statistics for given data_id dataframe.

    :param data_id: Dataframe which contains vehicle data for corresponding diagrashot.
    :return: None
    """
    # Path generation for results
    obd_radar_path = bt.save_path.joinpath('OBD_Statistik')
    if not obd_radar_path.exists():
        obd_radar_path.mkdir()

    # Excel list extraction
    config_diagnose_list = bt.config_excel_list[4].loc[:, 'ID in DEX'].to_list()
    config_label_list = bt.config_excel_list[4].loc[:, 'Label'].to_list()
    config_id_list = bt.config_excel_list[4].loc[:, 'ID'].to_list()
    config_klartext_list = bt.config_excel_list[4].loc[:, 'Klartext'].to_list()
    config_schwelle_list = bt.config_excel_list[4].loc[:, 'Schwelle'].to_list()

    # Data storage
    unique_vehicles = data_id['vehicle_name'].unique()
    obd_radar_dict = {key: {key: [] for key in config_diagnose_list} for key in unique_vehicles}
    worst_values_dict = {key: {key: None for key in config_diagnose_list} for key in unique_vehicles}
    latest_shot_diagnose_dict = {key: [] for key in unique_vehicles}
    
    # OBD Radar Data Extraction Loop
    for i, shot in enumerate(bt.diagra_list):
        vehicle_name = data_id.loc[i, 'vehicle_name']
        temp_value_dict = {}  # Empty dictionary for storing obd_radar diagnose-values data
        temp_worst_values_dict = {}  # Empty dictionary for storing obd_radar diagnose - worst values
        temp_values = []  # Storing obd-radar data values for corresponding shot
        temp_worst_values = []  # Storing obd-radar worst values for corresponding shot and label
        temp_labels = []  # Storing diagnose label order for corresponding shot
        values_start_index = 0
        values_end_index = 0
        worst_values_start_index = 0
        worst_values_end_index = 0
        diagnose_label_start_index = 0
        diagnose_label_end_index = 0

        for j, value in enumerate(shot[0].values):
            if 'SysDiag_cntrObdObsvrHstg_msg' in value:
                values_start_index = j + 1
            elif 'SysDiag_dataObdObsvrFrzFrm_msg' in value:
                values_end_index = j - 1
            elif 'SysDiag_idxObdObsvrDiagChAlc' in value:
                diagnose_label_start_index = j + 1
            elif 'SysDiag_nrObdObsvrFrzRst_msg' in value:
                diagnose_label_end_index = j - 1
            elif 'SysDiag_ratObdObsvrNrmFltMax_msg' in value:
                worst_values_start_index = j + 1
            elif 'SysDiag_ratObdObsvrNrmFltMaxCyc' in value:
                worst_values_end_index = j - 1
                break

        # If there is no OBD-Radar Data entry, skip corresponding shot
        if not values_start_index \
                or not values_end_index \
                or not diagnose_label_start_index \
                or not diagnose_label_end_index \
                or not worst_values_start_index \
                or not worst_values_end_index:
            continue

        # Storing obd radar values for corresponding shot
        for value_row in shot.loc[values_start_index:values_end_index, 0].values:
            numeric_values = np.array([float(x) for x in value_row.split()])
            temp_values.extend(numeric_values)

        # Storing diagnose labels for corresponding shot
        for label_row in shot.loc[diagnose_label_start_index:diagnose_label_end_index, 0].values:
            temp_labels.extend(label_row.split())

        # Storing obd radar worst values for corresponding shot and label
        for value_row in shot.loc[worst_values_start_index:worst_values_end_index, 0].values:
            numeric_values = np.array([float(x) for x in value_row.split()])
            temp_worst_values.extend(numeric_values)

        # Matching label-values
        value_counter = 0  # Addresses correct value index
        for worst_index, label in enumerate(temp_labels):
            if 'Diagnose' not in label:
                value_counter += 10
                continue
            else:
                temp_value_dict[label] = temp_values[value_counter: value_counter + 10]
                temp_worst_values_dict[label] = temp_worst_values[worst_index]
            value_counter += 10
        
        # Selecting latest diagnose names for further filtering
        temp_diagnose_df = pd.DataFrame(temp_labels, columns=['Diagnose Name'])
        if not ('nicht_verwendet' == temp_diagnose_df['Diagnose Name'].all):
            # Selecting only Diagnose Values
            diagnose_filter = temp_diagnose_df['Diagnose Name'].str.contains('Diagnose')
            latest_shot_diagnose_dict[vehicle_name] = temp_diagnose_df[diagnose_filter]['Diagnose Name'].to_list()
            
            # Adding corresponding shot values and worst values into storage dictionaries
            for diagnose in list(temp_value_dict.keys()):
                if diagnose in config_diagnose_list:
                    obd_radar_dict[vehicle_name][diagnose] = temp_value_dict[diagnose]
                    worst_values_dict[vehicle_name][diagnose] = temp_worst_values_dict[diagnose]

    # OBD-Radar Dataframe
    df_obd_radar = pd.DataFrame(obd_radar_dict)
    df_worst_values = pd.DataFrame(worst_values_dict)
    df_worst_values.fillna(0, inplace=True)
    
    # Filter OBD-Radar Dataframe according to latest shot diagnose values
    for column in df_obd_radar.columns:
        diagnose_filter = df_obd_radar[column].index.isin(latest_shot_diagnose_dict[column])
        for index in df_obd_radar[column][~diagnose_filter].index:
            df_obd_radar.loc[index, column] = []
            df_worst_values.loc[index, column] = 0

    ##### Graph Output #####
    klartext_holder = []
    background_color_holder = []  # Holds background color information for toc table
    for dn, diagnose in enumerate(df_obd_radar.index):
        sum_all_values = sum(np.concatenate(df_obd_radar.loc[diagnose, :]))
        if not sum_all_values:  # If there is no value for diagnose or sum of all values is 0, pass
            continue

        klartext_holder.append(f'ID({str(config_id_list[dn])}) {config_klartext_list[dn]}')
        colors = ['b', 'darkorange', 'olive', 'purple', 'darkred', 'c', 'r', 'pink',
                  'cyan', 'y', 'm', 'lightblue', 'k', 'lightyellow', 'g', 'gold']

        # Graphic generation
        colors = colors[0:len(df_obd_radar.columns)]
        fig_x, fig_y = 13.33, 5.69  # Figure size
        fig, ax = plt.subplots(figsize=(fig_x, fig_y))

        ax.grid(True)
        ax.set_xlabel('Klassierung [%]')
        ax.set_ylabel('Ereigniszähler [-]')
        
        # Bar graphs
        sum_values = np.zeros(10)
        for i, vehicle in enumerate(df_obd_radar.columns):
            temp_value_list = df_obd_radar.loc[diagnose, vehicle]
            if len(temp_value_list) == 10:
                for j, value in enumerate(temp_value_list):
                    ax.bar(j, value, color=colors[i], width=0.75, bottom=sum_values[j])
                    if value:
                        ax.text(x=j, s=str(value), y=sum_values[j] + temp_value_list[j], color='white', va='top',
                                 ha='center', fontsize=8)
                sum_values += temp_value_list

        # Conditional Background coloring and x-axis labels
        faded_red = (1.0, 0.5, 0.5)  # RGB values for faded red
        faded_yellow = (1.0, 1.0, 0.5)  # RGB values for faded yellow
        custom_ticks = np.arange(10)
        
        # Schwelle Type --> SYSDIAG_STNRMTYPUPANDLO_SC --> x-axis starts from -100
        if config_schwelle_list[dn] == 'SYSDIAG_STNRMTYPUPANDLO_SC':
            # Background color
            for i in range(10):
                if i == 0 or i == 9:
                    ax.axvspan(i - 0.5, i + 0.5, color=faded_red, alpha=0.5)
                elif i == 1 or i == 8:
                    ax.axvspan(i - 0.5, i + 0.5, color=faded_yellow, alpha=0.5)

            # X-Axis labels
            custom_x_labels = ['x <= -100'].__add__(
                [f'{-100 + i * 25:.0f} < x <= {-75 + i * 25:.0f}' for i in range(8)]).__add__(['x > 100'])

            # Background color for TOC Table
            if not sum_values[0] == 0 or not sum_values[9] == 0:
                background_color_holder.append('red')
            elif not sum_values[1] == 0 or not sum_values[8] == 0:
                background_color_holder.append('yellow')
            else:
                background_color_holder.append('green')
        # Schwelle Type --> SYSDIAG_STNRMTYPUP_SC --> x-axis starts from 0
        else:
            # Background color
            for i in range(10):
                if i == 7 or i == 8:
                    ax.axvspan(i - 0.5, i + 0.5, color=faded_yellow, alpha=0.5)
                elif i == 9:
                    ax.axvspan(i - 0.5, i + 0.5, color=faded_red, alpha=0.5)

            # X-Axis labels
            custom_x_labels = ['x <= 0'].__add__(
                [f'{i * 12.5:.1f} < x <= {i * 12.5 + 12.5:.1f}' for i in range(8)]).__add__(['x > 100'])

            # Background color for TOC Table
            if not sum_values[9] == 0:
                background_color_holder.append('red')
            elif not sum_values[7] == 0 or not sum_values[8] == 0:
                background_color_holder.append('yellow')
            else:
                background_color_holder.append('green')
                
        ax.set_xticks(custom_ticks, custom_x_labels, rotation=45,
                   ha="right")  # Set custom x-axis tick positions and labels
        
        # Title
        ax.set_title(f'{config_label_list[dn]}: {config_klartext_list[dn]} ID({str(config_id_list[dn])})')
        
        # Sum values under bar
        for i, sum_value in enumerate(sum_values):
            ax.text(x=i, s=f'Σ:{sum_value}', y=0, color='white', va='bottom', ha='center', fontsize=8,
                     fontweight='bold')

        # Legends
        handles = []
        for i, vehicle_name in enumerate(unique_vehicles):
            label = f'{vehicle_name}\nWorst[%] {str(df_worst_values.loc[diagnose, vehicle_name])}'
            handles.append(mpatches.Patch(color=colors[i], label=label))
        ax.legend(handles=handles, loc='center left', bbox_to_anchor=(1, 0.5))
        
        # Saving corresponding plot
        graph_path = obd_radar_path.joinpath(f'{dn:03}_Auswertung_OBD_Radar_{diagnose}.png')  # Path generation with zero padding
        fig.savefig(graph_path, bbox_inches='tight', dpi=200)
        plt.close(fig)
        
    # Table of Contents for OBD-Radar Graph
    label_names = klartext_holder  # Names of graphic labels
    number_of_toc_pages = int(len(label_names) / 20) + 1
    page_numbers = np.arange(len(label_names)) + bt.slide_positions[6] + number_of_toc_pages + 1
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
    for i, table in enumerate(toc_tables):
        ppt_table.create_ppt_table(6, table, background_color=True,
                                   color_indexes=background_color_holder, starting_color_index=i)

    # Graphs to PPT
    import Functions_Basic.create_ppt_graph as ppt_graph
    ppt_graph.create_ppt_graph(obd_radar_path, 6, x=0, y=1.5, width=fig_x, height=fig_y)
