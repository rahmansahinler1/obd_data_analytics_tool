import pandas as pd
import matplotlib.pyplot as plt
import itertools
import datetime as dt
import Functions_GUI.buttons as bt


def used_diagra_files(data_id: pd.DataFrame, unique_vehicle: pd.DataFrame):
    """
    Generates excel table for used diagrashots within evaluation. Creates unique software name vs date graph
    Calls create_ppt_table and appends report with used diagrashot information table
    :param data_id: Vehicle diagrashot information
    :param unique_vehicle: Unique vehicle names stored
    :return: None. Creates: Verwendete_DiagRA.xlsx, Auswertung_Softwarestaende.png, Diagrashot table within report
    """
    # Create file save paths
    path = bt.save_path.joinpath('Excel Table')
    if not path.exists():
        path.mkdir()
    excel_path = path.joinpath('Verwendete_DiagRA.xlsx')
    graph_path = bt.save_path.joinpath('Softwarestaende')
    if not graph_path.exists():
        graph_path.mkdir()
    data_id.rename(columns={'file_name':'Dateiname', 'vehicle_name':'Fahrzeug', 'software/data_status':'Softwarenummer',
                            'date':'Datum', 'km-state':'Km-Stand', 'time':'Uhrzeit', 'vehicle_id':'Fahrzeug-ID'}, inplace=True)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    data_id.to_excel(writer, sheet_name="Verwendete_DiagRA", index=False, na_rep='NaN')
    
    # Auto-adjust columns' width
    for column in data_id:
        column_width = max(data_id[column].astype(str).map(len).max(), len(column))
        col_idx = data_id.columns.get_loc(column)
        writer.sheets['Verwendete_DiagRA'].set_column(col_idx, col_idx, column_width)
    writer.save()

    # Append PowerPoint report with table
    import Functions_Basic.create_ppt_table as ppt_table
    ppt_table.create_ppt_table(8, data_id)

    # Softwarestaende Graphic
    # Data gathering
    unique_software = pd.DataFrame(data_id['Softwarenummer'].unique(), columns=['Softwarenummer'])
    data_id['Softwarenummer_id'] = '0'
    for i, unique_sn in enumerate(unique_software.loc[:, 'Softwarenummer']):
        data_id.loc[data_id['Softwarenummer'] == unique_sn, 'Softwarenummer_id'] = i + 1
    unique_vehicle_list = unique_vehicle['name'].to_list()  # Unique vehicles

    # Graphic generation
    color_cycle = itertools.cycle(['k', 'b', 'r', 'g', 'y', 'm', 'c', 'darkorange', 'brown', 'olive', 'pink', 'limegreen', 'lightyellow'])
    markers = itertools.cycle(['o', 's', 'D', 'v', '^', '>', '<', 'p', '*'])
    vehicle_last_index = []  # List to store last seen vehicle index within data_id

    evaluation_graphic = plt.figure(figsize=(15, 5), facecolor='w')

    date_list = data_id['Datum'].to_list()  # Holds the date within.
    date_plot = [dt.datetime.strptime(date_str, '%d.%m.%Y').date() for date_str in date_list]  # Change 'str' to 'date'

    for i in range(len(unique_vehicle.index)):
        vehicle_last_index.append(data_id.index[data_id['Fahrzeug'] == unique_vehicle.iloc[i, 0]].max())

    software_id_list = data_id['Softwarenummer_id'].to_list()
    software_id_list = [int(sid) for sid in software_id_list]

    plot_counter = 0
    line_counter = 0

    plt.grid(True)

    color_hold = []
    marker_hold = []

    for idx in vehicle_last_index:
        color = next(color_cycle)
        marker = next(markers)
        plt.plot(date_plot[plot_counter:idx + 1], software_id_list[plot_counter:idx + 1],
                 marker, color=color, markerfacecolor=color, linewidth=1)
        for j in range(idx - plot_counter):  # connect dots with lines when second dot created
            plt.plot([date_plot[line_counter], date_plot[line_counter + 1]],
                     [software_id_list[line_counter], software_id_list[line_counter + 1]],
                     color=color, linewidth=1)
            line_counter += 1

        plot_counter = idx + 1
        line_counter += 1

        color_hold.append(color)
        marker_hold.append(marker)

    # Title, X, Y labels
    plt.title('Übersicht der Datenstände', fontsize=16, color='blue')
    plt.xticks(rotation=45)
    plt.xlabel('Datum [Jahr/Monat/Tag]')
    plt.gca().set_yticks(data_id['Softwarenummer_id'].unique())
    plt.gca().set_yticklabels(unique_software['Softwarenummer'].to_list())

    # Legends
    leg = plt.legend(unique_vehicle_list, loc='center left', bbox_to_anchor=(1, 0.5))
    for i, text in enumerate(leg.get_texts()):
        leg.legendHandles[i].set_marker(marker_hold[i])
        leg.legendHandles[i].set_color(color_hold[i])
        leg.legendHandles[i].set_markerfacecolor(color_hold[i])
        leg.legendHandles[i].set_linewidth(0)
        text.set_text(text.get_text())
        text.set_color(color_hold[i])

    # Saving corresponding plot
    plt.savefig(str(graph_path.joinpath('Auswertung_Softwarestaende.png')), bbox_inches='tight', dpi=100)
    plt.close(evaluation_graphic)

    # Graphs to PPT
    import Functions_Basic.create_ppt_graph as ppt_graph
    ppt_graph.create_ppt_graph(graph_path, 0)
