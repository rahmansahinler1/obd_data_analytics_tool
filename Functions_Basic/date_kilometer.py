import pandas as pd
import matplotlib.pyplot as plt
from collections import Counter
import itertools
import datetime as dt
import Functions_GUI.buttons as bt


def date_kilometer(data_id: pd.DataFrame, unique_vehicle: pd.DataFrame):
    """
    Graphs mileage vs date
    :param data_id: Vehicle diagrashot information
    :param unique_vehicle: Unique vehicle names stored
    :return: Creates .png formatted graph
    """
    # Create .png output path
    path = bt.save_path.joinpath('Kilometer')
    if not path.exists():
        path.mkdir()

    # Data gathering
    vehicle_division = data_id['km-state'].to_list()  # Holds km-data for every diagrashot
    km_state = [float(km) for km in vehicle_division]  # Km data str to float conversion
    unique_vehicle_list = unique_vehicle['name'].to_list()  # Unique vehicles
    occurrence = Counter(data_id['vehicle_name'].to_list()).values()  # Unique vehicle occurrences
    color_cycle = itertools.cycle(['k', 'b', 'r', 'g', 'y', 'm', 'c', 'darkorange', 'brown', 'olive', 'pink', 'limegreen', 'lightyellow'])

    date_list = data_id['date'].to_list()  # Holds the date within.

    date_plot = [dt.datetime.strptime(date_str, '%d.%m.%Y').date() for date_str in date_list]  # Change 'str' to 'date'

    # Graphic generation
    evaluation_graphic = plt.figure(figsize=(15, 5), facecolor='w')
    plt.grid(True)
    plot_counter = 0
    line_counter = 0
    color_hold = []
    for i, occ in enumerate(occurrence):
        #  Draws dot for every kilometer data, different color for different vehicle
        color = next(color_cycle)
        plt.plot(date_plot[plot_counter:plot_counter + occ], km_state[plot_counter:plot_counter + occ],
                 'o', color=color, markerfacecolor=color, linewidth=1)
        for j in range(occ - 1):  # connect dots with lines when second dot created
            plt.plot([date_plot[line_counter], date_plot[line_counter + 1]],
                     [km_state[line_counter], km_state[line_counter + 1]],
                     color=color, linewidth=1)
            line_counter += 1
        plot_counter += occ
        line_counter = plot_counter
        color_hold.append(color)

    # Title, X, Y labels
    plt.title('Laufleistung', fontsize=16, color='blue')
    plt.xticks(rotation=45)
    plt.xlabel('Datum [Jahr/Monat/Tag]')
    plt.ylabel('Kilometerstand [km]')

    # Legends
    leg = plt.legend(unique_vehicle_list, loc='center left', bbox_to_anchor=(1, 0.5))
    for i, text in enumerate(leg.get_texts()):
        leg.legendHandles[i].set_marker('o')
        leg.legendHandles[i].set_color(color_hold[i])
        leg.legendHandles[i].set_markerfacecolor(color_hold[i])
        leg.legendHandles[i].set_linewidth(0)
        text.set_text(text.get_text())
        text.set_color(color_hold[i])

    # Saving corresponding plot
    plt.savefig(str(path.joinpath('000_date_kilometer.png')), bbox_inches='tight', dpi=100)
    plt.close(evaluation_graphic)

    # Graphs to PPT
    import Functions_Basic.create_ppt_graph as ppt_graph
    ppt_graph.create_ppt_graph(path, 0)
