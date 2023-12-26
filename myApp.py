
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d

import PySimpleGUI as sg


# test
def plot_excel_data(excel_files):
    # Define some line styles
    line_styles = ['-', '--', '-.', ':']

    # Initialize the plot
    fig, ax = plt.subplots(figsize=(10, 5))

    # Define a list of colors that you want to use
    colors = ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9']

    x_bar = np.array([])
    y_bar = np.array([])
    bar_names = np.array([])

    # Loop through each Excel file
    for file_idx, excel_file in enumerate(excel_files):
        # Open the Excel file
        xls = pd.ExcelFile(excel_file)

        # Choose a color for this file
        color = colors[file_idx % len(colors)]
        linestyle = line_styles[file_idx % len(line_styles)]

        interpolated_ys = []

        x_min = float('inf')
        x_max = float('-inf')

        # Iterate over each sheet in the Excel file starting from the 3rd one
        for sheet_idx, sheet_name in enumerate(xls.sheet_names[3:]):  # Assuming the first two sheets are not data sheets
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[0, 1], skiprows=[0, 1])

            # Assuming the first column is X and the second column is Y
            x_min = min(x_min, df.iloc[:, 0].min())
            x_max = max(x_max, df.iloc[:, 0].max())

        x_grid = np.linspace(x_min, x_max, 1000)
        # test

        for sheet_idx, sheet_name in enumerate(xls.sheet_names[3:]):  # Assuming the first two sheets are not data sheets
            x_values = df.iloc[:,0].values
            y_values = df.iloc[:,1].values

            interpolation_function = interp1d(x_values, y_values, kind='nearest', bounds_error=False, fill_value='extrapolate')

            interpolated_y = interpolation_function(x_grid)
            interpolated_ys.append(interpolated_y)


        mean_y = np.mean(interpolated_ys, axis=0)

        x_bar = np.append(x_bar, x_max)
        y_bar = np.append(y_bar, max(mean_y))
        bar_names = np.append(bar_names, excel_file.split('/')[-1][:-5])

        n_bars = len(x_bar)
        bar_width = 0.35
        # Calculate positions for the first set of bars
        positions = np.arange(n_bars)

        # The bar plot
        if file_idx == 0:
            ax.bar(positions - bar_width/2, x_bar, width=bar_width, label='Elongation', color='skyblue')
            ax.bar(positions + bar_width/2, y_bar, width=bar_width, label='Tensile Strength', color='orange')
        else:
            ax.bar(positions - bar_width/2, x_bar, width=bar_width, color='skyblue')
            ax.bar(positions + bar_width/2, y_bar, width=bar_width, color='orange')



        # Plot each sheet's data with the same color and linestyle for the current file
#        ax.plot(x_grid, mean_y, color=color, linestyle=linestyle, label='{}'.format(excel_file.split('/')[-1][:-5]))

    # Set the position and labels of the x-ticks
    ax.set_xticks(ticks=positions, labels=bar_names, rotation=45)
    ax.set_ylabel('Elongation and Tensile Strength')
#    ax.set_xlabel('Elongation (%)')
#    ax.set_ylabel('Standard Force (MPa)')
#    ax.set_title('Data from Multiple Excel Files')
# this is nur a test
    ax.legend()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.grid(True)

    # Display the plot
    plt.show()
    #fig.savefig("test.png")
    pass

def plot_excel_data_2(excel_files):

    # Define some line styles
    line_styles = ['-', '--', '-.', ':']

    # Initialize the plot
    fig, ax = plt.subplots(figsize=(10, 5))

    # Define a list of colors that you want to use
    colors = ['C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9']

    # Loop through each Excel file
    for file_idx, excel_file in enumerate(excel_files):
        # Open the Excel file
        xls = pd.ExcelFile(excel_file)

        # Choose a color for this file
        color = colors[file_idx % len(colors)]
        linestyle = line_styles[file_idx % len(line_styles)]

        # Iterate over each sheet in the Excel file starting from the 3rd one
        for sheet_idx, sheet_name in enumerate(xls.sheet_names[3:]):  # Assuming the first two sheets are not data sheets
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[0, 1], skiprows=[0, 1])

            # Assuming the first column is X and the second column is Y
            x = df.iloc[:, 0].dropna()
            y = df.iloc[:, 1].dropna()

            # Plot each sheet's data with the same color and linestyle for the current file
            if sheet_idx == 0:
                ax.plot(x, y, color=color, linestyle=linestyle, label='{}'.format(excel_file.split('/')[-1][:-5]))
            else:
                ax.plot(x, y, color=color, linestyle=linestyle)

    ax.set_xlabel('Elongation (%)')
    ax.set_ylabel('Standard Force (MPa)')
    ax.set_title('Data from Multiple Excel Files')
    ax.legend()
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    ax.grid(True)

    # Display the plot
    plt.show()

    pass

# Choose a color theme
sg.theme('DarkTeal9')

# Define the layout of the window with a larger size for text and buttons
layout = [
    [sg.Text('Excel Plotter', font=('Any', 24), size=(30, 1), justification='center')],
    [sg.Button('Bar plot', size=(15, 2), font=('Any', 14)), sg.Button('Line plot', size=(15, 2), font=('Any', 14))],
    [sg.Button('Exit', size=(10, 2), font=('Any', 14))]
]

# Create the window with a larger size
window = sg.Window('Excel Plotter', layout, size=(500, 200))

# Event loop
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    elif event == 'Bar plot':
        filenames = sg.popup_get_file('Select Excel Files', multiple_files=True, file_types=(("Excel Files", "*.xlsx"),))
        if filenames:
            filenames = filenames.split(';')  # Filenames is a string with multiple files separated by ;
            plot_excel_data(filenames)
    elif event == 'Line plot':
        filenames = sg.popup_get_file('Select Excel Files', multiple_files=True, file_types=(("Excel Files", "*.xlsx"),))
        if filenames:
            filenames = filenames.split(';')  # Filenames is a string with multiple files separated by ;
            plot_excel_data_2(filenames)

# Close the window
window.close()
