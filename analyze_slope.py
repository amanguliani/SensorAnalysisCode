from tkinter import Tk, filedialog

import matplotlib
matplotlib.use('Qt5Agg')
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from scipy.signal import find_peaks
from itertools import tee, islice, chain

def process_columns(data, time, col_number, sheet_name):
    plt.figure(figsize=(15, 5))
    plt.plot(time, data, label='data', color='blue')
    plt.xlabel('Time (Min)')
    plt.ylabel('Value')

    time_index = input("Time index of peak?")
    # type cast into integer
    time_index = int(time_index)

    rise_rate = np.mean(np.gradient(data[:time_index], time[:time_index]))  # Compute rate of change

    plt.plot(time[time_index], data[time_index], 'rX', label='Chosen Peak')
    plt.title(
        f"Plot of {sheet_name} column {col_number} data vs time with Peaks and their min values marked")
    plt.grid(True)
    plt.legend()
    plt.draw()

    return pd.DataFrame([{
        'Rate of Rise': rise_rate,
    }])


def process_sheet(file, sheet_number, sheet_name):
    data = pd.read_excel(file, sheet_name=sheet_number)
    time = data['Time (Min)']
    data_array = []
    for col_number in data.columns.values[1:]:
        col_data = data[col_number]
        # Remove the empty data from the end.
        last_idx = col_data.last_valid_index()
        fixed_col_data = col_data[:last_idx]
        fixed_col_time = time[:last_idx]

        data_array.append(process_columns(fixed_col_data, fixed_col_time, col_number, sheet_name))

    return data_array


if __name__ == "__main__":
    Tk().withdraw()  # Hide the main Tkinter window
    file_name = filedialog.askopenfilename(title="Select the input file", filetypes=[("Excel files", "*.xlsx")])
    if not file_name:
        print("No file selected. Exiting.")
        exit()

    output_name = filedialog.asksaveasfilename(title="Save output file as", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if not output_name:
        print("No output file selected. Exiting.")
        exit()

    result_fret = pd.concat(process_sheet(file_name, 0, 'FRET'))

    with pd.ExcelWriter(output_name) as writer:
        result_fret.T.to_excel(writer, sheet_name='Fret')
    plt.show()