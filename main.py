from tkinter import Tk, filedialog

import matplotlib
matplotlib.use('Qt5Agg')
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from scipy.signal import find_peaks
from itertools import tee, islice, chain

# Load configuration for the run
CONFIG = {
    "prominence": 0.08,
    "percent_drop_min_left": 10,
    "percent_drop_min_right": 10,
    "rate_of_change_left": 1,
    "fret_fall_percent": 0.3,
    "rhod_fall_percent": 0.6
}

def previous_and_next(some_iterable):
    prev, items, next = tee(some_iterable, 3)
    prev = chain([None], prev)
    next = chain(islice(next, 1, None), [None])
    return zip(prev, items, next)

def percent_change(new, previous):
    if previous == new:
        return 0
    try:
        return (abs(float(new) - previous) / previous) * 100.0 if previous != 0 else float('inf')
    except ZeroDivisionError:
        return float('inf')

def calculate_slope(point1, point2):
    x1, y1 = point1
    x2, y2 = point2
    return (y2 - y1) / (x2 - x1) if x2 != x1 else float('inf')

def find_peak_boundaries(data, peaks, fall_percentage_change):
    result = {}

    for previous, peak, next in previous_and_next(peaks):
        left_bound = previous if previous is not None else 0
        right_bound = next if next is not None else len(data) - 1

        # Look for the rise (local minimum before the peak)
        left_base = peak
        left_fall_found = False
        while left_base > left_bound and not left_fall_found:
            if percent_change(data[left_base - 1], data[peak]) <= CONFIG['percent_drop_min_left']:
                left_base -= 1
            else:
                left_fall_found = True
                while (left_base > left_bound
                       and data[left_base - 1] < data[left_base]
                       and percent_change(data[left_base - 1], data[left_base]) > CONFIG['rate_of_change_left']):
                    left_base -= 1

        # Look for the fall (local minimum after the peak)
        right_base = peak
        right_fall_found = False
        while right_base < right_bound and not right_fall_found:
            if percent_change(data[right_base + 1], data[peak]) <= CONFIG['percent_drop_min_right']:
                right_base += 1
            else:
                right_fall_found = True
                while (right_base < right_bound
                       and data[right_base + 1] < data[right_base]
                       and percent_change(data[right_base], data[right_base + 1]) > fall_percentage_change):
                    right_base += 1

        result[peak] = [left_base, right_base]

    return result

def calculate_single_col(data, time, col_number, sheet_name, fall_percentage_change):
    # Find peaks
    peaks, properties = find_peaks(data, prominence=CONFIG['prominence'])  # Increased prominence value

    peak_values = data.iloc[peaks]
    peak_times = time.iloc[peaks]

    # Plotting Peak data first.
    plt.figure(figsize=(15, 5))
    plt.plot(time, data, label='data', color='blue')
    plt.plot(peak_times, peak_values, 'rX', label='Peaks')
    plt.xlabel('Time (Min)')
    plt.ylabel('Value')

    peak_boundary_indexes = find_peak_boundaries(data, peaks, fall_percentage_change)

    amplitudes = []
    durations = []
    areas = []
    percentage_changes = []
    time_to_peak = []
    rise_rate = []
    decay_rate = []
    time_after_peak = []

    # Do something here to get the start, end value & time for each peak accurately

    for peak in peak_boundary_indexes:
        start_index = peak_boundary_indexes[peak][0]
        end_index = peak_boundary_indexes[peak][1]

        amplitudes.append(data.iloc[peak] - data.iloc[start_index])
        durations.append(time[end_index] - time[start_index])
        time_to_peak.append(time[peak] - time[start_index])
        rise_rate.append(calculate_slope((time[start_index], data[start_index]), (time[peak], data[peak])))
        decay_rate.append(calculate_slope((time[peak], data[peak]), (time[end_index], data[end_index])))
        time_after_peak.append(time[end_index])
        percentage_changes.append(
            (abs(data.iloc[end_index] - data.iloc[start_index]) / data.iloc[start_index]) * 100.0)

        # Using Simpson's rule for numerical integration
        x1, y1 = time.iloc[start_index], data.iloc[start_index]
        x2, y2 = time.iloc[end_index], data.iloc[end_index]
        peak_area = np.trapezoid(data.iloc[start_index:end_index], time.iloc[start_index:end_index])
        areas.append(peak_area)

        plt.plot([x1, x2], [y1, y2], marker='o')

    plt.title(
        f"Plot of {sheet_name} column {col_number} data vs time with Peaks and their min values marked")
    plt.grid(True)
    plt.legend()
    plt.draw()

    return pd.DataFrame({'Column No': col_number,
                         'Time of peak occurrence': peak_times[:6],
                         'Peak Values': peak_values[:6],
                         'Amplitude of Peak': amplitudes[:6],
                         'Time to Peak': time_to_peak[:6],
                         'Rate of Rise': rise_rate[:6],
                         'Rate of Decay': decay_rate[:6],
                         'Time after Peak': time_after_peak[:6],
                         'Duration of Peak': durations[:6],
                         '% Change from baseline': percentage_changes[:6],
                         'Area Under Curve': areas[:6]
                         })


def process_sheet(file_name, sheet_number, sheet_name, fall_percentage_change):
    data = pd.read_excel(file_name, sheet_name=sheet_number)
    time = data['Time (Min)']
    data_array = []
    for col_number in data.columns.values[1:]:
        col_data = data[col_number]
        # Remove the empty data from the end.
        last_idx = col_data.last_valid_index()
        fixed_col_data = col_data[:last_idx]
        fixed_col_time = time[:last_idx]

        data_array.append(calculate_single_col(fixed_col_data, fixed_col_time, col_number, sheet_name, fall_percentage_change))

    return data_array


def process_data(file_name, output_name):
    result_fret = pd.concat(process_sheet(file_name, 0, 'FRET', CONFIG['fret_fall_percent']))
    result_rhod = pd.concat(process_sheet(file_name, 1, 'RHOD', CONFIG['rhod_fall_percent']))

    with pd.ExcelWriter(output_name) as writer:
        result_fret.T.to_excel(writer, sheet_name='Fret')
        result_rhod.T.to_excel(writer, sheet_name='Rhod')


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

    process_data(file_name, output_name)
    plt.show()