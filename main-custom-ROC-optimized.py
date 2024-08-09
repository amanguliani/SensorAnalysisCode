import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from scipy.signal import find_peaks, peak_prominences
from scipy.integrate import simps


def percent_change(new, previous):
    if previous == new:
        return 0
    try:
        return (abs(float(new) - previous) / previous) * 100.0
    except ZeroDivisionError:
        return float('inf')


def find_start_end_indexes_for_peak(data, peaks):
    result = {}

    for peak in peaks:
        # Look for the rise (local minimum before the peak)
        left_base = peak
        left_fall_found = False
        while left_base > 0 and not left_fall_found:
            if percent_change(data[left_base - 1], data[peak]) <= 10:
                left_base -= 1
            else:
                left_fall_found = True
                while (left_base > 0
                       and data[left_base - 1] < data[left_base]
                       and percent_change(data[left_base - 1], data[left_base]) > 1):
                    left_base -= 1

        # Look for the fall (local minimum after the peak)
        right_base = peak
        right_fall_found = False
        while right_base < len(data) - 1 and not right_fall_found:
            if percent_change(data[right_base + 1], data[peak]) <= 10:
                right_base += 1
            else:
                right_fall_found = True
                while (right_base < len(data) - 1
                       and data[right_base + 1] < data[right_base]
                       and percent_change(data[right_base], data[right_base + 1]) > 0.6):
                    right_base += 1

        result[peak] = [left_base, right_base]

    return result


def calculate_single_col(data, time, col_number, sheet_name):
    # Find peaks
    peaks, properties = find_peaks(data, prominence=0.1)  # Increased prominence value

    peak_values = data.iloc[peaks]
    peak_times = time.iloc[peaks]

    # Plotting Peak data first.
    plt.figure(figsize=(15, 5))
    plt.plot(time, data, label='data', color='blue')
    plt.plot(peak_times, peak_values, 'rX', label='Peaks')
    plt.xlabel('Time (Min)')
    plt.ylabel('Value')

    peak_index_map = find_start_end_indexes_for_peak(data, peaks)

    amplitudes = []
    durations = []
    areas = []
    percentage_changes = []

    # Do something here to get the start, end value & time for each peak accurately

    for peak in peak_index_map:
        start_index = peak_index_map[peak][0]
        end_index = peak_index_map[peak][1]

        amplitudes.append(data.iloc[peak] - data.iloc[start_index])
        durations.append(time[end_index] - time[start_index])
        percentage_changes.append(
            (abs(data.iloc[end_index] - data.iloc[start_index]) / data.iloc[start_index]) * 100.0)

        # Using Simpson's rule for numerical integration
        x1, y1 = time.iloc[start_index], data.iloc[start_index]
        x2, y2 = time.iloc[end_index], data.iloc[end_index]
        area_under_peak = simps(data.iloc[start_index:end_index], time.iloc[start_index:end_index]) - (
                0.5 * (x2 - x1) * (y1 + y2))
        areas.append(abs(area_under_peak))

        plt.plot([x1, x2], [y1, y2], marker='o')

    plt.title(
        f"Plot of {sheet_name} column {col_number} data vs time with Peaks and their min values marked")
    plt.grid(True)
    plt.legend()
    plt.draw()
    plt.interactive(True)

    return pd.DataFrame({'Column No': col_number,
                         'Time of peak occurrence': peak_times[:6],
                         'Peak Values': peak_values[:6],
                         'Amplitude of Peak': amplitudes[:6],
                         'Duration of Peak': durations[:6],
                         '% Change from baseline': percentage_changes[:6],
                         'Area Under Curve': areas[:6]
                         })


def process_sheet(file_name, sheet_number, sheet_name):
    data = pd.read_excel(file_name, sheet_name=sheet_number)
    time = data['Time (Min)']
    data_array = []
    for col_number in data.columns.values[1:]:
        col_data = data[col_number]
        # Remove the empty data from the end.
        last_idx = col_data.last_valid_index()
        fixed_col_data = col_data[:last_idx]
        fixed_col_time = time[:last_idx]

        data_array.append(calculate_single_col(fixed_col_data, fixed_col_time, col_number, sheet_name))

    return data_array


def process_data(file_name, output_name):
    result_fret = pd.concat(process_sheet(file_name, 0, 'FRET'))
    result_rhod = pd.concat(process_sheet(file_name, 1, 'RHOD'))

    with pd.ExcelWriter(output_name) as writer:
        result_fret.T.to_excel(writer, sheet_name='Fret')
        result_rhod.T.to_excel(writer, sheet_name='Rhod')


# Load data
file_name = 'CaMKAR CY Sr2+ 1 mM DTT processes excel.xlsx'  # Update the file path accordingly
output_name = 'output_test.xlsx'

process_data(file_name, output_name)

plt.close()
plt.show()
