import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import find_peaks
from scipy.integrate import simps


def tighten_peaks(peak_data, s, e, peak):
    peak_pct = peak_data.pct_change()
    new_s = peak_pct[s:peak].idxmax() - 5
    new_e = peak_pct[peak:e].idxmin() + 5
    return [new_s, new_e]


def find_mins_for_peaks(min_data, peaks):
    result = {}
    start_index = 0
    last_index = len(min_data) - 1

    for i, peak in enumerate(peaks):
        s_index = min_data[start_index:peak].idxmin()
        if i < len(peaks) - 1:
            e_index = min_data[peak:peaks[i + 1]].idxmin()
        else:
            e_index = last_index
        start_index = peak

        result[peak] = tighten_peaks(min_data, s_index, e_index, peak)

    return result

def calc_single(d, t, col_num, sheet_name):
    # Identifying peaks with increased prominence

    # Remove the empty data from the end.
    last_idx = d.last_valid_index()
    col_data = d[:last_idx]
    col_time = t[:last_idx]

    # Find peaks
    peaks, properties = find_peaks(col_data, prominence=0.05)  # Increased prominence value

    peak_values = col_data.iloc[peaks]
    peak_times = col_time.iloc[peaks]

    # Plotting Peak data first.
    plt.figure(figsize=(12, 6))
    plt.plot(t, d, label='data', color='blue')
    plt.plot(peak_times, peak_values, 'rX', label='Peaks')  # 'X' markers for peaks
    plt.xlabel('Time (Min)')
    plt.ylabel('Value')

    peak_min_map = find_mins_for_peaks(col_data, peaks)

    # Initialize a list to store the minimum values and times between peaks
    amplitude = []
    duration = []
    areas = []
    percentage_change = []

    for i, peak in enumerate(peak_min_map):
        start_index = peak_min_map[peak][0]
        end_index = peak_min_map[peak][1]
        if start_index < 0:
            start_index = 0
        if end_index > len(col_data) - 1:
            end_index = len(col_data) - 1
        amplitude.append(col_data.iloc[peak] - col_data.iloc[start_index])
        duration.append(col_time[end_index] - col_time[start_index])
        percentage_change.append(
            (abs(col_data.iloc[end_index] - col_data.iloc[start_index]) / col_data.iloc[start_index]) * 100.0)

        # Using Simpson's rule for numerical integration
        x1, y1 = col_time[start_index], col_data[start_index]
        x2, y2 = col_time[end_index], col_data[end_index]
        area_under_peak = simps(col_data[start_index:end_index], col_time[start_index:end_index]) - (
                0.5 * (x2 - x1) * (y1 + y2))
        areas.append(abs(area_under_peak))

        plt.plot([x1, x2], [y1, y2], marker='o')

    # Combining time and peak values into a DataFrame for display
    peak_data = pd.DataFrame(
        {'Column': col_num,
         'Time of peak occurrence': peak_times[:5],
         'Peak Values': peak_values[:5],
         'Amplitude of Peak': amplitude[:5],
         'Duration of Peak': duration[:5],
         '% Change in baseline': percentage_change[:5],
         'Area Under Curve': areas[:5]
         })

    plt.title(
        f"Plot of {sheet_name} column {col_num} data vs time with Peaks and their min values marked")
    plt.grid(True)
    plt.legend()
    plt.draw()
    plt.interactive(False)
    return peak_data


# Load data
file_path = 'data.xlsx'  # Update the file path accordingly
output_path = 'output.xlsx'
fret_data = pd.read_excel(file_path, sheet_name=0)
rhod_data = pd.read_excel(file_path, sheet_name=1)

time = fret_data['Time (Min)']
df_fret = []
for c in fret_data.columns.values[1:]:
    data = fret_data[c]
    df_fret.append(calc_single(data, time, c, 'Fret'))

result_fret = pd.concat(df_fret)

time = rhod_data['Time (Min)']
df_rhod = []
for c in rhod_data.columns.values[1:]:
    data = rhod_data[c]
    df_rhod.append(calc_single(data, time, c, 'Rhod'))

result_rhod = pd.concat(df_rhod)

with pd.ExcelWriter(output_path) as writer:
    result_fret.T.to_excel(writer, sheet_name='Fret')
    result_rhod.T.to_excel(writer, sheet_name='Rhod')

plt.show()
