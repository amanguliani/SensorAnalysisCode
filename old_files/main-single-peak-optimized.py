import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import find_peaks, peak_widths
from scipy.integrate import simps

def calc_single(d, t, col_num, sheet_name):
    # Identifying peaks with increased prominence

    # Remove the empty data from the end.
    last_idx = d.last_valid_index()
    col_data = d[:last_idx]
    col_time = t[:last_idx]

    # Find peaks
    peaks, properties = find_peaks(col_data, prominence=0.05)  # Increased prominence value
    results_widths = peak_widths(col_data, peaks, rel_height=0.99)

    peak_values = col_data.iloc[peaks]
    peak_times = col_time.iloc[peaks]

    # Plotting Peak data first.
    plt.figure(figsize=(15, 5))
    plt.plot(t, d, label='data', color='blue')
    plt.plot(peak_times, peak_values, 'rX', label='Peaks')
    plt.xlabel('Time (Min)')
    plt.ylabel('Value')

    # Initialize a list to store the minimum values and times between peaks
    amplitude = []
    duration = []
    areas = []
    percentage_change = []

    for width, height, left, right in zip(results_widths[0], results_widths[1], results_widths[2], results_widths[3]):
        left = round(left)
        right = round(right)

        # Using Simpson's rule for numerical integration
        x1, y1 = col_time.iloc[left], col_data.iloc[left]
        x2, y2 = col_time.iloc[right], col_data.iloc[right]
        percentage_change.append(
            (abs(col_data.iloc[right] - col_data.iloc[left]) / col_data.iloc[left]) * 100.0)
        amplitude.append(height)
        duration.append(x2-x1)
        plt.plot([x1, x2], [y1, y2], marker='o')
        area_under_peak = simps(col_data[left:right], col_time[left:right]) - (
                0.5 * (x2 - x1) * (y1 + y2))
        areas.append(abs(area_under_peak))

    plt.title(
        f"Plot of {sheet_name} column {col_num} data vs time with Peaks and their min values marked")
    plt.grid(True)
    plt.legend()
    plt.draw()
    plt.interactive(False)
    peak_data = pd.DataFrame(
        {'Column': col_num,
         'Time of peak occurrence': peak_times[:6],
         'Peak Values': peak_values[:6],
         'Amplitude of Peak': amplitude[:6],
         'Duration of Peak': duration[:6],
         '% Change in baseline': percentage_change[:6],
         'Area Under Curve': areas[:6]
         })
    return peak_data


# Load data
file_path = '0.5 data.xlsx'  # Update the file path accordingly
output_path = '0.5 data output.xlsx'
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

plt.close()
plt.show()