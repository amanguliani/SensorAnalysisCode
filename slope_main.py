import matplotlib
matplotlib.use('Qt5Agg')

import matplotlib.pyplot as plt

from tkinter import Tk, filedialog

import pandas as pd
import numpy as np

def smooth(y, box_pts):
    box = np.ones(box_pts)/box_pts
    y_smooth = np.convolve(y, box, mode='same')
    return y_smooth


# Define a function to process each column of data
def analyze_and_plot(data, sheet_name):
    time_column = 'Time (Min)'
    time = data[time_column]
    results = []

    # Iterate through each column (excluding the time column)
    for column in data.columns:
        if column == time_column:
            continue

        # Extract column data and drop NaN values
        y = data[column].dropna()
        time_cleaned = time.iloc[y.index]

        if y.empty:
            continue

        smoothed_y = smooth(y, 10)
        # Plot the data
        plt.figure(figsize=(10, 5))
        plt.plot(time_cleaned, smoothed_y, label=f'{column}', marker='o')
        # plt.axvline(peak_time, color='r', linestyle='--', label='Peak Time')
        plt.title(f"{sheet_name} - {column} vs Time")
        plt.xlabel("Time (Min)")
        plt.ylabel(f"{column}")
        plt.ylim(0.8, 1.8)
        plt.xticks(time_cleaned)
        plt.legend()
        plt.grid()
        plt.show()

        print (f"Analyzing {column}...")
        time_index = input("Time index of peak?")
        # type cast into integer
        time_index = int(time_index)

        peak_value = y.iloc[time_index]
        peak_time = time_cleaned.loc[time_index]

        # Calculate the rate of rise and decay
        initial_value = y.iloc[0]
        final_value = y.iloc[-1]

        rise_rate = (peak_value - initial_value) / (peak_time - time_cleaned.iloc[0])
        decay_rate = (final_value - peak_value) / (time_cleaned.iloc[-1] - peak_time)

        # Store the results
        results.append({
            'Column': column,
            'Peak Value': peak_value,
            'Time of Peak (Min)': peak_time,
            'Rate of Rise': rise_rate,
            'Rate of Decay': decay_rate
        })



    # Return the results as a DataFrame
    return pd.DataFrame(results)



if __name__ == "__main__":
    Tk().withdraw()
    file_name = filedialog.askopenfilename(title="Select the input file", filetypes=[("Excel files", "*.xlsx")])
    if not file_name:
        print("No file selected. Exiting.")
        exit()

    output_name = filedialog.asksaveasfilename(title="Save output file as", defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])
    if not output_name:
        print("No output file selected. Exiting.")
        exit()

    data = pd.ExcelFile(file_name)

    # Display the sheet names to understand its structure
    fret_data = data.parse('FRET ')

    # Analyze and plot for each sheet
    fret_results = analyze_and_plot(fret_data, "FRET")

    with pd.ExcelWriter(output_name) as writer:
        fret_results.T.to_excel(writer, sheet_name='FRET')



