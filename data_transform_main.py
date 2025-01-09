import pandas as pd
from tkinter import Tk, filedialog

def extract_field(input_df, field_name):
    """
    Extract Amplitude of Peak values for all unique Column No values where:
    - Time of Peak Occurrence is not empty.
    """
    # Transpose the DataFrame for easier processing
    transposed = input_df.transpose()
    transposed.columns = transposed.iloc[0]  # Set the first row as column headers
    transposed = transposed[1:]  # Drop the first row

    # Initialize a dictionary to hold amplitudes for each Column No
    amplitude_dict = {}

    # Process for each unique Column No
    for col_no in transposed["Column No"].unique():
        filtered = transposed[
            (transposed["Column No"] == col_no) & (~transposed["Time of peak occurrence"].isna())
        ]
        amplitude_values = filtered[field_name].astype(float).dropna()
        amplitude_dict[col_no] = amplitude_values.tolist()

    return amplitude_dict


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

    # Load the uploaded Excel file
    fret_df = pd.read_excel(file_name, sheet_name='Fret')
    fret_transformed = extract_field(fret_df, field_name="Amplitude of Peak")
    fret_transformed_df = pd.DataFrame.from_dict(fret_transformed, orient='index')

    rhod_df = pd.read_excel(file_name, sheet_name='Rhod')
    rhod_transformed = extract_field(rhod_df, field_name="Amplitude of Peak")
    rhod_transformed_df = pd.DataFrame.from_dict(rhod_transformed, orient='index')

    with pd.ExcelWriter(output_name) as writer:
        fret_transformed_df.to_excel(writer, sheet_name='Fret')
        rhod_transformed_df.T.to_excel(writer, sheet_name='Rhod')

