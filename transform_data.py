from tkinter import Tk, filedialog

import pandas as pd

# Function to transform the data
def transform_data(df):
    # Identify the row containing 'Column No'
    col_no_row = df.iloc[0, 1:].astype(float).values
    transformed_data = []

    for measurement in measurement_types:
        # Find the row index of the measurement type
        row_idx = df[df.iloc[:, 0] == measurement].index
        if row_idx.empty:
            continue  # Skip if the measurement type is not found

        row_idx = row_idx[0]  # Get the row index
        data_values = df.iloc[row_idx, 1:].values  # Extract data values

        # Track changes in 'Column No'
        prev_col_no = col_no_row[0]
        temp_data = []

        for i, (col_no, value) in enumerate(zip(col_no_row, data_values)):
            if col_no != prev_col_no and temp_data:
                transformed_data.append([measurement, prev_col_no] + temp_data)
                temp_data = []  # Reset for new column number

            temp_data.append(value)
            prev_col_no = col_no

        # Append remaining data
        if temp_data:
            transformed_data.append([measurement, prev_col_no] + temp_data)

    # Convert to DataFrame
    transformed_df = pd.DataFrame(transformed_data)
    return transformed_df

if __name__ == "__main__":
    Tk().withdraw()
    file_name = filedialog.askopenfilename(title="Select the input file", filetypes=[("Excel files", "*.xlsx")])
    if not file_name:
        print("No file selected. Exiting.")
        exit()

    print(f"Processing Tranform for file {file_name}")

    xls = pd.ExcelFile(file_name)

    # Load the data from both sheets
    df_fret = pd.read_excel(xls, sheet_name='Fret')
    df_rhod = pd.read_excel(xls, sheet_name='Rhod')

    # Define the measurement types to extract
    measurement_types = ["Rate of Rise", "Rate of Decay"]

    # Transform data for both sheets
    transformed_fret = transform_data(df_fret)
    transformed_rhod = transform_data(df_rhod)

    output_name = filedialog.asksaveasfilename(title="Save output file as", defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])
    if not output_name:
        print("No output file selected. Exiting.")
        exit()
    # Save to a new Excel file

    with pd.ExcelWriter(output_name) as writer:
        transformed_fret.to_excel(writer, sheet_name="Fret", index=False)
        transformed_rhod.to_excel(writer, sheet_name="Rhod", index=False)

    print(f"Transformed data saved to {output_name}")

