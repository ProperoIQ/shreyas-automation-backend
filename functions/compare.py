import pandas as pd

# Function to compare two Excel files and extract mismatched rows
def compare_excel_sheets(file1, file2, output_file):
    # Read the two Excel sheets into DataFrames
    try:
        df1 = pd.read_excel(file1, sheet_name='NVB_ar_aging_details', skiprows=3)  # Skip first 3 rows
        df2 = pd.read_excel(file2)
    except Exception as e:
        print(f"Error reading files: {e}")
        return

    # Debug: Print column names of both DataFrames
    print(f"Columns in {file1}: {df1.columns.tolist()}")
    print(f"Columns in {file2}: {df2.columns.tolist()}")

    # Clean column names by stripping leading/trailing spaces
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Ensure that both DataFrames have the same set of columns
    missing_cols_df1 = [col for col in df2.columns if col not in df1.columns]
    missing_cols_df2 = [col for col in df1.columns if col not in df2.columns]

    if missing_cols_df1:
        print(f"Columns missing in {file1}: {missing_cols_df1}")
    if missing_cols_df2:
        print(f"Columns missing in {file2}: {missing_cols_df2}")

    # If there are missing columns in either DataFrame, print and return early
    if missing_cols_df1 or missing_cols_df2:
        print("The two sheets have different columns. Exiting comparison.")
        return

    # Check if 'transaction_number' exists in both DataFrames
    if 'transaction_number' not in df1.columns or 'transaction_number' not in df2.columns:
        print("Error: 'transaction_number' column is missing in one or both files.")
        return

    # Fill NaN values in 'transaction_number' column with a placeholder (e.g., '0')
    df1['transaction_number'] = df1['transaction_number'].fillna('0')
    df2['transaction_number'] = df2['transaction_number'].fillna('0')

    # Extract the numeric part of the 'transaction_number' column for sorting
    df1['transaction_number_numeric'] = df1['transaction_number'].str.extract('(\d+)').astype(int)
    df2['transaction_number_numeric'] = df2['transaction_number'].str.extract('(\d+)').astype(int)

    # Sort the DataFrames by the numeric part of 'transaction_number'
    df1 = df1.sort_values(by=['transaction_number_numeric']).reset_index(drop=True)
    df2 = df2.sort_values(by=['transaction_number_numeric']).reset_index(drop=True)

    # Ensure both DataFrames have the same columns in the same order
    df1 = df1[df2.columns]  # Reorder df1 columns to match df2's order

    # Compare the two DataFrames to find mismatched rows
    comparison = df1.compare(df2, keep_shape=True, keep_equal=False)

    # Extract the rows where there is a mismatch
    mismatched_rows_df1 = df1.loc[comparison.index]  # Rows from df1 with mismatches
    mismatched_rows_df2 = df2.loc[comparison.index]  # Rows from df2 with mismatches

    # Combine mismatched rows from both DataFrames
    mismatched_rows = pd.concat([mismatched_rows_df1, mismatched_rows_df2], ignore_index=True)

    # Save the mismatched rows to a new Excel file
    try:
        mismatched_rows.to_excel(output_file, index=False)
        print(f"Mismatched rows saved to {output_file}")
    except Exception as e:
        print(f"Error saving the file: {e}")

# Example usage
if __name__ == "__main__":
    compare_excel_sheets('output.xlsx', 'NVB_Age_Range_Columns.xlsx', 'mismatched_rows.xlsx')