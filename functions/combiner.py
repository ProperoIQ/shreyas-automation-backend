import pandas as pd

def combine_sheets(input_file1, input_file2, output_file, gap=5):
    """
    Combines the 'Summary' sheet from the first input file and the first sheet from the second input file.
    The second sheet's column names are added just above its data with a gap between the two sheets.
    Then, it applies color coding to the header columns and data cells as per the instructions.

    Args:
    - input_file1 (str): Path to the first input Excel file (with 'Summary' sheet).
    - input_file2 (str): Path to the second input Excel file (with the first sheet).
    - output_file (str): Path to the output Excel file to save the combined result.
    - gap (int): Number of empty rows between the sheets (default is 5).
    """
    try:
        # Read the 'Summary' sheet from the first input file
        df1 = pd.read_excel(input_file1, sheet_name='Summary')

        # Read the first sheet from the second input file
        df2 = pd.read_excel(input_file2, sheet_name=0)
        
        # Add 'gap' number of empty rows (NaN values) to create space between the sheets
        empty_rows = pd.DataFrame([[None] * len(df1.columns)] * gap, columns=df1.columns)

        # Insert column names of df2 just above its data
        df2_column_names = pd.DataFrame([df2.columns], columns=df2.columns)

        # Combine the two dataframes with the gap and column names
        combined_df = pd.concat([df1, empty_rows, df2_column_names, df2], ignore_index=True)
        
        # Reset index and drop it to ensure no index appears in the final result
        combined_df.reset_index(drop=True, inplace=True)

        # Write the combined dataframe to a new Excel file with formatting
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='CombinedSheet')
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['CombinedSheet']

            # Define formats for the color coding
            green_format = workbook.add_format({'bg_color': '#00FF00'})
            yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
            red_format = workbook.add_format({'bg_color': '#FF0000'})

            # Apply color coding to the header row (the second row)
            header_row_idx =  1 # Adjust for the gap and header row

            # Color the first set of "Invoice Balance", "Available Credits", "Closing Balance" green
            worksheet.write(header_row_idx, 1, 'Invoice Balance', green_format)
            worksheet.write(header_row_idx, 2, 'Available Credits', green_format)
            worksheet.write(header_row_idx, 3, 'Closing Balance', green_format)

            # Color the second set of columns yellow
            worksheet.write(header_row_idx, 4, 'Invoice Balance', yellow_format)
            worksheet.write(header_row_idx, 5, 'Available Credits', yellow_format)
            worksheet.write(header_row_idx, 6, 'Closing Balance', yellow_format)

            # Color the third set of columns red
            worksheet.write(header_row_idx, 7, 'Invoice Balance', red_format)
            worksheet.write(header_row_idx, 8, 'Available Credits', red_format)
            worksheet.write(header_row_idx, 9, 'Closing Balance', red_format)

            # Apply color formatting to all the data rows below the header until the data ends
            for row_idx in range(header_row_idx + 1, len(combined_df)  + 2):
                # Apply the green color format for first set of columns (Invoice Balance, Available Credits, Closing Balance)
                if pd.notnull(combined_df.iloc[row_idx  - 1, 1]):  # Check if cell is not empty
                    worksheet.write(row_idx, 1, combined_df.iloc[row_idx  - 1, 1], green_format)
                    worksheet.write(row_idx, 2, combined_df.iloc[row_idx  - 1, 2], green_format)
                    worksheet.write(row_idx, 3, combined_df.iloc[row_idx  - 1, 3], green_format)

                # Apply the yellow color format for second set of columns
                if pd.notnull(combined_df.iloc[row_idx  - 1, 4]):  # Check if cell is not empty
                    worksheet.write(row_idx, 4, combined_df.iloc[row_idx  - 1, 4], yellow_format)
                    worksheet.write(row_idx, 5, combined_df.iloc[row_idx  - 1, 5], yellow_format)
                    worksheet.write(row_idx, 6, combined_df.iloc[row_idx  - 1, 6], yellow_format)

                # Apply the red color format for third set of columns
                if pd.notnull(combined_df.iloc[row_idx  - 1, 7]):  # Check if cell is not empty
                    worksheet.write(row_idx, 7, combined_df.iloc[row_idx  - 1, 7], red_format)
                    worksheet.write(row_idx, 8, combined_df.iloc[row_idx  - 1, 8], red_format)
                    worksheet.write(row_idx, 9, combined_df.iloc[row_idx  - 1, 9], red_format)

        print(f"Sheets combined and saved to {output_file}")

    except Exception as e:
        print(f"Error: {e}")
