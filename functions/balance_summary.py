import pandas as pd

def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading the file: {e}")
        return None

def generate_sheets_by_balance(dataframe, column_index):
    if column_index < 0 or column_index >= len(dataframe.columns):
        print(f"Invalid column index: {column_index}")
        return {}

    column_name = dataframe.columns[column_index]
    dataframe[column_name] = pd.to_numeric(dataframe[column_name], errors='coerce')
    dataframe = dataframe.dropna(subset=[column_name])

    derived_sheets = {}
    ranges = [
        (float('-inf'), 0, "<0"),
        (0, 50000, "0-50K"),
        (50000, 200000, "50K-2L"),
        (200000, 500000, "2L-5L"),
        (500000, float('inf'), ">5L"),
    ]

    for lower, upper, label in ranges:
        filtered_df = dataframe[(dataframe[column_name] > lower) & (dataframe[column_name] <= upper)]
        if not filtered_df.empty:
            derived_sheets[label] = filtered_df

    return derived_sheets

def calculate_totals(df):
    df = df.apply(pd.to_numeric, errors='coerce').fillna(0)
    totals = {
        'svm_invoice_balance_sum': df.iloc[:, 4].sum(),
        'svm_available_credits_sum': df.iloc[:, 5].sum(),
        'svm_balance': df.iloc[:, 6].sum(),
        'nvb_invoice_balance_sum': df.iloc[:, 1].sum(),
        'nvb_available_credits_sum': df.iloc[:, 2].sum(),
        'nvb_balance': df.iloc[:, 3].sum(),
        'consolidated_invoice_balance_sum': df.iloc[:, 7].sum(),
        'consolidated_available_credits_sum': df.iloc[:, 8].sum(),
        'consolidated_cons_bal_os_sum': df.iloc[:, 9].sum(),
    }
    return totals

def create_summary_sheet(range_totals, ranges, summary_columns, writer):
    summary_data = []
    for range_label in ranges:
        row = range_totals.get(range_label, [0] * 9)
        summary_data.append(row)

    summary_df = pd.DataFrame(
        summary_data,
        columns=pd.MultiIndex.from_arrays(summary_columns),
        index=ranges
    )
    summary_df.to_excel(writer, sheet_name='Summary', index=True)

def apply_color_formatting(worksheet, num_rows, num_cols, workbook):
    # Define the formats
    smcs_format = workbook.add_format({'bg_color': '#FFFF00', 'border': 1})
    nvb_format = workbook.add_format({'bg_color': '#00FF00', 'border': 1})
    consolidated_format = workbook.add_format({'bg_color': '#FF0000', 'border': 1})
    border_format = workbook.add_format({'border': 1})

    # Apply specific column formats INCLUDING headers (row index 1)
    worksheet.conditional_format(0, 1, num_rows + 1, 3, {'type': 'no_blanks', 'format': smcs_format})
    worksheet.conditional_format(0, 4, num_rows + 1, 6, {'type': 'no_blanks', 'format': nvb_format})
    worksheet.conditional_format(0, 7, num_rows + 1, 9, {'type': 'no_blanks', 'format': consolidated_format})

    # Apply border to entire used area INCLUDING header rows
    worksheet.conditional_format(0, 0, num_rows + 2, num_cols - 1, {'type': 'no_blanks', 'format': border_format})

def process_file(input_file, output_file):
    consolidated_df = read_excel_file(input_file)
    if consolidated_df is None:
        return

    split_column_index = 9
    derived_sheets = generate_sheets_by_balance(consolidated_df, split_column_index)
    range_totals = {}

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Top headers and subheaders
        top_headers = [
            "", "SMCS Receivables", "SMCS Receivables", "SMCS Receivables",
            "NVB Receivables", "NVB Receivables", "NVB Receivables",
            "Consolidated Receivables", "Consolidated Receivables", "Consolidated Receivables"
        ]
        sub_headers = [
            "customer_name", "Invoice Balance", "Available Credits", "Closing Balance",
            "Invoice Balance", "Available Credits", "Closing Balance",
            "Invoice Balance", "Available Credits", "Closing Balance"
        ]

        # Write consolidated
        consolidated_df.to_excel(writer, sheet_name='Consolidated_Descending', index=False, header=False, startrow=2)
        worksheet = writer.sheets['Consolidated_Descending']

        for col_num, header in enumerate(top_headers):
            worksheet.write(0, col_num, header)
        for col_num, header in enumerate(sub_headers):
            worksheet.write(1, col_num, header)

        apply_color_formatting(worksheet, len(consolidated_df), len(consolidated_df.columns), workbook)

        # Derived Sheets
        for sheet_name, df in derived_sheets.items():
            df.reset_index(drop=True, inplace=True)
            for col_idx in range(1, 10):
                df.iloc[:, col_idx] = pd.to_numeric(df.iloc[:, col_idx], errors='coerce').fillna(0)

            totals = calculate_totals(df)
            range_totals[sheet_name] = list(totals.values())

            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False, header=False, startrow=2)
            worksheet = writer.sheets[safe_name]

            for col_num, header in enumerate(top_headers):
                worksheet.write(0, col_num, header)
            for col_num, header in enumerate(sub_headers):
                worksheet.write(1, col_num, header)

            apply_color_formatting(worksheet, len(df), len(df.columns), workbook)

        # Summary sheet
        ranges = [">5L", "2L-5L", "50K-2L", "0-50K", "<0"]
        summary_columns = [
            ["NVB Receivables"] * 3 + ["SMCS Receivables"] * 3 + ["Consolidated Receivables"] * 3,
            ["Invoice Balance", "Available Credits", "Closing Balance"] * 3
        ]

        create_summary_sheet(range_totals, ranges, summary_columns, writer)
        summary_ws = writer.sheets['Summary']
        apply_color_formatting(summary_ws, len(ranges), 11, workbook)

    print(f"✅ Successfully processed and saved to {output_file}")
