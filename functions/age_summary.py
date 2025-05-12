import pandas as pd
from openpyxl.utils import get_column_letter

def clean_sheet_name(name):
    """Clean sheet name to be Excel-compatible and hyperlink-safe"""
    replacements = {
        '>=': '_ge_',
        '<=': '_le_',
        '>': '_gt_',
        '<': '_lt_',
        '=': '_eq_',
        ':': '', '\\': '', '/': '', '?': '', '*': '', '[': '', ']': ''
    }
    for k, v in replacements.items():
        name = name.replace(k, v)
    return name[:31]  # Excel sheet name limit

def generate_summary(input_files, output_file):
    ageing_cols = [
        '3Yrs>=', '3Yr<=2Yr', '2Yr<=1Yr',
        '1Yr<=180days', '180<=90days', '90<=60days',
        '60<=30days', '>=30days', 'Age_Not_Provided'
    ]

    summary_data = []
    detail_sheets = []
    sheet_mapping = {}

    for file_name, file_path in input_files.items():
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            continue

        df.columns = df.columns.str.strip()
        df = df[~df.iloc[:, 0].astype(str).str.strip().str.lower().str.contains('total')]

        missing_cols = [col for col in ageing_cols if col not in df.columns]
        if missing_cols:
            print(f"Missing columns in {file_name}: {missing_cols}")
            continue

        df['Unpaid Invoices'] = df[ageing_cols].sum(axis=1)

        row_data = {'Ageing bucket': file_name}
        for col in ageing_cols:
            value = int(df[col].sum())  # Convert to integer here
            sheet_name = clean_sheet_name(f"{file_name}_{col}")
            sheet_mapping[(file_name, col)] = sheet_name
            row_data[col] = value
        row_data['Unpaid Invoices'] = int(df['Unpaid Invoices'].sum())  # Convert to integer
        summary_data.append(row_data)

        for col in ageing_cols:
            filtered_df = df[df[col] > 0].copy()
            if not filtered_df.empty:
                cols_to_write = [c for c in df.columns if c not in ageing_cols or c == col]
                detail_sheets.append((sheet_mapping[(file_name, col)], filtered_df[cols_to_write], file_name, col))

    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            workbook = writer.book
            header_format = workbook.add_format({'bold': True, 'bg_color': '#DCE6F1'})
            total_format = workbook.add_format({'bold': True, 'top': 1, 'bg_color': '#F2F2F2'})
            number_format = workbook.add_format({'num_format': '0'})  # No decimals format
            link_format = workbook.add_format({'font_color': 'blue', 'underline': 1})

            # Create summary DataFrame
            summary_df = pd.DataFrame(summary_data)
            
            # Calculate totals (converted to integers)
            totals = {k: int(v) if isinstance(v, (int, float)) else v 
                     for k, v in summary_df.drop(columns=['Ageing bucket']).sum().to_dict().items()}
            totals['Ageing bucket'] = 'Total'
            
            # Append totals
            summary_df = pd.concat([summary_df, pd.DataFrame([totals])], ignore_index=True)
            
            # Write Summary sheet
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            worksheet = writer.sheets['Summary']

            # Format headers and apply number formatting
            for col_idx, col_name in enumerate(summary_df.columns):
                worksheet.write(0, col_idx, col_name, header_format)
                if col_name != 'Ageing bucket':
                    worksheet.set_column(col_idx, col_idx, None, number_format)

            # Add hyperlinks with integer values
            for row_idx, row in summary_df.iloc[:-1].iterrows():  # Skip total row
                for col_idx, col in enumerate(summary_df.columns):
                    if col == 'Ageing bucket':
                        continue
                    value = row[col]
                    if col in ageing_cols and value > 0:
                        sheet_name = sheet_mapping.get((row['Ageing bucket'], col))
                        if sheet_name:
                            try:
                                worksheet.write_url(
                                    row_idx + 1, col_idx,
                                    f"internal:'{sheet_name}'!A1",
                                    string=str(int(value)),  # Ensure integer display
                                    tip=f"Go to {sheet_name}",
                                    cell_format=number_format  # Apply number format
                                )
                            except:
                                worksheet.write_number(row_idx + 1, col_idx, value, number_format)
                    elif col == 'Unpaid Invoices':
                        worksheet.write_number(row_idx + 1, col_idx, value, number_format)

            # Write detail sheets with return hyperlinks
            for sheet_name, df, source_file, source_col in detail_sheets:
                clean_sheet = clean_sheet_name(sheet_name)
                df.to_excel(writer, sheet_name=clean_sheet, index=False)
                
                # Add ONE hyperlink in A1 only
                ws_detail = writer.sheets[clean_sheet]
                ws_detail.write_url(
                    0, 0,  # First row, first column (A1)
                    f"internal:'Summary'!A1",
                    string="date",
                    cell_format=link_format
                )
                
                # If there's a date column, add links to each date cell
                date_cols = [col for col in df.columns if 'date' in col.lower()]
                for date_col in date_cols:
                    col_idx = df.columns.get_loc(date_col)
                    for row_idx in range(len(df)):
                        ws_detail.write_url(
                            row_idx + 1, col_idx,  # +1 to skip header
                            f"internal:'Summary'!A1",
                            string=str(df.iloc[row_idx, col_idx]),
                            cell_format=link_format
                        )

        print(f"Excel file saved with bidirectional hyperlinks at {output_file}")

    except Exception as e:
        print(f"Error saving Excel file: {e}")