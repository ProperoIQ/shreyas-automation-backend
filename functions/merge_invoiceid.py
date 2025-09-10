import pandas as pd
import os
import logging
import uuid

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def merge_invoice(client_name):
    """Merge Invoice ID into aging details and overwrite the original file for the given client."""
    # File paths
    invoices_file = f"csvdata/input_invoices_details_{client_name}.xlsx"
    aging_file = f"csvdata/input_invoice_aging_details_{client_name}.xlsx"
    output_file = f"csvdata/input_invoice_aging_details_{client_name}.xlsx"  # Overwrite original

    try:
        # Read invoices details (Invoice ID and Invoice Number)
        invoices_df = pd.read_excel(invoices_file, sheet_name="Invoices")
        logging.info(f"Loaded invoices file: {invoices_file}")
        logging.info(f"Invoices columns: {list(invoices_df.columns)}")
        
        # Read aging details (includes transaction_number)
        aging_df = pd.read_excel(aging_file, sheet_name=0)  # Assuming first sheet
        logging.info(f"Loaded aging file: {aging_file}")
        logging.info(f"Aging columns: {list(aging_df.columns)}")

        # Verify required columns
        if 'Invoice ID' not in invoices_df.columns or 'Invoice Number' not in invoices_df.columns:
            logging.error(f"Missing 'Invoice ID' or 'Invoice Number' in {invoices_file}")
            return
        if 'transaction_number' not in aging_df.columns:
            logging.error(f"Missing 'transaction_number' in {aging_file}")
            return

        # Merge to add Invoice ID to aging details
        merged_df = aging_df.merge(
            invoices_df[['Invoice Number', 'Invoice ID']],
            how='left',
            left_on='transaction_number',
            right_on='Invoice Number'
        )

        # Drop the redundant 'Invoice Number' column from merge
        if 'Invoice Number' in merged_df.columns:
            merged_df = merged_df.drop(columns=['Invoice Number'])

        # Debug: Log merge results
        logging.info(f"Merged DataFrame columns: {list(merged_df.columns)}")
        logging.info(f"Merged DataFrame head:\n{merged_df.head().to_string()}")
        logging.info(f"Number of rows in aging file: {len(aging_df)}")
        logging.info(f"Number of rows after merge: {len(merged_df)}")
        logging.info(f"Number of non-null Invoice IDs: {merged_df['Invoice ID'].notna().sum()}")

        # Save to the original Excel file (overwrite)
        os.makedirs("data", exist_ok=True)
        if os.path.exists(output_file):
            try:
                os.remove(output_file)
                logging.info(f"Removed existing file: {output_file}")
            except Exception as e:
                logging.error(f"Error removing existing file {output_file}: {e}")

        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Aging_Details")
            worksheet = writer.sheets["Aging_Details"]
            for idx, col in enumerate(merged_df.columns):
                max_len = max(merged_df[col].astype(str).map(len).max(), len(str(col))) + 2
                worksheet.set_column(idx, idx, min(max_len, 50))
        logging.info(f"Overwrote aging file: {output_file}")

    except FileNotFoundError as e:
        logging.error(f"File not found for {client_name}: {e}")
    except Exception as e:
        logging.error(f"Error processing files for {client_name}: {e}")

def merge_invoice_id():
    """Process both smcs and nvb clients."""
    clients = ["smcs", "nvb"]
    for client in clients:
        logging.info(f"Processing client: {client}")
        merge_invoice(client)

if __name__ == "__main__":
    merge_invoice_id()