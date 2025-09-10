import requests
import pandas as pd
from pathlib import Path
import logging
import json
from concurrent.futures import ThreadPoolExecutor
from tenacity import retry, stop_after_attempt, wait_exponential

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(client_name)s] - %(message)s'
)

class ClientLoggingAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        return f"{msg}", {"extra": {"client_name": self.extra.get("client_name", "Unknown")}}

logger = ClientLoggingAdapter(logging.getLogger(__name__), {"client_name": "Unknown"})

# Configuration (replace with your credentials)
CREDENTIALS = [
    {
        "CLIENT_ID": "1000.MAV029IW7FDMD5XO3BIVN83KDSP8LC",
        "CLIENT_SECRET": "c58044529b8a0e5d0895c4ee0b45adc2987048eadc",
        "REFRESH_TOKEN": "1000.c9ca462b4979baf18115b1ea8cc52910.ce3ac1ec195ce4891f29e4000e7441fd",
        "ORG_ID": "642273083",
        "Client": "smcs"
    },
    {
        "CLIENT_ID": "1000.MAV029IW7FDMD5XO3BIVN83KDSP8LC",
        "CLIENT_SECRET": "c58044529b8a0e5d0895c4ee0b45adc2987048eadc",
        "REFRESH_TOKEN": "1000.c9ca462b4979baf18115b1ea8cc52910.ce3ac1ec195ce4891f29e4000e7441fd",
        "ORG_ID": "693033731",
        "Client": "nvb"
    }
]

INVOICES_URL = (
    "https://www.zohoapis.com/books/v3/invoices?page={page}&per_page=200"
    "&filter_by=Status.All&sort_column=created_time&sort_order=D"
    "&requiredfields=created_time%2Clast_modified_time%2Cdate%2Cinvoice_number"
    "%2Creference_number%2Ccustomer_name%2Cstatus%2Cdue_date%2Ctotal%2Cbalance"
    "%2Clocation_name%2Cinvoice_id%2Ccustomer_id%2Ctype%2Cdue_days%2Cis_emailed"
    "%2Cis_viewed_by_client%2Cis_viewed_in_mail%2Cschedule_time"
    "%2Cunprocessed_payment_amount%2Cis_peppol_supported%2Cbbps"
    "%2Cis_square_transaction%2Cis_digitally_signed%2Ctax_source%2Cis_pre_gst"
    "%2Cach_payment_initiated%2Cclient_viewed_time%2Cclient_viewed_time_formatted"
    "%2Cmail_last_viewed_time%2Cmail_last_viewed_time_formatted%2Chas_attachment"
    "%2Cis_correction_invoice%2Ceinvoice_details&usestate=true&organization_id={ORG_ID}"
)

def save_json(json_data: dict, file_path: Path, client_name: str) -> None:
    """Save JSON data to a file."""
    try:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with file_path.open("w") as f:
            json.dump(json_data, f, indent=2)
        logger.info(f"Saved JSON to {file_path}", extra={"client_name": client_name})
    except Exception as e:
        logger.error(f"Failed to save JSON to {file_path}: {e}", extra={"client_name": client_name})

def write_excel(df: pd.DataFrame, file_path: Path, sheet_name: str, client_name: str) -> None:
    """Write DataFrame to Excel with dynamic column widths."""
    try:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
                worksheet.set_column(idx, idx, min(max_len, 50))
        logger.info(f"Saved Excel to {file_path}", extra={"client_name": client_name})
    except Exception as e:
        logger.error(f"Failed to save Excel to {file_path}: {e}", extra={"client_name": client_name})

@retry(stop=stop_after_attempt(3), wait=wait_exponential(multiplier=1, min=4, max=10))
def generate_access_token(client_id: str, client_secret: str, refresh_token: str, client_name: str) -> str:
    """Generate OAuth 2.0 access token with retry logic."""
    params = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
    }
    try:
        response = requests.post("https://accounts.zoho.com/oauth/v2/token", data=params)
        response.raise_for_status()
        return response.json().get("access_token")
    except requests.RequestException as e:
        logger.error(f"Failed to generate access token: {e}", extra={"client_name": client_name})
        raise

def fetch_and_merge_invoices_for_client(client_data: dict) -> None:
    """Fetch invoices and merge Invoice ID into aging details directly from JSON."""
    client_name = client_data["Client"]
    logger.extra["client_name"] = client_name

    if not all(client_data.get(key) for key in ["CLIENT_ID", "CLIENT_SECRET", "REFRESH_TOKEN", "ORG_ID"]):
        logger.error("Missing credentials", extra={"client_name": client_name})
        return

    # Fetch invoices
    access_token = generate_access_token(
        client_data["CLIENT_ID"], client_data["CLIENT_SECRET"], client_data["REFRESH_TOKEN"], client_name
    )
    if not access_token:
        return

    all_invoices = []
    page = 1
    while True:
        url = INVOICES_URL.format(page=page, ORG_ID=client_data["ORG_ID"])
        headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            json_data = response.json()
            if "invoices" in json_data and json_data["invoices"]:
                all_invoices.extend(json_data["invoices"])
                if len(json_data["invoices"]) < 200:
                    break
                page += 1
            else:
                logger.warning(f"No invoices found on page {page}", extra={"client_name": client_name})
                break
        except requests.RequestException as e:
            logger.error(f"Failed to fetch invoices: {e}", extra={"client_name": client_name})
            break

    if not all_invoices:
        logger.warning("No invoices to process", extra={"client_name": client_name})
        return

    # Save JSON for reference (optional)
    json_path = Path(f"csvdata/invoices_details_{client_name}.json")
    save_json({"invoices": all_invoices}, json_path, client_name)

    # Convert invoices to DataFrame for merging
    invoices_df = pd.json_normalize(all_invoices)
    if 'invoice_id' not in invoices_df.columns or 'invoice_number' not in invoices_df.columns:
        logger.error("Missing 'invoice_id' or 'invoice_number' in invoice data", extra={"client_name": client_name})
        return

    # Clean invoice_number to handle formatting issues
    invoices_df['invoice_number'] = invoices_df['invoice_number'].astype(str).str.strip().str.lower()

    # Check for duplicate invoice numbers
    if invoices_df['invoice_number'].duplicated().any():
        logger.warning("Duplicate invoice numbers found in invoice data", extra={"client_name": client_name})

    # Read aging details from Excel
    aging_file = Path(f"csvdata/invoice_aging_details_{client_name}.xlsx")
    try:
        aging_df = pd.read_excel(aging_file, sheet_name=0)
    except FileNotFoundError:
        logger.error(f"Aging file not found: {aging_file}", extra={"client_name": client_name})
        return
    except Exception as e:
        logger.error(f"Error reading aging file: {e}", extra={"client_name": client_name})
        return

    if 'transaction_number' not in aging_df.columns:
        logger.error(f"Missing 'transaction_number' in {aging_file}", extra={"client_name": client_name})
        return

    # Clean transaction_number to handle formatting issues
    aging_df['transaction_number'] = aging_df['transaction_number'].astype(str).str.strip().str.lower()

    # Check for duplicate transaction numbers
    if aging_df['transaction_number'].duplicated().any():
        logger.warning("Duplicate transaction numbers found in aging data", extra={"client_name": client_name})

    # Merge Invoice ID based on Invoice Number
    merged_df = aging_df.merge(
        invoices_df[['invoice_number', 'invoice_id']],
        how='left',
        left_on='transaction_number',
        right_on='invoice_number'
    ).drop(columns=['invoice_number'], errors='ignore')

    # Rename invoice_id to Invoice ID for consistency
    if 'invoice_id' in merged_df.columns:
        merged_df = merged_df.rename(columns={'invoice_id': 'Invoice ID'})

    # Log unmatched records
    unmatched = merged_df[merged_df['Invoice ID'].isna()]['transaction_number'].unique()
    if len(unmatched) > 0:
        logger.warning(f"{len(unmatched)} transaction numbers not matched: {', '.join(map(str, unmatched))}", 
                       extra={"client_name": client_name})

    # Save merged DataFrame to a new file to avoid overwriting original
    output_file = Path(f"csvdata/input_invoice_aging_details_{client_name}.xlsx")
    write_excel(merged_df, output_file, "Aging_Details", client_name)

def invoice_step():
    """Main function to fetch invoices and merge invoice IDs directly from JSON."""
    with ThreadPoolExecutor(max_workers=2) as executor:
        executor.map(fetch_and_merge_invoices_for_client, CREDENTIALS)

if __name__ == "__main__":
    invoice_step()