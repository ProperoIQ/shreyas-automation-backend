import requests
import pandas as pd
from pathlib import Path
import logging
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
from ratelimit import limits, sleep_and_retry
import time

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(client_name)s] - %(message)s'
)

class ClientLoggingAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        return f"{msg}", {"extra": {"client_name": self.extra.get("client_name", "Unknown")}}

logger = ClientLoggingAdapter(logging.getLogger(__name__), {"client_name": "Unknown"})

# Configuration
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

# Rate limit configuration (adjust based on Zoho's API limits, e.g., 100 calls per minute)
CALLS = 100
PERIOD = 60  # seconds

@sleep_and_retry
@limits(calls=CALLS, period=PERIOD)
@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=4, max=60),
    retry=retry_if_exception_type(requests.exceptions.HTTPError),
    after=lambda retry_state: logger.info(
        f"Retrying {retry_state.fn.__name__} after attempt {retry_state.attempt_number}",
        extra={"client_name": retry_state.args[-1]}
    )
)
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
        if response.status_code == 429:
            retry_after = int(response.headers.get("Retry-After", 60))
            logger.warning(f"Rate limit hit for access token. Waiting {retry_after} seconds.", 
                          extra={"client_name": client_name})
            time.sleep(retry_after)
            raise requests.exceptions.HTTPError("Rate limit exceeded")
        response.raise_for_status()
        return response.json().get("access_token")
    except requests.RequestException as e:
        logger.error(f"Failed to generate access token: {e}", extra={"client_name": client_name})
        raise

@sleep_and_retry
@limits(calls=CALLS, period=PERIOD)
@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=4, max=60),
    retry=retry_if_exception_type(requests.exceptions.HTTPError),
    after=lambda retry_state: logger.info(
        f"Retrying {retry_state.fn.__name__} after attempt {retry_state.attempt_number}",
        extra={"client_name": retry_state.args[-1]}
    )
)
def fetch_invoice_comments(invoice_id: str, org_id: str, access_token: str, client_name: str) -> list:
    """Fetch comments for a specific invoice based on the provided schema."""
    url = f"https://www.zohoapis.com/books/v3/invoices/{invoice_id}/comments?organization_id={org_id}"
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 429:
            retry_after = int(response.headers.get("Retry-After", 60))
            logger.warning(f"Rate limit hit for invoice {invoice_id}. Waiting {retry_after} seconds.", 
                          extra={"client_name": client_name})
            time.sleep(retry_after)
            raise requests.exceptions.HTTPError("Rate limit exceeded")
        response.raise_for_status()
        json_data = response.json()
        if json_data.get("code") == 0 and "comments" in json_data:
            return json_data["comments"]
        else:
            logger.warning(f"Invalid response for invoice {invoice_id}: {json_data.get('message', 'No message')}", 
                          extra={"client_name": client_name})
            return []
    except requests.RequestException as e:
        logger.error(f"Error fetching comments for invoice {invoice_id}: {e}", 
                    extra={"client_name": client_name})
        return []

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

def fetch_comments_for_client(client_data: dict) -> None:
    """Fetch comments for invoices and add as Description column in Excel."""
    client_name = client_data["Client"]
    logger.extra["client_name"] = client_name

    if not all(client_data.get(key) for key in ["CLIENT_ID", "CLIENT_SECRET", "REFRESH_TOKEN", "ORG_ID"]):
        logger.error("Missing credentials", extra={"client_name": client_name})
        return

    # Fetch access token
    access_token = generate_access_token(
        client_data["CLIENT_ID"], client_data["CLIENT_SECRET"], client_data["REFRESH_TOKEN"], client_name
    )
    if not access_token:
        return

    # Read the existing Excel file with Invoice IDs
    input_file = Path(f"csvdata/input_invoice_aging_details_{client_name}.xlsx")
    try:
        df = pd.read_excel(input_file, sheet_name="Aging_Details")
    except FileNotFoundError:
        logger.error(f"Input file not found: {input_file}", extra={"client_name": client_name})
        return
    except Exception as e:
        logger.error(f"Error reading input file: {e}", extra={"client_name": client_name})
        return

    if 'Invoice ID' not in df.columns:
        logger.error(f"Missing 'Invoice ID' column in {input_file}", extra={"client_name": client_name})
        return

    # Ensure 'Description' column exists
    if 'Description' not in df.columns:
        df['Description'] = ''

    # Move 'Description' column next to 'Invoice ID'
    if 'Invoice ID' in df.columns:
        cols = df.columns.tolist()
        invoice_id_idx = cols.index('Invoice ID')
        if 'Description' in cols:
            cols.remove('Description')
        cols.insert(invoice_id_idx + 1, 'Description')
        df = df[cols]

    # Fetch comments and update Description column with all comment fields
    def format_comment(comment):
        if not isinstance(comment, dict):
            return ''
        fields = [
            f"Comment ID: {comment.get('comment_id', '')}",
            f"Description: {comment.get('description', '')}",
            f"Commented By: {comment.get('commented_by', '')}",
            f"Comment Type: {comment.get('comment_type', '')}",
            f"Operation Type: {comment.get('operation_type', '')}",
            f"Date: {comment.get('date', '')}",
            f"Date Description: {comment.get('date_description', '')}",
            f"Time: {comment.get('time', '')}",
            f"Transaction ID: {comment.get('transaction_id', '')}",
            f"Transaction Type: {comment.get('transaction_type', '')}"
        ]
        return " | ".join([f for f in fields if f.split(": ")[1]])

    # Process invoices in batches to reduce API load
    batch_size = 50
    for i in range(0, len(df), batch_size):
        batch_df = df.iloc[i:i + batch_size].copy()
        logger.info(f"Processing batch {i // batch_size + 1} for {client_name}", 
                    extra={"client_name": client_name})

        # Fetch comments and update Description column
        batch_df['Description'] = batch_df['Invoice ID'].apply(
            lambda x: "; ".join([
                format_comment(comment)
                for comment in fetch_invoice_comments(str(x), client_data["ORG_ID"], access_token, client_name)
                if isinstance(comment, dict)
            ]) if pd.notna(x) else ''
        )
        df.iloc[i:i + batch_size] = batch_df

        # Add delay between batches to avoid rate limiting
        time.sleep(1)

    # Log invoices with no comments
    no_comments = df[df['Description'] == '']['Invoice ID'].count()
    if no_comments > 0:
        logger.info(f"{no_comments} invoices have no comments", extra={"client_name": client_name})

    # Save updated DataFrame to the same file
    write_excel(df, input_file, "Aging_Details", client_name)

def fetch_comments_step():
    """Main function to fetch comments for all clients."""
    for client_data in CREDENTIALS:
        fetch_comments_for_client(client_data)

if __name__ == "__main__":
    fetch_comments_step()