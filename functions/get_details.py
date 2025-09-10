import requests
import os
import logging
import pandas as pd

# Configure logging
logging.basicConfig(level=logging.INFO)

# Function to generate the access token
def generate_access_token(client_id, client_secret, refresh_token):
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
        logging.error(f"Error refreshing access token: {e}")
        return None

# Function to save response content as an Excel file
def save_excel_from_response(response, report_name, client_name):
    output_dir = "csvdata"
    os.makedirs(output_dir, exist_ok=True)
    excel_filename = os.path.join(output_dir, f"input_{report_name}_{client_name}.xlsx")

    try:
        with open(excel_filename, "wb") as f:
            f.write(response.content)  # Save Excel directly from response
        logging.info(f"Excel file saved: {excel_filename}")

        # Process the saved Excel file after saving
        process_excel_file(excel_filename)

    except Exception as e:
        logging.error(f"Error saving the file {excel_filename}: {e}")

# Function to process the Excel file (using the second row as header)
def process_excel_file(filepath):
    try:
        # Read the Excel file, skipping the first row and using the second row as the header (index 1)
        df = pd.read_excel(filepath, header=1)

        # Optional: Drop empty columns or rows if needed
        df.dropna(axis=1, how='all', inplace=True)
        df.dropna(axis=0, how='all', inplace=True)

        # Log the shape of the processed data
        logging.info(f"Processed file {filepath}, shape: {df.shape}")

        # Save the cleaned data back into the same Excel file
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

        logging.info(f"File {filepath} processed and saved.")

    except Exception as e:
        logging.error(f"Error processing {filepath}: {e}")

# Function to fetch the report and save it
def fetch_report(report_name, url_template, client_data, date_filter):
    access_token = generate_access_token(
        client_data["CLIENT_ID"], 
        client_data["CLIENT_SECRET"], 
        client_data["REFRESH_TOKEN"]
    )
    if not access_token:
        logging.error("Failed to obtain access token.")
        return
    
    url = url_template.format(ORG_ID=client_data["ORG_ID"], value=date_filter)
    headers = {"Authorization": f"Zoho-oauthtoken {access_token}"}
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()

        # Check if the response content is not empty
        if response.content:
            save_excel_from_response(response, report_name, client_data["Client"])
        else:
            logging.warning(f"Empty response for {report_name} from {client_data['Client']}.")

    except requests.RequestException as e:
        logging.error(f"Error fetching {report_name}: {e}")

# Function to get the appropriate URL based on client
def get_customer_balance_url(client_name):
    if client_name.lower() == "nvb":
        # URL for NVB customer
        return """https://www.zohoapis.com/books/v3/reports/customerbalancesummary?accept=xlsx&page=1&per_page=20000&sort_order=A&filter_by=TransactionDate.{value}&select_columns=%5B%7B%22field%22%3A%22customer_name%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22bcy_invoice_balance%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22bcy_available_credits%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22closing_balance%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22last_name%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22mobile_phone%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22email%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22custom_field_1941648000001342037%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22custom_field_1941648000003467027%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22custom_field_1941648000003146083%22%2C%22group%22%3A%22contact%22%7D%5D&is_for_date_range=false&usestate=true&group_by=%5B%7B%22field%22%3A%22none%22%2C%22group%22%3A%22report%22%7D%5D&sort_column=customer_name&can_ignore_zero_cb=false&response_option=1&x-zb-source=zbclient&formatneeded=true&paper_size=A4&orientation=portrait&font_family_for_body=opensans&margin_top=0.7&margin_bottom=0.7&margin_left=0.55&margin_right=0.2&table_size=classic&table_style=default&show_org_name=true&show_generated_date=false&show_generated_time=false&show_page_number=false&show_report_basis=true&show_generated_by=false&can_fit_to_page=true&watermark_opacity=50&show_org_logo_in_header=false&show_org_logo_as_watermark=false&watermark_position=center+center&watermark_zoom=50&file_name=Customer+Balance+Summary&organization_id={ORG_ID}&frameorigin=https%3A%2F%2Fbooks.zoho.com"""
    else:
        # URL for SMCS customer (existing URL)
        return """https://www.zohoapis.com/books/v3/reports/customerbalancesummary?accept=xlsx&page=1&per_page=20000&sort_order=A&filter_by=TransactionDate.{value}&select_columns=%5B%7B%22field%22%3A%22customer_name%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22bcy_invoice_balance%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22bcy_available_credits%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22closing_balance%22%2C%22group%22%3A%22report%22%7D%2C%7B%22field%22%3A%22last_name%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22mobile_phone%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22email%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22custom_field_544542000001383001%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22custom_field_544542000011019221%22%2C%22group%22%3A%22contact%22%7D%2C%7B%22field%22%3A%22custom_field_544542000010260003%22%2C%22group%22%3A%22contact%22%7D%5D&is_for_date_range=false&usestate=false&group_by=%5B%7B%22field%22%3A%22none%22%2C%22group%22%3A%22report%22%7D%5D&sort_column=customer_name&can_ignore_zero_cb=false&response_option=1&x-zb-source=zbclient&formatneeded=true&paper_size=A4&orientation=portrait&font_family_for_body=opensans&margin_top=0.7&margin_bottom=0.7&margin_left=0.55&margin_right=0.2&table_size=classic&show_generated_date=false&show_generated_time=false&show_page_number=false&show_report_basis=true&show_generated_by=false&can_fit_to_page=true&watermark_opacity=50&show_org_logo_in_header=false&show_org_logo_as_watermark=false&watermark_position=center+center&watermark_zoom=50&file_name=Customer+Balance+Summary&organization_id={ORG_ID}&frameorigin=https%3A%2F%2Fbooks.zoho.com"""

# Function to fetch all reports for different clients
def fetch_all_reports(date_filter):
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
    
    # Common reports configuration
    COMMON_REPORTS = {
        "invoice_aging": {
            "url": "https://www.zohoapis.com/books/v3/reports/aragingdetails?accept=xlsx&organization_id={ORG_ID}&page=1&per_page=100000&sort_order=A&sort_column=date&interval_range=15&number_of_columns=4&interval_type=days&group_by=none&filter_by=InvoiceDueDate.{value}&entity_list=invoice&is_new_flow=true&response_option=1",
            "key": "invoiceaging"
        }
    }
    
    for client_data in CREDENTIALS:
        # Fetch common reports (invoice_aging) for all clients
        for report_name, report_data in COMMON_REPORTS.items():
            fetch_report(report_name, report_data["url"], client_data, date_filter)
        
        # Fetch customer balance summary with client-specific URL
        customer_balance_url = get_customer_balance_url(client_data["Client"])
        fetch_report("customer_balance", customer_balance_url, client_data, date_filter)


if __name__ == "__main__":
    fetch_all_reports("")