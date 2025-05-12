
from segregator import process_multiple_files
from consolidater import process_and_merge_files
from balance_summary import process_file
from age_summary import generate_summary
from combiner import combine_sheets
from get_details import fetch_all_reports
def main():


    fetch_all_reports()
    # Process the AR aging details files
    file1 = 'csvdata/invoice_aging_details_nvb.xlsx'
    file2 = 'csvdata/invoice_aging_details_smcs.xlsx'
    process_multiple_files(file1, file2)
    input_files = {
    'SMCS': 'output/SMCS_Age_Range_Columns.xlsx',
    'NVB': 'output/NVB_Age_Range_Columns.xlsx'
     }
    output_file = 'output/Age_summary.xlsx'
    generate_summary(input_files, output_file)

    # Process and merge the customer balance files
    file1_path = 'csvdata/customer_balance_summary_details_nvb.xlsx'
    file2_path = 'csvdata/customer_balance_summary_details_smcs.xlsx'
    output_file_path = 'output/unified_file.csv'
    process_and_merge_files(file1_path, file2_path, output_file_path)

    # Process the consolidated file and generate summary
    input_file = 'output/unified_file.csv'
    output_file = 'output/balances_summary.xlsx'
    process_file(input_file, output_file)
    
    # Example usage
    input_file1 = 'output/balances_summary.xlsx'
    input_file2 = 'output/Age_summary.xlsx'
    output_file = 'output/Final.xlsx'

    combine_sheets(input_file1, input_file2, output_file)
    
    
if __name__ == "__main__":
   main()