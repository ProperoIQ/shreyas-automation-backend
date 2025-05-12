import streamlit as st
from time import sleep
from functions.segregator import process_multiple_files
from functions.consolidater import process_and_merge_files
from functions.balance_summary import process_file
from functions.age_summary import generate_summary
from functions.combiner import combine_sheets
from functions.get_details import fetch_all_reports

def process_files():
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    status_text.text("Fetching all reports...")
    fetch_all_reports()
    progress_bar.progress(10)
    
    status_text.text("Processing AR aging details files...")
    file1 = 'csvdata/invoice_aging_details_nvb.xlsx'
    file2 = 'csvdata/invoice_aging_details_smcs.xlsx'
    process_multiple_files(file1, file2)
    progress_bar.progress(30)
    
    status_text.text("Generating age summary...")
    input_files = {
        'SMCS': 'output/SMCS_Age_Range_Columns.xlsx',
        'NVB': 'output/NVB_Age_Range_Columns.xlsx'
    }
    output_file = 'output/Age_summary.xlsx'
    generate_summary(input_files, output_file)
    progress_bar.progress(50)
    
    status_text.text("Merging customer balance files...")
    file1_path = 'csvdata/customer_balance_summary_details_nvb.xlsx'
    file2_path = 'csvdata/customer_balance_summary_details_smcs.xlsx'
    output_file_path = 'output/unified_file.csv'
    process_and_merge_files(file1_path, file2_path, output_file_path)
    progress_bar.progress(70)
    
    status_text.text("Processing consolidated file...")
    input_file = 'output/unified_file.csv'
    output_file = 'output/balances_summary.xlsx'
    process_file(input_file, output_file)
    progress_bar.progress(85)
    
    status_text.text("Combining all summaries...")
    input_file1 = 'output/balances_summary.xlsx'
    input_file2 = 'output/Age_summary.xlsx'
    final_output = 'output/Final.xlsx'
    combine_sheets(input_file1, input_file2, final_output)
    progress_bar.progress(100)
    
    st.success("Final Report Generated Successfully!")
    st.download_button(label="Download Final Report", data=open(final_output, "rb").read(), file_name="Final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.session_state.processing = False

def main():
    st.title("Automated File Processing Application")
    
    if "processing" not in st.session_state:
        st.session_state.processing = False
    
    if not st.session_state.processing:
        if st.button("Start Processing"):
            st.session_state.processing = True
            process_files()

if __name__ == "__main__":
    main()
