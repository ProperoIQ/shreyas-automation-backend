import pandas as pd
import os

# Function to process each file and add new columns for age ranges
def process_file(input_file, output_file):
    # Ensure the output directory exists
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Read the file
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading {input_file}: {e}")
        return

    # Extract numeric value from the 'Age' column
    df['age'] = pd.to_numeric(df['age'], errors='coerce')

    # Create new columns for the age ranges
    df['3Yrs>='] = df['balance'].where(df['age'] >= 1095, 0)
    df['3Yr<=2Yr'] = df['balance'].where((df['age'] >= 730) & (df['age'] < 1095), 0)
    df['2Yr<=1Yr'] = df['balance'].where((df['age'] >= 365) & (df['age'] < 730), 0)
    df['1Yr<=180days'] = df['balance'].where((df['age'] >= 180) & (df['age'] < 365), 0)
    df['180<=90days'] = df['balance'].where((df['age'] >= 90) & (df['age'] < 180), 0)
    df['90<=60days'] = df['balance'].where((df['age'] >= 60) & (df['age'] < 90), 0)
    df['60<=30days'] = df['balance'].where((df['age'] >= 30) & (df['age'] < 60), 0)
    df['>=30days'] = df['balance'].where(df['age'] < 30, 0)
    df['Age_Not_Provided'] = df['balance'].where(df['age'].isna(), 0)

    # Calculate totals
    age_totals = df[['3Yrs>=', '3Yr<=2Yr', '2Yr<=1Yr', '1Yr<=180days', '180<=90days', '90<=60days', '60<=30days', '>=30days', 'Age_Not_Provided']].sum()
    total_balance = df['balance'].sum()
    age_counts = df[['3Yrs>=', '3Yr<=2Yr', '2Yr<=1Yr', '1Yr<=180days', '180<=90days', '90<=60days', '60<=30days', '>=30days', 'Age_Not_Provided']].astype(bool).sum()
    balance_count = df['balance'].notna().sum()

 

    # Save to file
    try:
        df.to_excel(output_file, index=False)
        print(f"Processed file saved to {output_file}")
    except Exception as e:
        print(f"Error saving the file: {e}")

# Function to process both files
def process_multiple_files(file1, file2):
    process_file(file1, 'output/NVB_Age_Range_Columns.xlsx')
    process_file(file2, 'output/SMCS_Age_Range_Columns.xlsx')
