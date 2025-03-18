import pandas as pd
import streamlit as st

# Function to load data
def load_data(uploaded_file):
    df = pd.read_csv(uploaded_file)
    return df

# Define the Per Balance Summary with the specific balance ranges
def calculate_balance_summary(df):
    # Defining the balance ranges based on your criteria
    balance_ranges = {
        "6,000.00 to 49,999.99": (6000.00, 49999.99),
        "50,000.00 to 99,999.99": (50000.00, 99999.99),
        "100,000.00 and above": (100000.00, float('inf')),
    }

    # Initialize a summary table for balance ranges
    balance_summary_table = pd.DataFrame(columns=[ 
        'Balance Range', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE (%)', 'CONNECTED #', 
        'CONNECTED RATE (%)', 'CONNECTED ACC', 'PTP ACC', 'PTP RATE', 'CALL DROP #', 
        'SYSTEM DROP', 'CALL DROP RATIO #', 'OVERALL COMBINED', 'OVERALL PREDICTIVE', 
        'OVERALL MANUAL', 'PREDICTIVE PER CYCLE', 'MANUAL PER CYCLE'
    ]) 

    # Iterate through the defined balance ranges
    for range_name, (min_balance, max_balance) in balance_ranges.items():
        # Filter data by balance range
        df_filtered = df[(df['Balance'] >= min_balance) & (df['Balance'] <= max_balance)]
        
        # Calculate the summary statistics for each balance range
        accounts = df_filtered['Account No.'].nunique()
        total_dialed = df_filtered['Account No.'].count()
        connected = df_filtered[df_filtered['Call Status'] == 'CONNECTED']['Account No.'].nunique()
        penetration_rate = (total_dialed / accounts * 100) if accounts != 0 else None
        connected_acc = df_filtered[df_filtered['Call Status'] == 'CONNECTED']['Account No.'].count()
        connected_rate = (connected_acc / total_dialed * 100) if total_dialed != 0 else None
        ptp_acc = df_filtered[(df_filtered['Status'].str.contains('PTP', na=False)) & (df_filtered['PTP Amount'] != 0)]['Account No.'].nunique()
        ptp_rate = (ptp_acc / connected * 100) if connected != 0 else None
        system_drop = df_filtered[(df_filtered['Status'].str.contains('DROPPED', na=False)) & (df_filtered['Remark By'] == 'SYSTEM')]['Account No.'].count()
        call_drop_count = df_filtered[(df_filtered['Status'].str.contains('NEGATIVE CALLOUTS - DROP CALL', na=False)) & 
                                      (~df_filtered['Remark By'].str.upper().isin(['SYSTEM']))]['Account No.'].count()
        call_drop_ratio = (system_drop / connected_acc * 100) if connected_acc != 0 else None
        
        # Additional predictive and manual calculations
        predictive_df = df_filtered[df_filtered['Call Type'] == 'Predictive']
        manual_df = df_filtered[df_filtered['Call Type'] == 'Manual']

        # Overall combined values
        overall_combined = total_dialed
        overall_predictive = predictive_df['Account No.'].count()
        overall_manual = manual_df['Account No.'].count()

        # Predictive and Manual per cycle (per total dialed)
        predictive_per_cycle = (overall_predictive / total_dialed * 100) if total_dialed != 0 else None
        manual_per_cycle = (overall_manual / total_dialed * 100) if total_dialed != 0 else None

        # Append the calculated summary for this balance range to the summary table
        balance_summary_table = pd.concat([balance_summary_table, pd.DataFrame([{
            'Balance Range': range_name,
            'ACCOUNTS': accounts,
            'TOTAL DIALED': total_dialed,
            'PENETRATION RATE (%)': f"{round(penetration_rate)}%" if penetration_rate is not None else None,
            'CONNECTED #': connected,
            'CONNECTED RATE (%)': f"{round(connected_rate)}%" if connected_rate is not None else None,
            'CONNECTED ACC': connected_acc,
            'PTP ACC': ptp_acc,
            'PTP RATE': f"{round(ptp_rate)}%" if ptp_rate is not None else None,
            'CALL DROP #': call_drop_count,
            'SYSTEM DROP': system_drop,
            'CALL DROP RATIO #': f"{round(call_drop_ratio)}%" if call_drop_ratio is not None else None,
            'OVERALL COMBINED': overall_combined,
            'OVERALL PREDICTIVE': overall_predictive,
            'OVERALL MANUAL': overall_manual,
            'PREDICTIVE PER CYCLE': f"{round(predictive_per_cycle)}%" if predictive_per_cycle is not None else None,
            'MANUAL PER CYCLE': f"{round(manual_per_cycle)}%" if manual_per_cycle is not None else None,
        }])], ignore_index=True)

    return balance_summary_table

# Streamlit app
def main():
    # Title of the app
    st.title('Per Balance Summary')
    
    # File uploader for CSV file
    uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])
    
    if uploaded_file is not None:
        # Load the data from the uploaded CSV
        df = load_data(uploaded_file)

        # Show a preview of the data
        st.write("### Preview of the Uploaded Data", df.head())

        # Ensure the column names are correct (Adjust according to your CSV column names)
        if 'Balance' in df.columns and 'Account No.' in df.columns and 'Call Type' in df.columns:
            # Add your other filtering logic here (e.g., excluding specific rows)
            st.write("## Per Balance Summary Table")
            balance_summary_table = calculate_balance_summary(df)
            st.write(balance_summary_table)
        else:
            st.write("### Error: Ensure 'Balance', 'Account No.', and 'Call Type' columns are present in your CSV file.")

if __name__ == "__main__":
    main()
