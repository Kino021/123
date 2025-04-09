import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pandas import ExcelWriter
from functools import lru_cache

st.set_page_config(layout="wide", page_title="BANK REPORT", page_icon="📊", initial_sidebar_state="expanded")

st.title('BANK REPORT')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file, usecols=lambda x: x.strip().upper() in 
                      ['DATE', 'REMARK BY', 'DEBTOR', 'STATUS', 'REMARK', 'CALL STATUS', 'CARD NO.', 
                       'ACCOUNT NO.', 'CALL DURATION', 'REMARK TYPE', 'PTP AMOUNT', 'BALANCE', 'TALK TIME DURATION', 'CLIENT'])
    df.columns = df.columns.str.strip().str.upper()
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    return df[df['DATE'].dt.weekday != 6]  # Exclude Sundays

@st.cache_data
def filter_dataframe(df):
    df = df[df['REMARK BY'] != 'SPMADRID']
    df = df[~df['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False)]
    df = df[~df['STATUS'].str.contains('ABORT', na=False)]
    df = df[~df['REMARK'].str.contains(r'1_\d{11} - PTP NEW', case=False, na=False, regex=True)]
    
    excluded_remarks = ["Broken Promise", "New files imported", "Updates when case reassign to another collector", 
                       "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", "New Assignment -"]
    mask = df['REMARK'].str.contains('|'.join(excluded_remarks), case=False, na=False)
    df = df[~mask]
    df = df[~df['CALL STATUS'].str.contains('OTHERS', case=False, na=False)]
    
    df['CARD NO.'] = df['CARD NO.'].astype(str)
    # Modified: Use full CARD NO. as CYCLE instead of first 2 characters
    df['CYCLE'] = df['CARD NO.'].fillna('Unknown')
    return df

@lru_cache(maxsize=128)
def format_seconds_to_hms(seconds):
    seconds = int(seconds)
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"

def sanitize_sheet_name(name):
    invalid_chars = r'[:\\/*?[\]]'
    name = ''.join(c for c in name if c not in invalid_chars)
    return name[:31]

@st.cache_data
def to_excel(df_dict):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00'}),
            'center': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'header': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': 'red', 'font_color': 'white', 'bold': True}),
            'comma': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'}),
            'percent': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.00%'}),
            'date': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'yyyy-mm-dd'}),
            'time': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'hh:mm:ss'})
        }
        
        for sheet_name, df in df_dict.items():
            sheet_name = sanitize_sheet_name(sheet_name)
            df_for_excel = df.copy()
            for col in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
                if col in df_for_excel.columns:
                    df_for_excel[col] = df_for_excel[col].str.rstrip('%').astype(float) / 100
            
            df_for_excel.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
            worksheet = writer.sheets[sheet_name]
            worksheet.merge_range('A1:' + chr(65 + len(df.columns) - 1) + '1', sheet_name, formats['title'])
            
            for col_num, col_name in enumerate(df_for_excel.columns):
                worksheet.write(1, col_num, col_name, formats['header'])
                max_len = max(df_for_excel[col_name].astype(str).str.len().max(), len(col_name)) + 2
                worksheet.set_column(col_num, col_num, max_len)
                
                col_data = df_for_excel[col_name]
                for row_num, value in enumerate(col_data, 2):
                    fmt = (formats['date'] if col_name == 'DATE' else
                           formats['comma'] if col_name in ['TOTAL PTP AMOUNT', 'TOTAL BALANCE'] else
                           formats['percent'] if col_name in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'PTP RATE', 'CALL DROP RATIO #'] else
                           formats['time'] if col_name in ['TOTAL TALK TIME', 'TALK TIME AVE'] else
                           formats['center'])
                    if col_name == 'DATE' and isinstance(value, (pd.Timestamp, datetime.date)):
                        worksheet.write_datetime(row_num, col_num, value, fmt)
                    else:
                        worksheet.write(row_num, col_num, value, fmt)
    
    return output.getvalue()

@st.cache_data
def calculate_summary(df, remark_types, manual_correction=False):
    # Filter by remark types
    df_filtered = df[df['REMARK TYPE'].isin(remark_types)].copy()
    
    # Define columns for the summary DataFrame
    summary_columns = ['DATE', 'CLIENT', 'COLLECTORS', 'ACCOUNTS', 'TOTAL DIALED', 
                       'PENETRATION RATE (%)', 'CONNECTED #', 'CONNECTED RATE (%)', 
                       'CONNECTED ACC', 'TOTAL TALK TIME', 'TALK TIME AVE', 'CONNECTED AVE', 
                       'PTP ACC', 'PTP RATE', 'TOTAL PTP AMOUNT', 'TOTAL BALANCE', 
                       'CALL DROP #', 'SYSTEM DROP', 'CALL DROP RATIO #']
    
    # Check if DataFrame is empty or missing 'DATE'
    if 'DATE' not in df_filtered.columns or df_filtered.empty:
        return pd.DataFrame(columns=summary_columns)
    
    # Ensure 'DATE' is in the correct format
    df_filtered['DATE'] = pd.to_datetime(df_filtered['DATE'], errors='coerce').dt.date
    
    summary_data = []
    for (date, client), group in df_filtered.groupby(['DATE', 'CLIENT']):
        collectors = group['REMARK BY'].nunique() if group['CALL DURATION'].notna().any() else 0
        if collectors == 0:
            continue
            
        accounts = group['ACCOUNT NO.'].nunique()
        total_dialed = len(group)
        connected = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].nunique()
        connected_acc = group[group['CALL STATUS'] == 'CONNECTED'].shape[0]
        
        penetration_rate = f"{(total_dialed / accounts * 100) if accounts else 0:.2f}%"
        connected_rate = f"{(connected_acc / total_dialed * 100) if total_dialed else 0:.2f}%"
        
        ptp_acc = group[(group['STATUS'].str.contains('PTP', na=False)) & (group['PTP AMOUNT'] != 0)]['ACCOUNT NO.'].nunique()
        ptp_rate = f"{(ptp_acc / connected * 100) if connected else 0:.2f}%"
        
        total_ptp_amount = group[group['PTP AMOUNT'] != 0]['PTP AMOUNT'].sum()
        total_balance = group[group['PTP AMOUNT'] != 0]['BALANCE'].sum()
        
        system_drop = group[(group['STATUS'].str.contains('DROPPED', na=False)) & (group['REMARK BY'] == 'SYSTEM')].shape[0]
        call_drop_count = group[(group['STATUS'].str.contains('NEGATIVE CALLOUTS - DROP CALL|NEGATIVE_CALLOUTS - DROPPED_CALL', na=False)) & 
                              (~group['REMARK BY'].str.upper().isin(['SYSTEM']))].shape[0]
        
        call_drop_ratio = f"{((call_drop_count if manual_correction else system_drop) / connected_acc * 100) if connected_acc else 0:.2f}%"
        
        total_talk_seconds = group['TALK TIME DURATION'].sum()
        total_talk_time = format_seconds_to_hms(total_talk_seconds)
        talk_time_ave = format_seconds_to_hms(total_talk_seconds / collectors) if collectors else "00:00:00"
        connected_ave = round(connected_acc / collectors, 2) if collectors else 0

        summary_data.append({
            'DATE': date, 'CLIENT': client, 'COLLECTORS': collectors, 'ACCOUNTS': accounts,
            'TOTAL DIALED': total_dialed, 'PENETRATION RATE (%)': penetration_rate,
            'CONNECTED #': connected, 'CONNECTED RATE (%)': connected_rate, 'CONNECTED ACC': connected_acc,
            'TOTAL TALK TIME': total_talk_time, 'TALK TIME AVE': talk_time_ave, 'CONNECTED AVE': connected_ave,
            'PTP ACC': ptp_acc, 'PTP RATE': ptp_rate, 'TOTAL PTP AMOUNT': total_ptp_amount,
            'TOTAL BALANCE': total_balance, 'CALL DROP #': call_drop_count, 'SYSTEM DROP': system_drop,
            'CALL DROP RATIO #': call_drop_ratio
        })
    
    # Create DataFrame and sort by 'DATE' if data exists
    summary_df = pd.DataFrame(summary_data, columns=summary_columns)
    if not summary_df.empty:
        summary_df = summary_df.sort_values(by=['DATE'])
    return summary_df

def get_cycle_summary(df, remark_types, manual_correction=False):
    return {f"Cycle {cycle}": calculate_summary(df[df['CYCLE'] == cycle], remark_types, manual_correction)
            for cycle in df['CYCLE'].unique() if cycle not in ['Unknown', 'na', 'NA']}

def get_balance_summary(df, remark_types):
    balance_bins = [(0.00, 9999.99, "0-9999.99"), (10000.00, 49999.99, "10K-49K"),
                   (50000.00, 99999.99, "50K-99K"), (100000.00, float('inf'), "100K+")]
    return {f"Balance {label}": calculate_summary(df[(df['BALANCE'] >= min_bal) & (df['BALANCE'] <= max_bal)], remark_types)
            for min_bal, max_bal, label in balance_bins if not df[(df['BALANCE'] >= min_bal) & (df['BALANCE'] <= max_bal)].empty}

uploaded_files = st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    # Process individual files
    all_dfs = []
    for file_idx, uploaded_file in enumerate(uploaded_files):
        st.write(f"### Processing File {file_idx + 1}: {uploaded_file.name}")
        df = filter_dataframe(load_data(uploaded_file))
        all_dfs.append(df)
        
        if not df.empty:
            summaries = {
                'combined': calculate_summary(df, ['Predictive', 'Follow Up', 'Outgoing']),
                'predictive': calculate_summary(df, ['Predictive', 'Follow Up']),
                'manual': calculate_summary(df, ['Outgoing'], manual_correction=True),
                'predictive_cycles': get_cycle_summary(df, ['Predictive', 'Follow Up']),
                'manual_cycles': get_cycle_summary(df, ['Outgoing'], manual_correction=True),
                'balance': get_balance_summary(df, ['Predictive', 'Follow Up', 'Outgoing'])
            }
            
            for title, data in [('Overall Combined', summaries['combined']),
                              ('Overall Predictive', summaries['predictive']),
                              ('Overall Manual', summaries['manual'])]:
                st.write(f"#### File {file_idx + 1} - {title} Summary Table")
                st.write(data)
            
            for cycle_type, cycles in [('Predictive', summaries['predictive_cycles']), 
                                     ('Manual', summaries['manual_cycles'])]:
                st.write(f"#### File {file_idx + 1} - Per Cycle {cycle_type} Summary Tables")
                for cycle, table in cycles.items():
                    with st.container():
                        st.subheader(f"Summary for {cycle}")
                        st.write(table)
            
            st.write(f"#### File {file_idx + 1} - Per Balance Overall Summary Tables")
            for balance, table in summaries['balance'].items():
                with st.container():
                    st.subheader(f"Summary for {balance}")
                    st.write(table)
            
            excel_data = {'Combined Summary': summaries['combined'], 'Predictive Summary': summaries['predictive'],
                         'Manual Summary': summaries['manual'], **summaries['predictive_cycles'],
                         **summaries['manual_cycles'], **summaries['balance']}
            st.download_button(f"Download Summaries for File {file_idx + 1}", to_excel(excel_data),
                            f"Daily_Remark_Summary_{uploaded_file.name.split('.')[0]}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.write("---")
        else:
            st.warning(f"No valid data found in file: {uploaded_file.name}")

    # Concatenated result
    if len(all_dfs) > 1:
        st.write("### Concatenated Results for All Files")
        combined_df = pd.concat(all_dfs, ignore_index=True)
        
        combined_summaries = {
            'combined': calculate_summary(combined_df, ['Predictive', 'Follow Up', 'Outgoing']),
            'predictive': calculate_summary(combined_df, ['Predictive', 'Follow Up']),
            'manual': calculate_summary(combined_df, ['Outgoing'], manual_correction=True),
            'predictive_cycles': get_cycle_summary(combined_df, ['Predictive', 'Follow Up']),
            'manual_cycles': get_cycle_summary(combined_df, ['Outgoing'], manual_correction=True),
            'balance': get_balance_summary(combined_df, ['Predictive', 'Follow Up', 'Outgoing'])
        }
        
        for title, data in [('Overall Combined', combined_summaries['combined']),
                          ('Overall Predictive', combined_summaries['predictive']),
                          ('Overall Manual', combined_summaries['manual'])]:
            st.write(f"#### {title} Summary Table (All Files)")
            st.write(data)
        
        excel_data = {'Combined Summary': combined_summaries['combined'],
                     'Predictive Summary': combined_summaries['predictive'],
                     'Manual Summary': combined_summaries['manual'],
                     **combined_summaries['predictive_cycles'],
                     **combined_summaries['manual_cycles'],
                     **combined_summaries['balance']}
        st.download_button("Download Concatenated Summaries", to_excel(excel_data),
                         f"Daily_Remark_Summary_Combined_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Please upload one or more Excel files to begin.")
