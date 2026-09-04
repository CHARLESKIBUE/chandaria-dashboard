import pandas as pd
import shutil
import time
import os
from pathlib import Path
from datetime import datetime
from dateutil.relativedelta import relativedelta
import warnings

# Suppress pandas warnings for cleaner terminal output
warnings.filterwarnings('ignore')

# ==========================================
# CONFIGURATION
# ==========================================
BASE_DIR = Path(__file__).parent.resolve()
INPUT_DIR = BASE_DIR / "Input"
OUTPUT_DIR = BASE_DIR / "Output"             # For Power BI (Full data, untouched)
WEB_OUTPUT_DIR = BASE_DIR / "Web_Output"     # For Streamlit / GitHub (Recent window)
ARCHIVE_DIR = BASE_DIR / "Monthly_Archive" 
DUMPS_DIR = BASE_DIR / "Raw_Dumps"         

for folder in [INPUT_DIR, OUTPUT_DIR, WEB_OUTPUT_DIR, ARCHIVE_DIR, DUMPS_DIR]:
    folder.mkdir(exist_ok=True)

REPORT_TYPES = [
    "Availability Report",
    "Coverage Report",
    "Planogram Report",
    "Share of Shelf Report"
]

DATE_COLUMN = "Date" 
GAP_REASON_COLUMN = "Gap Reason" 
BRAND_NAME_COLUMN = "Brand Name" 
SUPERVISOR_COLUMN = "Supervisor Name" 

# ==========================================
# SMART DATE PARSER & FORMATTERS
# ==========================================
def parse_dates_safely(date_series):
    parsed = pd.to_datetime(date_series, dayfirst=True, errors='coerce')
    today = pd.Timestamp(datetime.now().date())
    future_mask = parsed > today
    if future_mask.any():
        future_strings = date_series[future_mask]
        parsed.loc[future_mask] = pd.to_datetime(future_strings, dayfirst=False, errors='coerce')
    return parsed

def get_date_suffix(day):
    if 11 <= day <= 13: return 'TH'
    return {1: 'ST', 2: 'ND', 3: 'RD'}.get(day % 10, 'TH')

def format_date_custom(date_obj):
    day = date_obj.day
    suffix = get_date_suffix(day)
    month = date_obj.strftime('%B').upper()
    return f"{day}{suffix} {month}"

def get_max_width(df, col_name, col_index):
    header_len = max([len(str(line)) for line in str(col_name).split('\n')])
    if df.empty:
        data_len = 0
    else:
        data_len = df.iloc[:, col_index].astype(str).str.len().max()
        data_len = 0 if pd.isna(data_len) else int(data_len)
    return min(max(header_len, data_len) + 4, 50)

def safe_read_csv(file_path):
    """Safely reads CSVs, bypassing Google Drive virtual drive & Null Byte errors."""
    file_path = Path(file_path)
    
    if file_path.exists() and file_path.stat().st_size == 0:
        print(f"   ⚠️ File is physically empty: {file_path.name}")
        return pd.DataFrame()

    try:
        return pd.read_csv(file_path, engine='python', on_bad_lines='skip')
    except pd.errors.EmptyDataError:
        print(f"   ⚠️ No columns to parse in {file_path.name} (Empty Data). Skipping.")
        return pd.DataFrame()
    except Exception as e:
        print(f"   ⚠️ Standard read failed ({e}). Attempting safe fallback...")
        try:
            with open(file_path, 'r', encoding='utf-8-sig', errors='replace') as f:
                return pd.read_csv(f, engine='python', on_bad_lines='skip')
        except pd.errors.EmptyDataError:
            return pd.DataFrame()

# ==========================================
# EXCEL DASHBOARD BUILDER (REUSABLE)
# ==========================================
def build_dashboard_workbook(df_chronic, title_prefix, file_name, output_dir, sku_col, brand_col, start_date, end_date):
    if df_chronic.empty:
        print(f"  ⚠️ No data available to generate dashboard for {title_prefix}.")
        return
        
    report_path = output_dir / file_name
    print(f"  -> Building {title_prefix} Dashboard...")

    with pd.ExcelWriter(report_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- Theme Formats ---
        title_fmt = workbook.add_format({'bold': True, 'font_size': 18, 'bg_color': '#7F7F7F', 'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        kpi_head_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'align': 'center', 'border': 1})
        kpi_val_blue = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#0070C0', 'align': 'center', 'border': 1, 'valign': 'vcenter', 'text_wrap': True})
        sub_title_fmt = workbook.add_format({'bold': True, 'bg_color': '#B4C6E7', 'border': 1, 'align': 'center', 'font_color': 'black'})
        sub_head_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        cell_fmt = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        divider_fmt = workbook.add_format({'bg_color': '#EFEFEF'}) 
        
        # ----------------------------------------------------
        # DASHBOARD SHEET
        # ----------------------------------------------------
        dash_sheet = 'Summary'
        ws_dash = workbook.add_worksheet(dash_sheet)
        
        ws_dash.set_column('E:E', 3, divider_fmt)
        
        # 1. Main Title
        header_text = f"{title_prefix.upper()} SUMMARY ({format_date_custom(start_date)} TO {format_date_custom(end_date)})"
        ws_dash.merge_range('A1:I1', header_text, title_fmt)
        ws_dash.set_row(0, 30)
        
        # 2. Pre-Calculate Top Entities 
        top_brand = df_chronic[brand_col].mode()[0] if not df_chronic[brand_col].empty else "N/A"
        top_sku = df_chronic[sku_col].mode()[0] if not df_chronic[sku_col].empty else "N/A"

        # 3. Inject Hidden Dynamic Array Formulas
        end_row_excel = 19 + len(df_chronic)
        outlets_range = f'B20:B{end_row_excel}'
        skus_range = f'D20:D{end_row_excel}'
        
        f_unique_outlets = f'=IFERROR(ROWS(UNIQUE(FILTER({outlets_range}, SUBTOTAL(103, OFFSET(B20, ROW({outlets_range})-20, 0))))), 0)'
        f_unique_skus = f'=IFERROR(ROWS(UNIQUE(FILTER({skus_range}, SUBTOTAL(103, OFFSET(D20, ROW({skus_range})-20, 0))))), 0)'

        ws_dash.write_dynamic_array_formula('AA1', f_unique_outlets)
        ws_dash.write_dynamic_array_formula('AB1', f_unique_skus)
        ws_dash.set_column('AA:AB', None, None, {'hidden': True})

        # 4. Dynamic KPI Banner 
        ws_dash.merge_range('A2:B2', 'Count of Outlets', kpi_head_fmt)
        ws_dash.merge_range('A3:B3', '', kpi_val_blue) 
        ws_dash.write_formula('A3', '=AA1', kpi_val_blue) 
        
        ws_dash.merge_range('C2:D2', 'Count of SKUs', kpi_head_fmt)
        ws_dash.merge_range('C3:D3', '', kpi_val_blue)
        ws_dash.write_formula('C3', '=AB1', kpi_val_blue)
        
        ws_dash.merge_range('F2:G2', 'Most Affected Brand', kpi_head_fmt)
        ws_dash.merge_range('F3:G3', str(top_brand), kpi_val_blue)
        
        ws_dash.merge_range('H2:I2', 'Most Affected SKU', kpi_head_fmt)
        ws_dash.merge_range('H3:I3', str(top_sku), kpi_val_blue)
        
        ws_dash.set_row(2, 35)

        # 5. Top 10 Side-by-Side Summary Headers
        ws_dash.merge_range('A5:D5', 'Top Persistent Issues (SKU Summary)', sub_title_fmt)
        ws_dash.merge_range('F5:I5', 'Top Persistent Issues (Outlet Summary)', sub_title_fmt) 
        
        sku_headers = ['Brand', 'Product Name', 'Total Incidents', 'Outlets Affected']
        for col_num, col_name in enumerate(sku_headers):
            ws_dash.write(5, col_num, col_name, sub_head_fmt)
            
        outlet_headers = ['Account', 'Outlet Name', 'Total Incidents', 'SKUs Affected']
        for col_num, col_name in enumerate(outlet_headers):
            ws_dash.write(5, col_num + 5, col_name, sub_head_fmt) 

        # Calculate Static Top 10s via Pandas
        sku_breakdown = df_chronic.groupby([brand_col, sku_col]).agg(
            Incidents=(DATE_COLUMN, 'count'),
            Outlets_Affected=('Store Name', 'nunique')
        ).reset_index().sort_values(by='Incidents', ascending=False).head(10)
        
        outlet_breakdown = df_chronic.groupby(['Store Account', 'Store Name']).agg(
            Incidents=(DATE_COLUMN, 'count'),
            SKUs_Affected=(sku_col, 'nunique')
        ).reset_index().sort_values(by='Incidents', ascending=False).head(10)

        # Safely Write Static Data
        for row_num in range(10):
            if row_num < len(sku_breakdown):
                row_data = sku_breakdown.iloc[row_num].values
                for col_num, cell_data in enumerate(row_data):
                    ws_dash.write(6 + row_num, col_num, cell_data, cell_fmt)
            else:
                for col_num in range(4):
                    ws_dash.write_string(6 + row_num, col_num, "-", cell_fmt)
            
            if row_num < len(outlet_breakdown):
                row_data = outlet_breakdown.iloc[row_num].values
                for col_num, cell_data in enumerate(row_data):
                    ws_dash.write(6 + row_num, col_num + 5, cell_data, cell_fmt)
            else:
                for col_num in range(4):
                    ws_dash.write_string(6 + row_num, col_num + 5, "-", cell_fmt)

        # 6. Master Granular Table 
        start_row = 18 
        
        granular_cols = ['Store Account', 'Store Name', brand_col, sku_col, GAP_REASON_COLUMN, DATE_COLUMN]
        available_cols = [c for c in granular_cols if c in df_chronic.columns]
        
        dash_data = df_chronic[available_cols].copy()
        dash_data = dash_data.sort_values(by=['Store Account', 'Store Name', brand_col, sku_col, DATE_COLUMN], ascending=[True, True, True, True, False])
        dash_data = dash_data.drop_duplicates(subset=['Store Account', 'Store Name', brand_col, sku_col], keep='first')
        
        if DATE_COLUMN in dash_data.columns:
            dash_data['Last Date'] = dash_data[DATE_COLUMN].dt.strftime('%Y-%m-%d')
            dash_data = dash_data.drop(columns=[DATE_COLUMN])
            
        dash_data = dash_data.rename(columns={'Store Name': 'Outlet Name'})
        final_col_order = ['Store Account', 'Outlet Name', brand_col, sku_col, GAP_REASON_COLUMN, 'Last Date']
        dash_data = dash_data[[c for c in final_col_order if c in dash_data.columns]]
        
        table_data = dash_data.values.tolist()
        table_columns = [{'header': c} for c in dash_data.columns]
        end_row = start_row + len(dash_data)
        end_col = len(dash_data.columns) - 1
        
        ws_dash.add_table(start_row, 0, end_row, end_col, {
            'data': table_data,
            'columns': table_columns,
            'style': 'Table Style Medium 9', 
            'name': 'MasterTable'
        })
        
        # 7. Intelligent Auto-Fit Column Widths 
        col_widths = {i: 15 for i in range(10)}
        col_widths[4] = 3 
        
        for c_idx, col_name in enumerate(sku_headers):
            max_val = sku_breakdown.iloc[:, c_idx].astype(str).str.len().max() if c_idx < 2 else 5
            max_val = 0 if pd.isna(max_val) else int(max_val)
            col_widths[c_idx] = max(col_widths[c_idx], len(col_name), max_val)
            
        for c_idx, col_name in enumerate(outlet_headers):
            real_c = c_idx + 5
            max_val = outlet_breakdown.iloc[:, c_idx].astype(str).str.len().max() if c_idx < 2 else 5
            max_val = 0 if pd.isna(max_val) else int(max_val)
            col_widths[real_c] = max(col_widths[real_c], len(col_name), max_val)
            
        for c_idx, col_name in enumerate(dash_data.columns):
            if c_idx != 4: 
                max_val = dash_data[col_name].astype(str).str.len().max() if not dash_data.empty else 0
                max_val = 0 if pd.isna(max_val) else int(max_val)
                col_widths[c_idx] = max(col_widths[c_idx], len(col_name), max_val)
            
        for c_idx, width in col_widths.items():
            if c_idx == 4:
                continue 
            ws_dash.set_column(c_idx, c_idx, width + 4)

        # ----------------------------------------------------
        # DETAIL SHEET 
        # ----------------------------------------------------
        detail_sheet_name = 'Raw Detail'
        reason_df = df_chronic.drop(columns=['_clean_reason', '_group_tuple'], errors='ignore')
        
        if DATE_COLUMN in reason_df.columns:
            reason_df['Date'] = reason_df[DATE_COLUMN].dt.strftime('%Y-%m-%d')
            cols = ['Date'] + [c for c in reason_df.columns if c not in ['Date', DATE_COLUMN]]
            reason_df = reason_df[cols]

        reason_df.to_excel(writer, sheet_name=detail_sheet_name, index=False)
        ws_detail = writer.sheets[detail_sheet_name]
        
        std_head_fmt = workbook.add_format({'bold': True, 'bg_color': '#1F497D', 'font_color': 'white', 'border': 1})
        for col_num, col_name in enumerate(reason_df.columns):
            ws_detail.write(0, col_num, col_name, std_head_fmt)
            col_width = get_max_width(reason_df, col_name, col_num)
            ws_detail.set_column(col_num, col_num, col_width, cell_fmt)

    print(f"  ✅ Saved Dashboard: {file_name}")

# ==========================================
# GAPS DATA PREPARATION
# ==========================================
def generate_gap_reason_reports(df, output_dir):
    print("  -> Preparing Gap Reason data streams...")
    
    gap_col = GAP_REASON_COLUMN if GAP_REASON_COLUMN in df.columns else ("GapReason" if "GapReason" in df.columns else None)
    if not gap_col:
        print(f"  ⚠️ Could not find Gap Reason column. Skipping Gap Report.")
        return

    today = pd.Timestamp(datetime.now().date())
    three_weeks_ago = today - pd.Timedelta(days=21)
    two_days_ago = today - pd.Timedelta(days=2)

    if DATE_COLUMN in df.columns:
        df = df[df[DATE_COLUMN] >= three_weeks_ago].copy()

    if SUPERVISOR_COLUMN in df.columns:
        supervisors_to_exclude = ['jackie', 'glory']
        df = df[~df[SUPERVISOR_COLUMN].astype(str).str.lower().str.strip().isin(supervisors_to_exclude)]
        
    if BRAND_NAME_COLUMN in df.columns:
        df = df[~df[BRAND_NAME_COLUMN].astype(str).str.lower().str.endswith('ib')]

    df['_clean_reason'] = df[gap_col].astype(str).str.lower().str.strip()

    group_cols = ['Store Account', 'Store Name', 'Employee Name', 'Supervisor Name']
    sku_col = 'Product Name' if 'Product Name' in df.columns else 'Product Code'
    brand_col = BRAND_NAME_COLUMN if BRAND_NAME_COLUMN in df.columns else sku_col
    if sku_col in df.columns:
        group_cols.append(sku_col)

    df['_group_tuple'] = df[group_cols].apply(tuple, axis=1)

    latest_status = df.sort_values(DATE_COLUMN).drop_duplicates(subset=group_cols, keep='last')
    snl_latest_mask = latest_status['_clean_reason'].str.contains('sku not list', na=False)
    groups_now_snl = set(latest_status[snl_latest_mask]['_group_tuple'])

    # STREAM 1: ORDER DONE DELIVERY PENDING
    od_mask = df['_clean_reason'].str.contains('order don|delivery pend', na=False)
    df_od = df[od_mask].copy()
    df_od = df_od[~df_od['_group_tuple'].isin(groups_now_snl)]
    
    od_counts = df_od.groupby(group_cols).size().reset_index(name='od_count')
    chronic_od_keys = od_counts[od_counts['od_count'] >= 1][group_cols]
    df_od_chronic = df_od.merge(chronic_od_keys, on=group_cols, how='inner')
    df_od_chronic = df_od_chronic[df_od_chronic[DATE_COLUMN] >= two_days_ago]
    
    od_file_name = f"Order_Done_But_Delivery_Pending_Report_{datetime.now().strftime('%Y%m%d')}.xlsx"
    build_dashboard_workbook(df_od_chronic, "ORDER DONE BUT DELIVERY PENDING", od_file_name, output_dir, sku_col, brand_col, two_days_ago, today)

    # STREAM 2: SKU NOT LISTED
    snl_mask = df['_clean_reason'].str.contains('sku not list', na=False)
    df_snl = df[snl_mask].copy()
    
    snl_counts = df_snl.groupby(group_cols).size().reset_index(name='snl_count')
    chronic_snl_keys = snl_counts[snl_counts['snl_count'] > 1][group_cols]
    df_snl_chronic = df_snl.merge(chronic_snl_keys, on=group_cols, how='inner')

    snl_file_name = f"Gap_Reason_Report_SKU_Not_Listed_{datetime.now().strftime('%Y%m%d')}.xlsx"
    build_dashboard_workbook(df_snl_chronic, "SKU NOT LISTED", snl_file_name, output_dir, sku_col, brand_col, three_weeks_ago, today)

# ==========================================
# AUTOMATED GITHUB DEPLOYMENT HELPER
# ==========================================
def push_updates_to_github():
    print("\n---------------------------------------------------")
    print("-> Syncing Web_Output CSVs to GitHub...")
    try:
        os.system('git add Web_Output/*_Master.csv App.py requirements.txt')
        os.system('git commit -m "Auto-update web master CSVs"')
        os.system('git push origin main')
        print("✅ Live Streamlit Dashboard updated successfully!")
    except Exception as e:
        print(f"⚠️ Could not push to GitHub automatically: {e}")

# ==========================================
# MAIN MERGE PIPELINE
# ==========================================
def process_files():
    print("===================================================")
    print("STARTING DATA MERGE PIPELINE")
    print("===================================================")

    cutoff_date = datetime.now() - relativedelta(months=3)
    print(f"-> Power BI 3-Month Cutoff: {cutoff_date.strftime('%Y-%m-%d')}")
    print(f"-> Looking for files in: {INPUT_DIR.name}\n")

    all_files_in_input = list(INPUT_DIR.glob("*"))
    
    if not all_files_in_input:
        print(f"⚠️ The {INPUT_DIR.name} folder is COMPLETELY EMPTY. Please drop your CSVs in there!")
        return
    else:
        print(f"Files currently sitting in {INPUT_DIR.name}:")
        for f in all_files_in_input:
            print(f"  - {f.name}")
    print("-" * 50)

    for report_type in REPORT_TYPES:
        new_files = [
            f for f in all_files_in_input 
            if report_type.lower() in f.name.lower() and f.suffix.lower() == '.csv'
        ]
        
        if not new_files:
            continue
            
        print(f"\nProcessing {len(new_files)} new file(s) for: {report_type}")
        
        df_list = []
        for file in new_files:
            try:
                df = safe_read_csv(file)
                if not df.empty:
                    df_list.append(df)
            except Exception as e:
                print(f"   ⚠️ Error reading {file.name}: {e}")
                
        if not df_list:
            continue
            
        new_data = pd.concat(df_list, ignore_index=True)
        
        if DATE_COLUMN in new_data.columns:
            new_data[DATE_COLUMN] = parse_dates_safely(new_data[DATE_COLUMN])

        # ----------------------------------------------------
        # 1. POWER BI MASTER (FULL DATA, UNTOUCHED)
        # ----------------------------------------------------
        master_file_path = OUTPUT_DIR / f"{report_type}_Master.csv"
        if master_file_path.exists():
            existing_data = safe_read_csv(master_file_path)
            if not existing_data.empty:
                if DATE_COLUMN in existing_data.columns:
                    existing_data[DATE_COLUMN] = pd.to_datetime(existing_data[DATE_COLUMN], errors='coerce')
                master_data = pd.concat([existing_data, new_data], ignore_index=True)
            else:
                master_data = new_data
        else:
            master_data = new_data

        if DATE_COLUMN in master_data.columns:
            master_data = master_data[master_data[DATE_COLUMN] >= pd.Timestamp(cutoff_date)]
            
        master_data = master_data.drop_duplicates()
        master_data.to_csv(master_file_path, index=False)
        print(f"  ✅ Saved Power BI Master: {master_file_path.name}")

        # ----------------------------------------------------
        # 2. STREAMLIT / GITHUB MASTER (LIGHTWEIGHT: LAST 45 DAYS)
        # ----------------------------------------------------
        web_cutoff = datetime.now() - relativedelta(days=45)
        web_data = master_data[master_data[DATE_COLUMN] >= pd.Timestamp(web_cutoff)] if DATE_COLUMN in master_data.columns else master_data
        
        web_file_path = WEB_OUTPUT_DIR / f"{report_type}_Master.csv"
        web_data.to_csv(web_file_path, index=False)
        print(f"  📱 Saved Web Master for Streamlit: {web_file_path.name}")

        if report_type == "Availability Report":
            generate_gap_reason_reports(master_data, OUTPUT_DIR)

        if DATE_COLUMN in new_data.columns:
            valid_data = new_data.dropna(subset=[DATE_COLUMN]).copy()
            today_ts = pd.Timestamp(datetime.now().date())
            valid_data = valid_data[(valid_data[DATE_COLUMN] <= today_ts) & (valid_data[DATE_COLUMN].dt.year >= 2020)]
            
            valid_data['Archive_Month'] = valid_data[DATE_COLUMN].dt.strftime('%B_%Y')
            
            for month_label, group in valid_data.groupby('Archive_Month'):
                archive_file_path = ARCHIVE_DIR / f"{report_type}_{month_label}.csv"
                clean_group = group.drop(columns=['Archive_Month'])
                
                if archive_file_path.exists():
                    existing_archive = safe_read_csv(archive_file_path)
                    if not existing_archive.empty:
                        if DATE_COLUMN in existing_archive.columns:
                            existing_archive[DATE_COLUMN] = pd.to_datetime(existing_archive[DATE_COLUMN], errors='coerce')
                        combined_archive = pd.concat([existing_archive, clean_group], ignore_index=True)
                        combined_archive = combined_archive.drop_duplicates()
                    else:
                        combined_archive = clean_group
                else:
                    combined_archive = clean_group
                    
                combined_archive.to_csv(archive_file_path, index=False)
                print(f"  📂 Updated Monthly Archive: {archive_file_path.name}")

        for file in new_files:
            timestamp = datetime.now().strftime("%H%M%S")
            dump_path = DUMPS_DIR / f"{file.stem}_{timestamp}{file.suffix}"
            
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    shutil.move(str(file), str(dump_path))
                    break  
                except PermissionError:
                    if attempt < max_retries - 1:
                        print(f"  ⏳ File locked by system/Drive. Retrying move for {file.name} in 2 seconds...")
                        time.sleep(2)
                    else:
                        print(f"  ❌ Could not move {file.name}. Please ensure it is closed in Excel!")
            
    print("\n===================================================")
    print("PIPELINE COMPLETE")
    print("===================================================")
    
    # Automatically sync Web_Output to GitHub for live Streamlit dashboard
    push_updates_to_github()

if __name__ == "__main__":
    process_files()