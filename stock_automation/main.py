import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials
import re

# -------------------
# Configuration
# -------------------

# --- File and Column Names ---
ARENA_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1j006cY1xf_uQoa_tu0knxlrBa82lsoNk-pHqjCRVvQw/edit?gid=0#gid=0"
VIN_COLUMN = "Chassis No."
CHANNEL_COLUMN = "Channel"
DISPATCH_FILTER_COLUMN = "DELR"
WINGS_CHASSIS_COLUMN = "Chassis"
DMS_VIN_COLUMN = "VIN"

# --- Google API ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file"]
CREDS_FILE = 'credentials.json'

# -------------------
# Helper Functions
# -------------------

def get_file_names_for_today():
    """Generates the dynamic filenames for all required files based on the current date."""
    today_str = datetime.datetime.now().strftime("%d-%m-%Y")
    return {
        "main_stock": f"STOCK_AS_ON_{today_str}.xlsx",
        "dispatch_report": f"DISPATCHES {today_str}.xlsx",
        "dms_stock": f"DMS_STOCK_AS_ON_{today_str}.xlsx",
        "wings_stock": f"Wings stock as on {today_str}.xlsx",
        "sales_register": f"DMS SALES REGISTER AS ON {today_str}.xlsx",
        "output": f"UPDATED_STOCK_{today_str}.xlsx"
    }

def get_google_sheet_data(url):
    """Connects to a Google Sheet and returns the data from the first worksheet."""
    try:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(url)
        worksheet = sheet.get_worksheet(0)
        return pd.DataFrame(worksheet.get_all_records())
    except FileNotFoundError:
        print(f"ERROR: Credentials file ('{CREDS_FILE}') not found. Cannot access Google Sheets.")
        return None
    except Exception as e:
        print(f"ERROR: Could not read Google Sheet at {url}. Reason: {e}")
        return None

def load_all_data_sources(file_names):
    """Loads all necessary files and Google Sheets into pandas DataFrames."""
    print("--- Loading all data sources ---")
    data_sources = {}

    try:
        print(f"Loading Arena stock from Google Sheet...")
        data_sources['arena_gsheet'] = get_google_sheet_data(ARENA_GSHEET_URL)
        if data_sources['arena_gsheet'] is None: return None

        print(f"Loading VEHICLE sheet from '{file_names['main_stock']}'...")
        data_sources['vehicle_master'] = pd.read_excel(file_names['main_stock'], sheet_name='VEHICLE')

        print(f"Loading COLOR sheet from '{file_names['main_stock']}'...")
        data_sources['color_master'] = pd.read_excel(file_names['main_stock'], sheet_name='COLOR')

        print(f"Loading MASTER sheet from '{file_names['main_stock']}'...")
        data_sources['master_sheet'] = pd.read_excel(file_names['main_stock'], sheet_name='MASTER')

        print(f"Loading Dispatch report from '{file_names['dispatch_report']}'...")
        data_sources['dispatch_report'] = pd.read_excel(file_names['dispatch_report'])

        print(f"Loading DMS stock from '{file_names['dms_stock']}'...")
        data_sources['dms_stock'] = pd.read_excel(file_names['dms_stock'])

        print(f"Loading Wings stock from '{file_names['wings_stock']}'...")
        data_sources['wings_stock'] = pd.read_excel(file_names['wings_stock'])

        print(f"Loading Sales Register from '{file_names['sales_register']}'...")
        data_sources['sales_register'] = pd.read_excel(file_names['sales_register'])

        print("Successfully loaded all initial data sources.")
        return data_sources

    except FileNotFoundError as e:
        print(f"ERROR: A required file was not found: {e.filename}")
        return None
    except Exception as e:
        print(f"An error occurred during file loading: {e}")
        return None

# -------------------
# Processing Functions
# -------------------

def process_arena_data(arena_df):
    """Processes the data from the Arena Google Sheet."""
    print("--- Processing Arena Google Sheet data ---")
    if arena_df is None or arena_df.empty:
        print("Arena Google Sheet data is empty. Skipping.")
        return arena_df
    # Add 'Channel' column and an empty 'Color 2' column as a placeholder
    arena_df[CHANNEL_COLUMN] = 'ARENA'
    arena_df['Color 2'] = pd.NA
    return arena_df

def process_dispatch_data(dispatch_df, vehicle_master_df, color_master_df):
    """Processes the raw dispatch report data."""
    print("--- Processing Dispatch Report data ---")
    if dispatch_df is None or dispatch_df.empty: return None

    if DISPATCH_FILTER_COLUMN in dispatch_df.columns:
        dispatch_df[DISPATCH_FILTER_COLUMN] = dispatch_df[DISPATCH_FILTER_COLUMN].astype(str)
        dispatch_df = dispatch_df[~dispatch_df[DISPATCH_FILTER_COLUMN].str.contains('13NB|6ANA', na=False)]

    dispatch_df[VIN_COLUMN] = dispatch_df['CHASSISPREFIX'].astype(str) + dispatch_df['CHASSISNO'].astype(str)

    processed_df = pd.merge(dispatch_df, vehicle_master_df, left_on='MODELCODE', right_on='Vehicle code', how='left')
    processed_df = pd.merge(processed_df, color_master_df, left_on='COLOR', right_on='Color', how='left', suffixes=('', '_c'))

    column_mapping = {
        VIN_COLUMN: VIN_COLUMN, 'Model Desc': 'Model', 'Variant Desc': 'Varient',
        'Varriant S': 'Varriant s', 'ENGINENO': 'Engine', 'Description': 'Color',
        'INVOICEDATE': 'Purc. Dt.'
    }

    required_cols = list(column_mapping.keys())
    if not all(col in processed_df.columns for col in required_cols):
        print("ERROR: Not all required columns were found after merging dispatch data.")
        return None

    final_df = processed_df[required_cols].rename(columns=column_mapping)
    final_df[CHANNEL_COLUMN] = 'ARENA'
    final_df['Color 2'] = pd.NA # Add placeholder column
    return final_df

def process_dms_stock(main_df, dms_df, vehicle_master_df):
    """Processes DMS stock and finds new vehicles."""
    print("--- Processing DMS Stock data ---")
    if dms_df is None or dms_df.empty: return None

    if 'Dealer Location' in dms_df.columns:
        dms_df = dms_df[~dms_df['Dealer Location'].str.contains("(NEXA)", na=False)]

    existing_vins = main_df[VIN_COLUMN].unique()
    new_vehicles_df = dms_df[~dms_df[DMS_VIN_COLUMN].isin(existing_vins)].copy()

    if new_vehicles_df.empty: return None

    new_vehicles_df['MODELCODE'] = new_vehicles_df['Variant Code'].astype(str) + "00"
    merged_df = pd.merge(new_vehicles_df, vehicle_master_df, left_on='MODELCODE', right_on='Vehicle code', how='left')

    column_mapping = {
        'VIN': VIN_COLUMN, 'Model Desc_y': 'Model', 'Variant Desc_y': 'Varient',
        'Engine No': 'Engine', 'Colour': 'Color', 'MUL Inv Dt.': 'Purc. Dt.'
    }
    return merged_df.rename(columns=column_mapping)[list(column_mapping.values())]

def process_wings_stock(main_df, wings_df):
    """Processes the Wings stock file to update status and location."""
    print("--- Processing Wings Stock data & Reconciling ---")
    if wings_df is None or wings_df.empty: return main_df

    wings_df[WINGS_CHASSIS_COLUMN] = wings_df[WINGS_CHASSIS_COLUMN].str.replace('-', '', regex=False)

    main_df = pd.merge(main_df, wings_df[[WINGS_CHASSIS_COLUMN, 'location', 'Status']],
                       left_on=VIN_COLUMN, right_on=WINGS_CHASSIS_COLUMN, how='left')

    main_df['Instock/Transit'] = main_df['location'].fillna(main_df['Instock/Transit'])
    main_df = main_df[main_df['Status'] != 'SOLD']
    main_df.loc[main_df['Status'] == 'TRANSIT', 'Instock/Transit'] = 'Transit'

    return main_df.drop(columns=['location', 'Status', WINGS_CHASSIS_COLUMN])

def find_scattered_vin(row):
    """Helper function to find a VIN-like string in a row."""
    vin_pattern = re.compile(r'^(MA3|MBH|MAJ)[A-Z0-9]{14}$')
    for item in row:
        if isinstance(item, str) and vin_pattern.match(item):
            return item
    return None

def remove_invoiced_vehicles(main_df, sales_register_df):
    """Filters sales register and removes invoiced vehicles from main stock."""
    print("--- Removing invoiced vehicles using Sales Register ---")
    if sales_register_df is None or sales_register_df.empty: return main_df

    customer_col = 'Customer Name'
    if customer_col in sales_register_df.columns:
        sales_register_df[customer_col] = sales_register_df[customer_col].astype(str)

        is_fake_sale = sales_register_df[customer_col].str.split().str[0].str.upper().isin(['F', 'G']) & \
                       (sales_register_df[customer_col].str.split().str.len() > 1)
        sales_register_df = sales_register_df[~is_fake_sale]

    sold_vins = sales_register_df.apply(find_scattered_vin, axis=1).dropna().unique()

    if len(sold_vins) > 0:
        main_df = main_df[~main_df[VIN_COLUMN].isin(sold_vins)]
        print(f"Removed {len(sold_vins)} invoiced vehicles.")

    return main_df

def perform_final_enrichment(main_df, master_df):
    """Performs final calculations and cleaning on the DataFrame."""
    print("--- Performing final data enrichment ---")
    if main_df is None or main_df.empty: return main_df

    # --- Data Type Conversion ---
    # Convert Purc. Dt. to datetime for sorting and calculations. Coerce errors will turn failed parses into NaT.
    main_df['Purc. Dt.'] = pd.to_datetime(main_df['Purc. Dt.'], errors='coerce')

    # --- Sorting ---
    print("Sorting data by Model, Variant, Color, and Purchase Date...")
    sort_columns = ['Model', 'Varriant s', 'Color', 'Purc. Dt.']
    # Check if all sort columns exist to prevent errors
    existing_sort_columns = [col for col in sort_columns if col in main_df.columns]
    if len(existing_sort_columns) < len(sort_columns):
        print(f"WARNING: One or more sort columns not found. Sorting by available columns: {existing_sort_columns}")

    if existing_sort_columns:
        main_df.sort_values(by=existing_sort_columns, ascending=True, inplace=True)
        print("Sorting complete.")

    # --- VLOOKUP for Channel ---
    if master_df is not None and 'Model' in main_df.columns and 'MODEL' in master_df.columns and 'CHANEL' in master_df.columns:
        print("Updating 'Channel' column using MASTER sheet lookup...")
        # Create temporary, cleaned keys for a robust, case/whitespace-insensitive lookup
        main_df['temp_model_key'] = main_df['Model'].astype(str).str.strip().str.lower()
        master_df['temp_model_key'] = master_df['MODEL'].astype(str).str.strip().str.lower()

        # Create mapping dictionary from the cleaned master sheet
        channel_map = master_df.drop_duplicates(subset=['temp_model_key']).set_index('temp_model_key')['CHANEL'].to_dict()

        # Update the 'Channel' column, preserving existing values if lookup fails
        main_df[CHANNEL_COLUMN] = main_df['temp_model_key'].map(channel_map).fillna(main_df[CHANNEL_COLUMN])

        # Remove temporary columns
        main_df.drop(columns=['temp_model_key'], inplace=True)
        # Clean up master_df as well to prevent issues in other functions if it's reused
        if 'temp_model_key' in master_df.columns:
            master_df.drop(columns=['temp_model_key'], inplace=True)
        print("Updated 'Channel' column.")

    # --- VLOOKUP for Color 2 ---
    if master_df is not None and 'Color' in main_df.columns and 'COLOR' in master_df.columns and 'COLOR 2' in master_df.columns:
        print("Updating 'Color 2' column using robust MASTER sheet lookup.")
        # Create temporary, cleaned keys for a robust, case/whitespace-insensitive lookup
        main_df['temp_color_key'] = main_df['Color'].astype(str).str.strip().str.lower()
        master_df['temp_color_key'] = master_df['COLOR'].astype(str).str.strip().str.lower()

        # Create mapping dictionary from the cleaned master sheet
        color_map = master_df.drop_duplicates(subset=['temp_color_key']).set_index('temp_color_key')['COLOR 2'].to_dict()

        # Update the 'Color 2' column, preserving existing values if lookup fails
        main_df['Color 2'] = main_df['temp_color_key'].map(color_map).fillna(main_df['Color 2'])

        # Remove temporary columns
        main_df.drop(columns=['temp_color_key'], inplace=True)
        if 'temp_color_key' in master_df.columns:
            master_df.drop(columns=['temp_color_key'], inplace=True)
        print("Updated 'Color 2' column.")

    # --- Final Calculations & Cleaning ---
    main_df['Allocation Status'].fillna('AVAILABLE', inplace=True)
    main_df['AGEING'] = (datetime.datetime.now() - main_df['Purc. Dt.'].dt.tz_localize(None)).dt.days
    main_df['MAKE YEAR'] = main_df['Purc. Dt.'].dt.year.astype('Int64') # Use nullable integer
    main_df.drop_duplicates(subset=[VIN_COLUMN], keep='first', inplace=True)

    # Format date back to string for the final report
    main_df['Purc. Dt.'] = main_df['Purc. Dt.'].dt.strftime('%d-%m-%Y')

    # Reset index after sorting to ensure 'Sr. No.' is sequential
    main_df.reset_index(drop=True, inplace=True)
    main_df['Sr. No.'] = range(1, len(main_df) + 1)

    return main_df

def save_output_file(final_df, filename):
    """Saves the final DataFrame to an Excel file."""
    print(f"\n--- Saving final output to '{filename}' ---")
    if final_df is None or final_df.empty:
        print("Final DataFrame is empty. Nothing to save.")
        return
    try:
        final_df.to_excel(filename, index=False)
        print("Successfully saved the output file.")
    except Exception as e:
        print(f"ERROR: Could not save the output file. Reason: {e}")

# -------------------
# Main Execution
# -------------------

def main():
    """Main function to orchestrate the entire stock automation process."""
    print("Starting the stock automation script...")

    file_names = get_file_names_for_today()
    data_sources = load_all_data_sources(file_names)

    if data_sources is None:
        print("\nAborting script due to errors in loading data sources.")
        return

    processed_arena_df = process_arena_data(data_sources['arena_gsheet'])
    processed_dispatch_df = process_dispatch_data(
        data_sources['dispatch_report'],
        data_sources['vehicle_master'],
        data_sources['color_master']
    )

    try:
        main_stock_df = pd.read_excel(file_names['main_stock'], sheet_name='ARENA')
        main_stock_df = main_stock_df[main_stock_df[CHANNEL_COLUMN] != 'Arena']
    except FileNotFoundError:
        print(f"WARNING: Main stock file '{file_names['main_stock']}' not found. Starting fresh.")
        main_stock_df = pd.DataFrame()

    frames_to_combine = [main_stock_df, processed_arena_df, processed_dispatch_df]
    main_stock_df = pd.concat(frames_to_combine, ignore_index=True)

    new_dms_df = process_dms_stock(main_stock_df, data_sources['dms_stock'], data_sources['vehicle_master'])
    if new_dms_df is not None:
        main_stock_df = pd.concat([main_stock_df, new_dms_df], ignore_index=True)

    main_stock_df = process_wings_stock(main_stock_df, data_sources['wings_stock'])
    main_stock_df = remove_invoiced_vehicles(main_stock_df, data_sources['sales_register'])
    final_df = perform_final_enrichment(main_stock_df, data_sources['master_sheet'])

    save_output_file(final_df, file_names['output'])
    print("\nScript finished successfully!")

if __name__ == "__main__":
    main()
