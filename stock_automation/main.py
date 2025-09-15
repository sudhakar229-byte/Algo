import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials
import re

# -------------------
# Configuration
# -------------------

# --- File and Column Names ---
# These can be adjusted if the names change in the future.
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
    # Note: The user mentioned "DMS_STOCK..." and "Wings stock...". Using these exact names.
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
    print("--- Step 1: Loading all data sources ---")
    data_sources = {}

    # 1. Arena Google Sheet
    print(f"Loading Arena stock from Google Sheet...")
    data_sources['arena_gsheet'] = get_google_sheet_data(ARENA_GSHEET_URL)
    if data_sources['arena_gsheet'] is None: return None # Stop if we can't get this vital data

    try:
        # 2. Main Stock File (for VEHICLE and COLOR lookup sheets)
        print(f"Loading VEHICLE sheet from '{file_names['main_stock']}'...")
        data_sources['vehicle_master'] = pd.read_excel(file_names['main_stock'], sheet_name='VEHICLE')
        print(f"Loading COLOR sheet from '{file_names['main_stock']}'...")
        data_sources['color_master'] = pd.read_excel(file_names['main_stock'], sheet_name='COLOR')

        # 3. Dispatch Report Excel File
        print(f"Loading Dispatch report from '{file_names['dispatch_report']}'...")
        data_sources['dispatch_report'] = pd.read_excel(file_names['dispatch_report'])

        # 4. DMS Stock File
        print(f"Loading DMS stock from '{file_names['dms_stock']}'...")
        data_sources['dms_stock'] = pd.read_excel(file_names['dms_stock'])

        # 5. Wings Stock File
        print(f"Loading Wings stock from '{file_names['wings_stock']}'...")
        data_sources['wings_stock'] = pd.read_excel(file_names['wings_stock'])

        # 6. Sales Register File
        print(f"Loading Sales Register from '{file_names['sales_register']}'...")
        data_sources['sales_register'] = pd.read_excel(file_names['sales_register'])

        # We will load the main 'ARENA' sheet in the main function.

        print("Successfully loaded all initial data sources.")
        return data_sources

    except FileNotFoundError as e:
        print(f"ERROR: A required file was not found: {e.filename}")
        print("Please make sure all required files are in the same directory as the script.")
        return None
    except Exception as e:
        print(f"An error occurred during file loading: {e}")
        return None

# -------------------
# Processing Functions (To be implemented in subsequent steps)
# -------------------

def process_arena_data(arena_df):
    """
    Processes the data from the Arena Google Sheet.
    - Adds a 'Channel' column.
    - Adds and forward-fills the 'Color 2' column.
    """
    print("--- Step 2: Processing Arena Google Sheet data ---")
    if arena_df is None or arena_df.empty:
        print("Arena Google Sheet data is empty. Skipping processing.")
        return arena_df

    # Add a "Channel" column and fill it with "ARENA"
    arena_df[CHANNEL_COLUMN] = 'ARENA'
    print(f"Added '{CHANNEL_COLUMN}' column.")

    # Insert 'Color 2' column and fill it by copying from the cell above.
    if 'Color' in arena_df.columns:
        color_col_index = arena_df.columns.get_loc('Color')
        # Using .shift(1) is a literal interpretation of "copying from the above cell"
        arena_df.insert(color_col_index + 1, 'Color 2', arena_df['Color'].shift(1))
        print("Added 'Color 2' column by shifting 'Color' column down.")
    else:
        print("WARNING: 'Color' column not found, could not add 'Color 2'.")

    print("Finished processing Arena data.")
    return arena_df

def process_dispatch_data(dispatch_df, vehicle_master_df, color_master_df):
    """
    Processes the raw dispatch report data into a clean format.
    - Filters out specific rows.
    - Combines chassis columns.
    - Looks up data from the master sheets.
    """
    print("--- Step 3: Processing Dispatch Report data ---")
    if dispatch_df is None or dispatch_df.empty:
        print("Dispatch report data is empty. Skipping.")
        return None

    # 1. Filter out rows containing "13NB" and "6ANA"
    filter_values = ["13NB", "6ANA"]
    if DISPATCH_FILTER_COLUMN in dispatch_df.columns:
        dispatch_df[DISPATCH_FILTER_COLUMN] = dispatch_df[DISPATCH_FILTER_COLUMN].astype(str)
        original_rows = len(dispatch_df)
        dispatch_df = dispatch_df[~dispatch_df[DISPATCH_FILTER_COLUMN].str.contains('|'.join(filter_values), na=False)]
        print(f"Filtered {original_rows - len(dispatch_df)} rows based on '{DISPATCH_FILTER_COLUMN}' content.")
    else:
        print(f"WARNING: Filter column '{DISPATCH_FILTER_COLUMN}' not found in dispatch report.")

    # 2. Combine chassis columns
    dispatch_df[VIN_COLUMN] = dispatch_df['CHASSISPREFIX'].astype(str) + dispatch_df['CHASSISNO'].astype(str)
    print("Combined 'CHASSISPREFIX' and 'CHASSISNO' into 'Chassis No.'.")

    # 3. Replicate VLOOKUPs by merging with master data
    print("Merging with VEHICLE and COLOR master data to look up details...")
    # First merge for vehicle details, adding a suffix to avoid column name conflicts
    processed_df = pd.merge(dispatch_df, vehicle_master_df, left_on='MODELCODE', right_on='Vehicle code', how='left', suffixes=('', '_vmaster'))
    # Second merge for color description
    processed_df = pd.merge(processed_df, color_master_df, left_on='COLOR', right_on='Color', how='left', suffixes=('', '_cmaster'))

    # 4. Select and rename columns to match the final format
    # Now we use the looked-up columns from the master sheets
    column_mapping = {
        VIN_COLUMN: VIN_COLUMN,
        'Model Desc': 'Model',
        'Variant Desc': 'Varient',
        'Varriant S': 'Varriant s',
        'ENGINENO': 'Engine',
        'Description': 'Color', # This is the looked-up color description from color_master
        'INVOICEDATE': 'Purc. Dt.'
    }

    # Check if all needed source columns exist before selecting
    required_cols = list(column_mapping.keys())
    missing_cols = [col for col in required_cols if col not in processed_df.columns]
    if missing_cols:
        print(f"ERROR: The following required columns were not found after merging dispatch data: {missing_cols}. Skipping.")
        return None

    final_dispatch_df = processed_df[required_cols].rename(columns=column_mapping)
    print("Selected and renamed columns from dispatch data.")

    # 5. Add Channel and Color 2 columns
    final_dispatch_df[CHANNEL_COLUMN] = 'ARENA'
    if 'Color' in final_dispatch_df.columns:
        color_col_index = final_dispatch_df.columns.get_loc('Color')
        final_dispatch_df.insert(color_col_index + 1, 'Color 2', final_dispatch_df['Color'].shift(1))
        print("Added 'Color 2' column by shifting 'Color' column down.")

    print("Finished processing dispatch data.")
    return final_dispatch_df

def process_dms_stock(main_df, dms_df, vehicle_master_df):
    """
    Processes DMS stock, finds new vehicles, and looks up model/variant info.
    """
    print("--- Step 4: Processing DMS Stock data ---")
    if dms_df is None or dms_df.empty:
        print("DMS stock data is empty. Skipping.")
        return None

    # 1. Filter out NEXA locations
    if 'Dealer Location' in dms_df.columns:
        dms_df = dms_df[~dms_df['Dealer Location'].str.contains("(NEXA)", na=False)]
        print(f"Filtered out NEXA locations from DMS stock.")

    # 2. Find new vehicles by comparing VINs
    existing_vins = main_df[VIN_COLUMN].unique()
    new_vehicles_df = dms_df[~dms_df[DMS_VIN_COLUMN].isin(existing_vins)].copy()

    if new_vehicles_df.empty:
        print("No new vehicles found in DMS stock to add.")
        return None
    print(f"Found {len(new_vehicles_df)} new vehicles in DMS stock.")

    # 3. Create lookup key and merge with vehicle master
    new_vehicles_df['MODELCODE'] = new_vehicles_df['Variant Code'].astype(str) + "00"
    print("Created lookup key from 'Variant Code'.")

    merged_df = pd.merge(new_vehicles_df, vehicle_master_df, left_on='MODELCODE', right_on='Vehicle code', how='left', suffixes=('', '_vmaster'))
    print("Merged DMS stock with VEHICLE master data.")

    # 4. Map columns, now using the looked-up values
    column_mapping = {
        'VIN': VIN_COLUMN,
        'Model Desc_vmaster': 'Model', # Use looked-up value
        'Variant Desc_vmaster': 'Varient', # Use looked-up value
        'Engine No': 'Engine',
        'Colour': 'Color',
        'MUL Inv Dt.': 'Purc. Dt.'
    }

    # Select and rename columns
    new_vehicles_processed = merged_df[list(column_mapping.keys())].rename(columns=column_mapping)

    print("Finished processing DMS stock data.")
    return new_vehicles_processed

def process_wings_stock(main_df, wings_df):
    """
    Processes the Wings stock file to update the main stock DataFrame.
    - Updates 'Instock/Transit' column based on 'location'.
    - Removes 'SOLD' vehicles.
    - Updates 'Instock/Transit' to 'Transit' for 'TRANSIT' vehicles.
    """
    print("--- Step 5: Processing Wings Stock data & Reconciling ---")
    if wings_df is None or wings_df.empty:
        print("Wings stock data is empty. Skipping.")
        return main_df

    # 1. Clean Chassis column in Wings data
    wings_df[WINGS_CHASSIS_COLUMN] = wings_df[WINGS_CHASSIS_COLUMN].str.replace('-', '', regex=False)
    print("Cleaned 'Chassis' column in Wings data.")

    # 2. Merge to get 'location' and 'Status' into the main DataFrame
    # We use a left merge to keep all records from the main stock and add info from Wings where chassis numbers match.
    main_df_updated = pd.merge(
        main_df,
        wings_df[[WINGS_CHASSIS_COLUMN, 'location', 'Status']],
        left_on=VIN_COLUMN,
        right_on=WINGS_CHASSIS_COLUMN,
        how='left',
        suffixes=('', '_wings')
    )

    # 3. Update 'Instock/Transit' with 'location' from Wings
    # Where 'location' is not null, update 'Instock/Transit'
    main_df_updated['Instock/Transit'] = main_df_updated['location'].fillna(main_df_updated['Instock/Transit'])
    print("Updated 'Instock/Transit' column with location data from Wings file.")

    # 4. Remove SOLD vehicles
    original_rows = len(main_df_updated)
    main_df_updated = main_df_updated[main_df_updated['Status'] != 'SOLD']
    print(f"Removed {original_rows - len(main_df_updated)} SOLD vehicles.")

    # 5. Update 'Instock/Transit' for TRANSIT vehicles
    transit_mask = main_df_updated['Status'] == 'TRANSIT'
    main_df_updated.loc[transit_mask, 'Instock/Transit'] = 'Transit'
    print(f"Updated {transit_mask.sum()} rows to 'Transit' status.")

    # Drop the temporary columns from the merge
    main_df_updated = main_df_updated.drop(columns=['location', 'Status', WINGS_CHASSIS_COLUMN])

    print("Finished processing Wings stock data.")
    return main_df_updated

def find_scattered_vin(row):
    """Helper function to find a VIN-like string in a row."""
    vin_pattern = re.compile(r'^(MA3|MBH|MAJ)[A-Z0-9]{14}$')
    for item in row:
        if isinstance(item, str) and vin_pattern.match(item):
            return item
    return None

def remove_invoiced_vehicles(main_df, sales_register_df):
    """Filters sales register for real sales, then removes those VINs from main stock."""
    print("--- (New) Removing invoiced vehicles using Sales Register ---")
    if sales_register_df is None or sales_register_df.empty:
        print("Sales Register data is empty. Skipping.")
        return main_df

    # 1. Filter out "fake" sales
    customer_col = 'Customer Name' # Assuming this column name
    if customer_col in sales_register_df.columns:
        original_rows = len(sales_register_df)
        sales_register_df[customer_col] = sales_register_df[customer_col].astype(str)

        # Define a function to check for the "fake sale" pattern
        def is_fake_sale(name):
            parts = name.split()
            if len(parts) > 1 and len(parts[0]) == 1 and parts[0].upper() in ['F', 'G']:
                return True
            return False

        # Keep rows where the customer name does NOT match the fake sale pattern
        sales_register_df = sales_register_df[~sales_register_df[customer_col].apply(is_fake_sale)]
        print(f"Filtered {original_rows - len(sales_register_df)} fake sales from Sales Register.")
    else:
        print(f"WARNING: Customer column '{customer_col}' not found in Sales Register. Cannot filter fake sales.")

    # 2. Find VINs in the cleaned sales data
    sold_vins = sales_register_df.apply(find_scattered_vin, axis=1).dropna().unique()

    # 3. Remove sold VINs from main stock
    if len(sold_vins) > 0:
        print(f"Found {len(sold_vins)} unique invoiced VINs to remove.")
        original_rows = len(main_df)
        main_df = main_df[~main_df[VIN_COLUMN].isin(sold_vins)]
        print(f"Removed {original_rows - len(main_df)} invoiced vehicles from the main stock list.")
    else:
        print("No matching invoiced VINs found in the sales register to remove.")

    return main_df

def perform_final_enrichment(main_df):
    """
    Performs final calculations on the DataFrame.
    - Fills 'Allocation Status'.
    - Calculates 'AGEING'.
    - Calculates 'MAKE YEAR'.
    """
    print("--- Step 6: Performing final data enrichment ---")
    if main_df is None or main_df.empty:
        print("Main DataFrame is empty. Skipping enrichment.")
        return main_df

    # Fill 'Allocation Status' for new rows
    main_df['Allocation Status'].fillna('AVAILABLE', inplace=True)
    print("Filled 'Allocation Status' for new vehicles.")

    # Convert 'Purc. Dt.' to datetime
    main_df['Purc. Dt.'] = pd.to_datetime(main_df['Purc. Dt.'], errors='coerce')

    # Calculate 'AGEING'
    today = datetime.datetime.now()
    main_df['AGEING'] = (today - main_df['Purc. Dt.']).dt.days
    print("Calculated 'AGEING' column.")

    # Calculate 'MAKE YEAR'
    main_df['MAKE YEAR'] = main_df['Purc. Dt.'].dt.year
    print("Calculated 'MAKE YEAR' column.")

    # De-duplicate the final DataFrame
    original_rows = len(main_df)
    main_df.drop_duplicates(subset=[VIN_COLUMN], keep='first', inplace=True)
    print(f"Removed {original_rows - len(main_df)} duplicate chassis numbers, keeping the first entry.")

    # Format 'Purc. Dt.' to the desired string format
    main_df['Purc. Dt.'] = main_df['Purc. Dt.'].dt.strftime('%d-%m-%Y')
    print("Formatted 'Purc. Dt.' column to DD-MM-YYYY.")

    # Generate final serial numbers
    main_df['Sr. No.'] = range(1, len(main_df) + 1)
    print("Generated final 'Sr. No.' column.")

    print("Finished data enrichment.")
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

    # Get today's filenames
    file_names = get_file_names_for_today()
    print("\nGenerated filenames for today:")
    for key, name in file_names.items():
        print(f"  - {key}: {name}")

    # Step 1: Load all the initial source files
    data_sources = load_all_data_sources(file_names)

    if data_sources is None:
        print("\nAborting script due to errors in loading data sources.")
        return

    # Step 2: Process the Arena Google Sheet Data
    processed_arena_df = process_arena_data(data_sources['arena_gsheet'])

    # Step 3: Process the Dispatch Report Data
    processed_dispatch_df = process_dispatch_data(
        data_sources['dispatch_report'],
        data_sources['vehicle_master'],
        data_sources['color_master']
    )

    # --- Construct the initial main DataFrame ---
    print("\n--- Constructing initial main stock DataFrame ---")
    try:
        # Read the original ARENA sheet
        main_stock_df = pd.read_excel(file_names['main_stock'], sheet_name='ARENA')
        print(f"Loaded initial 'ARENA' sheet with {len(main_stock_df)} rows.")

        # Filter out old Arena data
        main_stock_df = main_stock_df[main_stock_df[CHANNEL_COLUMN] != 'Arena']
        print(f"Removed old Arena data, {len(main_stock_df)} rows remaining.")

        # Combine the base data with the new data from Google Sheets and Dispatch
        frames_to_combine = [main_stock_df]
        if processed_arena_df is not None:
            frames_to_combine.append(processed_arena_df)
        if processed_dispatch_df is not None:
            frames_to_combine.append(processed_dispatch_df)

        main_stock_df = pd.concat(frames_to_combine, ignore_index=True)
        print(f"Combined with new data. Total rows are now: {len(main_stock_df)}")

    except FileNotFoundError:
        print(f"WARNING: Main stock file '{file_names['main_stock']}' not found. Starting with an empty stock list.")
        # If the main file doesn't exist, we start fresh with the new data
        main_stock_df = pd.concat([processed_arena_df, processed_dispatch_df], ignore_index=True)
    except Exception as e:
        print(f"ERROR: Could not construct initial main stock DataFrame. Reason: {e}")
        return # Abort if we can't build the base frame

    # Step 4: Process DMS Stock Data
    new_dms_vehicles_df = process_dms_stock(main_stock_df, data_sources['dms_stock'], data_sources['vehicle_master'])
    if new_dms_vehicles_df is not None:
        main_stock_df = pd.concat([main_stock_df, new_dms_vehicles_df], ignore_index=True)
        print(f"Added new DMS vehicles. Total rows are now: {len(main_stock_df)}")

    # Step 5: Process Wings Stock Data & Reconcile
    main_stock_df = process_wings_stock(main_stock_df, data_sources['wings_stock'])

    # Step 5b: Remove invoiced vehicles from Sales Register
    main_stock_df = remove_invoiced_vehicles(main_stock_df, data_sources['sales_register'])

    # Step 6: Final Data Enrichment
    final_df = perform_final_enrichment(main_stock_df)

    # Final Step: Save the output file
    save_output_file(final_df, file_names['output'])

    print("\nScript finished successfully!")

if __name__ == "__main__":
    main()
