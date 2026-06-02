import pandas as pd
import os

from openpyxl import load_workbook
from datetime import datetime
from openpyxl.utils import get_column_letter

VERSION = "1.0"

# =========================
# FOLDER SETUP
# =========================

REPORTS_FOLDER = "reports"
DATA_FOLDER = "data"

# CSV reports from PowerShell app
SCANNED_FILE = os.path.join(REPORTS_FOLDER, "scanned_assets.csv")
MANUAL_FILE = os.path.join(REPORTS_FOLDER, "manual_assets.csv")

# Main Excel inventory file
INVENTORY_FILE = os.path.join(DATA_FOLDER, "inventory.xlsx")
REPORT_WORKBOOK = os.path.join(REPORTS_FOLDER, "asset_reports.xlsx")

# Create folders automatically
os.makedirs(REPORTS_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

# =========================
# VALID STATUSES
# =========================

VALID_STATUSES = {
    "Active",
    "Issued Out",
    "Returned",
    "Broken",
    "Retired"
}

# =========================
# SAVE / LOAD FUNCTIONS
# =========================

def save_inventory(df):
    """
    Save inventory while preserving:
    - filters
    - formatting
    - tables
    - column widths
    - colors
    """

    # Create workbook if missing
    if not os.path.exists(INVENTORY_FILE):

        with pd.ExcelWriter(
            INVENTORY_FILE,
            engine="openpyxl"
        ) as writer:

            df.to_excel(
                writer,
                sheet_name="Assets",
                index=False
            )

        return

    # Load existing workbook
    workbook = load_workbook(INVENTORY_FILE)

    # Ensure Assets sheet exists
    if "Assets" not in workbook.sheetnames:

        sheet = workbook.create_sheet("Assets")

        # Add headers
        sheet.append(list(df.columns))

    else:
        sheet = workbook["Assets"]

    # Delete old data rows only
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row)

    # Write updated rows
    for row in df.itertuples(index=False):

        sheet.append(list(row))

    # Save workbook
    workbook.save(INVENTORY_FILE)


def repair_asset(asset_id):

    df = load_inventory()

    if df is None:
        return
    
    if asset_id not in df["Asset ID"].values:

        print("Asset not found")
        return
    
    current_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]
    if current_status != "Broken":

        print(f"Asset {asset_id} is not marked as Broken, only Broken assets can be repaired.")
        pause()
        return
    
    repair_notes = input("Enter repair notes: ")

    old_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]

    current_user = ""

    if "User" in df.columns:

        current_user = df.loc[df["Asset ID"] == asset_id, "User"].iloc[0]

        df.loc[df["Asset ID"] == asset_id, "Status"] = "Active"

        save_inventory(df)

        log_history(
            asset_id,
            old_status,
            "Active",
            current_user)
        
        add_asset_log(
            asset_id,
            "Repaired",
            repair_notes)
        
        print(f"{asset_id} repaired and set to Active")
        
        pause()


def add_asset_log(asset_id, log_type, notes):
    if not os.path.exists(INVENTORY_FILE):
        return
    
    workbook = load_workbook(INVENTORY_FILE)

    if "Asset Logs" not in workbook.sheetnames:
        log_sheet = workbook.create_sheet("Asset Logs")

        log_sheet.append(["Timestamp", "Asset ID", "Asset Name", "Serial Number", "Log Type", "Notes"])
    else:
        log_sheet = workbook["Asset Logs"]

    df = load_inventory()
    asset_name = ""
    serial_number = ""

    if df is not None:
        asset_row = df[df["Asset ID"] == asset_id]

        if not asset_row.empty:
            if "AssetName" in asset_row.columns:
                asset_name = asset_row["AssetName"].iloc[0]
            if "SerialNumber" in asset_row.columns:
                serial_number = asset_row["SerialNumber"].iloc[0]

    log_sheet.append([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        asset_id,
        asset_name,
        serial_number,
        log_type,
        notes
    ])
    workbook.save(INVENTORY_FILE)

    print(f"Log added for {asset_id}: {log_type} - {notes}")

def view_asset_logs(asset_id):
    if not os.path.exists(INVENTORY_FILE):
        print("No logs found.")
        return
    workbook = load_workbook(INVENTORY_FILE)
    if "Asset Logs" not in workbook.sheetnames:
        print("No logs found.")
        return
    
    log_sheet = workbook["Asset Logs"]
    print(f"\nLogs for {asset_id}:\n")
    found = False

    for row in log_sheet.iter_rows(min_row=2, values_only=True):
        timestamp, log_asset_id, log_asset_name, log_serial_number, log_type, notes = row

        if str(log_asset_id) == str(asset_id):
            print(f"{timestamp} - {log_type} - {notes}")
            found = True

    if not found:
        print("No logs found.")

    pause()


def log_history(asset_id, old_status, new_status, user = ""):
    if not os.path.exists(INVENTORY_FILE):
        return
    
    workbook = load_workbook(INVENTORY_FILE)

    if "History" not in workbook.sheetnames:
        history_sheet = workbook.create_sheet("History")

        history_sheet.append(["Timestamp", "Asset ID", "Old Status", "New Status", "User"])

    else:
        history_sheet = workbook["History"]

    history_sheet.append([
        datetime.now().strftime("%Y-%m-%d%H:%M:%S"),
        asset_id,
        old_status,
        new_status,
        user,
    ])
    workbook.save(INVENTORY_FILE)


def load_inventory():
    """Load inventory dataframe from Excel."""

    if not os.path.exists(INVENTORY_FILE):
        print("Inventory file not found.")
        return None

    return pd.read_excel(INVENTORY_FILE)


# =========================
# CREATE / UPDATE INVENTORY
# =========================

def find_asset_match(inventory_df, asset):

    # Match by Serial Number
    if (
        "SerialNumber" in inventory_df.columns
        and pd.notna(asset.get("SerialNumber"))
        and str(asset.get("SerialNumber")).strip() != ""
    ):

        matches = inventory_df[
            inventory_df["SerialNumber"] == asset["SerialNumber"]
        ]

        if not matches.empty:
            return matches.index[0]

    # Match by MAC Address
    if (
        "MACAddress" in inventory_df.columns
        and pd.notna(asset.get("MACAddress"))
        and str(asset.get("MACAddress")).strip() != ""
    ):

        matches = inventory_df[
            inventory_df["MACAddress"] == asset["MACAddress"]
        ]

        if not matches.empty:
            return matches.index[0]

    # Match by Asset Name + Model
    if (
        "AssetName" in inventory_df.columns
        and "Model" in inventory_df.columns
    ):

        matches = inventory_df[
            (inventory_df["AssetName"] == asset.get("AssetName"))
            &
            (inventory_df["Model"] == asset.get("Model"))
        ]

        if not matches.empty:
            return matches.index[0]

    return None



def update_inventory():

    # Load CSV source files
    scanned_df = pd.read_csv(SCANNED_FILE)
    manual_df = pd.read_csv(MANUAL_FILE)

    # Combine import files
    import_df = pd.concat(
        [scanned_df, manual_df],
        ignore_index=True
    )

    # Load existing inventory
    inventory_df = load_inventory()

    if inventory_df is None:

        inventory_df = pd.DataFrame()

    # Process imported assets
    for _, asset in import_df.iterrows():

        match_index = None

        if not inventory_df.empty:

            match_index = find_asset_match(
                inventory_df,
                asset
            )

        # Existing asset found
        if match_index is not None:

            fields_to_update = [

                "AssetName",
                "User",
                "Location",
                "Date",
                "AssetType",
                "Model",
                "SerialNumber",
                "MACAddress",
                "Comments",
                "IMEI"

            ]

            for field in fields_to_update:

                if field in inventory_df.columns:

                    inventory_df.at[
                        match_index,
                        field
                    ] = asset.get(field)

        # New asset
        else:

            new_asset = asset.to_dict()

            if "Status" not in new_asset:
                new_asset["Status"] = "Active"

            if "Asset ID" not in new_asset:
                new_asset["Asset ID"] = ""

            inventory_df = pd.concat(
                [
                    inventory_df,
                    pd.DataFrame([new_asset])
                ],
                ignore_index=True
            )

    # Ensure Status column exists
    if "Status" not in inventory_df.columns:
        inventory_df["Status"] = "Active"

    # Ensure Asset ID column exists
    if "Asset ID" not in inventory_df.columns:
        inventory_df["Asset ID"] = ""

    # Find highest Asset ID
    highest_id = 0

    for asset_id in inventory_df["Asset ID"]:

        if pd.notna(asset_id):

            asset_id = str(asset_id).strip()

            if asset_id.startswith("A"):

                try:

                    number = int(asset_id[1:])

                    if number > highest_id:
                        highest_id = number

                except ValueError:
                    pass

    # Generate IDs only for blank assets
    for index in inventory_df.index:

        current_id = str(
            inventory_df.at[index, "Asset ID"]
        ).strip()

        if current_id == "" or current_id.lower() == "nan":

            highest_id += 1

            inventory_df.at[
                index,
                "Asset ID"
            ] = f"A{str(highest_id).zfill(3)}"

    save_inventory(inventory_df)

    print("Inventory updated successfully.")

    pause()
    clear_screen()



# =========================
# UPDATE ASSET STATUS
# =========================

def update_asset_status(asset_id, new_status):

    df = load_inventory()

    current_user = ""

    if df is None:
        return

    if new_status not in VALID_STATUSES:

        print(
            f"Invalid status. Use: "
            f"{', '.join(VALID_STATUSES)}"
        )

        return

    if asset_id not in df["Asset ID"].values:

        print(f"Asset ID {asset_id} not found.")
        return

    old_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]

    if "User" in df.columns:
        current_user = df.loc[df["Asset ID"] == asset_id, "User"].iloc[0]

    df.loc[df["Asset ID"] == asset_id, "Status"] = new_status

    save_inventory(df)

    log_history(
        asset_id, 
        old_status, 
        new_status,
        current_user
    )

    print(f"{asset_id} updated to {new_status}")

    pause()
    clear_screen()


# =========================
# SEARCH FUNCTIONS
# =========================

def search_assets_by_serial(serial):

    df = load_inventory()

    if df is None:
        return

    result = df[df["SerialNumber"].astype(str).str.contains(
        serial,
        case=False,
        na=False
    )]

    if result.empty:
        print("No matching serial numbers found")
    else:
        print(result)

    pause()
    clear_screen()


def search_assets_by_asset_name(asset_name):

    df = load_inventory()

    if df is None:
        return

    result = df[df["AssetName"].astype(str).str.contains(
        asset_name,
        case=False,
        na=False
    )]

    if result.empty:
        print("No matching assets found")
    else:
        print(result)

    pause()
    clear_screen()


def search_assets_by_model(model):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Model"].astype(str).str.contains(
        model,
        case=False,
        na=False
    )]

    if result.empty:
        print("No matching models found")
    else:
        print(result)

    pause()
    clear_screen()


def search_asset(asset_id):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Asset ID"] == asset_id]

    if result.empty:
        print(f"No asset found with ID {asset_id}")
    else:
        print(result)

    pause()
    clear_screen()


def search_assets_by_person(person):

    df = load_inventory()

    if df is None:
        return

    result = df[df["User"] == person]

    if result.empty:
        print(f"No assets found for {person}")
    else:
        print(result)

    pause()
    clear_screen()


def search_assets_by_location(location):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Location"] == location]

    if result.empty:
        print(f"No assets found at {location}")
    else:
        print(result)
    
    pause()
    clear_screen()


def search_assets_by_status(status):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Status"] == status]

    if result.empty:
        print(f"No assets found with status {status}")
    else:
        print(result)

    pause()
    clear_screen()


# =========================
# ISSUE OUT ASSET
# =========================

def issue_out_asset(asset_id, person, location):

    df = load_inventory()

    if df is None:
        return

    if asset_id not in df["Asset ID"].values:

        print("Asset not found")
        return
    
    old_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Active"

    df.loc[df["Asset ID"] == asset_id, "User"] = person

    df.loc[df["Asset ID"] == asset_id, "Location"] = location

    save_inventory(df)

    log_history(asset_id,
                old_status, 
                "Issued Out",
                person
    )

    print(f"{asset_id} issued to {person}")

    pause()
    clear_screen()


# =========================
# RETURN ASSET
# =========================

def return_asset(asset_id):

    df = load_inventory()

    if df is None:
        return

    if asset_id not in df["Asset ID"].values:

        print("Asset not found")
        return
    
    old_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Returned"

    current_user = df.loc[df["Asset ID"] == asset_id, "User"].iloc[0]

    df.loc[df["Asset ID"] == asset_id, "User"] = ""

    df.loc[df["Asset ID"] == asset_id, "Location"] = "IT Storage"

    save_inventory(df)

    log_history(
        asset_id, 
        old_status, 
        "Returned", 
        current_user
    )

    print(f"{asset_id} returned to IT Storage")

    pause()
    clear_screen()


# =========================
# MARK AS BROKEN
# =========================

def mark_asset_broken(asset_id):

    df = load_inventory()

    if df is None:
        return

    if asset_id not in df["Asset ID"].values:

        print("Asset not found")
        return
    
    old_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]

    current_user = df.loc[df["Asset ID"] == asset_id, "User"].iloc[0]

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Broken"

    fault_description = input("Enter fault description: ")

    save_inventory(df)

    log_history(
        asset_id,
        old_status,
        "Broken", 
        current_user
    )

    add_asset_log(
        asset_id,
        "Fault Reported",
        fault_description
    )

    print(f"{asset_id} marked as Broken")

    pause()
    clear_screen()


# =========================
# RETIRE ASSET
# =========================

def retire_asset(asset_id):

    df = load_inventory()

    if df is None:
        return

    if asset_id not in df["Asset ID"].values:

        print("Asset not found")
        return
    
    old_status = df.loc[df["Asset ID"] == asset_id, "Status"].iloc[0]

    current_user = df.loc[df["Asset ID"] == asset_id, "User"].iloc[0]

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Retired"

    save_inventory(df)

    log_history(
        asset_id,
        old_status,
        "Retired", 
        current_user
    )

    print(f"{asset_id} retired")

    pause()
    clear_screen()


def generate_reports():

    df = load_inventory()

    if df is None:
        return

    from openpyxl import Workbook

    if os.path.exists(REPORT_WORKBOOK):

        workbook = load_workbook(REPORT_WORKBOOK)

    else:

        workbook = Workbook()

        if "Sheet" in workbook.sheetnames:
            del workbook["Sheet"]

    report_sheets = [

        "Summary",
        "Assigned Assets",
        "Broken Assets",
        "Returned Assets",
        "Retired Assets",
        "IT Storage Assets",
        "User Assets"

    ]

    # Delete old report sheets

    for sheet_name in report_sheets:

        if sheet_name in workbook.sheetnames:

            del workbook[sheet_name]

    # =====================
    # SUMMARY
    # =====================

    summary_sheet = workbook.create_sheet("Summary")

    summary_sheet.append(["Metric", "Count"])

    summary_sheet.append(["Total Assets", len(df)])

    summary_sheet.append([
        "Active",
        len(df[df["Status"] == "Active"])
    ])

    summary_sheet.append([
        "Broken",
        len(df[df["Status"] == "Broken"])
    ])

    summary_sheet.append([
        "Returned",
        len(df[df["Status"] == "Returned"])
    ])

    summary_sheet.append([
        "Retired",
        len(df[df["Status"] == "Retired"])
    ])

    assigned_count = len(

        df[
            (df["Status"] == "Active")
            &
            (df["User"].notna())
            &
            (df["User"] != "")
        ]

    )

    summary_sheet.append([
        "Assigned Assets",
        assigned_count
    ])

    # =====================
    # ASSIGNED ASSETS
    # =====================

    assigned_df = df[
        (df["Status"] == "Active")
        &
        (df["User"].notna())
        &
        (df["User"] != "")
    ]

    assigned_sheet = workbook.create_sheet(
        "Assigned Assets"
    )

    assigned_sheet.append(
        list(assigned_df.columns)
    )

    for row in assigned_df.itertuples(index=False):

        assigned_sheet.append(list(row))

    # =====================
    # BROKEN ASSETS
    # =====================

    broken_df = df[
        df["Status"] == "Broken"
    ]

    broken_sheet = workbook.create_sheet(
        "Broken Assets"
    )

    broken_sheet.append(
        list(broken_df.columns)
    )

    for row in broken_df.itertuples(index=False):

        broken_sheet.append(list(row))

    # =====================
    # RETURNED ASSETS
    # =====================

    returned_df = df[
        df["Status"] == "Returned"
    ]

    returned_sheet = workbook.create_sheet(
        "Returned Assets"
    )

    returned_sheet.append(
        list(returned_df.columns)
    )

    for row in returned_df.itertuples(index=False):

        returned_sheet.append(list(row))

    # =====================
    # RETIRED ASSETS
    # =====================

    retired_df = df[
        df["Status"] == "Retired"
    ]

    retired_sheet = workbook.create_sheet(
        "Retired Assets"
    )

    retired_sheet.append(
        list(retired_df.columns)
    )

    for row in retired_df.itertuples(index=False):

        retired_sheet.append(list(row))

    # =====================
    # IT STORAGE
    # =====================

    storage_df = df[
        df["Location"] == "IT Storage"
    ]

    storage_sheet = workbook.create_sheet(
        "IT Storage Assets"
    )

    storage_sheet.append(
        list(storage_df.columns)
    )

    for row in storage_df.itertuples(index=False):

        storage_sheet.append(list(row))

    # =====================
    # USER ASSETS
    # =====================

    user_assets_df = df[
        (df["User"].notna())
        &
        (df["User"] != "")
    ]

    user_assets_df = user_assets_df.sort_values(
        by=["User", "AssetName"]
    )

    user_sheet = workbook.create_sheet(
        "User Assets"
    )

    user_sheet.append(
        list(user_assets_df.columns)
    )

    for row in user_assets_df.itertuples(index=False):

        user_sheet.append(list(row))

    # Add filters and freeze top row

    for sheet in workbook.worksheets:

        if sheet.max_row > 1:

            sheet.auto_filter.ref = sheet.dimensions
            sheet.freeze_panes = "A2"
            sheet.protection.sheet = True
            sheet.protection.autoFilter = False
            sheet.protection.sort = False


            for column in sheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                sheet.column_dimensions[column_letter].width = max_length + 2

    workbook.save(REPORT_WORKBOOK)

    print("Reports generated successfully.")

    pause()
    clear_screen()


def menu():

    while True:

        clear_screen()

        print(f"\n===== IT Inventory Management System {VERSION} =====")
        print("1. Update Inventory")
        print("2. Generate Reports")
        print("3. Search Asset")
        print("4. Issue Asset")
        print("5. Return Asset")
        print("6. Mark Asset Broken")
        print("7. Retire Asset")
        print("8. View Asset Logs")
        print("9. Repair Asset")
        print("10. Exit")

        choice = input("\nSelect option: ")

        if choice == "1":

            update_inventory()

        elif choice == "2":

            generate_reports()

        elif choice == "3":

            search_menu()

        elif choice == "4":

            asset_id = input("Asset ID: ")
            person = input("User: ")
            location = input("Location: ")

            issue_out_asset(
                asset_id,
                person,
                location
            )

        elif choice == "5":

            asset_id = input("Asset ID: ")
            return_asset(asset_id)

        elif choice == "6":

            asset_id = input("Asset ID: ")
            mark_asset_broken(asset_id)

        elif choice == "7":

            asset_id = input("Asset ID: ")
            retire_asset(asset_id)

        elif choice == "8":

            asset_logs_menu()

        elif choice == "9":

            asset_id = input("Asset ID: ")
            repair_asset(asset_id)

        elif choice == "10":

            print("Goodbye")
            break

        else:

            print("Invalid option")


def search_menu():

    while True:

        clear_screen()

        print("\n=== Asset Search ===")
        print("1. Asset ID")
        print("2. User")
        print("3. Location")
        print("4. Status")
        print("5. Serial Number")
        print("6. Asset Name")
        print("7. Model")
        print("8. Back")

        choice = input("Select option: ")

        if choice == "1":

            search_asset(
                input("Asset ID: ")
            )

        elif choice == "2":

            search_assets_by_person(
                input("User: ")
            )

        elif choice == "3":

            search_assets_by_location(
                input("Location: ")
            )

        elif choice == "4":

            search_assets_by_status(
                input("Status: ")
            )

        elif choice == "5":

            search_assets_by_serial(
                input("Serial Number: ")
            )

        elif choice == "6":

            search_assets_by_asset_name(
                input("Asset Name: ")
            )

        elif choice == "7":

            search_assets_by_model(
                input("Model: ")
            )

        elif choice == "8":

            break

        else:

            print("Invalid option")

def asset_logs_menu():
    while True:
        clear_screen()
        print("\n=== View Asset Logs ===")
        print("1. Add log entry")
        print("2. View logs for asset")
        print("3. Back")

        choice = input("\nSelect option: ")

        if choice == "1":
            asset_id = input("Asset ID: ")
            log_type = input("Log Type (e.g. Maintenance, Repair, Note): ")
            notes = input("Notes: ")
            add_asset_log(asset_id, log_type, notes)
            pause()
        elif choice == "2":
            asset_id = input("Asset ID: ")
            view_asset_logs(asset_id)
        elif choice == "3":
            break
        else:
            print("Invalid option")
            pause()
def pause():
    input("\nPress Enter to continue...")

def clear_screen():

    os.system(
        "cls" if os.name == "nt" else "clear"
    )



# =========================
# MAIN
# =========================

if __name__ == "__main__":

    menu()
    
    
