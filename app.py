import pandas as pd
import os

from openpyxl import load_workbook
from datetime import datetime

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


# =========================
# SEARCH FUNCTIONS
# =========================

def search_asset(asset_id):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Asset ID"] == asset_id]

    if result.empty:
        print(f"No asset found with ID {asset_id}")
    else:
        print(result)


def search_assets_by_person(person):

    df = load_inventory()

    if df is None:
        return

    result = df[df["User"] == person]

    if result.empty:
        print(f"No assets found for {person}")
    else:
        print(result)


def search_assets_by_location(location):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Location"] == location]

    if result.empty:
        print(f"No assets found at {location}")
    else:
        print(result)


def search_assets_by_status(status):

    df = load_inventory()

    if df is None:
        return

    result = df[df["Status"] == status]

    if result.empty:
        print(f"No assets found with status {status}")
    else:
        print(result)


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

    save_inventory(df)

    log_history(
        asset_id,
        old_status,
        "Broken", 
        current_user
    )

    print(f"{asset_id} marked as Broken")


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


def generate_reports():

    df = load_inventory()

    if df is None:
        return

    workbook = load_workbook(INVENTORY_FILE)

    report_sheets = [

        "Summary",
        "Assigned Assets",
        "Broken Assets",
        "Returned Assets",
        "IT Storage Assets"

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

    workbook.save(INVENTORY_FILE)

    print("Reports generated successfully.")




# =========================
# MAIN
# =========================

if __name__ == "__main__":

    generate_reports()

    print("Inventory System Ready")

