import pandas as pd
import os

# =========================
# FOLDER SETUP
# =========================

REPORTS_FOLDER = "reports"
DATA_FOLDER = "data"

SCANNED_FILE = os.path.join(REPORTS_FOLDER, "scanned_assets.csv")
MANUAL_FILE = os.path.join(REPORTS_FOLDER, "manual_assets.csv")

INVENTORY_FILE = os.path.join(DATA_FOLDER, "inventory.csv")

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
    """Save inventory dataframe."""
    df.to_csv(INVENTORY_FILE, index=False)


def load_inventory():
    """Load inventory dataframe."""

    if not os.path.exists(INVENTORY_FILE):
        print("Inventory file not found.")
        return None

    return pd.read_csv(INVENTORY_FILE)


# =========================
# CREATE / UPDATE INVENTORY
# =========================

def update_inventory():

    # Load source CSV files, files from scanner results can be dropped into here
    scanned_df = pd.read_csv(SCANNED_FILE)
    manual_df = pd.read_csv(MANUAL_FILE)

    # Merge both dataframes
    merged_df = pd.concat(
        [scanned_df, manual_df],
        ignore_index=True
    )

    # Add missing Status column
    if "Status" not in merged_df.columns:
        merged_df["Status"] = "Active"

    # Add missing Asset ID column
    if "Asset ID" not in merged_df.columns:
        merged_df["Asset ID"] = ""

    # Generate IDs only for missing rows
    for index in merged_df.index:

        current_id = merged_df.at[index, "Asset ID"]

        if pd.isna(current_id) or current_id == "":

            merged_df.at[index, "Asset ID"] = (
                f"A{str(index + 1).zfill(3)}"
            )

    save_inventory(merged_df)

    print("Inventory updated successfully.")


# =========================
# UPDATE ASSET STATUS
# =========================

def update_asset_status(asset_id, new_status):

    df = load_inventory()

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

    df.loc[df["Asset ID"] == asset_id, "Status"] = new_status

    save_inventory(df)

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

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Issued Out"

    df.loc[df["Asset ID"] == asset_id, "User"] = person

    df.loc[df["Asset ID"] == asset_id, "Location"] = location

    save_inventory(df)

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

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Returned"

    df.loc[df["Asset ID"] == asset_id, "User"] = ""

    df.loc[df["Asset ID"] == asset_id, "Location"] = "IT Storage"

    save_inventory(df)

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

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Broken"

    save_inventory(df)

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

    df.loc[df["Asset ID"] == asset_id, "Status"] = "Retired"

    save_inventory(df)

    print(f"{asset_id} retired")


# =========================
# MAIN
# =========================

if __name__ == "__main__":

    print("Inventory System Ready")
    mark_asset_broken(asset_id= "A001")


