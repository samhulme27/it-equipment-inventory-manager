import pandas as pd

VALID_STATUSES = {"Active", "Inactive", "Maintenance", "Retired"}


def update_inventory():
    #dataframes for scanned and manual assets
    scanned_df = pd.read_csv('scanned_assets.csv')
    manual_df = pd.read_csv('manual_assets.csv')

    #merging the two dataframes
    merged_df = pd.concat([scanned_df, manual_df])

    #adding a status column to indicate active assets
    if "status" not in merged_df.columns:
        merged_df["status"] = "Active"

    #create a set of unique asset IDs for each asset
    merged_df["Asset ID"] = [
        f"A{str(i + 1).zfill(3)}"
        for i in range(len(merged_df))
    ]

    #saving the merged dataframe to a new csv file
    merged_df.to_csv("inventory.csv", index=False)



def update_asset_status(asset_id, new_status):
    # Load the inventory
    df = load_inventory()

    if df is None:
        return

    if asset_id in df["Asset ID"].values:

        if new_status not in VALID_STATUSES:
            print(f"Invalid status '{new_status}'. Please use one of the following: {', '.join(VALID_STATUSES)}")
            return

        df.loc[df["Asset ID"] == asset_id, "status"] = new_status

        # Save the updated inventory
        df.to_csv("inventory.csv", index=False)
        print(f"Updated asset {asset_id} status to {new_status}.")

    else:
        print(f"Asset ID {asset_id} not found in inventory.")


def load_inventory():
    # Load the inventory from the CSV file and return it as a DataFrame if it exists, otherwise return None
    try:
        df = pd.read_csv("inventory.csv")
        print("Inventory loaded successfully.")
        return df
    except FileNotFoundError:
        print("Inventory file not found. Please run update_inventory() first.")
        return None




##search funtions for assets by ID, person, location, and status
def search_asset(asset_id):
    df = load_inventory()
    if df is None:
        return None
    result = df[df["Asset ID"] == asset_id]
    if result.empty:
        print(f"No asset found with ID {asset_id}.")
    else:
        print(result)

def search_assets_by_person(person):
    df = load_inventory()
    if df is None:
        return None
    result = df[df["User"] == person]
    if result.empty:
        print(f"No assets found for person {person}, try entering the full name or check for typos.")
    else:
        print(result)

def search_assets_by_location(location):
    df = load_inventory()
    if df is None:
        return None
    result = df[df["Location"] == location]
    if result.empty:
        print(f"No assets found at location {location}, try entering the full location or check for typos.")
    else:
        print(result)

def search_assets_by_status(status):
    df = load_inventory()
    if df is None:
        return None
    result = df[df["status"] == status]
    if result.empty:
        print(f"No assets found with status {status}, try entering a valid status: {', '.join(VALID_STATUSES)}.")
    else:
        print(result)



def main():

    # Example of updating an asset's status
    search_assets_by_status("Inactive")

if __name__ == "__main__":
    main()