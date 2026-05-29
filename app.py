import pandas as pd


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
    df = pd.read_csv("inventory.csv")

    if asset_id in df["Asset ID"].values:

        df.loc[df["Asset ID"] == asset_id, "status"] = new_status

        # Save the updated inventory
        df.to_csv("inventory.csv", index=False)
        print(f"Updated asset {asset_id} status to {new_status}.")

    else:
        print(f"Asset ID {asset_id} not found in inventory.")






def main():

    # Example of updating an asset's status
    update_asset_status("A002", "Inactive")

if __name__ == "__main__":
    main()