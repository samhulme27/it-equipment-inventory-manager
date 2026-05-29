import pandas as pd

#dataframes for scanned and manual assets
scanned_df = pd.read_csv('scanned_assets.csv')
manual_df = pd.read_csv('manual_assets.csv')

#merging the two dataframes
merged_df = pd.concat([scanned_df, manual_df])

#adding a status column to indicate active assets
merged_df["status"] = "Active"

#saving the merged dataframe to a new csv file
merged_df.to_csv("inventory.csv", index=False)

print(merged_df)

