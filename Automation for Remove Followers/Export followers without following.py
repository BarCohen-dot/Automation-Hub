"""
========================== Introduction: Instagram Data Preparation Script ==========================

This Python script processes Instagram data exported directly from Meta (Instagram)
and creates an Excel file that lists the accounts you follow who do *not* follow you back.

----------------------------------------------------------------------------------------
📦 How to Obtain the Data:
----------------------------------------------------------------------------------------
1. Log into your Instagram account.
2. Go to **Settings → Privacy and Security → Download Your Information**.
3. Request to download your data (choose JSON format if possible).
4. Once you receive the ZIP file, extract it locally.
5. Inside the extracted folder, navigate to **followers_and_following/**.
6. Locate and use the files:
   - `following.json`
   - `followers.json`

----------------------------------------------------------------------------------------
⚙️ What This Script Does:
----------------------------------------------------------------------------------------
✅ Loads the downloaded JSON files (`following.json` and `followers.json`).
✅ Extracts usernames, timestamps, and profile links from your “following” list.
✅ Compares against your “followers” list to find users who do not follow you back.
✅ Exports the results to an Excel file named **not_following_back_detailed.xlsx**.
✅ Adds clickable hyperlinks to each profile directly in Excel.

----------------------------------------------------------------------------------------
📁 Output Example:
----------------------------------------------------------------------------------------
The resulting Excel file includes the following columns:

| Username      | Followed At    | Profile URL               |
|---------------|----------------|---------------------------|
| user_example  | 1696943452     | (clickable link to user)  |

----------------------------------------------------------------------------------------
🧠 Notes:
----------------------------------------------------------------------------------------
- This script does **not** use Selenium or automation; it only processes your downloaded data.
- The resulting Excel file can later be used by the "Instagram Unfollow Bot" script
  to automate unfollowing these users directly on Instagram.
- Always verify the exported data before taking automated actions.

----------------------------------------------------------------------------------------
Author: Bar Cohen
Language: Python 3.13
Libraries: pandas, json, os
========================================================================================
"""

import json
import pandas as pd
import os

# ============= Load JSON safely =============
def load_json_file(filename):
    """
    Load a JSON file safely if it exists, otherwise raise an error.
    Ensures proper error handling for missing or invalid file paths.
    """
    if not os.path.exists(filename):
        raise FileNotFoundError(f"❌ File '{filename}' was not found. Please check the path.")
    with open(filename, encoding="utf-8") as f:
        return json.load(f)

# ============= Load following & followers data =============
# After downloading data from Meta/Instagram, load both JSON files from the extracted folder.
following_data = load_json_file("followers_and_following/following.json")
followers_data = load_json_file("followers_and_following/followers.json")

# ============= Extract Following Info =============
# Extract username, timestamp, and profile URL from your following list.
following = [
    (
        entry["string_list_data"][0]["value"],        # username
        entry["string_list_data"][0].get("timestamp", "N/A"),  # followed at (may be missing)
        entry["string_list_data"][0]["href"]          # profile link
    )
    for entry in following_data.get("relationships_following", [])
]

# ============= Extract Followers Info =============
# followers_data is a list; extract usernames directly.
followers = [
    entry["string_list_data"][0]["value"]
    for entry in followers_data
]

# ============= Find who is NOT following back =============
# Compare the lists and identify users who do not follow you back.
not_following_back = [
    {
        "Username": username,
        "Followed At": timestamp,
        "Profile URL": f'=HYPERLINK("{url}", "{username}")'  # clickable link in Excel
    }
    for username, timestamp, url in following
    if username not in followers
]

# ============= Export to Excel =============
# Create a DataFrame and export it to an Excel file.
df = pd.DataFrame(not_following_back)
output_file = "not_following_back_detailed.xlsx"
df.to_excel(output_file, index=False)

print(f"✅ Excel file created successfully: {output_file}")
