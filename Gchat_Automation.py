# add your conn details and query
import mysql.connector
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import requests

conn = mysql.connector.connect(
        host="",
        user="",
        password="",
        database="",   
        port=3306
    )



# ------------------ SQL QUERY ------------------
query = """

"""

# ------------------ RUN QUERY ------------------
df = pd.read_sql(query, conn)
conn.close()

# ------------------ FINAL REPORT LOGIC ------------------
final_df = pd.DataFrame()

final_df["Agent"] = df["Agent"]
final_df["Booked_Qty"] = df["Booked"].astype(int)

# ENACH %
final_df["ENACH_%"] = (
    df["Active_Enach"]
    .div(df["Booked"])
    .mul(100)
    .replace([float("inf"), -float("inf")], 0)
    .fillna(0)
    .round(0)
    .astype(int)
).astype(str) + "%"

# FORM FILLING %
final_df["Form_Filling_%"] = (
    df["Form_Filled"]
    .div(df["OTP_Eligible"])
    .mul(100)
    .replace([float("inf"), -float("inf")], 0)
    .fillna(0)
    .round(0)
    .astype(int)
).astype(str) + "%"

# ------------------ PRINT ------------------
print(final_df)

# ------------------ SAVE TABLE AS IMAGE ------------------
fig, ax = plt.subplots(figsize=(14, max(2, len(final_df) * 0.6)))
ax.axis('off')

table = ax.table(
    cellText=final_df.values,
    colLabels=final_df.columns,
    loc='center',
    cellLoc='center'
)

# Left-align Agent column
for row in range(len(final_df) + 1):
    table[(row, 0)]._loc = 'left'

table.auto_set_font_size(False)
table.set_fontsize(10)
table.scale(1, 1.4)

plt.savefig("agent_mis.png", bbox_inches="tight")
plt.close()

# ------------------ GOOGLE CHAT WEBHOOK ------------------
WEBHOOK_URL = ""

text_table = final_df.to_string(index=False)

message = {
    "text": (
        "*Agent-wise ENACH & Form Filling MIS*\n"
        "```\n"
        f"{text_table}\n"
        "```"
    )
}

response = requests.post(WEBHOOK_URL, json=message)

if response.status_code == 200:
    print("MIS sent to Google Chat successfully!")
else:
    print("Failed to send MIS:", response.text)
