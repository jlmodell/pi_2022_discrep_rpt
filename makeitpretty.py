import os
import pandas as pd

file = "PICOUNT_2023-01-03_15-21-11 with pretty tags.xlsx"
path = r"\\busse\Purchasing\Physical Inventory\2022\discrepancies"
sheet = "tags"

df = pd.read_excel(os.path.join(path, file), sheet_name=sheet)

df.item = df.item.astype(str)

df2 = pd.read_excel(os.path.join(path, file), sheet_name="All")

df = df[(df.item != "ORES.VALUE.TAG   15:1") | (df.item != "")]

tags = {}

for index, row in df.iterrows():
    if row["item"] not in tags:
        tags[row["item"]] = [
            dict(tag=row["tag"], loc=row["location"], qty=row["counted"])
        ]
    else:
        tags[row["item"]].append(
            dict(tag=row["tag"], loc=row["location"], qty=row["counted"])
        )

df2["tags"] = df2["part"].apply(lambda x: tags.get(x, []))


with pd.ExcelWriter(
    os.path.join(path, "DISCREP RPT WITH pretty_tags_2022.xlsx"), engine="xlsxwriter"
) as writer:
    df2.to_excel(writer, sheet_name="All", index=False)
    df.to_excel(writer, sheet_name="tags", index=False)

df.to_excel(os.path.join(path, "pretty_tags_2022.xlsx"), index=False)
