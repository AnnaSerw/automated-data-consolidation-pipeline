import pandas as pd
from pathlib import Path
import openpyxl
#notice: the .py file is saved in the same folder as .xlsx files
folder = Path(".")
files = folder.glob("*.xlsx")
df = pd.concat([pd.read_excel(file) for file in files],
               ignore_index=True
)
df["date"] = pd.to_datetime(df["date"], format= "%d.%m.%Y")
#check oryginally loaded data
print("\nOryginally loaded data as below:")
print(df)
print(df.info())
print(f"\n{df.shape[0]} rows and {df.shape[1]} columns.\n")
#change data types, data should be as below
#date as datetime
#id as integer
#project_number as text
#value as a value rounded to two decimal places
df["date"] = pd.to_datetime(df["date"], dayfirst=True)
df["id"] = df["id"].astype("int")
df["project_number"] = df["project_number"].astype("string")
df["value"] = df["value"].astype(float).round(2)
print("\nAfter changed date types:\n")
print(df.info())
print(df)
#check if there are any duplicates regarding to id
#id should be unique
df_duplicates = df["id"].duplicated().sum()
#show duplicated rows
duplicates = df[df["id"].duplicated(keep=False)]
duplicates = duplicates.sort_values(
    by=["id",]
)
if df_duplicates > 0:
    print(f"\nThere are {df_duplicates} duplicates in the id column.")
    print("\nPlease find below duplicated positions. The seconds will be dropped.")
    print(duplicates)
else:
    print("There aren't any duplicates in the id column.")
df = df.drop_duplicates(subset=["id"], keep="first")
print("\nData checked after duplicates are dropped:")
print(f"Now there are {df.shape[0]} rows.")
print(f"The number of duplicates {df["id"].duplicated().sum()}")
#check if there are any missing values
print("\nThe number of missing data regarding to coulmns as below:")
print(df.isna().sum())
#remove rows with missing values
df = df.dropna()
print("Rows with missing data removed. Rechecked the number of missing data:")
print(df.isna().sum())

#Data concolidation and nesessary data transformation and data cleanig are done.
#Export data to .xlsx.
date = df["date"].iloc[0]
month_year = date.strftime("%Y-%m")

with pd.ExcelWriter(
    f"Consolidated_data_{month_year}.xlsx",
    engine="openpyxl"
) as writer:
    df.to_excel(
        writer,
        sheet_name=f"{month_year}",
        index=False
    )
    ws = writer.sheets[f"{month_year}"]
    for cell in ws["A"][1:]:
        cell.number_format = "dd.mm.yyyy"
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    for cell in ws["D"][1:]:
        cell.number_format = '#,##0.00'
print("The new file with concatenated and cleaned data has been just created.")