# automated-data-consolidation-pipeline
The project aims to automate and accelerate the consolidation and preparation of data from multiple sources for subsequent processing.

## Overview

This case presents how to automate data processing activity that needs to be performed regularly, for example, every month. The purpose is to prepare a single file containing transformed and cleaned data, ready to the next step, such as uploading it to an ERP system or creating visualizations. The example uses only a few .xlsx files, whereas in practice, there can be far more extensive and numerous files. I created this code especially for such numerous cases to optimize work. Of course, the code can be modified accordingly, and other formats, such as csv, can be used as source files for the output data.

## Workflow
- Data acquisition and concatenation
- Data exploration 
- Data transformation and cleaning 
- Data validation 
- Final dataset

## Technologies

- Python
- Pandas
- OpenPyXL
- pathlib

## Usage

1. Place the input .xlsx files
2. Run the pipeline
3. Find the consolidated data

## Features

First, data should be loaded. Received files are saved in the same folder as .py file. All received .xlsx files (“file_A.xlsx”, “file_B.xlsx”, “file_C.xlsx”) consist of the same column and should concern the same period.

Start with concatenation at the beginning.
```python
folder = Path(".")
files = folder.glob("*.xlsx")
df = pd.concat([pd.read_excel(file) for file in files],
               ignore_index=True
)
```

It is very important that the date is in the correct format.
```python
df["date"] = pd.to_datetime(df["date"], format= "%d.%m.%Y")
```

Take a quick look at the loaded data:
```python
print("\nOryginally loaded data as below:")
print(df)
print(df.info())
print(f"\n{df.shape[0]} rows and {df.shape[1]} columns.\n")
```

Check how many rows and columns (57 rows and 4 columns in this case).
Also check date types and change them to appropriate ones.
```python
df["date"] = pd.to_datetime(df["date"], dayfirst=True)
df["id"] = df["id"].astype("int")
df["project_number"] = df["project_number"].astype("string")
df["value"] = df["value"].astype(float).round(2)
```

Make sure, there are no duplicates regarding to “id” column among the loaded data.
The decision is to drop the second and the other rows with duplicates.
```python
df_duplicates = df["id"].duplicated().sum()
duplicates = df[df["id"].duplicated(keep=False)]
duplicates = duplicates.sort_values(
    by=["id",]
)
if df_duplicates > 0:
    print(f"\nThere are {df_duplicates} duplicates in the id column.")
    print("\nPlease find below duplicated positions. The second ones will be dropped.")
    print(duplicates)
else:
    print("There aren't any duplicates in the id column.")
df = df.drop_duplicates(subset=["id"], keep="first")
```

Again, double check if there are no duplicates any more.
```python
print(f"Now there are {df.shape[0]} rows.")
print(f"The number of duplicates {df["id"].duplicated().sum()}")
```

Uploaded data could have also missing values. Check it and remove rows with missing data.
```python
print(df.isna().sum())
df = df.dropna()
```

The last step is to export data. New .xlsx file will be created. 
The file name and the sheet name is appropriate for the data period.
```python
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
```
        
The new file with concatenated and cleaned data has been created (please find attached the file “consolidated_data_2026-08.xlsx”).

