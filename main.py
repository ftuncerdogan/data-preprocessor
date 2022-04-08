import pandas as pd
import numpy as np


# ----- INPUT XLSX ----

# Define File Path
df_main_excel = pd.ExcelFile("define main excel's path")

# Get Sheets Names
df_main_excel.sheet_names

# ----- CONVERT EXCEL TO DF ----
df = pd.read_excel("main excel's path", sheet_name="which sheets you want to work on")

#Print info of the DataFrame.
df.info()

# Convert data type to operable

df = pd.read_excel("main excel's path",
    sheet_name="which sheets you want to work on",
    dtype={'BranchID': str,
            'Date': str})

# No need to math operation then set datatype string

df_brands_info = pd.read_excel("define path",
                dtype= {"Company-ID": str,
                        "Date": str})


#Print info of the DataFrame.
df_brands_info.info()

# Specify data type for columns.

df_brands_branchs = pd.read_excel("define path",
            dtype=str)

# COLUMN H's datatype is int. To fix missing char use zfill func. 
df['Branch ID']= df['Branch ID'].str.zfill(8)


# ----- DF CONFIG ----

# Specify values to consider as True.

# Drop unnecessary column to reduce size of df
df.drop(columns=['unnecessary column'])
# Convert object to datetime
df['Date']= pd.to_datetime(df['string date'])
# String extraction
df['extracted string'] = df['long string'].str[:6]
# Create boolean
df['HasAgreement'] = np.where(df['X'].isna(), False, True)
df['HasB'] = np.where(df['Y'].isna(), False, True)
df_brands_branchs.rename(columns={'Company': 'Company-ID'}, inplace=True)
# Convert object to datetime
df_brands_branchs['DueDate'] = pd.to_datetime(df_brands_info['DueDate'])
# Drop duplicates (if any)
df_dropped_na_dubs_brands_branchs = df_brands_info.dropna().drop_duplicates(subset=['Company-ID'], keep='first')

# CREATE DICT for MAPPING
dict_brand_info_by_branch_id = df_dropped_na_dubs_brands_branchs.groupby(['Company-ID'])['BranchID'].apply(list).to_dict()

dict_brand_info_by_branch_id = df_brands_info.groupby(['Company-ID'])['BranchID'].apply(list).to_dict()
reversed_dict_brand_info_by_branch_id = {val: key for key in dict_brand_info_by_branch_id for val in dict_brand_info_by_branch_id[key]}

company_id_by_due_date = df_brands_info.groupby(['Company-ID'])['DueDate'].apply(list).to_dict()
non_dubs_company_id_due_date = {a:list(set(b)) for a, b in company_id_by_due_date.items()}

# ----- MAPPING ----
df["Company-ID"] = df["BranchID"].map(reversed_dict_brand_info_by_branch_id)

# Descripton contains 'X' then set Company-ID value X's ID
df.loc[df['Description'].str.contains('X'), 'Company-ID'] = 'X ID'

# ----- CREATE ANOTHER DATE COLUMN FRO COMPARE DUE DATE ----
df["Company-ID_DueDate"] = df["Company-ID"].map(df_dropped_na_dubs_brands_branchs.set_index("Company-ID")["DueDate"])

# Check first 10 entry of DataFrame
df.head(10)
# Check last 10 entry of DataFrame
df.tail(10)

# ----- OUTPUT ----
df.to_excel('OUTPUT.xlsx')
