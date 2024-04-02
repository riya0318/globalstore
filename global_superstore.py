import openpyxl
import pandas as pd
import matplotlib.pyplot as plt

excel_file = "C:\\Users\\Riya Yadav\\Downloads\\global_superstore_2016.xlsx"
df = pd.read_excel(excel_file, sheet_name="Returns", header=0, usecols="A:E", nrows=1080)
print(df)
print(df.info())
print(df.drop_duplicates())
print(df.describe())


# print(df.ffill(inplace=True))
# print(df.info())


def remove_numeric(x):
    if isinstance(x, str):
        return x
    else:
        return None


df["Region"] = df["Region"].apply(remove_numeric)
print(df)
print(df.ffill(inplace=True))
print(df)

df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce', downcast='float')
print(df)
print(df.ffill(inplace=True))
print(df)
print(df.info())


def total_sales(row):
    return row["Quantity"] * row["Sales"]


df["total_sales"] = df.apply(total_sales, axis=1)
print(df)

df.to_excel("C:\\complete.xlsx", index=False)
