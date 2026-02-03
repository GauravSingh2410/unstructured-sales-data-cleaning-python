import pandas as pd

df = pd.read_excel("sales_voucher.xlsx")

print(df.head()) 
print("Shape of data (rows, columns):", df.shape)

print("\nColumn Names:")
print(df.columns)

print("\nData Types:")
print(df.dtypes)

print("\nFirst 5 Rows:")
print(df.head())

print("\nLast 5 Rows:")
print(df.tail())

df["Operator Number"] = df["Operator Number"].astype(str)

df["tax amount"] = df["tax amount"].fillna(0)

df["discount"] = df["discount"].fillna(0)

df["State"] = df["State"].fillna("unknown")

df["company name"] = df["company name"].fillna("Unknown")

print("\nMissing values in each column:")
print(df.isnull().sum())

df.to_excel("Gaurav_sales_voucher.xlsx", index=True)

