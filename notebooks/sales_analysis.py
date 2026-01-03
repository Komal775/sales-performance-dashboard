import pandas as pd

# Load Excel file (RAW DATA SHEET, NOT PIVOT)
df = pd.read_excel(
    r"C:\Users\Komal Nimbalkar\OneDrive\Desktop\Sales_Analytics_Project\data_clean\sales_clean.xlsx"
)


# View data
print("FIRST 5 ROWS")
print(df.head())

print("\nLAST 5 ROWS")
print(df.tail())

print("\nDATA INFO")
print(df.info())

print("\nDESCRIPTIVE STATS")
print(df.describe())

# Access columns
print("\nSALES COLUMN")
print(df["Sales"])

print("\nMONTH COLUMN")
print(df["Month"])

# Basic calculations
print("\nTOTAL SALES")
print(df["Sales"].sum())

print("\nAVERAGE SALES")
print(df["Sales"].mean())

print("\nMONTHLY SALES ANALYSIS")

monthly_sales = (
    df.groupby("Month")["Sales"]
    .sum()
    .reset_index()
    .sort_values("Month")
)

print(monthly_sales)

print("\nREGION-WISE SALES")

region_sales = (
    df.groupby("Region")["Sales"]
    .sum()
    .reset_index()
    .sort_values("Sales", ascending=False)
)

print(region_sales)

print("\nTOP 10 PRODUCTS BY SALES")

top_products = (
    df.groupby("Product Name")["Sales"]
    .sum()
    .reset_index()
    .sort_values("Sales", ascending=False)
    .head(10)
)

print(top_products)

# Export analysis results for Power BI (absolute path)
monthly_sales.to_csv(
    r"C:\Users\Komal Nimbalkar\OneDrive\Desktop\Sales_Analytics_Project\data_clean\monthly_sales.csv",
    index=False
)

region_sales.to_csv(
    r"C:\Users\Komal Nimbalkar\OneDrive\Desktop\Sales_Analytics_Project\data_clean\region_sales.csv",
    index=False
)

top_products.to_csv(
    r"C:\Users\Komal Nimbalkar\OneDrive\Desktop\Sales_Analytics_Project\data_clean\top_10_products.csv",
    index=False
)

print("\nCSV files exported successfully for Power BI")
