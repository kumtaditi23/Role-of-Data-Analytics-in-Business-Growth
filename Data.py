# ============================================
# Role of Data Analytics in Business Growth
# KPI Tracking + Business Intelligence (Python)
# ============================================

import pandas as pd
import numpy as np

# ------------------------------------------------
# Load Dataset
# ------------------------------------------------
df = pd.read_excel
("business_kpi_dataset.xlsx")

# Convert date column
df["OrderDate"] = pd.to_datetime(df["OrderDate"])

# Create additional time columns
df["Year"] = df["OrderDate"].dt.year
df["Month"] = df["OrderDate"].dt.to_period("M").astype(str)

# ------------------------------------------------
# Data Quality Checks
# ------------------------------------------------
print("Dataset Shape:", df.shape)
print("\nMissing Values:\n", df.isnull().sum())
print("\nData Preview:\n", df.head())

# Remove duplicates if any
df = df.drop_duplicates()

# ------------------------------------------------
# SALES & REVENUE KPIs
# ------------------------------------------------

kpi_summary = {
    "Total Revenue": df["Revenue"].sum(),
    "Total Profit": df["Profit"].sum(),
    "Average Order Value": df["Revenue"].mean(),
    "Total Orders": df["OrderID"].nunique(),
    "Unique Customers": df["CustomerID"].nunique(),
    "Average Profit Margin %": (df["Profit"].sum() / df["Revenue"].sum()) * 100
}

kpi_df = pd.DataFrame(kpi_summary, index=[0])
print("\n--- KPI Summary ---\n")
print(kpi_df)

# Export for Power BI
kpi_df.to_csv("output_kpi_summary.csv", index=False)

# ------------------------------------------------
# MONTHLY SALES TREND
# ------------------------------------------------
monthly_sales = (
    df.groupby("Month")
      .agg(Revenue=("Revenue", "sum"),
           Profit=("Profit", "sum"),
           Orders=("OrderID", "count"))
      .reset_index()
      .sort_values("Month")
)

monthly_sales["Profit Margin %"] = (monthly_sales["Profit"] / monthly_sales["Revenue"]) * 100

print("\n--- Monthly Sales Trend ---\n")
print(monthly_sales.head())

monthly_sales.to_csv("output_monthly_sales.csv", index=False)

# ------------------------------------------------
# REGION & CITY ANALYSIS
# ------------------------------------------------
region_sales = (
    df.groupby("Region")
      .agg(Revenue=("Revenue", "sum"),
           Profit=("Profit", "sum"),
           Orders=("OrderID", "count"))
      .reset_index()
)

region_sales["Profit Margin %"] = (region_sales["Profit"] / region_sales["Revenue"]) * 100

print("\n--- Region-wise Performance ---\n")
print(region_sales)

region_sales.to_csv("output_region_performance.csv", index=False)

city_sales = (
    df.groupby("City")
      .agg(Revenue=("Revenue", "sum"),
           Profit=("Profit", "sum"))
      .reset_index()
)

city_sales.to_csv("output_city_sales.csv", index=False)

# ------------------------------------------------
# PRODUCT & CATEGORY PERFORMANCE
# ------------------------------------------------
product_perf = (
    df.groupby("Product")
      .agg(Revenue=("Revenue", "sum"),
           Profit=("Profit", "sum"),
           Quantity=("Quantity", "sum"))
      .reset_index()
)

product_perf["Profit Margin %"] = (product_perf["Profit"] / product_perf["Revenue"]) * 100

print("\n--- Product Performance ---\n")
print(product_perf.sort_values("Revenue", ascending=False).head())

product_perf.to_csv("output_product_performance.csv", index=False)

category_perf = (
    df.groupby("Category")
      .agg(Revenue=("Revenue", "sum"),
           Profit=("Profit", "sum"))
      .reset_index()
)

category_perf["Profit Margin %"] = (category_perf["Profit"] / category_perf["Revenue"]) * 100

category_perf.to_csv("output_category_performance.csv", index=False)

# ------------------------------------------------
# CUSTOMER ANALYTICS + RFM SEGMENTATION
# ------------------------------------------------

# Latest purchase date reference
snapshot_date = df["OrderDate"].max() + pd.Timedelta(days=1)

rfm = (
    df.groupby("CustomerID")
      .agg(
          Recency=("OrderDate", lambda x: (snapshot_date - x.max()).days),
          Frequency=("OrderID", "count"),
          Monetary=("Revenue", "sum")
      )
      .reset_index()
)

# RFM Scoring
rfm["R_Score"] = pd.qcut(rfm["Recency"], 4, labels=[4,3,2,1])
rfm["F_Score"] = pd.qcut(rfm["Frequency"].rank(method="first"), 4, labels=[1,2,3,4])
rfm["M_Score"] = pd.qcut(rfm["Monetary"], 4, labels=[1,2,3,4])

rfm["RFM_Score"] = (
    rfm["R_Score"].astype(int)
  + rfm["F_Score"].astype(int)
  + rfm["M_Score"].astype(int)
)

# Customer Segments
def assign_segment(score):
    if score >= 10:
        return "Premium / Loyal"
    elif score >= 7:
        return "High-Value"
    elif score >= 5:
        return "Potential"
    else:
        return "At-Risk / Low Value"

rfm["Segment"] = rfm["RFM_Score"].apply(assign_segment)

print("\n--- RFM Customer Segmentation ---\n")
print(rfm.head())

rfm.to_csv("output_rfm_segmentation.csv", index=False)

segment_summary = (
    rfm.groupby("Segment")
       .agg(
           Customers=("CustomerID", "count"),
           AvgRevenue=("Monetary", "mean")
       )
       .reset_index()
)

print("\n--- Segment Summary ---\n")
print(segment_summary)

segment_summary.to_csv("output_segment_summary.csv", index=False)

# ------------------------------------------------
# SAVE CLEANED DATA FOR POWER BI
# ------------------------------------------------
df.to_csv("output_clean_dataset.csv", index=False)

print("\nAll analysis files exported successfully!")
