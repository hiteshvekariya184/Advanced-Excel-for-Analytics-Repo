import pandas as pd

# --------------------------------------------------
# 1. Read the input CSV file
# --------------------------------------------------

df = pd.read_csv("food_delivery_orders.csv")

# --------------------------------------------------
# 2. Convert numeric columns safely
# --------------------------------------------------

df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce")
df["DeliveryTime"] = pd.to_numeric(df["DeliveryTime"], errors="coerce")

# --------------------------------------------------
# 3. Check for missing Revenue
# --------------------------------------------------

missing_revenue = df["Revenue"].isna().sum()

print("Missing Revenue values:", missing_revenue)

# --------------------------------------------------
# 4. Check for invalid negative DeliveryTime
# --------------------------------------------------

negative_delivery = df[df["DeliveryTime"] < 0]

if not negative_delivery.empty:
    print("\nWarning: Negative DeliveryTime values found:")
    print(negative_delivery[["OrderID", "DeliveryTime"]])

# --------------------------------------------------
# 5. Create PerformanceFlag
# --------------------------------------------------

def performance_flag(delivery_time):

    if pd.isna(delivery_time):
        return "Invalid"

    elif delivery_time < 0:
        return "Invalid"

    elif delivery_time > 60:
        return "Delayed"

    else:
        return "On Time"


df["PerformanceFlag"] = df["DeliveryTime"].apply(performance_flag)

# --------------------------------------------------
# 6. Create cuisine-wise summary
# --------------------------------------------------

summary = (
    df.groupby("Cuisine")
    .agg(
        TotalRevenue=("Revenue", "sum"),
        AverageDeliveryTime=("DeliveryTime", "mean"),
        OrderCount=("OrderID", "count")
    )
    .reset_index()
)

# --------------------------------------------------
# 7. Save results to Excel
# --------------------------------------------------

with pd.ExcelWriter(
    "food_delivery_summary.xlsx",
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Cleaned Orders",
        index=False
    )

    summary.to_excel(
        writer,
        sheet_name="Cuisine Summary",
        index=False
    )

# --------------------------------------------------
# 8. Final message
# --------------------------------------------------

print("\nAnalysis completed successfully.")
print("Output file: food_delivery_summary.xlsx")