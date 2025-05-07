import streamlit as st
import pandas as pd

# Load Excel files
sales_df = pd.read_excel("sales.xlsx")
sku_df = pd.read_excel("price_list.xlsx")

# Clean sales_df
sales_df['Order Date'] = pd.to_datetime(sales_df['Order Date'], errors='coerce')
sales_df['Order Date'] = sales_df['Order Date'].dt.strftime('%d-%m-%Y')
sales_df['Sale Amount'] = sales_df['Sale Amount'].astype(float).round(2)
sales_df['Sale Quantity'] = sales_df['Sale Quantity'].astype(int)
sales_df['Sale Code'] = sales_df['Sale Code'].astype(str).str.strip().str.upper()

# Clean sku_df
sku_df['Sale Price'] = sku_df['Sale Price'].astype(float).round(2)
sku_df['Sale Code'] = sku_df['Sale Code'].astype(str).str.strip()
sku_df['Solution'] = sku_df['Solution'].astype(str).str.strip().str.lower()
sku_df["From"] = pd.to_datetime(sku_df["From"], errors='coerce').dt.strftime('%d-%m-%Y')
sku_df["To"] = pd.to_datetime(sku_df["To"], errors='coerce').dt.strftime('%d-%m-%Y')

# Add Join Keys
sku_df['Join Key'] = sku_df['Sale Code'] + " - " + sku_df['Sale Price'].round(2).astype(str)
sales_df['Price'] = (sales_df['Sale Amount'] / sales_df['Sale Quantity']).round(2)
sales_df['Join Key'] = sales_df['Sale Code'] + " - " + sales_df['Price'].astype(str)

# Split SKU into solution and not solution
bridge_solution = sku_df[sku_df['Solution'] != 'not solution'].copy()
bridge_notsolution = sku_df[sku_df['Solution'] == 'not solution'].copy()

# Merge Solution
merged_solution = pd.merge(bridge_solution, sales_df, how='inner', on='Join Key')

# Filter sales not in solution
solution_prices = bridge_solution[['Sale Code', 'Sale Price']].drop_duplicates()
solution_prices['Sale Price'] = solution_prices['Sale Price'].round(2)
sales_df['Price'] = sales_df['Price'].round(2)

sales_exclusive_notsolution = pd.merge(
    sales_df,
    solution_prices,
    how='left',
    left_on=['Sale Code', 'Price'],
    right_on=['Sale Code', 'Sale Price'],
    indicator=True
).query('_merge == "left_only"').drop(columns=['Sale Price', '_merge'])

# Merge Not Solution
merged_notsolution = pd.merge(
    bridge_notsolution,
    sales_exclusive_notsolution,
    how='inner',
    on='Sale Code'
)

# Combine
merged_all = pd.concat([merged_solution, merged_notsolution], ignore_index=True)

# Streamlit App Start
st.title("Sales Pivot Matrix")

# Filter for Solutions
available_solutions = sku_df['Solution'].dropna().unique()
selected_solutions = st.sidebar.multiselect("Filter by Solution", available_solutions, default=available_solutions)

# Filter merged_all
filtered_data = merged_all[merged_all['Solution'].isin(selected_solutions)]

# Create hierarchical row label
filtered_data['Solution_SaleCode'] = filtered_data['Solution'].str.title()

# Build Pivot Table
pivot_table = pd.pivot_table(
    filtered_data,
    index='Solution_SaleCode',
    columns='Order Date',
    values='Sale Amount',
    aggfunc='sum',
    fill_value=0,
    margins=True,
    margins_name='Total'
)


# Display
st.dataframe(pivot_table.style.format("{:,.0f}"), height=600)
st.download_button(
    label="Download Pivot Table",
    data=pivot_table.to_csv().encode('utf-8'),
    file_name='pivot_table.csv',
    mime='text/csv'
)
st.markdown("---")
st.markdown("### Note:")    