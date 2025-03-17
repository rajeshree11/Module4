import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import base64


# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    # Adjust the path if needed, e.g. "data/Sample - Superstore.xlsx"
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    # Convert Order Date to datetime if not already
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)

# Filter data by Region
if selected_region != "All":
    df_filtered_region = df_original[df_original["Region"] == selected_region]
else:
    df_filtered_region = df_original

# State Filter
all_states = sorted(df_filtered_region["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)

# Filter data by State
if selected_state != "All":
    df_filtered_state = df_filtered_region[df_filtered_region["State"] == selected_state]
else:
    df_filtered_state = df_filtered_region

# Category Filter
all_categories = sorted(df_filtered_state["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)

# Filter data by Category
if selected_category != "All":
    df_filtered_category = df_filtered_state[df_filtered_state["Category"] == selected_category]
else:
    df_filtered_category = df_filtered_state

# Sub-Category Filter
all_subcats = sorted(df_filtered_category["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)

# Final filter by Sub-Category
df = df_filtered_category.copy()
if selected_subcat != "All":
    df = df[df["Sub-Category"] == selected_subcat]

# ---- Sidebar Date Range (From and To) ----
if df.empty:
    # If there's no data after filters, default to overall min/max
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()
else:
    min_date = df["Order Date"].min()
    max_date = df["Order Date"].max()

from_date = st.sidebar.date_input(
    "From Date", value=min_date, min_value=min_date, max_value=max_date
)
to_date = st.sidebar.date_input(
    "To Date", value=max_date, min_value=min_date, max_value=max_date
)

# Ensure from_date <= to_date
if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

# Apply date range filter
df = df[
    (df["Order Date"] >= pd.to_datetime(from_date))
    & (df["Order Date"] <= pd.to_datetime(to_date))
]

# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

# Function to convert DataFrame to CSV (for download)
def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# Create columns for layout
col1, col2 = st.columns([0.9, 0.1])  # Adjust the ratio based on your layout needs

# Place the download button in the right column
with col2:
    if not df.empty:
        csv = convert_df_to_csv(df)
        st.download_button(
            label="Download CSV",
            data=csv,
            file_name='filtered_data.csv',
            mime='text/csv'
        )
    else:
        st.write("No data available to download.")  # Adjust as necessary

# ---- Information Box for New Users ----
st.info(
    """
    **Welcome to the SuperStore KPI Dashboard!**

    This dashboard provides an interactive way to explore sales performance across various dimensions including regions, states, product categories, and more. Use the sidebar filters to dive into specific data segments and discover insights through dynamic charts that update based on your selections.

    **Features include:**
    - **Dynamic Filters**: Adjust the view by selecting different options in the sidebar.
    - **KPI Metrics**: View key performance indicators such as Sales, Quantity Sold, Profit, and Margin Rate at a glance.
    - **Interactive Charts**: Explore trends over time and compare top products by selected KPIs.

    Get started by using the filters on the left to tailor the data displayed to your interests!
    """
)

# ---- Custom CSS for KPI Tiles ----
st.markdown(
    """
    <style>
    .kpi-box {
        background-color: #FFFFFF;
        border: 2px solid #EAEAEA;
        border-radius: 8px;
        padding: 16px;
        margin: 8px;
        text-align: center;
    }
    .kpi-title {
        font-weight: 600;
        color: #333333;
        font-size: 16px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 24px;
        color: #1E90FF;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Function to format numbers with 'K' for thousands, 'M' for millions
def format_number(num):
    if num >= 1_000_000:
        return f"${num / 1_000_000:.2f}M"
    elif num >= 1000:
        return f"${num / 1000:.1f}K"
    else:
        return f"${num:.2f}"

# Calculate KPIs
if df.empty:
    total_sales = format_number(0)
    total_quantity = '0'
    total_profit = format_number(0)
    margin_rate = "0.00%"
else:
    total_sales = format_number(df["Sales"].sum())
    total_quantity = f"{df['Quantity'].sum()/1000:.1f}K"  # Assuming large numbers for quantities
    total_profit = format_number(df["Profit"].sum())
    margin_rate = f"{(df['Profit'].sum() / df['Sales'].sum() * 100) if df['Sales'].sum() != 0 else 0:.2f}%"

# Display KPIs in columns
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
with kpi_col1:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Sales</div>
            <div class='kpi-value'>{total_sales}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col2:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            <div class='kpi-value'>{total_quantity}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col3:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit</div>
            <div class='kpi-value'>{total_profit}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col4:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Margin Rate</div>
            <div class='kpi-value'>{margin_rate}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---- KPI Selection (Affects Both Charts) ----
st.subheader("Visualize KPI Across Time & Top Products")

if df.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    # Radio button above both charts
    kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
    selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)

    # ---- Prepare Data for Charts ----
    # Daily grouping for line chart
    daily_grouped = df.groupby("Order Date").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    # Avoid division by zero
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)

    # Product grouping for top 10 chart
    product_grouped = df.groupby("Product Name").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)

    # Sort for top 10 by selected KPI
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)

    # ---- Side-by-Side Layout for Charts ----
    col_left, col_right = st.columns(2)

    with col_left:
        # Line Chart
        fig_line = px.line(
            daily_grouped,
            x="Order Date",
            y=selected_kpi,
            title=f"{selected_kpi} Over Time",
            labels={"Order Date": "Date", selected_kpi: selected_kpi},
            template="plotly_white",
        )
        fig_line.update_layout(height=400)
        st.plotly_chart(fig_line, use_container_width=True)

    with col_right:
        # Horizontal Bar Chart
        fig_bar = px.bar(
            top_10,
            x=selected_kpi,
            y="Product Name",
            orientation="h",
            title=f"Top 10 Products by {selected_kpi}",
            labels={selected_kpi: selected_kpi, "Product Name": "Product"},
            color=selected_kpi,
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_bar.update_layout(
            height=400,
            yaxis={"categoryorder": "total ascending"}
        )
        st.plotly_chart(fig_bar, use_container_width=True)

# Create two columns for side-by-side layout
col1, col2 = st.columns(2)

# Pie Chart for Segment
with col1:
    fig_segment = px.pie(df, names='Segment', values=selected_kpi, title=f"Distribution of {selected_kpi} by Segment")
    st.plotly_chart(fig_segment, use_container_width=True)

# Pie Chart for Category
with col2:
    fig_category = px.pie(df, names='Category', values=selected_kpi, title=f"Distribution of {selected_kpi} by Category")
    st.plotly_chart(fig_category, use_container_width=True)

# Treemap Visualization
with st.container():
    fig_treemap = px.treemap(
        df,
        path=[px.Constant("All"), 'Category', 'Sub-Category', 'Product Name'],
        values=selected_kpi,
        color=selected_kpi,
        color_continuous_scale='Viridis',
        title=f"Treemap of {selected_kpi} by Product Categories"
    )
    fig_treemap.update_layout(margin=dict(t=50, l=25, r=25, b=25))
    st.plotly_chart(fig_treemap, use_container_width=True)

