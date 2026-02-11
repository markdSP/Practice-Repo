import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(
    page_title="Sales Dashboard",
    page_icon="📈",
    layout="wide",
)

# ----------------------------
# Helpers
# ----------------------------
@st.cache_data
def load_sales_xlsx(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="SalesData")

    # Standardize column names (handles small variations)
    expected = ["Date", "Product", "Quantity", "Unit Price", "Total"]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in SalesData sheet: {missing}")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])
    df["Month"] = df["Date"].dt.to_period("M").astype(str)

    # Ensure numeric
    for col in ["Quantity", "Unit Price", "Total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "FilteredData") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# ----------------------------
# Load data
# ----------------------------
st.title("📈 Sales Dashboard")
st.caption("Interactive reporting dashboard (filters, KPIs, trends, breakdowns)")

# Option A: hardcode your file (recommended for you)
DATA_PATH = "sales_report2.xlsx"

# Option B: allow upload (uncomment if you want)
# uploaded = st.file_uploader("Upload sales Excel (.xlsx)", type=["xlsx"])
# if uploaded is None:
#     st.stop()
# df = pd.read_excel(uploaded, sheet_name="SalesData")

try:
    df = load_sales_xlsx(DATA_PATH)
except Exception as e:
    st.error(f"Could not load '{DATA_PATH}'. Put it in the same folder as this app.\n\nDetails: {e}")
    st.stop()

# ----------------------------
# Sidebar filters
# ----------------------------
st.sidebar.header("Filters")

min_date = df["Date"].min().date()
max_date = df["Date"].max().date()

date_range = st.sidebar.date_input(
    "Date range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)

if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

products = sorted(df["Product"].dropna().unique().tolist())
product_sel = st.sidebar.multiselect("Product", products, default=products)

# Apply filters
f = df[
    (df["Date"].dt.date >= start_date)
    & (df["Date"].dt.date <= end_date)
    & (df["Product"].isin(product_sel))
].copy()

# ----------------------------
# KPIs
# ----------------------------
k1, k2, k3, k4 = st.columns(4)

total_sales = float(f["Total"].sum())
total_units = int(f["Quantity"].sum())
avg_unit_price = float(f["Unit Price"].mean()) if len(f) else 0.0
tx_count = int(len(f))

k1.metric("Total Sales", f"{total_sales:,.2f}")
k2.metric("Units Sold", f"{total_units:,}")
k3.metric("Avg Unit Price", f"{avg_unit_price:,.2f}")
k4.metric("Transactions", f"{tx_count:,}")

st.divider()

# ----------------------------
# Charts (top row)
# ----------------------------
left, right = st.columns([1.2, 1])

with left:
    st.subheader("📆 Monthly Sales Trend")
    monthly = (
        f.groupby("Month", as_index=False)["Total"].sum()
         .sort_values("Month")
    )
    fig_month = px.line(
        monthly,
        x="Month",
        y="Total",
        markers=True,
        title="Total Sales per Month",
    )
    fig_month.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig_month, use_container_width=True)

with right:
    st.subheader("🧾 Sales by Product")
    by_prod = (
        f.groupby("Product", as_index=False)
         .agg(Sales=("Total", "sum"), Units=("Quantity", "sum"))
         .sort_values("Sales", ascending=False)
    )
    fig_prod = px.bar(
        by_prod,
        x="Product",
        y="Sales",
        title="Total Sales by Product",
        hover_data=["Units"],
    )
    fig_prod.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig_prod, use_container_width=True)

# ----------------------------
# Charts (bottom row)
# ----------------------------
c1, c2 = st.columns([1, 1])

with c1:
    st.subheader("📀 Daily Sales")
    daily = f.groupby("Date", as_index=False)["Total"].sum().sort_values("Date")
    fig_daily = px.area(daily, x="Date", y="Total", title="Daily Total Sales")
    fig_daily.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig_daily, use_container_width=True)

with c2:
    st.subheader("📌 Product Mix (Share of Sales)")
    mix = by_prod.copy()
    if mix["Sales"].sum() > 0:
        mix["Share"] = mix["Sales"] / mix["Sales"].sum()
    else:
        mix["Share"] = 0
    fig_pie = px.pie(mix, names="Product", values="Sales", title="Sales Mix")
    fig_pie.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig_pie, use_container_width=True)

st.divider()

# ----------------------------
# Data Table + download
# ----------------------------
tab1, tab2 = st.tabs(["📄 Data", "⬇️ Export"])

with tab1:
    st.subheader("Filtered Transactions")
    st.dataframe(
        f.sort_values("Date", ascending=False),
        use_container_width=True,
        height=420,
    )

with tab2:
    st.subheader("Download filtered data")
    st.write("Export what you are seeing (after filters) to Excel.")
    excel_bytes = to_excel_bytes(f, sheet_name="FilteredSales")
    st.download_button(
        "Download Excel",
        data=excel_bytes,
        file_name="filtered_sales.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
