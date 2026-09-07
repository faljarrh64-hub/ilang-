# ============================================================
# ILANG CLINIC - DATA ANALYSIS DASHBOARD
# ============================================================

import os
import pandas as pd
import streamlit as st
import plotly.express as px


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="ILANG Clinic Dashboard",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    .main {
        background-color: #f7f8fa;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    h1 {
        font-weight: 700;
    }

    h2, h3 {
        font-weight: 600;
    }

    [data-testid="stMetric"] {
        background-color: white;
        border: 1px solid #e6e8eb;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }

    [data-testid="stSidebar"] {
        background-color: #ffffff;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# CONSTANTS
# ============================================================

FILE_PATH = "ILANG_Clinic_Dataset.xlsx"
SHEET_NAME = "Appointments"

LASER_DEVICES = 12
LASER_HOURS_PER_DEVICE = 14
LASER_CAPACITY_PER_DAY = LASER_DEVICES * LASER_HOURS_PER_DEVICE


# ============================================================
# HEADER
# ============================================================

st.title("🏥 ILANG Clinic")
st.caption("Professional Clinic Data Analysis Dashboard")


# ============================================================
# CHECK FILE
# ============================================================

if not os.path.exists(FILE_PATH):

    st.error(
        f"""
        ❌ File not found:

        `{FILE_PATH}`

        Make sure that `app.py` and `ILANG_Clinic_Dataset.xlsx`
        are inside the same folder.
        """
    )

    st.stop()


# ============================================================
# LOAD DATA
# ============================================================

try:

    df = pd.read_excel(
        FILE_PATH,
        sheet_name=SHEET_NAME
    )

except Exception as e:

    st.error(f"❌ Error reading Excel file: {e}")

    st.stop()


# ============================================================
# CLEAN COLUMN NAMES
# ============================================================

df.columns = [
    str(col).strip()
    for col in df.columns
]


# ============================================================
# REMOVE UNNECESSARY COLUMNS
# ============================================================

columns_to_drop = [
    "Duplicate_Checkk",
    "Unnamed: 16"
]

df = df.drop(
    columns=columns_to_drop,
    errors="ignore"
)


# Remove columns with numeric names
df = df.loc[
    :,
    [
        col
        for col in df.columns
        if not str(col).replace(".", "", 1).isdigit()
    ]
]


# ============================================================
# REQUIRED COLUMNS
# ============================================================

required_columns = [
    "Appointment_ID",
    "Date",
    "Day",
    "Time",
    "Department",
    "Service",
    "Staff",
    "Device_ID",
    "Package_Type",
    "Session_Number",
    "Price",
    "Discount",
    "Final_Price",
    "Payment_Method",
    "Status"
]

missing_columns = [
    col
    for col in required_columns
    if col not in df.columns
]

if missing_columns:

    st.error(
        "❌ The following columns are missing from the Excel file:"
    )

    st.write(missing_columns)

    st.stop()


# ============================================================
# DATA CLEANING
# ============================================================

df["Date"] = pd.to_datetime(
    df["Date"],
    errors="coerce"
)

df["Price"] = pd.to_numeric(
    df["Price"],
    errors="coerce"
).fillna(0)

df["Discount"] = pd.to_numeric(
    df["Discount"],
    errors="coerce"
).fillna(0)

df["Final_Price"] = pd.to_numeric(
    df["Final_Price"],
    errors="coerce"
).fillna(0)

df["Session_Number"] = pd.to_numeric(
    df["Session_Number"],
    errors="coerce"
).fillna(0)

df["Department"] = (
    df["Department"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)

df["Service"] = (
    df["Service"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)

df["Staff"] = (
    df["Staff"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)

df["Status"] = (
    df["Status"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)

df["Package_Type"] = (
    df["Package_Type"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)

df["Payment_Method"] = (
    df["Payment_Method"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)

df["Device_ID"] = (
    df["Device_ID"]
    .fillna("No Device")
    .astype(str)
    .str.strip()
)

df["Time"] = (
    df["Time"]
    .fillna("Unknown")
    .astype(str)
    .str.strip()
)


# ============================================================
# DATE FEATURES
# ============================================================

df["Year"] = df["Date"].dt.year

df["Month"] = df["Date"].dt.month

df["Month_Name"] = df["Date"].dt.strftime("%B")

df["Day_Number"] = df["Date"].dt.day

df["Date_Only"] = df["Date"].dt.date


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("🎛️ Dashboard Filters")

st.sidebar.markdown("---")


# ============================================================
# DATE FILTER
# ============================================================

valid_dates = df["Date"].dropna()

if len(valid_dates) == 0:

    st.error("❌ No valid dates found in the dataset.")

    st.stop()


min_date = valid_dates.min().date()

max_date = valid_dates.max().date()


date_range = st.sidebar.date_input(
    "📅 Date Range",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)


# ============================================================
# DEPARTMENT FILTER
# ============================================================

department_options = sorted(
    df["Department"].unique().tolist()
)

selected_departments = st.sidebar.multiselect(
    "🏢 Department",
    department_options,
    default=department_options
)


# ============================================================
# STATUS FILTER
# ============================================================

status_options = sorted(
    df["Status"].unique().tolist()
)

selected_statuses = st.sidebar.multiselect(
    "📌 Status",
    status_options,
    default=status_options
)


# ============================================================
# SERVICE FILTER
# ============================================================

service_options = sorted(
    df["Service"].unique().tolist()
)

selected_services = st.sidebar.multiselect(
    "💆 Service",
    service_options,
    default=service_options
)


# ============================================================
# PAYMENT METHOD FILTER
# ============================================================

payment_options = sorted(
    df["Payment_Method"].unique().tolist()
)

selected_payments = st.sidebar.multiselect(
    "💳 Payment Method",
    payment_options,
    default=payment_options
)


# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = df.copy()


# Date filter

if len(date_range) == 2:

    start_date = pd.to_datetime(
        date_range[0]
    )

    end_date = (
        pd.to_datetime(date_range[1])
        + pd.Timedelta(days=1)
        - pd.Timedelta(seconds=1)
    )

    filtered_df = filtered_df[
        (filtered_df["Date"] >= start_date)
        &
        (filtered_df["Date"] <= end_date)
    ]


# Department

if selected_departments:

    filtered_df = filtered_df[
        filtered_df["Department"].isin(
            selected_departments
        )
    ]


# Status

if selected_statuses:

    filtered_df = filtered_df[
        filtered_df["Status"].isin(
            selected_statuses
        )
    ]


# Service

if selected_services:

    filtered_df = filtered_df[
        filtered_df["Service"].isin(
            selected_services
        )
    ]


# Payment

if selected_payments:

    filtered_df = filtered_df[
        filtered_df["Payment_Method"].isin(
            selected_payments
        )
    ]


# ============================================================
# EMPTY DATA CHECK
# ============================================================

if filtered_df.empty:

    st.warning(
        "⚠️ No data matches the selected filters."
    )

    st.stop()


# ============================================================
# KPI CALCULATIONS
# ============================================================

total_appointments = len(filtered_df)


completed_appointments = len(
    filtered_df[
        filtered_df["Status"].str.lower()
        == "completed"
    ]
)


cancelled_appointments = len(
    filtered_df[
        filtered_df["Status"].str.lower()
        == "cancelled"
    ]
)


no_show_appointments = len(
    filtered_df[
        filtered_df["Status"].str.lower()
        == "no-show"
    ]
)


if total_appointments > 0:

    completion_rate = (
        completed_appointments
        /
        total_appointments
    ) * 100

else:

    completion_rate = 0


total_revenue = filtered_df[
    "Final_Price"
].sum()


average_revenue = (
    total_revenue
    /
    total_appointments
    if total_appointments > 0
    else 0
)


# ============================================================
# LASER DATA
# ============================================================

laser_df = filtered_df[
    filtered_df["Department"]
    .str.lower()
    .str.contains("laser", na=False)
]


laser_appointments = len(laser_df)


# Number of days in selected date range

if len(date_range) == 2:

    selected_days = (
        pd.to_datetime(date_range[1])
        -
        pd.to_datetime(date_range[0])
    ).days + 1

else:

    selected_days = 1


# Laser capacity

laser_capacity = (
    selected_days
    *
    LASER_CAPACITY_PER_DAY
)


if laser_capacity > 0:

    laser_utilization = (
        laser_appointments
        /
        laser_capacity
    ) * 100

else:

    laser_utilization = 0


# ============================================================
# DASHBOARD KPIs
# ============================================================

st.subheader("📊 Key Performance Indicators")


kpi1, kpi2, kpi3, kpi4 = st.columns(4)


with kpi1:

    st.metric(
        "TOTAL APPOINTMENTS",
        f"{total_appointments:,}"
    )


with kpi2:

    st.metric(
        "COMPLETION RATE",
        f"{completion_rate:.1f}%"
    )


with kpi3:

    st.metric(
        "TOTAL REVENUE",
        f"{total_revenue:,.2f} JD"
    )


with kpi4:

    st.metric(
        "LASER UTILIZATION",
        f"{laser_utilization:.1f}%"
    )


# ============================================================
# SECOND KPI ROW
# ============================================================

st.subheader("📌 Additional KPIs")


k1, k2, k3, k4 = st.columns(4)


with k1:

    st.metric(
        "COMPLETED",
        f"{completed_appointments:,}"
    )


with k2:

    st.metric(
        "CANCELLED",
        f"{cancelled_appointments:,}"
    )


with k3:

    st.metric(
        "NO-SHOW",
        f"{no_show_appointments:,}"
    )


with k4:

    st.metric(
        "AVERAGE REVENUE",
        f"{average_revenue:,.2f} JD"
    )


# ============================================================
# DIVIDER
# ============================================================

st.markdown("---")


# ============================================================
# DEPARTMENT ANALYSIS
# ============================================================

st.subheader("🏢 Department Performance")


col1, col2 = st.columns(2)


# ------------------------------------------------------------
# Appointments by Department
# ------------------------------------------------------------

with col1:

    department_data = (
        filtered_df
        .groupby("Department")
        .size()
        .reset_index(
            name="Appointments"
        )
        .sort_values(
            "Appointments",
            ascending=False
        )
    )

    fig_department = px.bar(
        department_data,
        x="Department",
        y="Appointments",
        title="Appointments by Department",
        text="Appointments"
    )

    fig_department.update_layout(
        height=400,
        xaxis_title="Department",
        yaxis_title="Appointments"
    )

    st.plotly_chart(
        fig_department,
        use_container_width=True
    )


# ------------------------------------------------------------
# Revenue by Department
# ------------------------------------------------------------

with col2:

    revenue_department = (
        filtered_df
        .groupby("Department")["Final_Price"]
        .sum()
        .reset_index()
        .sort_values(
            "Final_Price",
            ascending=False
        )
    )

    fig_revenue_department = px.bar(
        revenue_department,
        x="Department",
        y="Final_Price",
        title="Revenue by Department",
        text_auto=".2f"
    )

    fig_revenue_department.update_layout(
        height=400,
        xaxis_title="Department",
        yaxis_title="Revenue (JD)"
    )

    st.plotly_chart(
        fig_revenue_department,
        use_container_width=True
    )


# ============================================================
# DAILY ANALYSIS
# ============================================================

st.subheader("📅 Daily Performance")


daily_data = (
    filtered_df
    .groupby("Date")
    .agg(
        Appointments=("Appointment_ID", "count"),
        Revenue=("Final_Price", "sum")
    )
    .reset_index()
    .sort_values("Date")
)


col1, col2 = st.columns(2)


# Daily appointments

with col1:

    fig_daily = px.line(
        daily_data,
        x="Date",
        y="Appointments",
        markers=True,
        title="Daily Appointments"
    )

    fig_daily.update_layout(
        height=400,
        xaxis_title="Date",
        yaxis_title="Appointments"
    )

    st.plotly_chart(
        fig_daily,
        use_container_width=True
    )


# Daily revenue

with col2:

    fig_daily_revenue = px.line(
        daily_data,
        x="Date",
        y="Revenue",
        markers=True,
        title="Daily Revenue"
    )

    fig_daily_revenue.update_layout(
        height=400,
        xaxis_title="Date",
        yaxis_title="Revenue (JD)"
    )

    st.plotly_chart(
        fig_daily_revenue,
        use_container_width=True
    )


# ============================================================
# SERVICE ANALYSIS
# ============================================================

st.subheader("💆 Service Performance")


service_data = (
    filtered_df
    .groupby("Service")
    .agg(
        Appointments=("Appointment_ID", "count"),
        Revenue=("Final_Price", "sum")
    )
    .reset_index()
    .sort_values(
        "Appointments",
        ascending=False
    )
)


col1, col2 = st.columns(2)


with col1:

    top_services = service_data.head(10)

    fig_services = px.bar(
        top_services,
        x="Appointments",
        y="Service",
        orientation="h",
        title="Top 10 Services by Appointments",
        text="Appointments"
    )

    fig_services.update_layout(
        height=500,
        yaxis=dict(
            categoryorder="total ascending"
        )
    )

    st.plotly_chart(
        fig_services,
        use_container_width=True
    )


with col2:

    top_service_revenue = (
        service_data
        .sort_values(
            "Revenue",
            ascending=False
        )
        .head(10)
    )

    fig_service_revenue = px.bar(
        top_service_revenue,
        x="Revenue",
        y="Service",
        orientation="h",
        title="Top 10 Services by Revenue",
        text_auto=".2f"
    )

    fig_service_revenue.update_layout(
        height=500,
        yaxis=dict(
            categoryorder="total ascending"
        )
    )

    st.plotly_chart(
        fig_service_revenue,
        use_container_width=True
    )


# ============================================================
# STATUS ANALYSIS
# ============================================================

st.subheader("📌 Appointment Status")


status_data = (
    filtered_df
    .groupby("Status")
    .size()
    .reset_index(
        name="Appointments"
    )
)


col1, col2 = st.columns(2)


with col1:

    fig_status = px.pie(
        status_data,
        names="Status",
        values="Appointments",
        hole=0.45,
        title="Appointment Status Distribution"
    )

    fig_status.update_layout(
        height=420
    )

    st.plotly_chart(
        fig_status,
        use_container_width=True
    )


with col2:

    fig_status_bar = px.bar(
        status_data,
        x="Status",
        y="Appointments",
        text="Appointments",
        title="Appointments by Status"
    )

    fig_status_bar.update_layout(
        height=420
    )

    st.plotly_chart(
        fig_status_bar,
        use_container_width=True
    )


# ============================================================
# LASER ANALYSIS
# ============================================================

st.markdown("---")

st.subheader("🔬 Laser Department Analysis")


if not laser_df.empty:

    laser_col1, laser_col2, laser_col3, laser_col4 = st.columns(4)


    with laser_col1:

        st.metric(
            "Laser Appointments",
            f"{laser_appointments:,}"
        )


    with laser_col2:

        st.metric(
            "Laser Capacity",
            f"{laser_capacity:,}"
        )


    with laser_col3:

        st.metric(
            "Utilization",
            f"{laser_utilization:.1f}%"
        )


    with laser_col4:

        laser_revenue = laser_df[
            "Final_Price"
        ].sum()

        st.metric(
            "Laser Revenue",
            f"{laser_revenue:,.2f} JD"
        )


    # --------------------------------------------------------
    # Device Performance
    # --------------------------------------------------------

    device_data = (
        laser_df[
            laser_df["Device_ID"] != "No Device"
        ]
        .groupby("Device_ID")
        .agg(
            Appointments=("Appointment_ID", "count"),
            Revenue=("Final_Price", "sum")
        )
        .reset_index()
        .sort_values(
            "Appointments",
            ascending=False
        )
    )


    col1, col2 = st.columns(2)


    # Device appointments

    with col1:

        fig_device = px.bar(
            device_data,
            x="Device_ID",
            y="Appointments",
            text="Appointments",
            title="Appointments by Laser Device"
        )

        fig_device.update_layout(
            height=450,
            xaxis_title="Device",
            yaxis_title="Appointments"
        )

        st.plotly_chart(
            fig_device,
            use_container_width=True
        )


    # Device revenue

    with col2:

        fig_device_revenue = px.bar(
            device_data,
            x="Device_ID",
            y="Revenue",
            text_auto=".2f",
            title="Revenue by Laser Device"
        )

        fig_device_revenue.update_layout(
            height=450,
            xaxis_title="Device",
            yaxis_title="Revenue (JD)"
        )

        st.plotly_chart(
            fig_device_revenue,
            use_container_width=True
        )


    # --------------------------------------------------------
    # Device Utilization
    # --------------------------------------------------------

    device_data["Capacity"] = (
        selected_days
        *
        LASER_HOURS_PER_DEVICE
    )


    device_data["Utilization"] = (
        device_data["Appointments"]
        /
        device_data["Capacity"]
        *
        100
    )


    st.write("### 📊 Laser Device Utilization")


    device_display = device_data.copy()

    device_display["Utilization"] = (
        device_display["Utilization"]
        .round(1)
        .astype(str)
        + "%"
    )


    st.dataframe(
        device_display,
        use_container_width=True,
        hide_index=True
    )


    # --------------------------------------------------------
    # Laser Peak Hours
    # --------------------------------------------------------

    st.write("### ⏰ Laser Appointments by Hour")


    laser_hour_data = (
        laser_df
        .groupby("Time")
        .size()
        .reset_index(
            name="Appointments"
        )
        .sort_values("Time")
    )


    fig_laser_hours = px.bar(
        laser_hour_data,
        x="Time",
        y="Appointments",
        text="Appointments",
        title="Laser Demand by Hour"
    )


    fig_laser_hours.update_layout(
        height=450,
        xaxis_title="Time",
        yaxis_title="Appointments"
    )


    st.plotly_chart(
        fig_laser_hours,
        use_container_width=True
    )


else:

    st.info(
        "No Laser appointments found with the current filters."
    )


# ============================================================
# STAFF ANALYSIS
# ============================================================

st.markdown("---")

st.subheader("👩‍⚕️ Staff Performance")


staff_data = (
    filtered_df
    .groupby("Staff")
    .agg(
        Appointments=("Appointment_ID", "count"),
        Revenue=("Final_Price", "sum")
    )
    .reset_index()
    .sort_values(
        "Appointments",
        ascending=False
    )
)


col1, col2 = st.columns(2)


with col1:

    fig_staff = px.bar(
        staff_data.head(10),
        x="Appointments",
        y="Staff",
        orientation="h",
        text="Appointments",
        title="Top Staff by Appointments"
    )

    fig_staff.update_layout(
        height=500,
        yaxis=dict(
            categoryorder="total ascending"
        )
    )

    st.plotly_chart(
        fig_staff,
        use_container_width=True
    )


with col2:

    staff_revenue = (
        staff_data
        .sort_values(
            "Revenue",
            ascending=False
        )
        .head(10)
    )

    fig_staff_revenue = px.bar(
        staff_revenue,
        x="Revenue",
        y="Staff",
        orientation="h",
        text_auto=".2f",
        title="Top Staff by Revenue"
    )

    fig_staff_revenue.update_layout(
        height=500,
        yaxis=dict(
            categoryorder="total ascending"
        )
    )

    st.plotly_chart(
        fig_staff_revenue,
        use_container_width=True
    )


# ============================================================
# PACKAGE ANALYSIS
# ============================================================

st.subheader("📦 Package Analysis")


package_data = (
    filtered_df
    .groupby("Package_Type")
    .agg(
        Appointments=("Appointment_ID", "count"),
        Revenue=("Final_Price", "sum")
    )
    .reset_index()
)


col1, col2 = st.columns(2)


with col1:

    fig_package = px.pie(
        package_data,
        names="Package_Type",
        values="Appointments",
        hole=0.45,
        title="Single vs Package Appointments"
    )

    st.plotly_chart(
        fig_package,
        use_container_width=True
    )


with col2:

    fig_package_revenue = px.bar(
        package_data,
        x="Package_Type",
        y="Revenue",
        text_auto=".2f",
        title="Revenue by Package Type"
    )

    st.plotly_chart(
        fig_package_revenue,
        use_container_width=True
    )


# ============================================================
# SESSION ANALYSIS
# ============================================================

st.subheader("🔢 Laser Package Sessions")


session_data = (
    laser_df
    .groupby("Session_Number")
    .size()
    .reset_index(
        name="Appointments"
    )
    .sort_values("Session_Number")
)


if not session_data.empty:

    fig_sessions = px.bar(
        session_data,
        x="Session_Number",
        y="Appointments",
        text="Appointments",
        title="Laser Appointments by Session Number"
    )

    fig_sessions.update_layout(
        height=400,
        xaxis_title="Session Number",
        yaxis_title="Appointments"
    )

    st.plotly_chart(
        fig_sessions,
        use_container_width=True
    )


# ============================================================
# PAYMENT METHOD ANALYSIS
# ============================================================

st.subheader("💳 Payment Methods")


payment_data = (
    filtered_df
    .groupby("Payment_Method")
    .agg(
        Transactions=("Appointment_ID", "count"),
        Revenue=("Final_Price", "sum")
    )
    .reset_index()
    .sort_values(
        "Revenue",
        ascending=False
    )
)


col1, col2 = st.columns(2)


with col1:

    fig_payment = px.pie(
        payment_data,
        names="Payment_Method",
        values="Transactions",
        hole=0.45,
        title="Payment Method Distribution"
    )

    st.plotly_chart(
        fig_payment,
        use_container_width=True
    )


with col2:

    fig_payment_revenue = px.bar(
        payment_data,
        x="Payment_Method",
        y="Revenue",
        text_auto=".2f",
        title="Revenue by Payment Method"
    )

    st.plotly_chart(
        fig_payment_revenue,
        use_container_width=True
    )


# ============================================================
# MONTHLY ANALYSIS
# ============================================================

st.subheader("📆 Monthly Performance")


monthly_data = (
    filtered_df
    .groupby(
        ["Year", "Month", "Month_Name"]
    )
    .agg(
        Appointments=("Appointment_ID", "count"),
        Revenue=("Final_Price", "sum")
    )
    .reset_index()
    .sort_values(
        ["Year", "Month"]
    )
)


if not monthly_data.empty:

    monthly_data["Period"] = (
        monthly_data["Year"]
        .astype(str)
        + "-"
        + monthly_data["Month_Name"]
    )


    col1, col2 = st.columns(2)


    with col1:

        fig_monthly = px.bar(
            monthly_data,
            x="Period",
            y="Appointments",
            text="Appointments",
            title="Monthly Appointments"
        )

        st.plotly_chart(
            fig_monthly,
            use_container_width=True
        )


    with col2:

        fig_monthly_revenue = px.bar(
            monthly_data,
            x="Period",
            y="Revenue",
            text_auto=".2f",
            title="Monthly Revenue"
        )

        st.plotly_chart(
            fig_monthly_revenue,
            use_container_width=True
        )


# ============================================================
# DATA TABLE
# ============================================================

st.markdown("---")

st.subheader("📋 Filtered Appointments Data")


st.write(
    f"Showing **{len(filtered_df):,}** appointments"
)


display_columns = [
    "Appointment_ID",
    "Date",
    "Day",
    "Time",
    "Department",
    "Service",
    "Staff",
    "Device_ID",
    "Package_Type",
    "Session_Number",
    "Price",
    "Discount",
    "Final_Price",
    "Payment_Method",
    "Status"
]


st.dataframe(
    filtered_df[display_columns],
    use_container_width=True,
    height=500,
    hide_index=True
)


# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================

st.subheader("⬇️ Export Data")


csv_data = filtered_df[
    display_columns
].to_csv(
    index=False
).encode("utf-8")


st.download_button(
    label="📥 Download Filtered Data CSV",
    data=csv_data,
    file_name="ILANG_Filtered_Appointments.csv",
    mime="text/csv"
)


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(
    "ILANG Clinic | Data Analysis Dashboard | Built with Python & Streamlit"
)