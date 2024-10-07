# app.py

# Import necessary libraries
import streamlit as st
import pandas as pd
import numpy as np
import seaborn as sns
from datetime import datetime, timedelta
import io
import time
import plotly.express as px
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image as ReportLabImage,
    Table,
    TableStyle,
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors

# Configure Streamlit page
st.set_page_config(
    page_title="TTU Purchase Orders Log",
    layout="wide",  # 'wide' layout to accommodate sidebar elements
    initial_sidebar_state="expanded",
)

# Apply Seaborn style
sns.set_style("whitegrid")

# Inject CSS for styling
st.markdown(
    """
    <style>
    /* Overall background */
    body {
        background-color: #f5f5f5;
    }
    .title-container {
        display: flex;
        align-items: center;
        justify-content: flex-start;
    }
    .title {
        font-size: 36px;
        font-weight: bold;
        color: #1f4e79;
        text-align: right;
        text-transform: uppercase;
        text-shadow: 1px 1px 2px #aaa;
        margin-right: 20px;
    }
    /* Style for index cards */
    .card {
        background-color: #ffffff;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 2px 4px 12px rgba(0,0,0,0.1);
        min-height: 100px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        margin: 5px;
        word-wrap: break-word;
    }
    .card h3 {
        margin-bottom: 5px;
        text-align: center;
        font-size: 16px;
        color: #333333;
    }
    .card p {
        margin: 0;
        font-size: 20px;
        text-align: center;
        font-weight: bold;
        white-space: pre-wrap;
        color: #1f4e79;
    }
    /* Style for percentage value in index cards */
    .percentage-value {
        font-size: 36px;
        color: #ff5722;
        font-weight: bold;
        margin: 0;
    }
    /* Custom colors for Streamlit elements */
    .stButton > button {
        background-color: #1f4e79;
        color: white;
    }
    .metrics-section h2 {
        text-align: center;
        color: #1f4e79;
        margin-bottom: 10px;
    }
    /* Style for tables */
    table {
        margin-left: auto;
        margin-right: auto;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Title and Logo
def display_logo_title():
    st.markdown('<h1 class="title">TTU Purchase Orders Log</h1>', unsafe_allow_html=True)

# Callback function to reset date filters
def reset_dates():
    # Ensure that initial dates are stored in session state
    if 'initial_order_min_date' in st.session_state and 'initial_order_max_date' in st.session_state:
        st.session_state.order_start_date = st.session_state.initial_order_min_date
        st.session_state.order_end_date = st.session_state.initial_order_max_date
    if 'initial_request_min_date' in st.session_state and 'initial_request_max_date' in st.session_state:
        st.session_state.request_start_date = st.session_state.initial_request_min_date
        st.session_state.request_end_date = st.session_state.initial_request_max_date

# Callback function to reset selectbox filters
def reset_filters():
    st.session_state.selected_purchase_account = 'All'
    st.session_state.selected_requisitioner = 'All'

# Caching the data loading function
@st.cache_data
def load_and_process_data(uploaded_file):
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl')

        # Rename 'Acct' to 'Purchase Account'
        df.rename(columns={'Acct': 'Purchase Account'}, inplace=True)

        # Convert date columns to datetime and normalize time
        date_columns = [col for col in df.columns if 'date' in col.lower()]
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.normalize()

        # Verify 'OrderDate' column
        if 'OrderDate' in df.columns:
            df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce').dt.normalize()
            df = df.dropna(subset=['OrderDate'])
        else:
            st.error("'OrderDate' column is missing.")
            return pd.DataFrame(), [], []

        # Filter rows where 'OrderDate' >= 2022-01-01
        df_filtered = df[df['OrderDate'] >= pd.to_datetime('2022-01-01')].copy()

        # Handle missing values in critical columns
        critical_columns = ['OrderDate', 'PONumber', 'Total']
        df_filtered.dropna(subset=critical_columns, inplace=True)

        # Remove duplicates
        df_filtered.drop_duplicates(inplace=True)

        # Ensure numerical columns are correct
        numerical_columns = ['Total', 'Amt', 'QtyOrdered', 'QtyRemaining']
        for col in numerical_columns:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)

        # Format 'Purchase Account'
        if 'Purchase Account' in df_filtered.columns:
            df_filtered = df_filtered.dropna(subset=['Purchase Account'])
            df_filtered['Purchase Account'] = df_filtered['Purchase Account'].astype(int).astype(str).str.zfill(8)
            df_filtered['Purchase Account'] = df_filtered['Purchase Account'].str.replace(
                r'(\d{4})(\d{4})', r'\1-\2', regex=True
            )
        else:
            st.error("'Purchase Account' column is missing.")

        # Rename 'QtyRemaining'
        df_filtered.rename(columns={'QtyRemaining': 'qty on order/backordered'}, inplace=True)

        return df_filtered, [], date_columns
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame(), [], []

# Map POStatus codes
def map_po_status(df):
    po_status_mapping = {'NN': 'NEW', 'AN': 'OPEN', 'F': 'RECEIVED', 'BN': 'BACKORDERED'}
    if 'POStatus' in df.columns:
        df['POStatus'] = df['POStatus'].map(po_status_mapping).fillna(df['POStatus'])
    else:
        st.write("'POStatus' column is missing.")
    return df

# Create index cards
def display_index_cards(metrics):
    for metric_name, metric_info in metrics.items():
        bg_color = metric_info.get('bg_color', '#FFFFFF')
        value = metric_info.get(metric_name, 'N/A')
        value_formatted = value.replace("<br>", "<br/>")
        st.markdown(
            f"""
            <div class="card" style='background-color: {bg_color}; width: 100%;'>
                <h3>{metric_name}</h3>
                <p>{value_formatted}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

# Calculate outliers using IQR method
def filter_outliers(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR
    df_filtered = df[(df[column] >= lower_bound) & (df[column] <= upper_bound)]
    return df_filtered

# Main application logic
def main():
    display_logo_title()

    with st.sidebar:
        # Display the logo
        st.image("TTU_LOGO.jpg", width=150)
        uploaded_file = st.file_uploader("Choose an Excel file here", type=["xlsx"])

    if uploaded_file:
        processing_start_time = time.time()
        df_filtered, _, date_columns = load_and_process_data(uploaded_file)

        if not df_filtered.empty:
            df_filtered = map_po_status(df_filtered)

            # Store the initial min and max dates
            initial_order_min_date = df_filtered['OrderDate'].min().date()
            initial_order_max_date = df_filtered['OrderDate'].max().date()
            if 'RequestDate' in df_filtered.columns:
                initial_request_min_date = df_filtered['RequestDate'].min().date()
                initial_request_max_date = df_filtered['RequestDate'].max().date()
            else:
                initial_request_min_date = None
                initial_request_max_date = None

            # Store initial dates in session state
            st.session_state.initial_order_min_date = initial_order_min_date
            st.session_state.initial_order_max_date = initial_order_max_date
            if initial_request_min_date and initial_request_max_date:
                st.session_state.initial_request_min_date = initial_request_min_date
                st.session_state.initial_request_max_date = initial_request_max_date

            # Initialize session state for date inputs
            if 'order_start_date' not in st.session_state:
                st.session_state.order_start_date = initial_order_min_date
            if 'order_end_date' not in st.session_state:
                st.session_state.order_end_date = initial_order_max_date
            if 'request_start_date' not in st.session_state and initial_request_min_date is not None:
                st.session_state.request_start_date = initial_request_min_date
            if 'request_end_date' not in st.session_state and initial_request_max_date is not None:
                st.session_state.request_end_date = initial_request_max_date

            with st.sidebar:
                # Filter by OrderDate
                st.header("Filter by OrderDate")

                # Get current min and max dates from df_filtered
                order_min_date = df_filtered['OrderDate'].min().date()
                order_max_date = df_filtered['OrderDate'].max().date()

                # Ensure default values are within min and max
                order_start_default = st.session_state.order_start_date
                order_end_default = st.session_state.order_end_date

                if order_start_default < order_min_date or order_start_default > order_max_date:
                    order_start_default = order_min_date
                if order_end_default < order_min_date or order_end_default > order_max_date:
                    order_end_default = order_max_date

                order_start_date = st.date_input(
                    "Order Start Date",
                    order_start_default,
                    min_value=order_min_date,
                    max_value=order_max_date,
                    key='order_start_date_input'
                )
                order_end_date = st.date_input(
                    "Order End Date",
                    order_end_default,
                    min_value=order_min_date,
                    max_value=order_max_date,
                    key='order_end_date_input'
                )

                # Update session state
                st.session_state.order_start_date = order_start_date
                st.session_state.order_end_date = order_end_date

                if order_start_date > order_end_date:
                    st.error("Error: Order End Date must fall after Order Start Date.")

                # Filter df_filtered based on OrderDate range
                df_filtered = df_filtered[
                    (df_filtered['OrderDate'] >= pd.to_datetime(order_start_date)) &
                    (df_filtered['OrderDate'] <= pd.to_datetime(order_end_date))
                    ].copy()

                # After applying OrderDate filter, update the min and max for RequestDate
                if 'RequestDate' in df_filtered.columns and not df_filtered.empty:
                    request_min_date = df_filtered['RequestDate'].min().date()
                    request_max_date = df_filtered['RequestDate'].max().date()

                    # Add a checkbox for enabling the RequestDate filter
                    filter_request_date = st.checkbox("Filter by RequestDate", value=False)

                    # Show date pickers only if checkbox is selected
                    if filter_request_date:
                        st.header("Filter by RequestDate")

                        # Ensure default values are within min and max
                        request_start_default = st.session_state.request_start_date
                        request_end_default = st.session_state.request_end_date

                        if request_start_default < request_min_date or request_start_default > request_max_date:
                            request_start_default = request_min_date
                        if request_end_default < request_min_date or request_end_default > request_max_date:
                            request_end_default = request_max_date

                        request_start_date = st.date_input(
                            "Request Start Date",
                            request_start_default,
                            min_value=request_min_date,
                            max_value=request_max_date,
                            key='request_start_date_input'
                        )
                        request_end_date = st.date_input(
                            "Request End Date",
                            request_end_default,
                            min_value=request_min_date,
                            max_value=request_max_date,
                            key='request_end_date_input'
                        )

                        # Update session state
                        st.session_state.request_start_date = request_start_date
                        st.session_state.request_end_date = request_end_date

                        if request_start_date > request_end_date:
                            st.error("Error: Request End Date must fall after Request Start Date.")

                        # Filter df_filtered based on RequestDate range
                        df_filtered = df_filtered[
                            (df_filtered['RequestDate'] >= pd.to_datetime(request_start_date)) &
                            (df_filtered['RequestDate'] <= pd.to_datetime(request_end_date))
                            ].copy()
                elif 'RequestDate' in df_filtered.columns:
                    st.error("No data available for the selected OrderDate range.")
                else:
                    st.error("'RequestDate' column is missing in the data.")

                # Reset All Dates Button
                st.button("Reset All Dates", on_click=reset_dates)

                # Filter by Purchase Account
                if 'Purchase Account' in df_filtered.columns:
                    purchase_accounts = sorted(df_filtered['Purchase Account'].dropna().unique().tolist())
                    options_purchase = ['All'] + purchase_accounts

                    # Initialize session state for purchase account
                    if 'selected_purchase_account' not in st.session_state:
                        st.session_state.selected_purchase_account = 'All'

                    # Determine the index for the selectbox
                    if st.session_state.selected_purchase_account in options_purchase:
                        index_purchase = options_purchase.index(st.session_state.selected_purchase_account)
                    else:
                        index_purchase = 0  # Default to 'All' if not found

                    selected_purchase_account = st.selectbox(
                        'Select Purchase Account',
                        options=options_purchase,
                        index=index_purchase,
                        key='selected_purchase_account'
                    )
                else:
                    selected_purchase_account = 'All'
                    st.error("'Purchase Account' column is missing.")

                # Filter by Requisitioner
                if 'Requisitioner' in df_filtered.columns:
                    requisitioners = sorted(df_filtered['Requisitioner'].dropna().unique().tolist())
                    options_requisitioner = ['All'] + requisitioners

                    # Initialize session state for requisitioner
                    if 'selected_requisitioner' not in st.session_state:
                        st.session_state.selected_requisitioner = 'All'

                    # Determine the index for the selectbox
                    if st.session_state.selected_requisitioner in options_requisitioner:
                        index_requisitioner = options_requisitioner.index(st.session_state.selected_requisitioner)
                    else:
                        index_requisitioner = 0  # Default to 'All' if not found

                    selected_requisitioner = st.selectbox(
                        'Select Requisitioner',
                        options=options_requisitioner,
                        index=index_requisitioner,
                        key='selected_requisitioner'
                    )
                else:
                    selected_requisitioner = 'All'
                    st.error("'Requisitioner' column is missing.")

                # Reset All Filters Button
                st.button("Reset All", on_click=reset_filters)

                # Apply filters
                if selected_requisitioner != 'All':
                    df_filtered = df_filtered[df_filtered['Requisitioner'] == selected_requisitioner]

                if selected_purchase_account != 'All':
                    df_filtered = df_filtered[df_filtered['Purchase Account'] == selected_purchase_account]

                # Handle case when no results are found after filtering
                if df_filtered.empty:
                    st.warning("No results found for the selected filters. Please adjust the filters and try again.")
                    return  # Stop further processing

                # KPI Title
                kpi_title = f"Key Performance Indicators"

                st.header(kpi_title)

                # Calculate metrics
                metrics = {}

                # Total Open Orders Amt
                if 'Amt' in df_filtered.columns and 'POStatus' in df_filtered.columns:
                    total_open_orders_amt = df_filtered[df_filtered['POStatus'] == 'OPEN']['Amt'].sum()
                    total_open_orders_amt_formatted = f"${total_open_orders_amt:,.2f}"
                    metrics['Total Open Orders Amt'] = {
                        'Total Open Orders Amt': total_open_orders_amt_formatted,
                        'bg_color': '#90CAF9',
                    }
                else:
                    metrics['Total Open Orders Amt'] = {'Total Open Orders Amt': "$0.00", 'bg_color': '#90CAF9'}

                # Total Orders Placed
                if 'PONumber' in df_filtered.columns:
                    total_orders_placed = df_filtered['PONumber'].nunique()
                    metrics['Total Orders Placed'] = {
                        'Total Orders Placed': f"{total_orders_placed}",
                        'bg_color': '#FFCDD2',
                    }

                # Total Lines Ordered
                total_lines_ordered = df_filtered.shape[0]
                metrics['Total Lines Ordered'] = {
                    'Total Lines Ordered': f"{total_lines_ordered}",
                    'bg_color': '#C8E6C9',
                }

                # Most Expensive Order
                if {'Total', 'PONumber', 'VendorName', 'Requisitioner'}.issubset(df_filtered.columns):
                    max_total_row = df_filtered.loc[df_filtered['Total'].idxmax()]
                    max_total_formatted = f"${max_total_row['Total']:,.2f}"
                    most_expensive_order_info = (
                        f"PO Number: {max_total_row['PONumber']}<br/>"
                        f"Vendor: {max_total_row['VendorName']}<br/>"
                        f"Requisitioner: {max_total_row['Requisitioner']}<br/>"
                        f"Total: {max_total_formatted}"
                    )
                    metrics['Most Expensive Order'] = {
                        'Most Expensive Order': most_expensive_order_info,
                        'bg_color': '#FFD54F',
                    }
                else:
                    metrics['Most Expensive Order'] = {'Most Expensive Order': "N/A", 'bg_color': '#FFD54F'}

                # Display Key Metrics
                display_index_cards(metrics)

            # Main Content Area

            # Under the title, if a requisitioner is selected, display last order details index card
            if selected_requisitioner != 'All':
                # Get last order details for the selected requisitioner
                last_order = df_filtered.sort_values(
                    by='OrderDate', ascending=False
                ).head(1)
                if not last_order.empty:
                    last_order_row = last_order.iloc[0]
                    last_order_info = (
                        f"Order Date: {last_order_row['OrderDate'].date()}<br/>"
                        f"PO Number: {last_order_row['PONumber']}<br/>"
                        f"Vendor: {last_order_row['VendorName']}<br/>"
                        f"Total: ${last_order_row['Total']:,.2f}<br/>"
                        f"Status: {last_order_row['POStatus']}"
                    )
                    st.markdown(
                        f"""
                        <div class="card" style='background-color: #BBDEFB; width: 100%;'>
                            <h3>Last Order for {selected_requisitioner}</h3>
                            <p>{last_order_info}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            # Removed 'Detailed Analyses' subheader as per instruction

            pdf_elements = []

            # Only show On-Time Delivery Performance if no requisitioner is selected
            if selected_requisitioner == 'All':
                # On-Time Delivery Performance
                if 'RecDate' in df_filtered.columns and 'RequestDate' in df_filtered.columns:
                    # Remove time from 'RecDate' and 'RequestDate' columns
                    df_filtered['RecDate'] = pd.to_datetime(df_filtered['RecDate']).dt.date
                    df_filtered['RequestDate'] = pd.to_datetime(df_filtered['RequestDate']).dt.date

                    on_time_pos = df_filtered[df_filtered['RecDate'] <= df_filtered['RequestDate']]
                    late_pos = df_filtered[df_filtered['RecDate'] > df_filtered['RequestDate']]
                    on_time_count = on_time_pos['PONumber'].nunique()
                    late_count = late_pos['PONumber'].nunique()
                    total_pos = on_time_count + late_count
                    if total_pos > 0:
                        on_time_percentage = (on_time_count / total_pos) * 100
                    else:
                        on_time_percentage = 0

                    on_time_percentage_formatted = f"{on_time_percentage:.2f}%"
                    late_percentage = 100 - on_time_percentage
                    late_percentage_formatted = f"{late_percentage:.2f}%"

                    # Display index card with On-Time Delivery Percentage
                    st.markdown(
                        f"""
                        <div class="card" style='background-color: #AED581; width: 100%;'>
                            <h3>On Time Delivery</h3>
                            <p class="percentage-value">{on_time_percentage_formatted}</p>
                            <p>On-Time: {on_time_percentage_formatted} | Late: {late_percentage_formatted}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

                    delivery_data = pd.DataFrame({
                        'Status': ['On-Time', 'Late'],
                        'Count': [on_time_count, late_count],
                    })

                    # Instead of a pie chart, use a bar chart
                    col1, col2 = st.columns(2)
                    with col1:
                        fig_delivery = px.bar(
                            delivery_data,
                            x='Status',
                            y='Count',
                            title='On-Time Delivery Performance',
                            color='Status',
                            color_discrete_map={'On-Time': 'green', 'Late': 'red'},
                            text='Count',
                        )
                        fig_delivery.update_layout(showlegend=False)
                        st.plotly_chart(fig_delivery, use_container_width=True)
                        pdf_elements.append(("On-Time Delivery Performance", delivery_data, fig_delivery.to_image(format="png")))

                    # List of Late Purchase Orders
                    with col2:
                        st.markdown("#### List of Late Purchase Orders by Request Date")
                        if not late_pos.empty:
                            late_pos['Days Late'] = (pd.to_datetime(late_pos['RecDate']) - pd.to_datetime(late_pos['RequestDate'])).dt.days
                            late_pos_display = late_pos[['OrderDate', 'PONumber', 'Total', 'Days Late', 'Requisitioner']].copy()
                            late_pos_display.rename(columns={'Total': 'Total Amt'}, inplace=True)
                            late_pos_display['Total Amt'] = late_pos_display['Total Amt'].apply(lambda x: f"${x:,.2f}")

                            # Remove time from 'OrderDate' column
                            late_pos_display['OrderDate'] = pd.to_datetime(late_pos_display['OrderDate']).dt.date

                            # Reset index and drop it
                            late_pos_display.reset_index(drop=True, inplace=True)

                            st.dataframe(late_pos_display)
                            pdf_elements.append(("List of Late Purchase Orders by Request Date", late_pos_display, None))
                        else:
                            st.write("No late purchase orders found.")
                            pdf_elements.append(("List of Late Purchase Orders by Request Date", pd.DataFrame(), None))
                else:
                    st.error("'RecDate' and/or 'RequestDate' columns are missing.")

            # PO Counts per Requisitioner
            st.markdown("### PO Count per Requisitioner by Order Date")
            po_counts = df_filtered.groupby('Requisitioner')['PONumber'].nunique().reset_index()
            po_counts.rename(columns={'PONumber': 'PO Count'}, inplace=True)
            po_amount = df_filtered.groupby('Requisitioner')['Total'].sum().reset_index()
            po_amount.rename(columns={'Total': 'Total Open PO Amount'}, inplace=True)
            po_counts_with_amount = pd.merge(po_counts, po_amount, on='Requisitioner', how='left')
            po_counts_with_amount['Total Open PO Amount'] = po_counts_with_amount['Total Open PO Amount'].apply(lambda x: f"${x:,.2f}")
            po_numbers = df_filtered.groupby('Requisitioner')['PONumber'].apply(lambda x: ', '.join(x.unique())).reset_index()
            po_numbers.rename(columns={'PONumber': 'PO Numbers'}, inplace=True)
            po_counts_final = pd.merge(po_counts_with_amount, po_numbers, on='Requisitioner', how='left')

            # Remove index and reset it
            po_counts_final.reset_index(drop=True, inplace=True)

            # Make table page-wide
            st.write(po_counts_final)
            pdf_elements.append(("PO Count per Requisitioner by Order Date", po_counts_final, None))

            # Last Orders for the period
            st.markdown("### Last Orders for the period")
            last_orders = df_filtered.sort_values(by='OrderDate', ascending=False)
            if not last_orders.empty:
                last_orders_display = last_orders.copy()
                if 'Total' in last_orders_display.columns:
                    last_orders_display['Total'] = last_orders_display['Total'].apply(lambda x: f"${x:,.2f}")
                    last_orders_display.rename(columns={'Total': 'Open Orders Amt'}, inplace=True)

                # Remove specified columns
                columns_to_remove = ['Responsibility Key', 'Open Lines Amt']
                for col in columns_to_remove:
                    if col in last_orders_display.columns:
                        last_orders_display.drop(columns=[col], inplace=True)

                # Remove time from date columns
                date_columns_in_last_orders = last_orders_display.select_dtypes(include=['datetime64[ns]']).columns
                for col in date_columns_in_last_orders:
                    last_orders_display[col] = last_orders_display[col].dt.date

                # Reset index and drop it
                last_orders_display.reset_index(drop=True, inplace=True)

                st.dataframe(last_orders_display)
                pdf_elements.append(("Last Orders for the period", last_orders_display, None))
            else:
                st.write("No orders found.")
                pdf_elements.append(("Last Orders for the period", pd.DataFrame(), None))

            # Open Orders Amount per Vendor (Only Table)
            st.markdown("### Open Orders Amount per Vendor")
            vendor_amount = df_filtered[df_filtered['POStatus'] == 'OPEN'].groupby('VendorName')['Total'].sum().reset_index()
            vendor_amount_no_outliers = filter_outliers(vendor_amount, 'Total').sort_values(by='Total', ascending=False)
            vendor_amount_no_outliers_display = vendor_amount_no_outliers.copy()
            vendor_amount_no_outliers_display['Total'] = vendor_amount_no_outliers_display['Total'].apply(lambda x: f"${x:,.2f}")

            # Reset index and drop it
            vendor_amount_no_outliers_display.reset_index(drop=True, inplace=True)

            st.dataframe(vendor_amount_no_outliers_display)
            pdf_elements.append(("Open Orders Amount per Vendor", vendor_amount_no_outliers_display, None))

            # Top 5 Vendors by Amount
            st.markdown("### Top 5 Vendors by Amount")
            top_vendors = df_filtered[df_filtered['POStatus'] == 'OPEN'].groupby('VendorName')['Total'].sum().reset_index()
            top_vendors = top_vendors.sort_values(by='Total', ascending=False).head(5)
            fig_top_vendors = px.bar(
                top_vendors,
                x='VendorName',
                y='Total',
                title='Top 5 Vendors by Amount',
                labels={'Total': 'Total Amount ($)', 'VendorName': 'Vendor'},
                color='Total',
                color_continuous_scale=px.colors.sequential.Plasma
            )

            # Arrange chart and table side by side
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(fig_top_vendors, use_container_width=True)
            with col2:
                top_vendors_display = top_vendors.copy()
                top_vendors_display['Total'] = top_vendors_display['Total'].apply(lambda x: f"${x:,.2f}")
                top_vendors_display.reset_index(drop=True, inplace=True)
                st.dataframe(top_vendors_display)
                img_bytes = fig_top_vendors.to_image(format="png", width=800, height=600)
                img_buf = BytesIO(img_bytes)
                pdf_elements.append(("Top 5 Vendors by Amount", top_vendors_display, img_buf))

            # Top Items by QtyOrdered
            st.markdown("### Top Items by QtyOrdered")
            top_items = df_filtered.groupby(['ItemDescription', 'VendorName'])['QtyOrdered'].sum().reset_index()
            top_items_no_outliers = filter_outliers(top_items, 'QtyOrdered').sort_values(by='QtyOrdered', ascending=False).head(10)
            fig_top_items = px.bar(
                top_items_no_outliers,
                x='ItemDescription',
                y='QtyOrdered',
                title='Top Items by QtyOrdered (Filtered)',
                labels={'QtyOrdered': 'Quantity Ordered', 'ItemDescription': 'Item Description'},
                color='QtyOrdered',
                color_continuous_scale='Agsunset'
            )

            # Arrange chart and table side by side
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(fig_top_items, use_container_width=True)
            with col2:
                top_items_no_outliers_display = top_items_no_outliers.copy()
                top_items_no_outliers_display['QtyOrdered'] = top_items_no_outliers_display['QtyOrdered'].apply(lambda x: f"{x:,.0f}")
                top_items_no_outliers_display.reset_index(drop=True, inplace=True)
                st.dataframe(top_items_no_outliers_display)
                img_bytes = fig_top_items.to_image(format="png", width=1000, height=600)
                img_buf = BytesIO(img_bytes)
                pdf_elements.append(("Top Items by QtyOrdered", top_items_no_outliers_display, img_buf))

            # Processing time
            processing_end_time = time.time()
            total_processing_time = processing_end_time - processing_start_time
            st.markdown(
                f"<p style='text-align: center; color: lightgray; font-size: 10px;'>Total processing time: {total_processing_time:.2f} seconds</p>",
                unsafe_allow_html=True,
            )

            # Generate PDF Report
            if st.button("Generate PDF Report"):
                buffer = io.BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter)
                elements = []
                styles = getSampleStyleSheet()

                try:
                    title = f"TTU Purchase Orders Log Report ({order_start_date} to {order_end_date})"
                    elements.append(Paragraph(title, styles['Title']))
                    elements.append(Spacer(1, 12))

                    # Add metrics
                    for metric_name, metric_info in metrics.items():
                        text_content = metric_info.get(metric_name, 'N/A').replace('<br/>', '\n').replace('<br>', '\n')
                        text = f"<b>{metric_name}:</b>\n{text_content}"
                        elements.append(Paragraph(text, styles['Normal']))
                        elements.append(Spacer(1, 12))

                    # Add analyses
                    for title_text, data, img_buf in pdf_elements:
                        elements.append(Paragraph(title_text, styles['Heading2']))
                        elements.append(Spacer(1, 12))
                        if isinstance(data, pd.DataFrame) and not data.empty:
                            # Convert date columns to strings to avoid issues in PDF table
                            date_cols = data.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns
                            for col in date_cols:
                                data[col] = data[col].astype(str)

                            table_data = [list(data.columns)] + data.values.tolist()
                            t = Table(table_data, repeatRows=1)
                            t.setStyle(
                                TableStyle(
                                    [
                                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                                        ('GRID', (0, 0), (-1, -1), 1, colors.black),
                                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                                    ]
                                )
                            )
                            elements.append(t)
                            elements.append(Spacer(1, 12))
                        elif isinstance(data, pd.DataFrame) and data.empty:
                            elements.append(Paragraph("No data available for this analysis.", styles['Normal']))
                            elements.append(Spacer(1, 12))
                        if img_buf and isinstance(img_buf, BytesIO):
                            img_buf.seek(0)
                            img = ReportLabImage(img_buf)
                            img.drawHeight = 4 * inch
                            img.drawWidth = 6 * inch
                            elements.append(img)
                            elements.append(Spacer(1, 12))

                    # Build the PDF
                    doc.build(elements)
                    pdf = buffer.getvalue()
                    buffer.close()

                    # Download button
                    st.download_button(
                        label="Download PDF Report",
                        data=pdf,
                        file_name="TTU_Purchase_Orders_Log_Report.pdf",
                        mime="application/pdf",
                    )
                except Exception as e:
                    st.error(f"An error occurred while generating the PDF: {e}")
        else:
            st.write("Please upload an Excel file to proceed.")
    else:
        st.write("Please upload an Excel file to proceed.")

if __name__ == "__main__":
    main()
