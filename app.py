import streamlit as st
import pandas as pd
import io
from datetime import datetime, date
import calendar

# ==================== PAGE CONFIG ==================== #
st.set_page_config(
    page_title="Australian Tax Depreciation Calculator",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("💰 Australian Tax Depreciation Calculator")
st.write("Calculate depreciation deductions for Australian tax purposes")

# ==================== SIDEBAR ==================== #
st.sidebar.title("⚙️ Settings")
calculation_mode = st.sidebar.radio(
    "Select Calculation Mode:",
    ["Single Asset", "Asset Register"],
    help="Single for one asset, Register for multiple assets"
)

# ==================== HELPER FUNCTIONS ==================== #

def calculate_depreciation(asset_cost, purchase_date, depreciation_method, annual_rate, calculation_date=None):
    """
    Calculate depreciation schedule for an asset

    Parameters:
    - asset_cost: float - original cost of asset
    - purchase_date: date - when asset was purchased/first used
    - depreciation_method: str - 'diminishing_value' or 'prime_cost'
    - annual_rate: float - annual depreciation rate (as decimal, e.g. 0.20 for 20%)
    - calculation_date: date - date to calculate up to (defaults to today)

    Returns:
    - dict with depreciation schedule
    """
    if calculation_date is None:
        calculation_date = date.today()

    # Initialize variables
    schedule = []
    remaining_value = asset_cost
    accumulated_depreciation = 0
    current_date = purchase_date

    # Calculate number of years to process
    years_diff = calculation_date.year - purchase_date.year
    if calculation_date.month < purchase_date.month or (calculation_date.month == purchase_date.month and calculation_date.day < purchase_date.day):
        years_diff -= 1

    for year in range(years_diff + 1):
        year_start = date(current_date.year + year, current_date.month, current_date.day)
        year_end = date(current_date.year + year + 1, current_date.month, current_date.day - 1)

        # Adjust for final year
        if year_end > calculation_date:
            year_end = calculation_date

        # Calculate days held in this year
        days_in_year = (year_end - year_start).days + 1
        days_in_financial_year = 365 + (1 if calendar.isleap(year_start.year) else 0)

        # Calculate depreciation for this year
        if depreciation_method == 'diminishing_value':
            # Diminishing value: rate applied to remaining value
            yearly_depreciation = remaining_value * annual_rate * (days_in_year / days_in_financial_year)
        else:
            # Prime cost: rate applied to original cost
            yearly_depreciation = asset_cost * annual_rate * (days_in_year / days_in_financial_year)

        # Ensure we don't depreciate below zero
        if yearly_depreciation > remaining_value:
            yearly_depreciation = remaining_value

        remaining_value -= yearly_depreciation
        accumulated_depreciation += yearly_depreciation

        schedule.append({
            'Year': year + 1,
            'Financial_Year_Start': year_start,
            'Financial_Year_End': year_end,
            'Days_Held': days_in_year,
            'Opening_Value': remaining_value + yearly_depreciation,
            'Depreciation_Amount': yearly_depreciation,
            'Accumulated_Depreciation': accumulated_depreciation,
            'Closing_Value': remaining_value
        })

    return schedule

def create_depreciation_excel(asset_name, asset_cost, purchase_date, depreciation_schedule, method, rate):
    """Create formatted Excel file with depreciation schedule"""
    df = pd.DataFrame(depreciation_schedule)

    # Format dates
    df['Financial_Year_Start'] = df['Financial_Year_Start'].dt.strftime('%d/%m/%Y')
    df['Financial_Year_End'] = df['Financial_Year_End'].dt.strftime('%d/%m/%Y')

    # Format currency columns
    currency_cols = ['Opening_Value', 'Depreciation_Amount', 'Accumulated_Depreciation', 'Closing_Value']
    for col in currency_cols:
        df[col] = df[col].apply(lambda x: f"${x:,.2f}")

    # Create workbook
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Summary sheet
        summary_data = {
            'Asset Name': [asset_name],
            'Original Cost': [f"${asset_cost:,.2f}"],
            'Purchase Date': [purchase_date.strftime('%d/%m/%Y')],
            'Depreciation Method': [method.replace('_', ' ').title()],
            'Annual Rate': [f"{rate*100:.1f}%"],
            'Total Depreciation': [f"${df['Accumulated_Depreciation'].iloc[-1]:,.2f}"],
            'Current Value': [f"${df['Closing_Value'].iloc[-1]:,.2f}"]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Schedule sheet
        df.to_excel(writer, sheet_name='Depreciation_Schedule', index=False)

        # Auto-adjust column widths
        for sheet_name in ['Summary', 'Depreciation_Schedule']:
            sheet = writer.sheets[sheet_name]
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                sheet.column_dimensions[column_letter].width = adjusted_width

    output.seek(0)
    return output

# ==================== MAIN INTERFACE ==================== #

if calculation_mode == "Single Asset":
    st.subheader("📊 Single Asset Depreciation Calculator")

    col1, col2 = st.columns(2)

    with col1:
        asset_name = st.text_input("Asset Name/Description:", value="Office Computer")
        asset_cost = st.number_input("Asset Cost ($):", min_value=0.0, value=2000.0, step=100.0)
        purchase_date = st.date_input("Purchase/First Use Date:", value=date.today().replace(day=1))

    with col2:
        depreciation_method = st.selectbox(
            "Depreciation Method:",
            ["diminishing_value", "prime_cost"],
            format_func=lambda x: "Diminishing Value" if x == "diminishing_value" else "Prime Cost"
        )

        # ATO effective life guide
        effective_life_options = {
            "Computers/Office Equipment": 4,
            "Motor Vehicles": 8,
            "Plant & Equipment": 10,
            "Buildings": 40,
            "Custom Rate": None
        }

        selected_life = st.selectbox("Asset Type (for rate guide):", list(effective_life_options.keys()))

        if selected_life == "Custom Rate":
            annual_rate_pct = st.number_input("Annual Depreciation Rate (%):", min_value=0.0, max_value=100.0, value=25.0)
        else:
            effective_life = effective_life_options[selected_life]
            if depreciation_method == "diminishing_value":
                annual_rate_pct = (1 / effective_life) * 200  # Diminishing value uses 2x the prime cost rate
            else:
                annual_rate_pct = (1 / effective_life) * 100
            st.info(f"Suggested rate for {selected_life}: {annual_rate_pct:.1f}% per annum")

    # Calculate button
    if st.button("🧮 Calculate Depreciation", use_container_width=True):
        if asset_cost <= 0:
            st.error("Asset cost must be greater than zero")
        else:
            annual_rate = annual_rate_pct / 100

            # Calculate depreciation
            schedule = calculate_depreciation(
                asset_cost=asset_cost,
                purchase_date=purchase_date,
                depreciation_method=depreciation_method,
                annual_rate=annual_rate
            )

            # Display results
            st.success("✅ Depreciation calculation complete!")

            # Summary
            st.subheader("📈 Summary")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Original Cost", f"${asset_cost:,.2f}")
            with col2:
                st.metric("Total Depreciation", f"${schedule[-1]['Accumulated_Depreciation']:,.2f}")
            with col3:
                st.metric("Current Value", f"${schedule[-1]['Closing_Value']:,.2f}")
            with col4:
                st.metric("Depreciation Rate", f"{annual_rate_pct:.1f}%")

            # Depreciation schedule
            st.subheader("📅 Depreciation Schedule")
            df = pd.DataFrame(schedule)
            df_display = df.copy()
            df_display['Financial_Year_Start'] = df_display['Financial_Year_Start'].dt.strftime('%d/%m/%Y')
            df_display['Financial_Year_End'] = df_display['Financial_Year_End'].dt.strftime('%d/%m/%Y')
            df_display['Opening_Value'] = df_display['Opening_Value'].apply(lambda x: f"${x:,.2f}")
            df_display['Depreciation_Amount'] = df_display['Depreciation_Amount'].apply(lambda x: f"${x:,.2f}")
            df_display['Accumulated_Depreciation'] = df_display['Accumulated_Depreciation'].apply(lambda x: f"${x:,.2f}")
            df_display['Closing_Value'] = df_display['Closing_Value'].apply(lambda x: f"${x:,.2f}")

            st.dataframe(df_display, use_container_width=True)

            # Export to Excel
            if st.button("📥 Export to Excel", use_container_width=True):
                excel_file = create_depreciation_excel(
                    asset_name=asset_name,
                    asset_cost=asset_cost,
                    purchase_date=purchase_date,
                    depreciation_schedule=schedule,
                    method=depreciation_method,
                    rate=annual_rate
                )

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"depreciation_{asset_name.replace(' ', '_')}_{timestamp}.xlsx"

                st.download_button(
                    label="💾 Download Excel File",
                    data=excel_file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success(f"✅ Excel file ready: {filename}")

else:  # Asset Register
    st.subheader("📋 Asset Register Calculator")

    # File upload for asset register
    uploaded_file = st.file_uploader(
        "Upload Asset Register (Excel/CSV)",
        type=['xlsx', 'xls', 'csv'],
        help="Upload a spreadsheet with columns: Asset_Name, Cost, Purchase_Date, Method, Rate"
    )

    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.success(f"✅ Loaded {len(df)} assets from {uploaded_file.name}")

            # Display loaded data
            st.subheader("📊 Loaded Assets")
            st.dataframe(df, use_container_width=True)

            # Process all assets
            if st.button("🧮 Calculate All Depreciation", use_container_width=True):
                all_results = []

                progress_bar = st.progress(0)
                status_text = st.empty()

                for idx, row in df.iterrows():
                    status_text.text(f"Processing asset {idx + 1}/{len(df)}: {row.get('Asset_Name', f'Asset {idx+1}')}")

                    try:
                        asset_name = row.get('Asset_Name', f'Asset {idx+1}')
                        asset_cost = float(row.get('Cost', 0))
                        purchase_date = pd.to_datetime(row.get('Purchase_Date', date.today())).date()
                        method = row.get('Method', 'diminishing_value').lower().replace(' ', '_')
                        rate = float(row.get('Rate', 0)) / 100  # Convert from percentage

                        if asset_cost > 0:
                            schedule = calculate_depreciation(
                                asset_cost=asset_cost,
                                purchase_date=purchase_date,
                                depreciation_method=method,
                                annual_rate=rate
                            )

                            # Add asset name to each row
                            for item in schedule:
                                item['Asset_Name'] = asset_name
                                item['Original_Cost'] = asset_cost
                                all_results.append(item)

                    except Exception as e:
                        st.warning(f"Error processing asset {idx + 1}: {str(e)}")

                    progress_bar.progress((idx + 1) / len(df))

                status_text.text("✅ Processing complete!")

                if all_results:
                    # Display summary
                    results_df = pd.DataFrame(all_results)
                    st.subheader("📈 Depreciation Summary")
                    summary = results_df.groupby('Asset_Name').agg({
                        'Original_Cost': 'first',
                        'Accumulated_Depreciation': 'last',
                        'Closing_Value': 'last'
                    }).reset_index()

                    summary['Total_Depreciation'] = summary['Accumulated_Depreciation']
                    summary['Current_Value'] = summary['Closing_Value']
                    summary = summary[['Asset_Name', 'Original_Cost', 'Total_Depreciation', 'Current_Value']]

                    st.dataframe(summary, use_container_width=True)

                    # Export all results
                    if st.button("📥 Export All to Excel", use_container_width=True):
                        # Format for Excel
                        excel_df = results_df.copy()
                        excel_df['Financial_Year_Start'] = excel_df['Financial_Year_Start'].dt.strftime('%d/%m/%Y')
                        excel_df['Financial_Year_End'] = excel_df['Financial_Year_End'].dt.strftime('%d/%m/%Y')

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            summary.to_excel(writer, sheet_name='Summary', index=False)
                            excel_df.to_excel(writer, sheet_name='Detailed_Schedule', index=False)

                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"asset_register_depreciation_{timestamp}.xlsx"

                        st.download_button(
                            label="💾 Download Excel File",
                            data=output,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success(f"✅ Excel file ready: {filename}")
                else:
                    st.error("No valid assets found to process")

        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            st.info("Expected columns: Asset_Name, Cost, Purchase_Date, Method, Rate")
    else:
        st.info("📤 Upload an Excel or CSV file with your asset register to calculate depreciation for multiple assets at once.")

# ==================== FOOTER ==================== #
st.divider()
st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**✨ Features:**\n- Diminishing Value & Prime Cost\n- ATO compliant calculations\n- Single & batch processing")
with col2:
    st.markdown("**📋 Methods:**\n- Diminishing Value (default)\n- Prime Cost method\n- Custom rates supported")
with col3:
    st.markdown("**💾 Export:**\n- Formatted Excel\n- Depreciation schedules\n- Tax-ready reports")
