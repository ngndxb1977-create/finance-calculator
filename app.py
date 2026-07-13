import streamlit as st
import pandas as pd
import numpy as np
import os
import io

st.set_page_config(page_title="Mitsubishi Finance Calculator", layout="wide")

# ------------------------------------------------------------------
# AUTOMATIC EXCEL DATA EXTRACTION ENGINE
# ------------------------------------------------------------------
@st.cache_data
def load_all_vehicle_data(file_2025_path, file_2026_path):
    # Nested mapping structure: catalog[year][vehicle_name][variant_code] = data
    catalog = {"2025": {}, "2026": {}}
    
    # --- 1. Parse 2025 File ---
    if os.path.exists(file_2025_path):
        xls_25 = pd.ExcelFile(file_2025_path)
        for sheet in xls_25.sheet_names:
            if sheet in ['Structure', 'MY-2025', 'MY-2026']:
                continue
            try:
                df = pd.read_excel(xls_25, sheet_name=sheet, header=None)
                name = str(df.iloc[4, 2]).strip()
                code = str(df.iloc[4, 3]).strip()
                
                if not name or name == "nan": name = sheet
                if not code or code == "nan": code = sheet.replace(" 25", "")
                
                base_price = float(df.iloc[6, 1])          
                vat_charges = float(df.iloc[6, 5])         
                interest_rate = float(df.iloc[18, 3])       
                
                accessories = {}
                for r_idx in range(10, 16):
                    acc_name = str(df.iloc[r_idx, 1]).strip()
                    acc_status = str(df.iloc[r_idx, 2]).strip().upper()
                    try:
                        acc_val = float(df.iloc[r_idx, 3])
                    except:
                        acc_val = 0.0
                    
                    if acc_name and acc_name != "nan":
                        accessories[acc_name] = {
                            "price": acc_val,
                            "default_checked": True if acc_status == "YES" else False
                        }
                
                if name not in catalog["2025"]:
                    catalog["2025"][name] = {}
                    
                catalog["2025"][name][code] = {
                    "base_price": base_price,
                    "interest_rate": interest_rate,
                    "vat_charges": vat_charges,
                    "accessories": accessories
                }
            except:
                continue

    # --- 2. Parse 2026 File ---
    if os.path.exists(file_2026_path):
        xls_26 = pd.ExcelFile(file_2026_path)
        for sheet in xls_26.sheet_names:
            if sheet in ['Structure', 'MY-2025', 'MY-2026']:
                continue
            try:
                df = pd.read_excel(xls_26, sheet_name=sheet, header=None)
                name = str(df.iloc[4, 2]).strip()
                code = str(df.iloc[4, 3]).strip()
                
                if not name or name == "nan": name = sheet
                if not code or code == "nan": code = sheet.replace(" 26", "")
                
                base_price = float(df.iloc[6, 1])          
                vat_charges = float(df.iloc[6, 5])         
                interest_rate = float(df.iloc[18, 3])       
                
                accessories = {}
                for r_idx in range(10, 16):
                    acc_name = str(df.iloc[r_idx, 1]).strip()
                    acc_status = str(df.iloc[r_idx, 2]).strip().upper()
                    try:
                        acc_val = float(df.iloc[r_idx, 3])
                    except:
                        acc_val = 0.0
                    
                    if acc_name and acc_name != "nan":
                        accessories[acc_name] = {
                            "price": acc_val,
                            "default_checked": True if acc_status == "YES" else False
                        }
                
                if name not in catalog["2026"]:
                    catalog["2026"][name] = {}
                    
                catalog["2026"][name][code] = {
                    "base_price": base_price,
                    "interest_rate": interest_rate,
                    "vat_charges": vat_charges,
                    "accessories": accessories
                }
            except:
                continue
                
    return catalog

# Load catalog mapping using both tracked files
FILE_2025 = "NFC New VRI Project (2).xlsx"
FILE_2026 = "NFC New VRI Project (1).xlsx"
VEHICLE_CATALOG = load_all_vehicle_data(FILE_2025, FILE_2026)

# ------------------------------------------------------------------
# INTERFACE LAYOUT (SIDEBAR CASCADING DROPDOWNS)
# ------------------------------------------------------------------
if not VEHICLE_CATALOG["2025"] and not VEHICLE_CATALOG["2026"]:
    st.error("Could not load vehicle datasets. Please verify your repository file allocations.")
else:
    with st.sidebar:
        st.header("🚗 Vehicle Selection")
        
        # Dropdown 1: Choose Year
        available_years = sorted(list(VEHICLE_CATALOG.keys()))
        selected_year = st.selectbox("Select Model Year:", available_years)
        
        # Dropdown 2: Choose Vehicle Name (Filtered by selected year)
        available_names = sorted(list(VEHICLE_CATALOG[selected_year].keys()))
        selected_name = st.selectbox("Select Vehicle Name:", available_names)
        
        # Dropdown 3: Choose Variant Code (Filtered by selected year + name)
        available_codes = sorted(list(VEHICLE_CATALOG[selected_year][selected_name].keys()))
        selected_code = st.selectbox("Select Variant Code:", available_codes)
        
        # Retrieve the final filtered record data points
        v_data = VEHICLE_CATALOG[selected_year][selected_name][selected_code]
        
        st.markdown("---")
        st.header("⚙️ Adjustments")
        base_vehicle_price = st.number_input("Base Price (AED):", value=v_data["base_price"], step=500.0)
        bank_rate = st.number_input("Flat Interest Rate:", value=v_data["interest_rate"], format="%.4f", step=0.0001)
        
        down_payment_pct = st.slider("Down Payment Percentage (%):", 0, 100, 20) / 100.0
        calculated_downpayment = base_vehicle_price * down_payment_pct
        st.write(f"**Down Payment Amount:** {calculated_downpayment:,.2f} AED")
        
        st.markdown("---")
        st.header("➕ Add-ons & Accessories")
        selected_addons_total = 0.0
        
        for addon_name, info in v_data["accessories"].items():
            is_checked = st.checkbox(
                f"{addon_name} (+{info['price']:,.0f} AED)", 
                value=info["default_checked"]
            )
            if is_checked:
                selected_addons_total += info["price"]

    # ------------------------------------------------------------------
    # MAIN AREA: CLEAN MATRIX PERFORMANCE DISPLAY
    # ------------------------------------------------------------------
    st.title("Mitsubishi Financial Matrix Calculator")
    st.markdown(f"### Currently Viewing: **{selected_name} - {selected_code} ({selected_year})**")
    st.markdown("---")
    
    total_financed_amount = (base_vehicle_price + selected_addons_total) - calculated_downpayment
    
    tenures = [2, 3, 4, 5]
    emi_results = []
    
    for years in tenures:
        months = years * 12
        total_interest = total_financed_amount * bank_rate * years
        total_repayable = total_financed_amount + total_interest
        monthly_emi = total_repayable / months
        
        emi_results.append({
            "Tenure Period": f"{years} Years ({months} Months)",
            "Financed Principal (AED)": round(total_financed_amount, 2),
            "Total Interest (AED)": round(total_interest, 2),
            "Estimated Monthly EMI (AED)": round(monthly_emi, 2)
        })
        
    df_output = pd.DataFrame(emi_results)
    
    df_display = df_output.copy()
    df_display["Financed Principal (AED)"] = df_display["Financed Principal (AED)"].map('{:,.2f}'.format)
    df_display["Total Interest (AED)"] = df_display["Total Interest (AED)"].map('{:,.2f}'.format)
    df_display["Estimated Monthly EMI (AED)"] = df_display["Estimated Monthly EMI (AED)"].map('{:,.2f}'.format)
    
    st.table(df_display)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # --- EXPORT FUNCTIONALITY ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_output.to_excel(writer, index=False, sheet_name="EMI Matrix")
    
    st.download_button(
        label="📥 Export Financial Matrix to Excel",
        data=buffer.getvalue(),
        file_name=f"Finance_Matrix_{selected_name.replace(' ', '_')}_{selected_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
