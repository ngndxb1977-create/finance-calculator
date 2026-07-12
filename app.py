import streamlit as st
import pandas as pd
import numpy as np
import os

st.set_page_config(page_title="Mitsubishi Finance Calculator", layout="wide")

# ------------------------------------------------------------------
# AUTOMATIC EXCEL DATA EXTRACTION ENGINE
# ------------------------------------------------------------------
@st.cache_data
def load_all_vehicle_data(file_path):
    if not os.path.exists(file_path):
        return {}
    
    xls = pd.ExcelFile(file_path)
    catalog = {}
    
    # Loop through all individual variant sheets
    for sheet in xls.sheet_names:
        # Ignore main interactive template sheets
        if sheet in ['Structure', 'MY-2025', 'MY-2026']:
            continue
            
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            # 1. Extract Model Information
            model_name = str(df.iloc[3, 2]).strip()  # Cell C4
            model_code = str(df.iloc[3, 3]).strip()  # Cell D4
            
            # Fallback to sheet name if cell parsing is blank
            if not model_name or model_name == "nan":
                model_name = sheet
                
            # Determine Year based on sheet name suffix
            year = "2026" if sheet.endswith("26") else "2025"
            display_label = f"{model_name} ({year})"
            
            # 2. Extract Key Financial Constants
            base_price = float(df.iloc[5, 1])          # Cell B6 (Vehicle Price)
            vat_charges = float(df.iloc[5, 5])         # Cell F6 (Total VAT Charges)
            interest_rate = float(df.iloc[17, 3])       # Cell D18 (Bank Interest Rate)
            
            # 3. Extract Accessories Matrix
            accessories = {}
            # Loop over rows 9 through 14 (0-indexed indices 9 to 14)
            for r_idx in range(9, 15):
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
            
            # Store everything systematically under its clear UI name
            catalog[display_label] = {
                "base_price": base_price,
                "interest_rate": interest_rate,
                "vat_charges": vat_charges,
                "accessories": accessories,
                "year": year
            }
        except Exception as e:
            # Silently skip incomplete or malformed diagnostic sheets
            continue
            
    return catalog

# Load the catalog dataset
EXCEL_FILE = "NFC New VRI Project (2).xlsx"
VEHICLE_CATALOG = load_all_vehicle_data(EXCEL_FILE)

# ------------------------------------------------------------------
# INTERFACE & USER SELECTION INTERACTION LAYER
# ------------------------------------------------------------------
st.title("Mitsubishi Financial Matrix Calculator")

if not VEHICLE_CATALOG:
    st.error(f"Could not find or load '{EXCEL_FILE}' in your repository directory. Please verify your filename match.")
else:
    # 1. Selection Tier
    sorted_options = sorted(list(VEHICLE_CATALOG.keys()))
    selected_variant = st.selectbox("Select Vehicle Model Variant:", sorted_options)
    
    # Retrieve pre-parsed structural data points
    v_data = VEHICLE_CATALOG[selected_variant]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Vehicle Pricing Configuration")
        base_vehicle_price = st.number_input("Base Vehicle Price (AED):", value=v_data["base_price"], step=500.0)
        bank_rate = st.number_input("Flat Bank Interest Rate (Flat):", value=v_data["interest_rate"], format="%.4f", step=0.0001)
        
        # Down Payment Control
        down_payment_pct = st.slider("Down Payment Percentage (%):", 0, 100, 20) / 100.0
        calculated_downpayment = base_vehicle_price * down_payment_pct
        st.info(f"Calculated Down Payment: {calculated_downpayment:,.2f} AED")

    with col2:
        st.subheader("Add-ons & Accessories")
        selected_addons_total = 0.0
        
        # Dynamically build checkbox UI entries from what was found inside the Excel sheets
        for addon_name, info in v_data["accessories"].items():
            is_checked = st.checkbox(
                f"{addon_name} (+{info['price']:,.2f} AED)", 
                value=info["default_checked"]
            )
            if is_checked:
                selected_addons_total += info["price"]
                
    # ------------------------------------------------------------------
    # LOAN CALCULATION MATRIX ENGINE
    # ------------------------------------------------------------------
    st.markdown("---")
    st.subheader("Financing Monthly Options Execution")
    
    # Financial Formula Math Tiers
    total_financed_amount = (base_vehicle_price + selected_addons_total) - calculated_downpayment
    
    tenures = [2, 3, 4, 5]
    emi_results = []
    
    for years in tenures:
        months = years * 12
        # Simple Flat Rate Pricing Model Rule
        total_interest = total_financed_amount * bank_rate * years
        total_repayable = total_financed_amount + total_interest
        monthly_emi = total_repayable / months
        
        emi_results.append({
            "Tenure Period": f"{years} Years ({months} Months)",
            "Financed Principal (AED)": f"{total_financed_amount:,.2f}",
            "Total Interest (AED)": f"{total_interest:,.2f}",
            "Estimated Monthly EMI (AED)": f"{monthly_emi:,.2f}"
        })
        
    st.table(pd.DataFrame(emi_results))
