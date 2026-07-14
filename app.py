import streamlit as st
import pandas as pd
import numpy as np
import os
import io

st.set_page_config(page_title="Mitsubishi Finance Calculator", layout="wide")

# Helper function to extract a clean parent model name
def get_clean_model_name(raw_name):
    raw_name_upper = raw_name.upper()
    if 'ATTRAGE' in raw_name_upper: return 'Attrage'
    elif 'MIRAGE' in raw_name_upper: return 'Mirage'
    elif 'ASX' in raw_name_upper: return 'ASX'
    elif 'ECLIPSE' in raw_name_upper: return 'Eclipse Cross'
    elif 'XPANDER' in raw_name_upper: return 'Xpander'
    elif 'OUTLANDER' in raw_name_upper: return 'Outlander'
    elif 'MONTERO' in raw_name_upper: return 'Montero Sport'
    elif 'DESTINATOR' in raw_name_upper: return 'Destinator'
    return raw_name

# ------------------------------------------------------------------
# AUTOMATIC DATA EXTRACTION ENGINE
# ------------------------------------------------------------------
@st.cache_data
def load_supplementary_data(file_path):
    """Loads the Bank details and RMC tiers from the standalone file."""
    bank_data = {}
    rmc_data = {}
    
    if os.path.exists(file_path):
        # 1. Parse Bank Details
        try:
            df_bank = pd.read_excel(file_path, sheet_name='Bank Details')
            for _, row in df_bank.iterrows():
                bank_name = str(row['Bank Name']).strip()
                if not bank_name or bank_name == "nan": continue
                
                bank_data[bank_name] = {}
                for i in range(1, 6):
                    sb_col = f"Salary Bracket.{i}" if i > 1 else "Salary Bracket"
                    roi_col = f"ROI.{i}" if i > 1 else "ROI"
                    
                    if sb_col in df_bank.columns and roi_col in df_bank.columns:
                        sb_val = str(row[sb_col]).strip()
                        roi_val = row[roi_col]
                        if sb_val and sb_val != "nan" and pd.notna(roi_val):
                            bank_data[bank_name][sb_val] = float(roi_val)
        except Exception as e:
            st.warning(f"Could not parse Bank Details tab: {e}")

        # 2. Parse RMC Tiers
        try:
            df_rmc = pd.read_excel(file_path, sheet_name='RMC')
            for _, row in df_rmc.iterrows():
                v_code = str(row['Variant Codes']).strip()
                if not v_code or v_code == "nan": continue
                
                rmc_data[v_code] = {
                    "RMC-10-40": float(row['RMC-10-40']) if pd.notna(row['RMC-10-40']) else 0.0,
                    "RMC-10-60": float(row['RMC-10-60']) if pd.notna(row['RMC-10-60']) else 0.0,
                    "RMC-10-70": float(row['RMC-10-70']) if pd.notna(row['RMC-10-70']) else 0.0,
                    "RMC-10-100": float(row['RMC-10-100']) if pd.notna(row['RMC-10-100']) else 0.0,
                }
        except Exception as e:
            st.warning(f"Could not parse RMC tab: {e}")
            
    return bank_data, rmc_data

@st.cache_data
def load_all_vehicle_data(vehicle_file_path):
    """Parses the 80 individual variant sheets from the core project workbook."""
    catalog = {"2025": {}, "2026": {}}
    
    if not os.path.exists(vehicle_file_path):
        return catalog

    xls = pd.ExcelFile(vehicle_file_path)
    for sheet in xls.sheet_names:
        # Skip tracking index sheets if they exist
        if sheet in ['Structure', 'MY-2025', 'MY-2026', 'Combined 2025-2026', 'Bank Details', 'RMC']: 
            continue
            
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            raw_name = str(df.iloc[4, 2]).strip()
            code = str(df.iloc[4, 3]).strip()
            year_string = str(df.iloc[4, 4]).strip()
            
            if not raw_name or raw_name == "nan": raw_name = sheet
            if not code or code == "nan": code = sheet
            
            # Determine Model Year dynamically
            year_key = "2026" if "2026" in year_string or "26" in sheet else "2025"
            
            base_price = float(df.iloc[6, 1])          
            interest_rate = float(df.iloc[18, 3])  
            
            # Extract accessories (Rows 10 to 15)
            accessories = {}
            for r_idx in range(10, 16):
                acc_name = str(df.iloc[r_idx, 1]).strip()
                acc_status = str(df.iloc[r_idx, 2]).strip().upper()
                try: acc_val = float(df.iloc[r_idx, 3])
                except: acc_val = 0.0
                
                if acc_name and acc_name != "nan" and acc_name != "Accessory":
                    accessories[acc_name] = {"price": acc_val, "default_checked": acc_status == "YES"}
            
            model_name = get_clean_model_name(raw_name)
            
            if model_name not in catalog[year_key]: 
                catalog[year_key][model_name] = {}
                
            catalog[year_key][model_name][code] = {
                "base_price": base_price, 
                "interest_rate": interest_rate, 
                "accessories": accessories
            }
        except: 
            continue
                
    return catalog

# ------------------------------------------------------------------
# ENVIRONMENT INITIALIZATION
# ------------------------------------------------------------------
FILE_VEHICLES = "NFC New VRI Project (2).xlsx"
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"

VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)

# ------------------------------------------------------------------
# INTERFACE LAYOUT (SIDEBAR CONSOLE)
# ------------------------------------------------------------------
if not VEHICLE_CATALOG["2025"] and not VEHICLE_CATALOG["2026"]:
    st.error(f"Could not load vehicle datasets from '{FILE_VEHICLES}'. Please verify it exists in your repository.")
else:
    with st.sidebar:
        st.header("🚗 Vehicle Selection")
        selected_year = st.selectbox("Select Model Year:", sorted(list(VEHICLE_CATALOG.keys())))
        
        available_names = sorted(list(VEHICLE_CATALOG[selected_year].keys()))
        selected_name = st.selectbox("Select Vehicle Name:", available_names)
        
        if selected_name not in VEHICLE_CATALOG[selected_year]:
            selected_name = available_names[0] if available_names else None

        if selected_name:
            available_codes = sorted(list(VEHICLE_CATALOG[selected_year][selected_name].keys()))
            selected_code = st.selectbox("Select Variant Code:", available_codes)
            v_data = VEHICLE_CATALOG[selected_year][selected_name][selected_code]
        else:
            st.error("No variants located matching specifications.")
            st.stop()
        
        st.markdown("---")
        st.header("🏦 Financial Provider Setup")
        
        # Load bank data dynamically if the sheet exists, otherwise fall back gracefully
        if BANK_RULES:
            bank_options = sorted(list(BANK_RULES.keys()))
            selected_bank = st.selectbox("Select Financial Institution:", bank_options)
            
            bracket_options = sorted(list(BANK_RULES[selected_bank].keys()))
            selected_bracket = st.selectbox("Select Base Income Bracket:", bracket_options)
            
            fetched_roi = BANK_RULES[selected_bank][selected_bracket]
        else:
            selected_bank = "Sheet Baseline Default"
            fetched_roi = v_data["interest_rate"]
            st.info("Upload 'Bank & RMC Details.xlsx' to activate bank bracket dropdowns.")
            
        bank_rate = st.number_input("Flat Interest Rate (ROI):", value=fetched_roi, format="%.4f", step=0.0001)
        
        st.markdown("---")
        st.header("⚙️ Adjustment Matrix")
        base_vehicle_price = st.number_input("Base Price (AED):", value=v_data["base_price"], step=500.0)
        Calculate accessories total
accessories_total = (
    ceramic_coating +
    extended_warranty +
    service_contract +
    insurance +
    vri
)

# Total vehicle value including everything
total_vehicle_value = (
    base_price +
    vat +
    accessories_total
)

# Correct down payment calculation
down_payment_amount = total_vehicle_value * (down_payment_percentage / 100)
        st.markdown("---")
        st.header("➕ Optional Agreements")
        
        # Pull RMC options dynamically if the file is present
        rmc_cost = 0.0
        if RMC_RULES and selected_code in RMC_RULES:
            st.write("**Regional Maintenance Contracts (RMC)**")
            rmc_packages = ["None"] + list(RMC_RULES[selected_code].keys())
            chosen_rmc = st.selectbox("Select Service Tier:", rmc_packages)
            if chosen_rmc != "None":
                rmc_cost = RMC_RULES[selected_code][chosen_rmc]
                st.write(f"*Appended Cost: +{rmc_cost:,.2f} AED*")
        
        st.write("**Accessories Checklists:**")
        selected_addons_total = 0.0
        for addon_name, info in v_data["accessories"].items():
            if st.checkbox(f"{addon_name} (+{info['price']:,.0f} AED)", value=info["default_checked"]):
                selected_addons_total += info["price"]

    # ------------------------------------------------------------------
    # MAIN RENDERING MATRIX
    # ------------------------------------------------------------------
    st.title("Mitsubishi Financial Matrix Calculator")
    st.markdown(f"### Current Node: **{selected_name} — {selected_code} ({selected_year})**")
    if BANK_RULES:
        st.caption(f"Provider Context: {selected_bank} ({selected_bracket})")
    st.markdown("---")
    
    # Calculate overall aggregate
    total_financed_amount = (base_vehicle_price + selected_addons_total + rmc_cost) - calculated_downpayment
    
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
    
    # Format for clean display visual
    df_display = df_output.copy()
    df_display["Financed Principal (AED)"] = df_display["Financed Principal (AED)"].map('{:,.2f}'.format)
    df_display["Total Interest (AED)"] = df_display["Total Interest (AED)"].map('{:,.2f}'.format)
    df_display["Estimated Monthly EMI (AED)"] = df_display["Estimated Monthly EMI (AED)"].map('{:,.2f}'.format)
    
    st.table(df_display)
    st.markdown("<br>", unsafe_allow_html=True)
    
    # --- STREAM EXPORT UTILITY ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_output.to_excel(writer, index=False, sheet_name="EMI Matrix")
    
    st.download_button(
        label="📥 Export Financial Matrix to Excel",
        data=buffer.getvalue(),
        file_name=f"Finance_Matrix_{selected_name.replace(' ', '_')}_{selected_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
