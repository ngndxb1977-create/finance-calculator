import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import re

# Set configuration at the absolute top
st.set_page_config(page_title="Mitsubishi Financial Matrix Calculator", layout="wide")

# ==========================================
# CUSTOM UX/UI STYLING ENGINE
# ==========================================
st.markdown(
   """
   <style>
   @import url('https://fonts.googleapis.com/css2?family=Karma:wght=400;600&family=Amethysta&family=Quicksand:wght=500;700&display=swap');
 
   .stApp {
       background-color: #FBF9F6 !important;
       color: #191919 !important;
       font-family: 'Karma', serif !important;
       font-size: 16px !important;
       line-height: 1.6 !important;
    }
 
   h1, [data-testid="stHeader"] {
       font-family: 'Quicksand', sans-serif !important;
       font-weight: 700 !important;
       font-size: 2.25rem !important;
       color: #191919 !important;
       letter-spacing: -0.02em;
    }
 
   h2, h3, h4, h5, h6 {
       font-family: 'Amethysta', serif !important;
       font-weight: 400 !important;
       font-size: 1.5rem !important;
       color: #383838 !important;
       margin-top: 1.5rem !important;
    }
 
   [data-testid="stSidebar"] {
       background-color: #F4F0EA !important;
    }
    
   .requirement-box {
       background-color: #F4F0EA;
       padding: 1.5rem;
       border-radius: 8px;
       border-left: 4px solid #191919;
       margin-top: 1.5rem;
    }
   .requirement-box ul {
       margin-bottom: 1rem;
       padding-left: 1.25rem;
    }
   .disclaimer-text {
       font-size: 0.85rem;
       color: #555555;
       margin-top: 0.5rem;
       line-height: 1.4;
    }
   </style>
   """,
   unsafe_allow_html=True
)

def get_clean_model_name(raw_name):
    raw_name_upper = raw_name.upper()
    if 'ATTRAGE' in raw_name_upper or 'ATG' in raw_name_upper: return 'Attrage'
    elif 'MIRAGE' in raw_name_upper or 'MIG' in raw_name_upper: return 'Mirage'
    elif 'ASX' in raw_name_upper: return 'ASX'
    elif 'ECLIPSE' in raw_name_upper: return 'Eclipse Cross'
    elif 'XPANDER' in raw_name_upper: return 'Xpander'
    elif 'OUTLANDER' in raw_name_upper: return 'Outlander'
    elif 'MONTERO' in raw_name_upper: return 'Montero Sport'
    elif 'DESTINATOR' in raw_name_upper: return 'Destinator'
    return raw_name

def normalize_bracket_string(raw_val):
   if not raw_val or pd.isna(raw_val): return ""
   val = str(raw_val).strip().replace(",", "")
   val = re.sub(r'\s+', '', val)  
   return val.replace("to", "-").replace("—", "-").replace("–", "-")

# ------------------------------------------------------------------
# MASTER EXCEL EXTRACTION ENGINE
# ------------------------------------------------------------------
@st.cache_data
def load_supplementary_data(file_path):
   bank_data, rmc_data = {}, {}
   if os.path.exists(file_path):
       try:
           df_bank = pd.read_excel(file_path, sheet_name='Bank Details', header=None)
           for _, row in df_bank.iterrows():
                try:
                    bank_name = ""
                    for i in range(4):
                        if i >= len(row): break
                        val = str(row.iloc[i]).strip()
                        if val and val.lower() != "nan" and not val.replace('.', '', 1).isdigit():
                            if val.upper() not in ["S.NO", "BANK NAME"]:
                                bank_name = val
                                break
                    if not bank_name: continue
                    if bank_name not in bank_data: bank_data[bank_name] = {}
                    # Simple parser logic...
                except: continue
       except: pass
       try:
           df_rmc = pd.read_excel(file_path, sheet_name='RMC')
           for _, row in df_rmc.iterrows():
                v_code = str(row['Variant Codes']).strip()
                if v_code != "nan": rmc_data[v_code] = {"RMC-10-40": float(row['RMC-10-40'])}
       except: pass
   return bank_data, rmc_data

@st.cache_data
def load_all_vehicle_data(vehicle_file_path):
   catalog = {"2025": {}, "2026": {}}
   if not os.path.exists(vehicle_file_path): return catalog
   xls = pd.ExcelFile(vehicle_file_path)
   for sheet in xls.sheet_names:
       if sheet in ['Structure', 'Bank Details', 'RMC']: continue
       try:
           df = pd.read_excel(xls, sheet_name=sheet, header=None)
           code = str(df.iloc[4, 3]).strip()
           year_key = "2026" if "26" in sheet else "2025"
           model_name = get_clean_model_name(str(df.iloc[4, 2]))
           if model_name not in catalog[year_key]: catalog[year_key][model_name] = {}
           catalog[year_key][model_name][code] = {
                "base_price": float(df.iloc[6, 1]), 
                "interest_rate": float(df.iloc[18, 3]),
                "registration_fee": float(df.iloc[11, 7]),
                "reservation_fee": 1000.0
           }
       except: continue
   return catalog

FILE_VEHICLES = "NFC New VRI Project (2) (2).xlsx"
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"
VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)

# ------------------------------------------------------------------
# APP LOGIC
# ------------------------------------------------------------------
if "view_state" not in st.session_state: st.session_state.view_state = "input"

with st.sidebar:
    st.header("🚗 Configuration Console")
    selected_year = st.selectbox("Model Year:", sorted(list(VEHICLE_CATALOG.keys())))
    selected_name = st.selectbox("Vehicle Name:", sorted(list(VEHICLE_CATALOG[selected_year].keys())))
    selected_code = st.selectbox("Variant Code:", sorted(list(VEHICLE_CATALOG[selected_year][selected_name].keys())))
    v_data = VEHICLE_CATALOG[selected_year][selected_name][selected_code]
    
    down_payment_pct = st.slider("Down Payment Percentage (%):", 0, 100, 20) / 100.0
    finance_dp_option = st.toggle("Finance the Down Payment?", value=False)
    
    if st.button("Generate Summary"): st.session_state.view_state = "summary"

if st.session_state.view_state == "summary":
    st.title("📄 Mitsubishi Financial Matrix Calculator")
    
    # Simple calculation display
    finance_amount = v_data["base_price"] * (1 - down_payment_pct)
    st.metric("Finance Amount", f"{finance_amount:,.2f} AED")

    # SECTION 5: DISCLOSURES (FIXED)
    st.header("5. Application Requirements & Disclosures")
    st.markdown(
        """
        <div class="requirement-box">
            <strong style="font-family: 'Quicksand', sans-serif; font-size: 1.1rem; color: #191919; display: block; margin-bottom: 0.75rem;">📋 Required Documentation Checklist:</strong>
            <ul>
                <li>Passport Copy, Digital Visa & Address Page.</li>
                <li>Emirates ID Card Copy.</li>
                <li>Salary Certificate & Last 3 Months Pay Slips.</li>
                <li>IBAN.</li>
            </ul>
            <p style="font-weight: 600; margin-top: 1rem; color: #191919;">✍️ Security Cheque for Bank:</p>
            <p style="font-size: 0.95rem; padding-left: 0.5rem; color: #383838;">One Security Cheque required from your salary account after finance approval.</p>
            
            <hr style="border: 0; border-top: 1px solid #D1C9BE; margin: 1.25rem 0;">
            
            <p class="disclaimer-text">Note: Please bring this Calculation Sheet at the time of submission of full documents. Interest rate may vary subject to bank approval.</p>
            <p class="disclaimer-text">Note: This sheet is only communication between dealer and customer.</p>
            
            <p style="margin-top: 1.25rem; font-weight: 600; font-size: 0.95rem; color: #191919;">
                Finance Clarifications: Kindly forward an email to <a href="mailto:naveen@habtoormotors.com" style="color: #191919; font-weight:700;">naveen@habtoormotors.com</a> or reach us on <strong>04-608-4000</strong>
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
