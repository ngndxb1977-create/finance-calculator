import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import re

# Set configuration at the absolute top
st.set_page_config(page_title="Mitsubishi Financial Matrix Calculator", layout="wide")

# ==========================================
# CUSTOM UX/UI STYLING ENGINE (AMETHYSTA UPDATE)
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
    
    [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        font-family: 'Quicksand', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1.2rem !important;
    }

    [data-testid="stMetricValue"] {
        font-family: 'Amethysta', serif !important;
        font-size: 1.6rem !important;
        font-weight: 400 !important;
        color: #191919 !important;
    }
    [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] p {
        font-family: 'Quicksand', sans-serif !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        color: #555555 !important;
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
    if not raw_val or pd.isna(raw_val):
        return ""
    val = str(raw_val).strip().replace(",", "")
    val = re.sub(r'\s+', '', val)  
    val = val.replace("to", "-").replace("—", "-").replace("–", "-")
    return val

# ------------------------------------------------------------------
# MASTER EXCEL EXTRACTION ENGINE (INTELLIGENT COLUMN LOCATOR)
# ------------------------------------------------------------------
@st.cache_data
def load_supplementary_data(file_path):
    bank_data = {}
    rmc_data = {}
    if os.path.exists(file_path):
        try:
            df_bank = pd.read_excel(file_path, sheet_name='Bank Details', header=None)
            for _, row in df_bank.iterrows():
                try:
                    bank_name = ""
                    start_col = 1
                    ignore_headers = ["S.NO", "S. NO", "SR NO", "SR. NO", "SL NO", "SL. NO", "SERIAL", "#"]
                    for i in range(4):
                        if i >= len(row): break
                        val = str(row.iloc[i]).strip()
                        if val and val.lower() != "nan" and not val.replace('.', '', 1).isdigit():
                            if val.upper() not in ignore_headers:
                                bank_name = val
                                start_col = i + 1  
                                break
                    if not bank_name or "BANK NAME" in bank_name.upper(): 
                        continue
                    if bank_name not in bank_data:
                        bank_data[bank_name] = {}
                    num_cols = len(row)
                    for col_idx in range(start_col, num_cols):
                        cell_val = row.iloc[col_idx]
                        if pd.isna(cell_val) or str(cell_val).strip() == "":
                            continue
                        cell_str = str(cell_val).strip()
                        is_salary_pattern = re.search(r'\d+', cell_str) is not None and any(x in cell_str.lower() for x in ['-', 'to', '+', 'above', 'k', 'min'])
                        if is_salary_pattern:
                            for look_ahead in range(1, 3):
                                if col_idx + look_ahead < num_cols:
                                    next_val = row.iloc[col_idx + look_ahead]
                                    if pd.notna(next_val) and str(next_val).strip() != "":
                                        try:
                                            roi_str = str(next_val).replace("%", "").strip()
                                            parsed_roi = float(roi_str)
                                            if parsed_roi > 1.0:
                                                parsed_roi = parsed_roi / 100.0
                                            norm_sb_val = normalize_bracket_string(cell_val)
                                            if "-" in norm_sb_val:
                                                parts = norm_sb_val.split("-")
                                                display_label = f"{parts[0]}-{parts[1]}"
                                            else:
                                                display_label = norm_sb_val
                                            bank_data[bank_name][display_label] = parsed_roi
                                            break 
                                        except:
                                            continue
                except:
                    continue
        except Exception as e:
            st.error(f"Critical error initializing Bank Details sheet: {e}")

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
        except:
            pass
    return bank_data, rmc_data

@st.cache_data
def load_all_vehicle_data(vehicle_file_path):
    catalog = {"2025": {}, "2026": {}}
    if not os.path.exists(vehicle_file_path):
        return catalog

    xls = pd.ExcelFile(vehicle_file_path)
    for sheet in xls.sheet_names:
        if sheet in ['Structure', 'MY-2025', 'MY-2026', 'Combined 2025-2026', 'Bank Details', 'RMC']: 
            continue
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            raw_name = str(df.iloc[4, 2]).strip()
            code = str(df.iloc[4, 3]).strip()
            year_string = str(df.iloc[4, 4]).strip()
            
            if not raw_name or raw_name == "nan": raw_name = sheet
            if not code or code == "nan": code = sheet
            
            year_key = "2026" if "2026" in year_string or "26" in sheet else "2025"
            
            base_price = float(df.iloc[6, 1])          
            interest_rate = float(df.iloc[18, 3])  
            
            registration_fee = float(df.iloc[11, 7]) if pd.notna(df.iloc[11, 7]) else 600.0
            processing_fee_dp = float(df.iloc[12, 7]) if pd.notna(df.iloc[12, 7]) else 315.0
            
            model_name = get_clean_model_name(raw_name)
            sheet_upper = sheet.upper()
            
            if model_name in ['Attrage', 'Mirage'] or 'MIG' in sheet_upper or 'ATG' in sheet_upper or 'MOP' in sheet_upper or 'MOG' in sheet_upper:
                reservation_fee = 500.0
            else:
                reservation_fee = 1000.0

            accessories = {}
            row_labels = {
                10: "Custom Accessories", 11: "Ceramic Gold Window Tint", 12: "FO PPF Gold Package",
                13: "Extended Warranty", 14: "VRI", 15: "Vehicle Insurance", 16: "RMC"
            }
            
            for row_idx, default_label in row_labels.items():
                cell_label = str(df.iloc[row_idx - 1, 1]).strip()
                status = str(df.iloc[row_idx - 1, 2]).strip().upper()
                label = cell_label if (cell_label and cell_label != "nan") else default_label
                is_checked = (status == "YES")
                try:
                    price_val = float(df.iloc[row_idx - 1, 3])
                except:
                    price_val = 0.0
                
                accessories[label] = {
                    "price_raw": price_val,
                    "default_checked": is_checked,
                    "type_tag": "VRI" if row_idx == 14 else ("INSURANCE" if row_idx == 15 else ("RMC" if row_idx == 16 else "STANDARD"))
                }
            
            if model_name not in catalog[year_key]: 
                catalog[year_key][model_name] = {}
                
            catalog[year_key][model_name][code] = {
                "base_price": base_price, 
                "interest_rate": interest_rate, 
                "accessories": accessories,
                "registration_fee": registration_fee,
                "processing_fee_dp": processing_fee_dp,
                "reservation_fee": reservation_fee
            }
        except: 
            continue
    return catalog

FILE_VEHICLES = "NFC New VRI Project (2) (2).xlsx"  
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"

VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)

if "view_state" not in st.session_state:
    st.session_state.view_state = "input"

if not VEHICLE_CATALOG["2025"] and not VEHICLE_CATALOG["2026"]:
    st.error(f"Could not load vehicle datasets from '{FILE_VEHICLES}'.")
else:
    with st.sidebar:
        st.header("🚗 Configuration Console")
        selected_year = st.selectbox("Model Year:", sorted(list(VEHICLE_CATALOG.keys())))
        available_names = sorted(list(VEHICLE_CATALOG[selected_year].keys()))
        selected_name = st.selectbox("Vehicle Name:", available_names)
        
        if selected_name:
            available_codes = sorted(list(VEHICLE_CATALOG[selected_year][selected_name].keys()))
            selected_code = st.selectbox("Variant Code:", available_codes)
            v_data = VEHICLE_CATALOG[selected_year][selected_name][selected_code]
        else:
            st.stop()
            
        st.markdown("---")
        st.subheader("🏦 Financial Provider Rates")
        
        if BANK_RULES:
            bank_options = sorted(list(BANK_RULES.keys()))
            selected_bank = st.selectbox("Select Institution:", bank_options)
            bracket_options = sorted(list(BANK_RULES[selected_bank].keys()))
            selected_bracket = st.selectbox("Income Bracket Selection:", bracket_options)
            fetched_roi = BANK_RULES[selected_bank][selected_bracket]
        else:
            selected_bank = "Sheet Benchmark"
            fetched_roi = v_data["interest_rate"]
            
        bank_rate = st.number_input("Flat Interest Rate (ROI):", value=fetched_roi, format="%.4f", step=0.0001)
        
        st.markdown("---")
        st.subheader("⚙️ Calculations Adjuster")
        base_vehicle_price = st.number_input("Base Vehicle Price (AED):", value=v_data["base_price"], step=500.0)
        down_payment_pct = st.slider("Down Payment Percentage (%):", 0, 100, 20) / 100.0

        # EXCEL EXACT DOWN PAYMENT FINANCE RULES
        st.markdown("---")
        st.subheader("💳 Down Payment Financing Plan")
        finance_dp_option = st.toggle("Finance the Down Payment?", value=False)
        
        if finance_dp_option:
            dp_tenor_selection = st.selectbox("Select DP Term Structure:", ["3 Months (0.00% ROI)", "12 Months (5.25% ROI)", "24 Months (6.30% ROI)"])
            if "3 Months" in dp_tenor_selection:
                dp_interest_rate = 0.0000
                dp_months = 3
            elif "12 Months" in dp_tenor_selection:
                dp_interest_rate = 0.0525
                dp_months = 12
            else:
                dp_interest_rate = 0.0630
                dp_months = 24
        else:
            dp_interest_rate = 0.0
            dp_months = 0

        st.markdown("---")
        st.subheader("➕ Custom Accessories Checklists")
        acc_selected_price = 0.0   
        ceramic_selected_price = 0.0 
        foppfgoldpackage_selected_price = 0.0 
        warranty_selected_price = 0.0 
        rmc_selected_cost = 0.0     
        
        override_rmc_active = (RMC_RULES and selected_code in RMC_RULES)
        checked_addons_list = []
        
        for name, info in v_data["accessories"].items():
            if info["type_tag"] == "STANDARD":
                display_name = name
                if "FOPPF" in name.upper() or "GOLD PACKAGE" in name.upper(): display_name = "FO PPF GOLD PACKAGE"
                elif name.strip().upper() in ["ACCESSORY", "ACCESSORIES"]: display_name = "Custom Accessories"

                checked = st.checkbox(f"{display_name} (+{info['price_raw']:,.2f} AED)", value=info["default_checked"], key=f"cb_{name}")
                if checked:
                    if "CUSTOM" in name.upper() and "ACCESSORIES" in name.upper(): acc_selected_price = info["price_raw"]
                    elif "CERAMIC" in name.upper() and "WINDOW" in name.upper(): ceramic_selected_price = info["price_raw"]
                    elif "FOPPF" in name.upper() or "GOLD PACKAGE" in name.upper(): foppfgoldpackage_selected_price = info["price_raw"]
                    elif "WARRANTY" in name.upper(): warranty_selected_price = info["price_raw"]
                    else: acc_selected_price += info["price_raw"]
                    checked_addons_list.append({"name": display_name, "price": info["price_raw"], "vat_taxable": True})

            elif info["type_tag"] == "RMC" and not override_rmc_active:
                checked = st.checkbox(f"{name} (+{info['price_raw']:,.2f} AED)", value=info["default_checked"], key=f"cb_{name}")
                if checked:
                    rmc_selected_cost = info["price_raw"]
                    checked_addons_list.append({"name": name, "price": rmc_selected_cost, "vat_taxable": False})

        if override_rmc_active:
            rmc_packages = ["None"] + list(RMC_RULES[selected_code].keys())
            chosen_rmc = st.selectbox("Routine Maintenance Contract (RMC):", rmc_packages)
            if chosen_rmc != "None":
                rmc_selected_cost = RMC_RULES[selected_code][chosen_rmc]
                checked_addons_list.append({"name": f"Routine Maintenance Contract ({chosen_rmc})", "price": rmc_selected_cost, "vat_taxable": False})

        u19_valuation_base = (base_vehicle_price + acc_selected_price + ceramic_selected_price + foppfgoldpackage_selected_price + warranty_selected_price + (rmc_selected_cost / 1.05)) * 1.05

        is_vri_selected = False
        is_insurance_selected = False
        display_vri_price = u19_valuation_base * 3.15 * 1.05 / 100
        
        code_clean = str(selected_code).strip().upper()
        name_clean = str(selected_name)
        
        if code_clean.startswith(('PR', 'HLP')) or code_clean in ["OTF03", "OTP03", "OTF06", "OTP06", "OTP06-01", "OTF08", "OTP08", "G03", "G05", "G06", "G08", "G09","G10", "G12", "P03", "P06", "P12", "G31"]:
            display_ins_price = (u19_valuation_base * 0.03 + 510) * 1.05
        elif code_clean.startswith(('H', 'P','MOP', 'MOG')) and any(num in code_clean for num in ["57", "59", "61", "62", "64"]):
            display_ins_price = (u19_valuation_base * 0.0275 + 510) * 1.05
        elif code_clean.startswith('EH') and any(num in code_clean for num in ["40", "41", "43"]):
            display_ins_price = (u19_valuation_base * 0.03 + 450) * 1.05
        else:
            display_ins_price = 3690.0 if "Xpander" in name_clean or "Destinator" in name_clean else 3625.0

        for name, info in v_data["accessories"].items():
            if info["type_tag"] == "VRI":
                is_vri_selected = st.checkbox(f"Vehicle Replacement Insurance (VRI) (+{display_vri_price:,.2f} AED)", value=info["default_checked"], key="cb_vri_insurance")
            elif info["type_tag"] == "INSURANCE":
                is_insurance_selected = st.checkbox(f"Vehicle Insurance (+{display_ins_price:,.2f} AED)", value=info["default_checked"], key="cb_car_insurance")

        vehicle_insurance_cost = display_ins_price if is_insurance_selected else 0.0
        vri_calculated_cost = display_vri_price if is_vri_selected else 0.0

        if is_vri_selected: checked_addons_list.append({"name": "Vehicle Replacement Insurance (VRI)", "price": vri_calculated_cost, "vat_taxable": False})
        if is_insurance_selected: checked_addons_list.append({"name": "Vehicle Insurance", "price": vehicle_insurance_cost, "vat_taxable": False})

        excel_addons_total = (acc_selected_price + ceramic_selected_price + foppfgoldpackage_selected_price + warranty_selected_price + vri_calculated_cost + vehicle_insurance_cost + rmc_selected_cost)
        total_vat_charges = (base_vehicle_price + acc_selected_price + ceramic_selected_price + foppfgoldpackage_selected_price + warranty_selected_price) * 0.05

        full_vehicle_value_including_addons = base_vehicle_price + excel_addons_total + total_vat_charges
        calculated_downpayment = full_vehicle_value_including_addons * down_payment_pct
        finance_amount = full_vehicle_value_including_addons - calculated_downpayment

        dp_processing_fee = 315.0  
        bank_processing_fee = finance_amount * 0.0105
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Generate Summary", use_container_width=True):
            st.session_state.view_state = "summary"

    # ------------------------------------------------------------------
    # MAIN WORKSPACE RENDERING
    # ------------------------------------------------------------------
    if st.session_state.view_state == "input":
        st.title("Mitsubishi Financial Dashboard")
        st.info("Configure variables in the sidebar panel. Then click **'Generate Summary'** to view the calculation report.")
        
    elif st.session_state.view_state == "summary":
        st.title("📄 Mitsubishi Financial Matrix Calculator")
        st.subheader(f"Unit Selected: {selected_name} — Variant {selected_code} ({selected_year})")
        st.markdown("---")

        # SECTION 1: SUMMARY SECTION
        st.header("Financial Overview")
        col_s1, col_s2, col_s3 = st.columns(3)
        col_s1.metric("Total Vehicle Value", f"{full_vehicle_value_including_addons:,.2f} AED") 
        col_s2.metric("Gross Down Payment Req.", f"{calculated_downpayment:,.2f} AED")
        col_s3.metric("Vehicle Finance Amount", f"{finance_amount:,.2f} AED")
        st.markdown("---")

        # SECTION 2: EMI BREAKDOWN MATRIX (SEPARATED TABLES ONLY)
        st.header("2. Loan Installment Breakdowns")
        
        # Primary Vehicle Loan Table (Always displayed standalone)
        st.subheader("🟢 Primary Asset Vehicle Financing")
        tenures = [1, 2, 3, 4, 5]
        vehicle_emi_results = []
        
        for years in tenures:
            months = years * 12
            total_interest = finance_amount * bank_rate * years
            monthly_emi = (finance_amount + total_interest) / months
            
            vehicle_emi_results.append({
                "Asset Term (Years)": f"{years} Years ({months} Mos)",
                "Flat ROI %": f"{bank_rate*100:.4f}%",
                "Principal Loan Block": f"{finance_amount:,.2f} AED",
                "Total Interest Accrued": f"{total_interest:,.2f} AED",
                "Monthly Vehicle EMI": f"{monthly_emi:,.2f} AED"
            })
        st.table(pd.DataFrame(vehicle_emi_results))
        
        # Down Payment Loan Table (Only appears if feature is selected)
        # Down Payment Loan Table (Displays all 3 options automatically)
        if finance_dp_option:
            st.markdown("<br>", unsafe_allow_html=True)
            st.subheader("🔵 Down Payment Loan Financing Options")
            vehicle_reservation_fee = v_data["reservation_fee"]
            dp_financed_base = max(0.0, calculated_downpayment - vehicle_reservation_fee)
            
            # Define the 3 options
            dp_options = [
                {"months": 3, "rate": 0.0000, "label": "3 Months (0.00% ROI)"},
                {"months": 12, "rate": 0.0525, "label": "12 Months (5.25% ROI)"},
                {"months": 24, "rate": 0.0630, "label": "24 Months (6.30% ROI)"}
            ]
            
            dp_results = []
            for opt in dp_options:
                total_interest = dp_financed_base * opt["rate"] * (opt["months"] / 12.0)
                monthly_emi = (dp_financed_base + total_interest) / opt["months"]
                
                dp_results.append({
                    "Term": opt["label"],
                    "Financed Balance": f"{dp_financed_base:,.2f} AED",
                    "Total Interest": f"{total_interest:,.2f} AED",
                    "Monthly EMI": f"{monthly_emi:,.2f} AED"
                })
            
            st.table(pd.DataFrame(dp_results))
            st.caption("ℹ️ *Comparison of available Down Payment loan financing terms.*")
            st.caption(f"ℹ️ *This Down Payment loan runs independently and stops entirely after Month {dp_months}.*")

        st.markdown("---")

        # SECTION 3: ACCESSORIES BREAKDOWN
        st.header("3. Accessories Breakdown")
        if checked_addons_list:
            addons_table_data = []
            total_display_addons_price = 0.0
            for addon in checked_addons_list:
                item_price = addon["price"]
                item_vat = (item_price * 0.05) if addon["vat_taxable"] else 0.0
                total_display_addons_price += (item_price + item_vat)
                addons_table_data.append({
                    "Selected Accessories / Services": addon["name"],
                    "Individual Price (Base)": f"{item_price:,.2f} AED",
                    "VAT Amount (5%)": f"{item_vat:,.2f} AED" if addon["vat_taxable"] else "0.00 AED (VAT Pre-incl.)",
                    "Total Cost (incl. VAT)": f"{(item_price + item_vat):,.2f} AED"
                })
            st.table(pd.DataFrame(addons_table_data))
        else:
            st.write("*No optional accessories selected.*")
        st.markdown("---")

        # SECTION 4: TOTAL CASH OUTLAY REQUIRED
        st.header("4. Out-of-Pocket Cash Outlay Summary")
        registration_fee = v_data["registration_fee"]
        
        if finance_dp_option:
            vehicle_reservation_fee = v_data["reservation_fee"]
            grand_total_cash_outlay = vehicle_reservation_fee + registration_fee + dp_processing_fee + bank_processing_fee
        else:
            grand_total_cash_outlay = calculated_downpayment + registration_fee + dp_processing_fee + bank_processing_fee
        
        col_out1, col_out2 = st.columns(2)
        with col_out1:
            if finance_dp_option:
                st.write(f"**Showroom Reservation Fee:** {v_data['reservation_fee']:,.2f} AED (Paid Upfront)")
                st.write(f"**Remaining Down Payment Balance:** {max(0.0, calculated_downpayment - v_data['reservation_fee']):,.2f} AED (Financed via Loan Plan)")
            else:
                st.write(f"**Full Down Payment Amount:** {calculated_downpayment:,.2f} AED (Upfront Out-of-Pocket)")
            st.write(f"**Registration Documentation Fee:** {registration_fee:,.2f} AED")
        with col_out2:
            st.write(f"**DP Processing Fee (DP PF):** {dp_processing_fee:,.2f} AED")
            st.write(f"**Bank Processing Fee (Bank PF):** {bank_processing_fee:,.2f} AED")
            
        st.markdown(
            f"""
            <div style="background-color: #F4F0EA; padding: 1.25rem 1.5rem; border-radius: 8px; border-left: 4px solid #191919; margin-top: 1.5rem;">
                <span style="font-family: 'Quicksand', sans-serif; font-weight: 700; font-size: 0.9rem; color: #555555; display: block; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.25rem;">
                    🔑 Actual Upfront Cash Required at Showroom Handover
                </span>
                <span style="font-family: 'Amethysta', serif; font-size: 1.8rem; color: #191919;">
                    {grand_total_cash_outlay:,.2f} <span style="font-size: 1.2rem;">AED</span>
                </span>
            </div>
            """, 
            unsafe_allow_html=True
        )
        st.markdown("---")

        # SECTION 5: DOCUMENTATION REQUIREMENTS & DISCLOSURES
        st.header("5. Application Requirements & Disclosures")
        st.markdown(r"""
        <div style="background-color: #F4F0EA; padding: 1.5rem; border-radius: 8px; border-left: 4px solid #191919; margin-top: 1.5rem;">
            <strong style="font-family: sans-serif; font-size: 1.1rem; color: #191919; display: block; margin-bottom: 0.75rem;">📋 Required Documentation Checklist:</strong>
            <ul style="margin-bottom: 1rem; padding-left: 1.25rem;">
                <li>Passport Copy, Digital Visa & Address Page For Indian Passport, Page #44 For Philippines Passport.</li>
                <li>Emirates ID Card Copy Both Sides.</li>
                <li>Labour Card / Free Zone / Employer ID.</li>
                <li>Copy of the UAE Driver's License Both Sides.</li>
                <li>Current Dated Salary Certificate from The Employer.</li>
                <li>Pay Slips For The Last 3 Months - [If Variance In Salary].</li>
                <li>IBAN.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)

        col_space_left, col_btn1, col_btn2, col_space_right = st.columns([1, 2, 2, 1])
        with col_btn1:
            if st.button("⬅️ Back to Input", use_container_width=True):
                st.session_state.view_state = "input"
                st.rerun()
        with col_btn2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pd.DataFrame(vehicle_emi_results).to_excel(writer, index=False, sheet_name="Vehicle Loan Matrix")
            st.download_button(
                label="💾 Save as Excel Spreadsheet",
                data=buffer.getvalue(),
                file_name=f"{selected_name.replace(' ', '_')}_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

