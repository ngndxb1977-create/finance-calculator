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
# MASTER EXCEL EXTRACTION ENGINE
# ------------------------------------------------------------------
@st.cache_data
def load_supplementary_data(file_path):
    bank_data = {}
    rmc_data = {}
    if os.path.exists(file_path):
        # Parse Bank Details
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
        except:
            pass

        # Parse RMC Pricing Map
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
            
            # Base Coordinates
            base_price = float(df.iloc[6, 1])          
            interest_rate = float(df.iloc[18, 3])  
            
            # Mandatory Static Fees from Sheet (H12 & H13)
            registration_fee = float(df.iloc[11, 7]) if pd.notna(df.iloc[11, 7]) else 600.0
            processing_fee_dp = float(df.iloc[12, 7]) if pd.notna(df.iloc[12, 7]) else 315.0
            
            # Map Row 10 to 16 Accessories
            accessories = {}
            row_labels = {
                10: "Accessory",
                11: "Ceramic Gold Window Tint",
                12: "Exterior Scotchguard Protection",
                13: "Extended Warranty",
                14: "VRI",
                15: "Vehicle Insurance",
                16: "RMC"
            }
            
            for row_idx, default_label in row_labels.items():
                cell_label = str(df.iloc[row_idx - 1, 1]).strip()
                status = str(df.iloc[row_idx - 1, 2]).strip().upper()
                
                label = cell_label if (cell_label and cell_label != "nan") else default_label
                is_checked = (status == "YES")
                
                # Retrieve raw sheet values as baseline rates
                try:
                    price_val = float(df.iloc[row_idx - 1, 3])
                except:
                    price_val = 0.0
                
                # Tag system features
                accessories[label] = {
                    "price_raw": price_val,
                    "default_checked": is_checked,
                    "type_tag": "VRI" if row_idx == 14 else ("INSURANCE" if row_idx == 15 else ("RMC" if row_idx == 16 else "STANDARD"))
                }
            
            model_name = get_clean_model_name(raw_name)
            if model_name not in catalog[year_key]: 
                catalog[year_key][model_name] = {}
                
            catalog[year_key][model_name][code] = {
                "base_price": base_price, 
                "interest_rate": interest_rate, 
                "accessories": accessories,
                "registration_fee": registration_fee,
                "processing_fee_dp": processing_fee_dp
            }
        except: 
            continue
                
    return catalog

# ------------------------------------------------------------------
# CONFIG & FILE TARGETS
# ------------------------------------------------------------------
FILE_VEHICLES = "NFC New VRI Project (2).xlsx"
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"

VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)

if "view_state" not in st.session_state:
    st.session_state.view_state = "input"

# ------------------------------------------------------------------
# SIDEBAR - CONFIGURATION INTERFACE
# ------------------------------------------------------------------
if not VEHICLE_CATALOG["2025"] and not VEHICLE_CATALOG["2026"]:
    st.error(f"Could not load vehicle datasets from '{FILE_VEHICLES}'. Please confirm it is stored in your repo.")
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

        st.markdown("---")
        st.subheader("➕ Accessories & Services Checklists")
        
        # Build live variables to replicate exact Sheet Row-Formulas
        acc_selected_price = 0.0   # Row 10
        ceramic_selected_price = 0.0 # Row 11
        exterior_selected_price = 0.0 # Row 12
        warranty_selected_price = 0.0 # Row 13
        rmc_selected_cost = 0.0     # Row 16
        
        # Override RMC with lookup-table rule if present
        override_rmc_active = (RMC_RULES and selected_code in RMC_RULES)
        
        checked_addons_list = []
        is_vri_selected = False
        is_insurance_selected = False
        
        for name, info in v_data["accessories"].items():
            checked = st.checkbox(f"{name} (+{info['price_raw']:,.2f} AED)", value=info["default_checked"])
            
            if checked:
                p = info["price_raw"]
                if info["type_tag"] == "STANDARD":
                    if "CERAMIC" in name.upper() and "WINDOW" in name.upper():
                        ceramic_selected_price = p
                    elif "EXTERIOR" in name.upper() or "SCOTCH" in name.upper():
                        exterior_selected_price = p
                    elif "WARRANTY" in name.upper():
                        warranty_selected_price = p
                    else:
                        acc_selected_price = p
                    checked_addons_list.append({"name": name, "price": p, "vat_taxable": True})
                elif info["type_tag"] == "VRI":
                    is_vri_selected = True
                elif info["type_tag"] == "INSURANCE":
                    is_insurance_selected = True
                elif info["type_tag"] == "RMC" and not override_rmc_active:
                    rmc_selected_cost = p
                    checked_addons_list.append({"name": name, "price": p, "vat_taxable": True})
                    
        # Apply external dynamic RMC table if detected
        if override_rmc_active:
            rmc_packages = ["None"] + list(RMC_RULES[selected_code].keys())
            chosen_rmc = st.selectbox("Dynamic Regional Maintenance Contract (RMC):", rmc_packages)
            if chosen_rmc != "None":
                rmc_selected_cost = RMC_RULES[selected_code][chosen_rmc]
                checked_addons_list.append({"name": f"RMC Package ({chosen_rmc})", "price": rmc_selected_cost, "vat_taxable": True})

        # --------------------------------------------------------------
        # HIGH-FIDELITY EXCEL CORE MATH REPLICATION ENGINE
        # --------------------------------------------------------------
        # V19 Formula: (Base Price + Acc + Ceramic + Exterior + Warranty + RMC) * 1.05
        total_taxable_base = (
            base_vehicle_price + 
            acc_selected_price + 
            ceramic_selected_price + 
            exterior_selected_price + 
            warranty_selected_price + 
            rmc_selected_cost
        )
        v19_valuation = total_taxable_base * 1.05
        
        # Row 14 VRI Formula: V19 * 3.15% * 1.05
        vri_calculated_cost = (v19_valuation * 0.0315 * 1.05) if is_vri_selected else 0.0
        if is_vri_selected:
            checked_addons_list.append({"name": "Value Retention Insurance (VRI)", "price": vri_calculated_cost, "vat_taxable": False})
            
        # Row 15 Vehicle Insurance Formula: Base Price * 3% + 130 AED (replicates =(B7*3%)+10+120)
        vehicle_insurance_cost = (base_vehicle_price * 0.03 + 130) if is_insurance_selected else 0.0
        if is_insurance_selected:
            checked_addons_list.append({"name": "Vehicle Insurance", "price": vehicle_insurance_cost, "vat_taxable": False})

        # Total Add-Ons Cost (Row 17 Value: SUM D10:D16)
        excel_addons_total = (
            acc_selected_price + 
            ceramic_selected_price + 
            exterior_selected_price + 
            warranty_selected_price + 
            vri_calculated_cost + 
            vehicle_insurance_cost + 
            rmc_selected_cost
        )
        
        # Total VAT Charges: 5% VAT on base vehicle price and standard taxable add-ons
        vat_vehicle = base_vehicle_price * 0.05
        vat_addons_taxable = (
            acc_selected_price + 
            ceramic_selected_price + 
            exterior_selected_price + 
            warranty_selected_price + 
            rmc_selected_cost
        ) * 0.05
        total_vat_charges = vat_vehicle + vat_addons_taxable
        
        # G7: Full Vehicle Value Including Add-Ons (B7 + E7 [Add-Ons] + F7 [Total VAT])
        full_vehicle_value_including_addons = base_vehicle_price + excel_addons_total + total_vat_charges
        
        # H7 Down Payment: G7 * DP %
        calculated_downpayment = full_vehicle_value_including_addons * down_payment_pct
        
        # I7 Finance Amount: G7 - H7
        finance_amount = full_vehicle_value_including_addons - calculated_downpayment
        
        # Row 16 Bank Processing Fee Formula: Finance Amount * 1.05%
        bank_processing_fee = finance_amount * 0.0105

        # Controls Action Button
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📊 Generate Complete Summary Report", use_container_width=True):
            st.session_state.view_state = "summary"

    # ------------------------------------------------------------------
    # MAIN WORKSPACE RENDERING
    # ------------------------------------------------------------------
    if st.session_state.view_state == "input":
        st.title("Mitsubishi Financial Dashboard")
        st.info("Configure your vehicle specs, accessories, and bank details in the sidebar panel. Then click **'Generate Complete Summary Report'** to run the matching Excel calculation engine.")
        
    elif st.session_state.view_state == "summary":
        st.title("📄 High-Fidelity Financial Summary Report")
        st.subheader(f"Unit Selected: {selected_name} — Variant {selected_code} ({selected_year})")
        st.markdown("---")

        # SECTION 1: SUMMARY SECTION
        st.header("1. Summary Section")
        cols1 = st.columns(4)
        cols1[0].metric("Vehicle Model", f"{selected_name}")
        cols1[1].metric("Total Vehicle Value (incl. VAT & Add-ons)", f"{full_vehicle_value_including_addons:,.2f} AED")
        cols1[2].metric("Down Payment Amount", f"{calculated_downpayment:,.2f} AED")
        cols1[3].metric("Total Finance Amount", f"{finance_amount:,.2f} AED")
        st.markdown("---")

        # SECTION 2: EMI BREAKDOWN
        st.header("2. EMI Breakdown")
        tenures = [2, 3, 4, 5]
        emi_results = []
        
        for years in tenures:
            months = years * 12
            total_interest = finance_amount * bank_rate * years
            total_repayable = finance_amount + total_interest
            monthly_emi = total_repayable / months
            
            emi_results.append({
                "Term (Years)": f"{years} Years ({months} Months)",
                "Applied Interest Rate": f"{bank_rate*100:.4f}%",
                "Finance Principal Amount": f"{finance_amount:,.2f} AED",
                "Total Interest Costs": f"{total_interest:,.2f} AED",
                "Monthly EMI Plan": f"{monthly_emi:,.2f} AED"
            })
        st.table(pd.DataFrame(emi_results))
        st.markdown("---")

        # SECTION 3: ACCESSORIES BREAKDOWN
        st.header("3. Accessories Breakdown")
        if checked_addons_list:
            addons_table_data = []
            for addon in checked_addons_list:
                item_price = addon["price"]
                item_vat = (item_price * 0.05) if addon["vat_taxable"] else 0.0
                addons_table_data.append({
                    "Add-on Item": addon["name"],
                    "Price (AED)": f"{item_price:,.2f} AED",
                    "VAT Amount (5%)": f"{item_vat:,.2f} AED" if addon["vat_taxable"] else "0.00 AED (VAT Pre-incl./Exempt)"
                })
            st.table(pd.DataFrame(addons_table_data))
            st.write(f"**Total Accessories (Sum of Add-ons):** {excel_addons_total:,.2f} AED")
        else:
            st.write("*No optional accessories selected.*")
        st.markdown("---")

        # SECTION 4: TOTAL CASH OUTLAY REQUIRED
        st.header("4. Total Cash Outlay")
        registration_fee = v_data["registration_fee"]
        processing_fee_dp = v_data["processing_fee_dp"]
        
        # Grand Total Required to Take Car: Down-Payment + Registration + DP PF + Bank PF
        grand_total_cash_outlay = calculated_downpayment + registration_fee + processing_fee_dp + bank_processing_fee
        
        col_out1, col_out2 = st.columns(2)
        with col_out1:
            st.write(f"**Down Payment Required:** {calculated_downpayment:,.2f} AED")
            st.write(f"**Registration Fee:** {registration_fee:,.2f} AED")
        with col_out2:
            st.write(f"**DP Processing Fee:** {processing_fee_dp:,.2f} AED")
            st.write(f"**Bank Processing Fee (1.05% of Principal):** {bank_processing_fee:,.2f} AED")
            
        st.markdown(f"### 🔑 **Grand Total Cash Required to Take Delivery:** {grand_total_cash_outlay:,.2f} AED")
        st.markdown("---")

        # SECTION 5: ACTIONS
        st.header("5. Action Toolbar")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("⬅️ Modify Parameters (Back to Input)", use_container_width=True):
                st.session_state.view_state = "input"
                st.rerun()
        with col_btn2:
            # Excel export script
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pd.DataFrame(emi_results).to_excel(writer, index=False, sheet_name="EMI Matrix")
            st.download_button(
                label="📥 Save Summary as Excel Document",
                data=buffer.getvalue(),
                file_name=f"{selected_name.replace(' ', '_')}_{selected_code}_Summary_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_btn3:
            st.button("✉️ Email Summary PDF (Placeholder)", use_container_width=True, disabled=True)
