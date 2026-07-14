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
                
                try:
                    price_val = float(df.iloc[row_idx - 1, 3])
                except:
                    price_val = 0.0
                
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
FILE_VEHICLES = "NFC New VRI Project Separated Tabs.xlsx"
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"

VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)

if "view_state" not in st.session_state:
    st.session_state.view_state = "input"

# ------------------------------------------------------------------
# SIDEBAR - CONFIGURATION INTERFACE
# ------------------------------------------------------------------
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

        st.markdown("---")
        st.subheader("➕ Accessories & Services Checklists")
        
        acc_selected_price = 0.0   
        ceramic_selected_price = 0.0 
        exterior_selected_price = 0.0 
        warranty_selected_price = 0.0 
        rmc_selected_cost = 0.0     
        
        # Check if dynamic RMC catalog rules override standard inputs
        override_rmc_active = (RMC_RULES and selected_code in RMC_RULES)
        
        checked_addons_list = []
        is_vri_selected = False
        is_insurance_selected = False
        
        # We need a quick pass to parse standard accessories to establish the U19 Base for dynamic checkbox pricing
        temp_acc_price = 0.0
        temp_ceramic_price = 0.0
        temp_exterior_price = 0.0
        temp_warranty_price = 0.0
        temp_rmc_price = 0.0

        for name, info in v_data["accessories"].items():
            if info["type_tag"] == "STANDARD":
                if "CERAMIC" in name.upper() and "WINDOW" in name.upper():
                    temp_ceramic_price = info["price_raw"]
                elif "EXTERIOR" in name.upper() or "SCOTCH" in name.upper():
                    temp_exterior_price = info["price_raw"]
                elif "WARRANTY" in name.upper():
                    temp_warranty_price = info["price_raw"]
                else:
                    temp_acc_price = info["price_raw"]
            elif info["type_tag"] == "RMC" and not override_rmc_active:
                temp_rmc_price = info["price_raw"]

        # Formulate dynamic U19 reference base to calculate real-time label values
        temp_u19 = (base_vehicle_price + temp_acc_price + temp_ceramic_price + temp_exterior_price + temp_warranty_price + temp_rmc_price) * 1.05

        # Render checkboxes with matched calculated prices instead of raw template prices
        for name, info in v_data["accessories"].items():
            if info["type_tag"] == "RMC" and override_rmc_active:
                continue
            
            # Determine dynamic checkbox label cost
            if info["type_tag"] == "VRI":
                display_price = temp_u19 * 3.15 * 1.05 / 100
            elif info["type_tag"] == "INSURANCE":
                if selected_code in ["PR", "PRP", "HLP"]:
                    display_price = (temp_u19 * 0.03 + 510) * 1.05
                elif selected_code in ["H57", "P57", "H64", "H59", "P59", "H61", "P61"]:
                    display_price = (temp_u19 * 0.0275 + 510) * 1.05
                elif selected_code in ["EH40", "EH43"]:
                    display_price = (temp_u19 * 0.03 + 450) * 1.05
                else:
                    display_price = 3690.0 if "Xpander" in selected_name else 3625.0
            else:
                display_price = info["price_raw"]

            checked = st.checkbox(f"{name} (+{display_price:,.2f} AED)", value=info["default_checked"])
            
            if checked:
                if info["type_tag"] == "STANDARD":
                    if "CERAMIC" in name.upper() and "WINDOW" in name.upper():
                        ceramic_selected_price = display_price
                    elif "EXTERIOR" in name.upper() or "SCOTCH" in name.upper():
                        exterior_selected_price = display_price
                    elif "WARRANTY" in name.upper():
                        warranty_selected_price = display_price
                    else:
                        acc_selected_price = display_price
                    checked_addons_list.append({"name": name, "price": display_price, "vat_taxable": True})
                elif info["type_tag"] == "VRI":
                    is_vri_selected = True
                elif info["type_tag"] == "INSURANCE":
                    is_insurance_selected = True
                elif info["type_tag"] == "RMC":
                    rmc_selected_cost = display_price
                    checked_addons_list.append({"name": name, "price": display_price, "vat_taxable": False})
                    
        # Render dynamic drop-down selection only if active
        if override_rmc_active:
            rmc_packages = ["None"] + list(RMC_RULES[selected_code].keys())
            chosen_rmc = st.selectbox("Routine Maintenance Contract (RMC):", rmc_packages)
            if chosen_rmc != "None":
                rmc_selected_cost = RMC_RULES[selected_code][chosen_rmc]
                checked_addons_list.append({"name": f"Routine Maintenance Contract ({chosen_rmc})", "price": rmc_selected_cost, "vat_taxable": False})

        # ==========================================
        # HIGH-FIDELITY EXCEL MATCHING MATH ENGINE (U19-BASED)
        # ==========================================
        # Step 1: Replicate the dynamic U19/V19 total asset valuation formula from the sheet
        u19_valuation_base = (
            base_vehicle_price + 
            acc_selected_price + 
            ceramic_selected_price + 
            exterior_selected_price + 
            warranty_selected_price + 
            rmc_selected_cost
        ) * 1.05

        # Step 2: Calculate Vehicle Insurance directly using the U19 base value
        if is_insurance_selected:
            if selected_code in ["PR", "PRP", "HLP"]:
                vehicle_insurance_cost = (u19_valuation_base * 0.03 + 510) * 1.05
            elif selected_code in ["H57", "P57", "H64", "H59", "P59", "H61", "P61"]:
                vehicle_insurance_cost = (u19_valuation_base * 0.0275 + 510) * 1.05
            elif selected_code in ["EH40", "EH43"]:
                vehicle_insurance_cost = (u19_valuation_base * 0.03 + 450) * 1.05
            else:
                vehicle_insurance_cost = 3690.0 if "Xpander" in selected_name else 3625.0
        else:
            vehicle_insurance_cost = 0.0

        # Step 3: Calculate VRI premium directly using the U19 base value
        vri_calculated_cost = (u19_valuation_base * 3.15 * 1.05 / 100) if is_vri_selected else 0.0

        # Inject Insurance and VRI into checked_addons_list for reporting visibility
        if is_vri_selected:
            checked_addons_list.append({"name": "Value Retention Insurance (VRI)", "price": vri_calculated_cost, "vat_taxable": False})
        if is_insurance_selected:
            checked_addons_list.append({"name": "Vehicle Insurance", "price": vehicle_insurance_cost, "vat_taxable": False})

        # Step 4: Aggregate Final Balances
        excel_addons_total = (
            acc_selected_price + 
            ceramic_selected_price + 
            exterior_selected_price + 
            warranty_selected_price + 
            vri_calculated_cost + 
            vehicle_insurance_cost + 
            rmc_selected_cost
        )

        # Total 5% VAT tracking for standard taxable accessories
        total_vat_charges = (base_vehicle_price + acc_selected_price + ceramic_selected_price + exterior_selected_price + warranty_selected_price) * 0.05

        # Final Contract Values
        full_vehicle_value_including_addons = base_vehicle_price + excel_addons_total + total_vat_charges
        calculated_downpayment = full_vehicle_value_including_addons * down_payment_pct
        finance_amount = full_vehicle_value_including_addons - calculated_downpayment

        # Bank Fees (1.05% of final net financed principal)
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
        col_s1, col_s2, col_s3, col_s4 = st.columns(4)
        col_s1.metric("Vehicle Model", f"{selected_name} ({selected_code})")
        col_s2.metric("Total Vehicle Value", f"{full_vehicle_value_including_addons:,.2f} AED") # Fixed: Now shows full value including accessories/insurance/VAT
        col_s3.metric("Down Payment Amount", f"{calculated_downpayment:,.2f} AED")
        col_s4.metric("Finance Amount", f"{finance_amount:,.2f} AED")
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
            st.write(f"**Total Accessories Cost:** {total_display_addons_price:,.2f} AED")
        else:
            st.write("*No optional accessories selected.*")
        st.markdown("---")

        # SECTION 4: TOTAL CASH OUTLAY REQUIRED
        st.header("4. Total Cash Outlay")
        registration_fee = v_data["registration_fee"]
        processing_fee_dp = v_data["processing_fee_dp"]
        total_insurance_and_vri = vehicle_insurance_cost + vri_calculated_cost
        
        # Grand total required to take the car
        grand_total_cash_outlay = calculated_downpayment + registration_fee + processing_fee_dp + bank_processing_fee
        
        col_out1, col_out2 = st.columns(2)
        with col_out1:
            st.write(f"**Down Payment:** {calculated_downpayment:,.2f} AED")
            st.write(f"**Accessories Total (Gross):** {total_display_addons_price:,.2f} AED")
        with col_out2:
            st.write(f"**Insurance Costs (Vehicle + VRI):** {total_insurance_and_vri:,.2f} AED")
            st.write(f"**Processing Fees (DP PF + Bank PF):** {(processing_fee_dp + bank_processing_fee):,.2f} AED")
            
        st.markdown(f"### 🔑 **Grand Total Required to Take the Car:** {grand_total_cash_outlay:,.2f} AED")
        st.markdown("---")

        # SECTION 5: BUTTONS
        st.header("5. Buttons")
        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("⬅️ Back to Input", use_container_width=True):
                st.session_state.view_state = "input"
                st.rerun()
        with col_btn2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                pd.DataFrame(emi_results).to_excel(writer, index=False, sheet_name="EMI Matrix")
            st.download_button(
                label="💾 Save as Excel / PDF Document",
                data=buffer.getvalue(),
                file_name=f"{selected_name.replace(' ', '_')}_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_btn3:
            st.button("✉️ Email Results", use_container_width=True, disabled=True)
