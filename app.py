import streamlit as st
import pandas as pd
import numpy as np
import os
import io

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(page_title="Mitsubishi Financial Matrix Calculator", layout="wide")

# ------------------------------------------------------------
# CUSTOM UI STYLING
# ------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Karma:wght@400;600&family=Amethysta&family=Quicksand:wght@500;700&display=swap');

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

[data-testid="stMetricValue"] {
    font-family: 'Amethysta', serif !important;
    font-size: 1.6rem !important;
    font-weight: 400 !important;
    color: #191919 !important;
}

[data-testid="stMetricLabel"] {
    font-family: 'Quicksand', sans-serif !important;
    font-size: 0.95rem !important;
    font-weight: 600 !important;
    color: #555555 !important;
}
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# CLEAN MODEL NAME
# ------------------------------------------------------------
def get_clean_model_name(raw_name):
    raw = raw_name.upper()
    if "ATTRAGE" in raw: return "Attrage"
    if "MIRAGE" in raw: return "Mirage"
    if "ASX" in raw: return "ASX"
    if "ECLIPSE" in raw: return "Eclipse Cross"
    if "XPANDER" in raw: return "Xpander"
    if "OUTLANDER" in raw: return "Outlander"
    if "MONTERO" in raw: return "Montero Sport"
    if "DESTINATOR" in raw: return "Destinator"
    return raw_name

# ------------------------------------------------------------
# LOAD SUPPLEMENTARY BANK & RMC DATA
# ------------------------------------------------------------
@st.cache_data
def load_supplementary_data(file_path):
    bank_data = {}
    rmc_data = {}

    if os.path.exists(file_path):
        try:
            df_bank = pd.read_excel(file_path, sheet_name="Bank Details")
            for _, row in df_bank.iterrows():
                bank = str(row["Bank Name"]).strip()
                if not bank or bank == "nan": continue
                bank_data[bank] = {}

                for i in range(1, 6):
                    sb = f"Salary Bracket.{i}" if i > 1 else "Salary Bracket"
                    roi = f"ROI.{i}" if i > 1 else "ROI"
                    if sb in df_bank.columns and roi in df_bank.columns:
                        sb_val = str(row[sb]).strip()
                        roi_val = row[roi]
                        if sb_val and sb_val != "nan" and pd.notna(roi_val):
                            bank_data[bank][sb_val] = float(roi_val)
        except:
            pass

        try:
            df_rmc = pd.read_excel(file_path, sheet_name="RMC")
            for _, row in df_rmc.iterrows():
                code = str(row["Variant Codes"]).strip()
                if not code or code == "nan": continue
                rmc_data[code] = {
                    "RMC-10-40": float(row["RMC-10-40"]) if pd.notna(row["RMC-10-40"]) else 0.0,
                    "RMC-10-60": float(row["RMC-10-60"]) if pd.notna(row["RMC-10-60"]) else 0.0,
                    "RMC-10-70": float(row["RMC-10-70"]) if pd.notna(row["RMC-10-70"]) else 0.0,
                    "RMC-10-100": float(row["RMC-10-100"]) if pd.notna(row["RMC-10-100"]) else 0.0,
                }
        except:
            pass

    return bank_data, rmc_data

# ------------------------------------------------------------
# LOAD VEHICLE CATALOG
# ------------------------------------------------------------
@st.cache_data
def load_all_vehicle_data(vehicle_file_path):
    catalog = {"2025": {}, "2026": {}}
    if not os.path.exists(vehicle_file_path):
        return catalog

    xls = pd.ExcelFile(vehicle_file_path)

    for sheet in xls.sheet_names:
        if sheet in ["Structure", "MY-2025", "MY-2026", "Combined 2025-2026", "Bank Details", "RMC"]:
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
                    "type_tag": (
                        "VRI" if row_idx == 14 else
                        "INSURANCE" if row_idx == 15 else
                        "RMC" if row_idx == 16 else
                        "STANDARD"
                    )
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

# ------------------------------------------------------------
# LOAD FILES
# ------------------------------------------------------------
FILE_VEHICLES = "NFC New VRI Project (2).xlsx"
FILE_SUPPLEMENT = "Bank & RMC Details.xlsx"

VEHICLE_CATALOG = load_all_vehicle_data(FILE_VEHICLES)
BANK_RULES, RMC_RULES = load_supplementary_data(FILE_SUPPLEMENT)

if "view_state" not in st.session_state:
    st.session_state.view_state = "input"

# ------------------------------------------------------------
# SIDEBAR CONFIG
# ------------------------------------------------------------
with st.sidebar:
    st.header("Configuration Console")

    selected_year = st.selectbox("Model Year:", sorted(list(VEHICLE_CATALOG.keys())))
    available_names = sorted(list(VEHICLE_CATALOG[selected_year].keys()))
    selected_name = st.selectbox("Vehicle Name:", available_names)

    available_codes = sorted(list(VEHICLE_CATALOG[selected_year][selected_name].keys()))
    selected_code = st.selectbox("Variant Code:", available_codes)

    v_data = VEHICLE_CATALOG[selected_year][selected_name][selected_code]

    st.markdown("---")
    st.subheader("Financial Provider Rates")

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
    st.subheader("Calculations Adjuster")

    base_vehicle_price = st.number_input("Base Vehicle Price (AED):", value=v_data["base_price"], step=500.0)
    down_payment_pct = st.slider("Down Payment Percentage (%):", 0, 100, 20) / 100.0

    st.markdown("---")
    st.subheader("+ Custom Accessories & Services Checklists")

# ------------------------------------------------------------
# ACCESSORY SELECTION
# ------------------------------------------------------------
acc_selected_price = 0.0
ceramic_selected_price = 0.0
exterior_selected_price = 0.0
warranty_selected_price = 0.0
rmc_selected_cost = 0.0

override_rmc_active = (RMC_RULES and selected_code in RMC_RULES)
checked_addons_list = []
is_vri_selected = False
is_insurance_selected = False

# ------------------------------------------------------------
# DYNAMIC U19 BASE
# ------------------------------------------------------------
def compute_insurance_cost(code, model_name, u19):
    code = code.strip().upper()

    if code.startswith(("PR", "HLP")) or code in [
        "G08","G03","G05","G06","G09","G10","G12",
        "P03","P05","P06","P08","P09","P10","P12",
        "G31"
    ]:
        return (u19 * 0.03 + 510) * 1.05

    elif code.startswith(("H", "P")) and any(x in code for x in ["57","59","61","62","64"]):
        return (u19 * 0.0275 + 510) * 1.05

    elif code.startswith("EH") and any(x in code for x in ["40","41","43"]):
        return (u19 * 0.03 + 450) * 1.05

    if "XPANDER" in model_name.upper(): return 3690.0
    if "DESTINATOR" in model_name.upper(): return 3690.0
    return 3625.0

def compute_vri_cost(u19):
    return (u19 * 3.15 * 1.05 / 100)

# ------------------------------------------------------------
# FIRST PASS: TEMPLATE ACCESSORY VALUES
# ------------------------------------------------------------
temp_acc_price = 0.0
temp_ceramic_price = 0.0
temp_exterior_price = 0.0
temp_warranty_price = 0.0
temp_rmc_price = 0.0

for name, info in v_data["accessories"].items():
    if info["type_tag"] == "STANDARD":
        if "CERAMIC" in name.upper():
            temp_ceramic_price = info["price_raw"]
        elif "EXTERIOR" in name.upper() or "SCOTCH" in name.upper():
            temp_exterior_price = info["price_raw"]
        elif "WARRANTY" in name.upper():
            temp_warranty_price = info["price_raw"]
        else:
            temp_acc_price = info["price_raw"]

    elif info["type_tag"] == "RMC" and not override_rmc_active:
        temp_rmc_price = info["price_raw"]

temp_u19 = (base_vehicle_price + temp_acc_price + temp_ceramic_price +
            temp_exterior_price + temp_warranty_price + temp_rmc_price) * 1.05

# ------------------------------------------------------------
# SECOND PASS: RENDER CHECKBOXES USING UNIFIED PRICING
# ------------------------------------------------------------
for name, info in v_data["accessories"].items():

    if info["type_tag"] == "RMC" and override_rmc_active:
        continue

    # Unified dynamic U19 base
    u19_valuation_base = (
        base_vehicle_price +
        acc_selected_price +
        ceramic_selected_price +
        exterior_selected_price +
        warranty_selected_price +
        (rmc_selected_cost / 1.05)
    ) * 1.05

    if info["type_tag"] == "VRI":
        display_price = compute_vri_cost(u19_valuation_base)

    elif info["type_tag"] == "INSURANCE":
        display_price = compute_insurance_cost(selected_code, selected_name, u19_valuation_base)

    else:
        display_price = info["price_raw"]

    checked = st.sidebar.checkbox(f"{name} (+{display_price:,.2f} AED)", value=info["default_checked"])

    if checked:
        if info["type_tag"] == "STANDARD":
            if "CERAMIC" in name.upper():
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

# ------------------------------------------------------------
# RMC DROPDOWN (IF OVERRIDE ACTIVE)
# ------------------------------------------------------------
if override_rmc_active:
    rmc_packages = ["None"] + list(RMC_RULES[selected_code].keys())
    chosen_rmc = st.sidebar.selectbox("Routine Maintenance Contract (RMC):", rmc_packages)

    if chosen_rmc != "None":
        rmc_selected_cost = RMC_RULES[selected_code][chosen_rmc]
        checked_addons_list.append({
            "name": f"Routine Maintenance Contract ({chosen_rmc})",
            "price": rmc_selected_cost,
            "vat_taxable": False
        })

# ------------------------------------------------------------
# FINAL INSURANCE + VRI VALUES (MATCH SIDEBAR)
# ------------------------------------------------------------
vehicle_insurance_cost = compute_insurance_cost(selected_code, selected_name, u19_valuation_base) if is_insurance_selected else 0.0
vri_calculated_cost = compute_vri_cost(u19_valuation_base) if is_vri_selected else 0.0

if is_vri_selected:
    checked_addons_list.append({"name": "Value Replacement Insurance (VRI)", "price": vri_calculated_cost, "vat_taxable": False})

if is_insurance_selected:
    checked_addons_list.append({"name": "Vehicle Insurance", "price": vehicle_insurance_cost, "vat_taxable": False})

# ------------------------------------------------------------
# TOTAL ADDONS
# ------------------------------------------------------------
excel_addons_total = (
    acc_selected_price +
    ceramic_selected_price +
    exterior_selected_price +
    warranty_selected_price +
    vri_calculated_cost +
    vehicle_insurance_cost +
    rmc_selected_cost
)

total_vat_charges = (base_vehicle_price + acc_selected_price + ceramic_selected_price +
                     exterior_selected_price + warranty_selected_price) * 0.05

full_vehicle_value_including_addons = base_vehicle_price + excel_addons_total + total_vat_charges
calculated_downpayment = full_vehicle_value_including_addons * down_payment_pct
finance_amount = full_vehicle_value_including_addons - calculated_downpayment

dp_processing_fee = 315.0
bank_processing_fee = finance_amount * 0.0105

total_cash_outlay = calculated_downpayment + dp_processing_fee + bank_processing_fee

# ------------------------------------------------------------
# MAIN VIEW
# ------------------------------------------------------------
if st.session_state.view_state == "input":
    st.title("Mitsubishi Financial Dashboard")
    st.info("Configure your vehicle specs, accessories, and bank details in the sidebar panel. Then click **'Generate Summary'**.")

    if st.button("Generate Summary"):
        st.session_state.view_state = "summary"
        st.rerun()

elif st.session_state.view_state == "summary":
    st.title("Mitsubishi Financial Matrix Calculator")
    st.subheader(f"Unit Selected: {selected_name} - Variant {selected_code} ({selected_year})")
    st.markdown("---")

    # ------------------------------------------------------------
    # SECTION 1: FINANCIAL OVERVIEW
    # ------------------------------------------------------------
    st.header("Financial Overview")
    col1, col2, col3 = st.columns
