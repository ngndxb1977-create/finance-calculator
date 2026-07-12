import streamlit as st
import pandas as pd

# Page theme configuration mapped to the modern editorial light color palette
st.set_page_config(
    page_title="Automotive Finance Calculator Dashboard", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- 1. DATA CATALOGS EXTRACTED FROM THE UPDATED TEMPLATES ---
VEHICLE_CATALOG = {
    "Attrage G16 (MY-2025)": {"base_price": 40400, "year": "2025", "vat_charges_fixed": 2290},
    "Destinator PR (MY-2026)": {"base_price": 95900, "year": "2026", "vat_charges_fixed": 5360}
}

# Add-ons items extracted explicitly from rows 8-14 of the new model sheets
ADDONS_CATALOG = {
    "2025": [
        {"name": "FO Ceramic+ Intr&Extr CeramicGold WdwTnt", "default_price": 0, "checked": False},
        {"name": "FO Exterior ceramic All Cars SCOTCHGUARD", "default_price": 0, "checked": False},
        {"name": "Extended Warranty", "default_price": 0, "checked": False},
        {"name": "VRI", "default_price": 1590.58, "checked": True},
        {"name": "Vehicle Insurance", "default_price": 3625.00, "checked": True},
        {"name": "RMC-10-70KMS", "default_price": 5400.00, "checked": True}
    ],
    "2026": [
        {"name": "FO Ceramic+ Intr&Extr CeramicGold WdwTnt", "default_price": 2700.00, "checked": True},
        {"name": "FO Exterior ceramic All Cars SCOTCHGUARD", "default_price": 0, "checked": False},
        {"name": "Extended Warranty", "default_price": 2000.00, "checked": True},
        {"name": "VRI", "default_price": 3722.92, "checked": True},
        {"name": "Vehicle Insurance", "default_price": 4081.14, "checked": True},
        {"name": "RMC-10-70KMS", "default_price": 6600.00, "checked": True}
    ]
}

# Precise interest configurations mapping standard flat rates and subvention discounts
STANDARD_INTEREST_RATE = 0.0249  # Updated baseline flat rate structure from sheet formulas
SUBVENTION_MULTIPLIERS = {1: 0.0000, 2: 0.0339, 3: 0.0339, 4: 0.0678, 5: 0.1017}

# --- 2. ADVANCED FINANCIAL ENGINE ---
def calculate_deal_metrics(base_price, year, selected_addons, dp_percentage):
    # Summing all active add-ons
    total_addons = sum([item['price'] for item in selected_addons])
    
    # Capitalization steps mirroring rows 4-7
    vat_amount = base_price * 0.05
    vehicle_with_vat = base_price + vat_amount
    
    # Calculated values including accessories and contracts
    full_value = vehicle_with_vat + total_addons
    down_payment = full_value * (dp_percentage / 100)
    finance_amount = full_value - down_payment
    
    tenure_matrix = []
    for years in [1, 2, 3, 4, 5]:
        months = years * 12
        
        # Standard Plan Amortization
        total_interest = finance_amount * STANDARD_INTEREST_RATE * years
        standard_emi = (finance_amount + total_interest) / months
        
        # Subvention Tier Adjustments
        sub_factor = SUBVENTION_MULTIPLIERS[years]
        subvention_discount = finance_amount * sub_factor
        
        # Mapping exact structural subvention monthly rates
        if years == 1:
            subvention_emi = standard_emi
        else:
            # Subvention adjustments tracking row 6 formulas
            subvention_emi = standard_emi - (subvention_discount / months) if sub_factor > 0 else standard_emi
            
        monthly_savings = standard_emi - subvention_emi
        
        tenure_matrix.append({
            "Tenure Loop": f"{years} Year(s) ({months} Mos)",
            "Standard EMI (AED)": f"{standard_emi:,.2f}",
            "Subvention Offer EMI (AED)": f"{subvention_emi:,.2f}",
            "Monthly Profit/Savings (AED)": f"{monthly_savings:,.2f}",
            "Total Plan Interest (AED)": f"{total_interest:,.2f}"
        })
        
    return {
        "summary": {
            "base_price": base_price,
            "vat_amount": vat_amount,
            "vehicle_with_vat": vehicle_with_vat,
            "total_addons": total_addons,
            "full_value": full_value,
            "down_payment": down_payment,
            "finance_amount": finance_amount
        },
        "matrix": pd.DataFrame(tenure_matrix)
    }

# --- 3. STREAMLIT FRONTEND USER INTERFACE ---
st.title("🚗 Automotive Portfolio Finance Builder")
st.markdown("Updated platform mapped directly to the new MY-2025 and MY-2026 system sheets.")
st.divider()

# Left Parameters Control Panel
st.sidebar.header("🛠️ Setup & Parameters")

selected_model = st.sidebar.selectbox("Select Vehicle Model Variant", list(VEHICLE_CATALOG.keys()))
model_meta = VEHICLE_CATALOG[selected_model]
model_year = model_meta["year"]

custom_price = st.sidebar.number_input(
    "Base Vehicle Price (AED)", 
    value=int(model_meta["base_price"]), 
    step=500
)

st.sidebar.subheader("📦 Add-Ons & Contract Inclusions")
active_addons = []
available_addons = ADDONS_CATALOG[model_year]

for addon in available_addons:
    is_active = st.sidebar.checkbox(f"{addon['name']}", value=addon['checked'])
    custom_addon_price = st.sidebar.number_input(
        f"Price: {addon['name']} (AED)", 
        value=float(addon['default_price']), 
        step=50.0, 
        key=f"input_{addon['name']}"
    )
    if is_active:
        active_addons.append({"name": addon['name'], "price": custom_addon_price})

down_payment_pct = st.sidebar.slider("Down Payment Allocation (%)", min_value=10, max_value=80, value=20, step=5)

# Engine Execution
deal = calculate_deal_metrics(custom_price, model_year, active_addons, down_payment_pct)
s = deal["summary"]

# Content Grid Layout Display
col1, col2 = st.columns([1, 1.6])

with col1:
    st.subheader("📊 Transaction Value Stack")
    
    summary_data = {
        "Financial Component": [
            "Base Vehicle Price", 
            "Vehicle VAT (5%)",
            "Vehicle Price (Inc. VAT)",
            "Total Add-Ons / Contracts", 
            "Full Capitalized Asset Value", 
            f"Down Payment Required ({down_payment_pct}%)", 
            "Total Capital Financed Pool"
        ],
        "Value (AED)": [
            f"{s['base_price']:,.2f}",
            f"{s['vat_amount']:,.2f}",
            f"{s['vehicle_with_vat']:,.2f}",
            f"{s['total_addons']:,.2f}",
            f"**{s['full_value']:,.2f}**",
            f"**{s['down_payment']:,.2f}**",
            f"**{s['finance_amount']:,.2f}**"
        ]
    }
    st.table(pd.DataFrame(summary_data))

with col2:
    st.subheader("📉 Amortization Matrix Grid")
    st.markdown("Tenure mapping factoring updated baseline interest and promotional subventions:")
    st.dataframe(deal["matrix"], use_container_width=True, hide_index=True)
    
    st.success(f"💡 **Subvention Notice:** 1-Year terms evaluate with zero promotional multiplier tiers, while 2-5 year timelines calculate promotional subvention deductions automatically.")
