import streamlit as st
import math

st.set_page_config(page_title="Finance Calculator", layout="wide")

st.title("Vehicle Finance Calculator")

# --- INPUT SECTION ---
st.header("Vehicle Pricing Details")

col1, col2 = st.columns(2)

with col1:
    base_price = st.number_input("Base Price", min_value=0.0, value=0.0)
    vat_percentage = st.number_input("VAT (%)", min_value=0.0, value=5.0)
    accessories = st.number_input("Accessories", min_value=0.0, value=0.0)

with col2:
    insurance_percentage = st.number_input("Insurance (%)", min_value=0.0, value=3.0)
    vri_percentage = st.number_input("VRI (%)", min_value=0.0, value=1.0)
    down_payment_percentage = st.number_input("Down Payment (%)", min_value=0.0, value=20.0)
    interest_rate = st.number_input("Interest Rate (%)", min_value=0.0, value=3.0)
    tenure_months = st.number_input("Tenure (Months)", min_value=1, value=60)

# --- CALCULATIONS ---
st.header("Calculation Summary")

# VAT calculation
vat_amount = base_price * (vat_percentage / 100)

# Base for Insurance & VRI (EXCLUDES Insurance & VRI)
insurance_vri_base = base_price + vat_amount + accessories

# Insurance calculation
insurance_amount = insurance_vri_base * (insurance_percentage / 100)

# VRI calculation
vri_amount = insurance_vri_base * (vri_percentage / 100)

# Total cost including all components
total_cost = base_price + vat_amount + accessories + insurance_amount + vri_amount

# Down payment based on total cost
down_payment_amount = total_cost * (down_payment_percentage / 100)

# Loan amount
loan_amount = total_cost - down_payment_amount

# Monthly interest rate
monthly_interest_rate = interest_rate / 100 / 12

# EMI calculation
if monthly_interest_rate > 0:
    emi = loan_amount * monthly_interest_rate * math.pow(1 + monthly_interest_rate, tenure_months) / \
          (math.pow(1 + monthly_interest_rate, tenure_months) - 1)
else:
    emi = loan_amount / tenure_months

# --- OUTPUT SECTION ---
col3, col4 = st.columns(2)

with col3:
    st.subheader("Price Breakdown")
    st.write(f"Base Price: AED {base_price:,.2f}")
    st.write(f"VAT ({vat_percentage}%): AED {vat_amount:,.2f}")
    st.write(f"Accessories: AED {accessories:,.2f}")
    st.write(f"Insurance Base: AED {insurance_vri_base:,.2f}")
    st.write(f"Insurance ({insurance_percentage}%): AED {insurance_amount:,.2f}")
    st.write(f"VRI ({vri_percentage}%): AED {vri_amount:,.2f}")
    st.write(f"**Total Cost: AED {total_cost:,.2f}**")

with col4:
    st.subheader("Finance Summary")
    st.write(f"Down Payment ({down_payment_percentage}%): AED {down_payment_amount:,.2f}")
    st.write(f"Loan Amount: AED {loan_amount:,.2f}")
    st.write(f"Interest Rate: {interest_rate}%")
    st.write(f"Tenure: {tenure_months} months")
    st.write(f"**Monthly EMI: AED {emi:,.2f}**")

st.success("Calculation completed successfully!")
