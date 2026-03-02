import os
import joblib
import numpy as np
import pandas as pd
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Egyptian Property Price Predictor",
    layout="wide",
)

# ── Load artifacts ─────────────────────────────────────────────────────────────
ARTIFACTS_PATH = os.path.join(os.path.dirname(__file__), "model_artifacts.pkl")

@st.cache_resource
def load_artifacts():
    return joblib.load(ARTIFACTS_PATH)

try:
    art = load_artifacts()
except FileNotFoundError:
    st.error(
        "Model artifacts not found. Please run **Step 13** in `HousePrices.ipynb` first "
        "to save the trained model, then restart this app."
    )
    st.stop()

model          = art["model"]
city_te        = art["city_te_map"]
compound_te    = art["compound_te_map"]
citytype_te    = art.get("citytype_te_map", {})
le             = art["le_dict"]
feat_cols      = art["feature_cols"]
gm             = art["global_mean_log_price"]
rmse           = art["model_rmse"]
r2             = art["model_r2"]
model_name     = art.get("model_name", "LightGBM")

cities         = art["cities"]
compounds      = art["compounds"]
types_list     = art["types"]
levels_list    = art["levels"]
del_terms      = art["delivery_terms"]
pay_opts       = art["payment_options"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def encode_safe(encoder, value):
    try:
        return int(encoder.transform([str(value)])[0])
    except (ValueError, KeyError):
        return int(encoder.transform(["Missing"])[0])

def build_input(prop_type, bedrooms, bathrooms, area,
                furnished, level, compound, payment,
                del_term, del_tier, city):
    row = {}
    for col in feat_cols:
        if   col == "Bedrooms":          row[col] = float(bedrooms)
        elif col == "Bathrooms":         row[col] = float(bathrooms)
        elif col == "Area":              row[col] = float(area)
        elif col == "log_Area":          row[col] = np.log1p(float(area))
        elif col == "Type_enc":          row[col] = encode_safe(le["Type"], prop_type)
        elif col == "Furnished_enc":     row[col] = encode_safe(le["Furnished"], furnished)
        elif col == "Level_enc":         row[col] = encode_safe(le["Level"], level)
        elif col == "Payment_Option_enc":row[col] = encode_safe(le["Payment_Option"], payment)
        elif col == "Delivery_Term_enc": row[col] = encode_safe(le["Delivery_Term"], del_term)
        elif col == "Delivery_Tier_enc": row[col] = encode_safe(le["Delivery_Tier"], del_tier)
        elif col == "City_te":           row[col] = city_te.get(city, gm)
        elif col == "City_enc":          row[col] = encode_safe(le["City"], city)
        elif col == "Compound_te":       row[col] = compound_te.get(compound, gm)
        elif col == "Compound_enc":      row[col] = encode_safe(le["Compound"], compound)
        elif col == "Beds_per_100m2":   row[col] = float(bedrooms) / (float(area) / 100)
        elif col == "Bath_bed_ratio":   row[col] = float(bathrooms) / max(float(bedrooms), 1)
        elif col == "CityType_te":      row[col] = citytype_te.get(f"{city} | {prop_type}", gm)
        else:                            row[col] = 0.0
    return pd.DataFrame([row])[feat_cols]

# ── Header ────────────────────────────────────────────────────────────────────
st.title("Egyptian Property Price Predictor")
st.markdown(
    "Enter the property details below to get an instant price estimate powered "
    "by a gradient-boosted machine learning model trained on 26,000+ listings."
)
st.markdown("---")

# ── Input form ────────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("Property")
    prop_type = st.selectbox("Type", types_list)
    bedrooms  = st.number_input("Bedrooms",  min_value=1, max_value=15, value=3, step=1)
    bathrooms = st.number_input("Bathrooms", min_value=1, max_value=15, value=2, step=1)
    area      = st.number_input("Area (m²)", min_value=30, max_value=1000, value=150, step=5)
    furnished = st.selectbox("Furnished", ["No", "Yes"])

with col2:
    st.subheader("Location")
    city     = st.selectbox("City", cities)
    compound = st.selectbox(
        "Compound",
        ["Not in a compound"] + sorted(compounds),
        help="Select the gated compound, or 'Not in a compound' if standalone."
    )
    level    = st.selectbox(
        "Floor / Level",
        ["Ground"] + [str(i) for i in range(1, 21)] + ["Missing"],
        help="Floor number or 'Missing' if unknown / not applicable."
    )

with col3:
    st.subheader("Purchase Details")
    payment  = st.selectbox("Payment Option", pay_opts)
    del_term = st.selectbox("Delivery Condition", del_terms)
    del_tier = st.selectbox(
        "Delivery Timeline",
        ["Ready", "Near_Term", "Mid_Term", "Long_Term", "Unknown"],
        format_func=lambda x: {
            "Ready":     "Ready to move",
            "Near_Term": "Near term (within ~1 year)",
            "Mid_Term":  "Mid term (1–2 years)",
            "Long_Term": "Long term (2+ years)",
            "Unknown":   "Unknown",
        }[x],
    )

st.markdown("---")

# ── Predict ───────────────────────────────────────────────────────────────────
if st.button("Predict Price", type="primary", use_container_width=True):
    compound_key = "Missing" if compound == "Not in a compound" else compound
    level_key    = level if level else "Missing"

    X = build_input(
        prop_type, bedrooms, bathrooms, area,
        furnished, level_key, compound_key,
        payment, del_term, del_tier, city
    )

    log_price     = model.predict(X)[0]
    price         = np.exp(log_price)
    price_per_m2  = price / area
    lower         = np.exp(log_price - rmse)
    upper         = np.exp(log_price + rmse)

    st.markdown("### Predicted Price")

    m1, m2, m3 = st.columns(3)
    m1.metric(
        "Estimated Price",
        f"EGP {price / 1e6:.2f}M",
        help="Central estimate from the model."
    )
    m2.metric(
        "Price per m²",
        f"EGP {price_per_m2:,.0f}",
    )
    m3.metric(
        "68% Confidence Range",
        f"EGP {lower / 1e6:.2f}M — {upper / 1e6:.2f}M",
        help="Based on ±1 RMSE in log space. True price falls in this range ~68% of the time."
    )

    st.markdown("---")

    with st.expander("Model information"):
        col_a, col_b, col_c, col_d = st.columns(4)
        col_a.metric("Model",          model_name)
        col_b.metric("R²",             f"{r2:.3f}")
        col_c.metric("CV RMSE",        f"{rmse:.3f} (log scale)")
        col_d.metric("Training rows",  "26,000+")
        st.markdown(
            f"An RMSE of **{rmse:.3f}** in log scale means predictions are typically within "
            f"**±{(np.exp(rmse) - 1) * 100:.0f}%** of the true price."
        )
