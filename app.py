import io
import datetime
import requests
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.metrics import mean_absolute_error, r2_score
import streamlit as st

# Export dependencies
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -----------------------------------------------------------------------------
# PAGE CONFIGURATION
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Hispanic Heritage Historical Data Explorer",
    page_icon="🌎",
    layout="wide"
)

# -----------------------------------------------------------------------------
# METADATA & CONSTANTS
# -----------------------------------------------------------------------------
COUNTRIES = {
    "Mexico": "MX",
    "Brazil": "BR",
    "Argentina": "AR",
    "Colombia": "CO",
    "Chile": "CL",
    "Peru": "PE",
    "Venezuela": "VE",
    "Ecuador": "EC",
    "Uruguay": "UY"
}

INDICATORS = {
    "Life expectancy at birth": {
        "code": "SP.DYN.LE00.IN",
        "units": "Years; total population",
        "min": 0, "max": 120
    },
    "Unemployment rate": {
        "code": "SL.UEM.TOTL.NE.ZS",
        "units": "% of total labor force; national estimate",
        "min": 0, "max": 100
    },
    "GDP per capita": {
        "code": "NY.GDP.PCAP.CD",
        "units": "Current US$; economic output per person",
        "min": 0, "max": 1000000
    },
    "Crude birth rate": {
        "code": "SP.DYN.CBRT.IN",
        "units": "Births per 1,000 people",
        "min": 0, "max": 1000
    },
    "Net migration": {
        "code": "SM.POP.NETM",
        "units": "People; immigration minus emigration",
        "min": -100000000, "max": 100000000
    },
    "Intentional homicide rate": {
        "code": "VC.IHR.PSRC.P5",
        "units": "Intentional homicides per 100,000 people",
        "min": 0, "max": 100000
    },
    "Consumer price inflation": {
        "code": "FP.CPI.TOTL.ZG",
        "units": "Annual % change",
        "min": -100, "max": 100000
    },
    "Central government debt": {
        "code": "GC.DOD.TOTL.GD.ZS",
        "units": "% of GDP",
        "min": 0, "max": 1000
    },
    "Total reserves (includes gold)": {
        "code": "FI.RES.TOTL.CD",
        "units": "Current US$",
        "min": 0, "max": 10000000000000
    },
    "External balance on goods/services": {
        "code": "NE.RSB.GNFS.ZS",
        "units": "% of GDP",
        "min": -100, "max": 100
    }
}

# -----------------------------------------------------------------------------
# DATA RETRIEVAL & CACHING
# -----------------------------------------------------------------------------
@st.cache_data(ttl=3600)
def fetch_world_bank_data(country_code, indicator_code):
    """
    Fetches historical data for target 70-year window: (Current Year - 70) to (Current Year - 1)
    """
    current_year = datetime.datetime.now().year
    start_yr = current_year - 70
    end_yr = current_year - 1
    
    url = f"http://api.worldbank.org/v2/country/{country_code}/indicator/{indicator_code}?date={start_yr}:{end_yr}&format=json&per_page=100"
    
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        json_data = response.json()
        
        if len(json_data) < 2 or json_data[1] is None:
            return pd.DataFrame(columns=["Year", "Value"])
            
        records = json_data[1]
        data = []
        for r in records:
            val = r.get("value")
            yr = r.get("date")
            if val is not None and yr is not None:
                data.append({"Year": int(yr), "Value": float(val)})
                
        df = pd.DataFrame(data)
        if not df.empty:
            df = df.sort_values("Year").reset_index(drop=True)
        return df
    except Exception:
        return pd.DataFrame(columns=["Year", "Value"])

# -----------------------------------------------------------------------------
# HELPER MATHEMATICS FUNCTIONS
# -----------------------------------------------------------------------------
def fit_centered_polynomial(df, degree):
    """Fits degree-d polynomial using centered time t = year - base_year."""
    if len(df) < degree + 1:
        return None
    
    base_year = int(df["Year"].min())
    t = df["Year"].values - base_year
    y = df["Value"].values
    
    coeffs = np.polyfit(t, y, degree)
    poly_func = np.poly1d(coeffs)
    
    y_pred = poly_func(t)
    mae = mean_absolute_error(y, y_pred)
    r2 = r2_score(y, y_pred) if len(y) > 1 else 0.0
    residuals = y - y_pred
    
    return {
        "base_year": base_year,
        "coeffs": coeffs,
        "poly": poly_func,
        "mae": mae,
        "r2": r2,
        "residuals": residuals,
        "min_year": int(df["Year"].min()),
        "max_year": int(df["Year"].max()),
        "count": len(df),
        "mean": float(y.mean())
    }

def poly_to_string(poly, base_year):
    """Formats polynomial equation into readable string notation."""
    deg = poly.order
    terms = []
    for i, c in enumerate(poly.coefficients):
        power = deg - i
        if abs(c) < 1e-9:
            continue
        c_str = f"{c:+.4e}" if (abs(c) < 1e-3 or abs(c) > 1e4) else f"{c:+.4f}"
        if power > 1:
            terms.append(f"{c_str}·(year - {base_year})^{power}")
        elif power == 1:
            terms.append(f"{c_str}·(year - {base_year})")
        else:
            terms.append(f"{c_str}")
    return "y = " + " ".join(terms) if terms else "y = 0"

# -----------------------------------------------------------------------------
# EXPORT GENERATORS
# -----------------------------------------------------------------------------
def generate_excel(df_working, model_info, country_name, ind_name, batch_df):
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Sheet 1: Summary
    ws_summary = workbook.add_worksheet("Summary")
    ws_summary.write(0, 0, "Hispanic Heritage Historical Data Explorer Summary")
    ws_summary.write(2, 0, "Country")
    ws_summary.write(2, 1, country_name)
    ws_summary.write(3, 0, "Indicator")
    ws_summary.write(3, 1, ind_name)
    ws_summary.write(4, 0, "Units")
    ws_summary.write(4, 1, INDICATORS[ind_name]["units"])
    ws_summary.write(5, 0, "Source")
    ws_summary.write(5, 1, "World Bank API")
    ws_summary.write(6, 0, "Retrieval Date")
    ws_summary.write(6, 1, datetime.date.today().strftime("%Y-%m-%d"))
    
    if model_info:
        ws_summary.write(8, 0, "Polynomial Degree")
        ws_summary.write(8, 1, len(model_info["coeffs"]) - 1)
        ws_summary.write(9, 0, "Base Year")
        ws_summary.write(9, 1, model_info["base_year"])
        ws_summary.write(10, 0, "Equation")
        ws_summary.write(10, 1, poly_to_string(model_info["poly"], model_info["base_year"]))
        ws_summary.write(11, 0, "In-Sample R-squared")
        ws_summary.write(11, 1, model_info["r2"])
        ws_summary.write(12, 0, "In-Sample MAE")
        ws_summary.write(12, 1, model_info["mae"])

    # Sheet 2: Data
    ws_data = workbook.add_worksheet("Data")
    ws_data.write(0, 0, "Year")
    ws_data.write(0, 1, "Value")
    for row_idx, row in df_working.iterrows():
        ws_data.write(row_idx + 1, 0, int(row["Year"]))
        ws_data.write(row_idx + 1, 1, float(row["Value"]))
        
    # Sheet 3: Batch Results
    if batch_df is not None and not batch_df.empty:
        ws_batch = workbook.add_worksheet("Calculations")
        for col_idx, col in enumerate(batch_df.columns):
            ws_batch.write(0, col_idx, col)
        for row_idx, row in batch_df.iterrows():
            for col_idx, col in enumerate(batch_df.columns):
                ws_batch.write(row_idx + 1, col_idx, str(row[col]))
                
    workbook.close()
    return output.getvalue()

def generate_pdf(df_working, model_info, country_name, ind_name, fig, batch_df):
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=16, leading=20)
    story.append(Paragraph(f"Hispanic Heritage Research Report: {country_name}", title_style))
    story.append(Spacer(1, 10))

    meta_text = f"<b>Indicator:</b> {ind_name}<br/><b>Units:</b> {INDICATORS[ind_name]['units']}<br/><b>Generated:</b> {datetime.date.today()}"
    story.append(Paragraph(meta_text, styles['Normal']))
    story.append(Spacer(1, 10))

    if model_info:
        eq_str = poly_to_string(model_info["poly"], model_info["base_year"])
        model_text = f"<b>Polynomial Model:</b> {eq_str}<br/><b>R² Score:</b> {model_info['r2']:.4f} | <b>MAE:</b> {model_info['mae']:.4f}"
        story.append(Paragraph(model_text, styles['Normal']))
        story.append(Spacer(1, 10))

    img_buf = io.BytesIO()
    fig.savefig(img_buf, format='png', dpi=150, bbox_inches='tight')
    img_buf.seek(0)
    story.append(Image(img_buf, width=500, height=250))
    story.append(Spacer(1, 10))

    # Findings section
    story.append(Paragraph("<b>Key Findings & Limitations</b>", styles['Heading2']))
    findings_text = (
        "• Model fit reflects historical in-sample alignment and does not represent uncertainty bounds.<br/>"
        "• Polynomial forecasts beyond observed dates tend to diverge rapidly; interpret extrapolations with caution.<br/>"
        "• Contextual factors (economic reforms, global crises) should complement quantitative results."
    )
    story.append(Paragraph(findings_text, styles['Normal']))
    story.append(Spacer(1, 10))

    # Batch estimates table preview
    if batch_df is not None and not batch_df.empty:
        story.append(Paragraph("<b>Calculated Estimates (Preview)</b>", styles['Heading3']))
        preview_df = batch_df.head(10)
        table_data = [preview_df.columns.tolist()] + preview_df.astype(str).values.tolist()
        t = Table(table_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 8),
        ]))
        story.append(t)

    doc.build(story)
    return output.getvalue()

# -----------------------------------------------------------------------------
# APPLICATION INTERFACE
# -----------------------------------------------------------------------------
def main():
    st.title("🌎 Hispanic Heritage Historical Data Explorer")
    st.markdown("*A mathematical exploration tool for historical country indicators across Latin America.*")

    # App Tabs
    tab_explorer, tab_calculus = st.tabs(["📊 Main Data Explorer", "📈 Advanced Calculus & Simulation"])

    # -------------------------------------------------------------------------
    # TAB 1: MAIN DATA EXPLORER
    # -------------------------------------------------------------------------
    with tab_explorer:
        st.sidebar.header("Data & Control Settings")
        
        # Country & Indicator Controls
        selected_country = st.sidebar.selectbox("Primary Country", list(COUNTRIES.keys()), index=0)
        selected_indicator = st.sidebar.selectbox("Indicator", list(INDICATORS.keys()), index=0)
        
        # Context note for Brazil
        if selected_country == "Brazil":
            st.sidebar.info("Note: Brazil is included as a Latin American regional comparison; it is not a Spanish-speaking Hispanic country.")
            
        comparison_countries = st.sidebar.multiselect(
            "Comparison Countries",
            [c for c in COUNTRIES.keys() if c != selected_country]
        )
        
        # Modeling parameters
        st.sidebar.subheader("Model Parameters")
        degree = st.sidebar.slider("Polynomial Degree", min_value=3, max_value=8, value=3)
        sample_inc = st.sidebar.slider("Graph Sample Increment (years)", min_value=1, max_value=10, value=1)
        horizon = st.sidebar.slider("Forward Display Horizon (years)", min_value=0, max_value=50, value=5)
        printer_mode = st.sidebar.checkbox("Printer-Friendly View")

        # Load Data
        country_code = COUNTRIES[selected_country]
        ind_code = INDICATORS[selected_indicator]["code"]
        raw_df = fetch_world_bank_data(country_code, ind_code)

        st.subheader(f"Metadata & Provenance: {selected_country} - {selected_indicator}")
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("World Bank Code", ind_code)
        col_m2.metric("Units", INDICATORS[selected_indicator]["units"])
        col_m3.metric("Retrieval Date", datetime.date.today().strftime("%Y-%m-%d"))

        if raw_df.empty:
            st.warning("No official data found for this specific indicator/country combination in the 70-year window.")
            return

        # Editable Data Table Setup
        st.write("---")
        st.write("### Data Table & Scenario Editing")
        
        if "working_df" not in st.session_state or st.sidebar.button("Reset to Source Data"):
            st.session_state.working_df = raw_df.copy()
            st.session_state.is_edited = False

        edited_df = st.data_editor(
            st.session_state.working_df,
            num_rows="dynamic",
            key="data_editor",
            use_container_width=True
        )

        # Validation Logic
        cleaned_df = edited_df.dropna().copy()
        cleaned_df["Year"] = pd.to_numeric(cleaned_df["Year"], errors="coerce")
        cleaned_df["Value"] = pd.to_numeric(cleaned_df["Value"], errors="coerce")
        cleaned_df = cleaned_df.dropna().drop_duplicates(subset=["Year"]).sort_values("Year")
        
        # Check domain restrictions
        val_min = INDICATORS[selected_indicator]["min"]
        val_max = INDICATORS[selected_indicator]["max"]
        cleaned_df = cleaned_df[(cleaned_df["Value"] >= val_min) & (cleaned_df["Value"] <= val_max)]

        if len(cleaned_df) != len(raw_df):
            st.session_state.is_edited = True

        if st.session_state.is_edited:
            st.info("⚠️ Using user-modified scenario data.")

        # Fit Polynomial Model
        model_info = fit_centered_polynomial(cleaned_df, degree)

        if not model_info:
            st.error(f"Insufficient distinct valid data points to fit polynomial of degree {degree}. Need at least {degree+1} points.")
            return

        # Model Summary & Metrics
        st.write("---")
        st.write("### Mathematical Model & Metrics")
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("Base Year (t=0)", model_info["base_year"])
        m_col2.metric("In-Sample R²", f"{model_info['r2']:.4f}")
        m_col3.metric("In-Sample MAE", f"{model_info['mae']:.4f}")
        m_col4.metric("Fitted Observations", model_info["count"])

        eq_str = poly_to_string(model_info["poly"], model_info["base_year"])
        st.code(f"Fitted Equation: {eq_str}", language="text")

        # Visual Plotting
        st.write("---")
        st.write("### Visualization & Comparisons")

        fig, ax = plt.subplots(figsize=(10, 5) if not printer_mode else (8, 4))
        if printer_mode:
            plt.style.use('grayscale')

        # Plot Historical & Fitted
        ax.scatter(cleaned_df["Year"], cleaned_df["Value"], color="black", label=f"Observed ({selected_country})", zorder=3)
        
        min_yr = model_info["min_year"]
        max_yr = model_info["max_year"] + horizon
        dense_years = np.linspace(min_yr, max_yr, 300)
        dense_t = dense_years - model_info["base_year"]
        dense_preds = model_info["poly"](dense_t)

        # Separate historical from extrapolation domain
        hist_mask = dense_years <= model_info["max_year"]
        ax.plot(dense_years[hist_mask], dense_preds[hist_mask], label="Fitted Model (Historical)", color="blue" if not printer_mode else "black", linewidth=2)
        ax.plot(dense_years[~hist_mask], dense_preds[~hist_mask], '--', label="Extrapolated Model", color="red" if not printer_mode else "gray", linewidth=2)
        ax.axvline(model_info["max_year"], color="gray", linestyle=":", label="Historical Domain Boundary")

        # Sampling Markers
        sample_years = np.arange(min_yr, max_yr + 1, sample_inc)
        sample_t = sample_years - model_info["base_year"]
        ax.plot(sample_years, model_info["poly"](sample_t), 's', alpha=0.5, label=f"Sample Markers ({sample_inc}yr step)")

        # Plot Comparison Countries
        for comp in comparison_countries:
            comp_df = fetch_world_bank_data(COUNTRIES[comp], ind_code)
            if not comp_df.empty:
                ax.plot(comp_df["Year"], comp_df["Value"], label=f"Comparison: {comp}", alpha=0.7)

        ax.set_xlabel("Year")
        ax.set_ylabel(INDICATORS[selected_indicator]["units"])
        ax.set_title(f"{selected_indicator} - {selected_country}")
        ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.grid(True, linestyle='--', alpha=0.6)

        st.pyplot(fig)

        # Single Year & Batch Estimation
        st.write("---")
        st.write("### Single & Batch Estimates")
        
        calc_tab1, calc_tab2 = st.tabs(["Single Year Estimation", "Batch Year Estimation"])
        
        with calc_tab1:
            target_year = st.number_input("Target Year", min_value=1950, max_value=2100, value=2030, step=1)
            if st.button("Calculate Value"):
                t_val = target_year - model_info["base_year"]
                est_val = model_info["poly"](t_val)
                
                # Classify year
                if target_year < model_info["min_year"] or target_year > model_info["max_year"]:
                    cls = "Extrapolation (Outside Fitted Range)"
                else:
                    cls = "Interpolation / Within Range"

                st.metric("Model Estimate", f"{est_val:.4f} {INDICATORS[selected_indicator]['units']}")
                st.caption(f"Classification: **{cls}** | Fitted Range: {model_info['min_year']} - {model_info['max_year']}")

        batch_results_df = None
        with calc_tab2:
            batch_mode = st.radio("Batch Mode", ["Year Range", "Custom List"])
            valid_years = []
            
            if batch_mode == "Year Range":
                b_start = st.number_input("Start Year", 1950, 2100, 2020)
                b_end = st.number_input("End Year", 1950, 2100, 2030)
                b_step = st.number_input("Step (Years)", 1, 10, 1)
                if b_start < b_end:
                    valid_years = list(range(int(b_start), int(b_end) + 1, int(b_step)))
            else:
                custom_str = st.text_input("Enter comma-separated years", "2022, 2025, 2030")
                try:
                    valid_years = sorted(list(set([int(y.strip()) for y in custom_str.split(",") if y.strip().isdigit()])))
                except Exception:
                    st.error("Invalid custom year list format.")

            if valid_years:
                st.write(f"Years to evaluate: `{valid_years}`")
                if st.button("Calculate Batch"):
                    batch_data = []
                    for yr in valid_years:
                        t_v = yr - model_info["base_year"]
                        e_v = model_info["poly"](t_v)
                        c_v = "Extrapolation" if (yr < model_info["min_year"] or yr > model_info["max_year"]) else "Interpolation"
                        batch_data.append({
                            "Year": yr,
                            "Estimate": e_v,
                            "Units": INDICATORS[selected_indicator]["units"],
                            "Classification": c_v
                        })
                    batch_results_df = pd.DataFrame(batch_data)
                    st.dataframe(batch_results_df, use_container_width=True)

        # Downloads Section
        st.write("---")
        st.write("### Export Reports & Data")
        d_col1, d_col2, d_col3 = st.columns(3)

        # CSV Download
        csv_bytes = cleaned_df.to_csv(index=False).encode('utf-8')
        d_col1.download_button("📥 Download CSV Data", csv_bytes, "indicator_data.csv", "text/csv")

        # Excel Download
        xlsx_bytes = generate_excel(cleaned_df, model_info, selected_country, selected_indicator, batch_results_df)
        d_col2.download_button("📥 Download Excel Workbook", xlsx_bytes, "indicator_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # PDF Download
        pdf_bytes = generate_pdf(cleaned_df, model_info, selected_country, selected_indicator, fig, batch_results_df)
        d_col3.download_button("📥 Download PDF Report", pdf_bytes, "indicator_report.pdf", "application/pdf")

    # -------------------------------------------------------------------------
    # TAB 2: ADVANCED CALCULUS & CONTEXTUAL SIMULATION
    # -------------------------------------------------------------------------
    with tab_calculus:
        st.header("📈 Interactive Calculus & Context Simulation Mode")
        
        # Pathway selection
        pathway = st.selectbox("Course Level / Analysis Mode", ["Honors Algebra 2", "Precalculus", "Calculus Mode"])
        
        context_type = st.selectbox(
            "Select Context Domain",
            ["Economic Forecasting", "Business Growth", "Medical/Drug Concentration", "Epidemiology/Population", "Custom"]
        )

        # Functions available
        f_option = st.selectbox("Function Source", ["Use Fitted Polynomial from Main Explorer", "Custom Cubic Polynomial"])

        if f_option == "Use Fitted Polynomial from Main Explorer" and model_info is not None:
            poly = model_info["poly"]
            base_yr = model_info["base_year"]
        else:
            # Default cubic $f(x) = x^3 - 3x^2 - 9x + 5$
            poly = np.poly1d([1, -3, -9, 5])
            base_yr = 0

        # Compute Analytical Derivatives
        d1 = np.polyder(poly, 1)
        d2 = np.polyder(poly, 2)

        st.write("---")
        st.write("### Function & Derivative Representations")
        st.latex(f"f(t) = {poly_to_string(poly, base_yr).replace('y = ', '')}")
        st.latex(f"f'(t) = {poly_to_string(d1, base_yr).replace('y = ', '')}")
        st.latex(f"f''(t) = {poly_to_string(d2, base_yr).replace('y = ', '')}")

        # Interactive Controls
        st.write("---")
        st.write("### Rate of Change Analysis Controls")
        
        eval_year = st.slider("Select Point of Tangency (Year)", min_value=1960, max_value=2050, value=2010)
        int_start, int_end = st.slider("Select Secant Interval [a, b]", min_value=1960, max_value=2050, value=(1980, 2010))

        t_pt = eval_year - base_yr
        f_val = poly(t_pt)
        d1_val = d1(t_pt)
        d2_val = d2(t_pt)

        # Calculate Secant Average Rate of Change
        t_a = int_start - base_yr
        t_b = int_end - base_yr
        f_a = poly(t_a)
        f_b = poly(t_b)
        avg_rate = (f_b - f_a) / (t_b - t_a) if t_b != t_a else 0.0

        # Calculus Interpretation
        st.write("---")
        st.write("### Contextual Meaning & Analysis")
        c1, c2, c3 = st.columns(3)
        c1.metric("Function Value f(t)", f"{f_val:.4f}")
        c2.metric("Instantaneous Rate f'(t)", f"{d1_val:.4f}")
        c3.metric("Concavity/Acceleration f''(t)", f"{d2_val:.4f}")

        st.info(f"**Average Rate of Change** over [{int_start}, {int_end}]: **{avg_rate:.4f}** units/year.")

        if pathway == "Calculus Mode":
            st.write("#### Derivative & Critical Point Analysis")
            roots_d1 = d1.roots
            real_roots = [r.real + base_yr for r in roots_d1 if np.isreal(r)]
            st.write(f"**Stationary Points (f'(t) = 0):** Evaluated at years {np.round(real_roots, 2)}")
            
            # Concavity check
            if d2_val > 0:
                concavity_str = "Concave Upward (f''(t) > 0) -> Rate of change is increasing."
            elif d2_val < 0:
                concavity_str = "Concave Downward (f''(t) < 0) -> Rate of change is decreasing."
            else:
                concavity_str = "Possible Inflection Point (f''(t) = 0)."
            st.write(f"**Concavity Status at {eval_year}:** {concavity_str}")

        # Visualization
        fig_c, ax_c = plt.subplots(figsize=(10, 5))
        t_domain = np.linspace(min(1960, int_start) - base_yr, max(2050, int_end) - base_yr, 200)
        years_domain = t_domain + base_yr

        # Plot f(t)
        ax_c.plot(years_domain, poly(t_domain), label="f(t) Function", color="blue")
        
        # Tangent line: y = f'(t_pt)*(t - t_pt) + f_val
        tangent_y = d1_val * (t_domain - t_pt) + f_val
        ax_c.plot(years_domain, tangent_y, 'r--', label=f"Tangent Line at {eval_year} (Slope = {d1_val:.2f})")

        # Secant line
        secant_slope = avg_rate
        secant_y = secant_slope * (t_domain - t_a) + f_a
        ax_c.plot(years_domain, secant_y, 'g:', label=f"Secant Line [{int_start}, {int_end}] (Slope = {avg_rate:.2f})")

        ax_c.scatter([eval_year], [f_val], color="red", zorder=5)
        ax_c.scatter([int_start, int_end], [f_a, f_b], color="green", zorder=5)

        ax_c.set_xlabel("Year")
        ax_c.set_ylabel("Value")
        ax_c.set_title("Function, Tangent, and Secant Analysis")
        ax_c.grid(True, linestyle="--", alpha=0.5)
        ax_c.legend()

        st.pyplot(fig_c)

if __name__ == "__main__":
    main()