import io
import math
from datetime import datetime, timezone

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="Hispanic Heritage Historical Data Explorer", layout="wide")

COUNTRIES = {
    "Mexico": "MEX", "Brazil": "BRA", "Argentina": "ARG", "Colombia": "COL",
    "Chile": "CHL", "Peru": "PER", "Venezuela": "VEN", "Ecuador": "ECU",
    "Uruguay": "URY",
}

INDICATORS = {
    "Life expectancy at birth": {
        "code": "SP.DYN.LE00.IN", "units": "years",
        "definition": "Life expectancy at birth, total population.",
    },
    "Unemployment rate (national estimate)": {
        "code": "SL.UEM.TOTL.NE.ZS", "units": "% of total labor force",
        "definition": "Unemployment as a percentage of the total labor force, national estimate.",
    },
    "GDP per capita": {
        "code": "NY.GDP.PCAP.CD", "units": "current US$ per person",
        "definition": "Economic output per person; this is not average wealth or household income.",
    },
    "Crude birth rate": {
        "code": "SP.DYN.CBRT.IN", "units": "births per 1,000 people",
        "definition": "Annual number of births per 1,000 people, not a total birth count.",
    },
    "Net migration": {
        "code": "SM.POP.NETM", "units": "people",
        "definition": "Net migration: immigration minus emigration for the source period.",
    },
    "Intentional homicide rate": {
        "code": "VC.IHR.PSRC.P5", "units": "intentional homicides per 100,000 people",
        "definition": "Intentional homicide rate per 100,000 people.",
    },
    "Consumer price inflation": {
        "code": "FP.CPI.TOTL.ZG", "units": "% annual change",
        "definition": "Annual percentage change in consumer prices.",
    },
    "Central government debt": {
        "code": "GC.DOD.TOTL.GD.ZS", "units": "% of GDP",
        "definition": "Central government debt as a percentage of GDP; not all public-sector debt.",
    },
    "Total reserves, including gold": {
        "code": "FI.RES.TOTL.CD", "units": "current US$",
        "definition": "Total reserves including gold, in current US dollars.",
    },
    "External balance: goods and services": {
        "code": "NE.RSB.GNFS.ZS", "units": "% of GDP",
        "definition": "External balance on goods and services as a percentage of GDP; not current account.",
    },
}

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_world_bank(country_code, indicator_code, start_year, end_year):
    url = (
        f"https://api.worldbank.org/v2/country/{country_code}/indicator/"
        f"{indicator_code}?date={start_year}:{end_year}&format=json&per_page=200"
    )
    response = requests.get(url, timeout=20)
    response.raise_for_status()
    payload = response.json()
    if not isinstance(payload, list) or len(payload) < 2 or not isinstance(payload[1], list):
        raise ValueError("The World Bank API returned an unexpected response.")
    rows = []
    for item in payload[1]:
        year = item.get("date")
        value = item.get("value")
        if year is not None and value is not None:
            try:
                rows.append({"Year": int(year), "Value": float(value)})
            except (TypeError, ValueError):
                continue
    frame = pd.DataFrame(rows, columns=["Year", "Value"])
    if not frame.empty:
        frame = frame.drop_duplicates("Year").sort_values("Year").reset_index(drop=True)
    return frame

def fit_model(frame, degree):
    clean = frame.dropna(subset=["Year", "Value"]).copy()
    clean["Year"] = clean["Year"].astype(int)
    clean["Value"] = pd.to_numeric(clean["Value"], errors="coerce")
    clean = clean.dropna(subset=["Value"]).sort_values("Year")
    if clean["Year"].duplicated().any():
        raise ValueError("Each year must appear only once.")
    if len(clean) < degree + 1:
        raise ValueError(f"Degree {degree} requires at least {degree + 1} distinct observations; only {len(clean)} are available.")
    base_year = int(clean["Year"].min())
    t = clean["Year"].to_numpy(dtype=float) - base_year
    y = clean["Value"].to_numpy(dtype=float)
    coeff = np.polyfit(t, y, degree)
    predicted = np.polyval(coeff, t)
    residuals = y - predicted
    ss_res = float(np.sum(residuals ** 2))
    ss_tot = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1 - ss_res / ss_tot if ss_tot > 0 else float("nan")
    mae = float(np.mean(np.abs(residuals)))
    return clean, base_year, coeff, predicted, residuals, r2, mae

def equation_text(coeff, base_year):
    terms = []
    degree = len(coeff) - 1
    for i, c in enumerate(coeff):
        power = degree - i
        if abs(c) < 1e-12:
            continue
        if power == 0:
            terms.append(f"{c:.6g}")
        elif power == 1:
            terms.append(f"({c:.6g})t")
        else:
            terms.append(f"({c:.6g})t^{power}")
    return " + ".join(terms).replace("+ -", "- ") if terms else "0"

def make_csv(frame):
    return frame.to_csv(index=False).encode("utf-8")

def make_xlsx(summary, data, results):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pd.DataFrame(list(summary.items()), columns=["Field", "Value"]).to_excel(writer, sheet_name="Summary", index=False)
        data.to_excel(writer, sheet_name="Working Data", index=False)
        if results is not None and not results.empty:
            results.to_excel(writer, sheet_name="Calculated Years", index=False)
    return output.getvalue()

def make_pdf(title, summary, data, results, fig):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(letter), rightMargin=0.45*inch,
                            leftMargin=0.45*inch, topMargin=0.4*inch, bottomMargin=0.4*inch)
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles["Title"]), Spacer(1, 8)]
    summary_rows = [["Field", "Value"]] + [[str(k), str(v)] for k, v in summary.items()]
    table = Table(summary_rows, colWidths=[2.0*inch, 7.4*inch], repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1f4e5f")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("LEFTPADDING", (0,0), (-1,-1), 5), ("RIGHTPADDING", (0,0), (-1,-1), 5),
    ]))
    story += [table, Spacer(1, 10)]
    image_buf = io.BytesIO()
    fig.savefig(image_buf, format="png", dpi=140, bbox_inches="tight")
    image_buf.seek(0)
    story += [Image(image_buf, width=8.7*inch, height=4.2*inch), Spacer(1, 8),
              Paragraph("Limitations: R-squared and MAE describe the in-sample fit, not proven forecasting accuracy. "
                        "Statistically validated prediction intervals are not provided. Polynomial extrapolations can be implausible.",
                        styles["BodyText"]), Spacer(1, 8),
              Paragraph("Working data preview", styles["Heading2"])]
    preview = [["Year", "Value"]] + data.head(18).astype({"Year": int}).values.tolist()
    dt = Table(preview, colWidths=[1.0*inch, 1.6*inch], repeatRows=1)
    dt.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.3,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("FONTSIZE",(0,0),(-1,-1),8)]))
    story.append(dt)
    if results is not None and not results.empty:
        story += [Spacer(1, 8), Paragraph("Calculated-year results", styles["Heading2"])]
        rr = [["Year", "Estimate", "Classification"]] + results[["Year", "Estimate", "Classification"]].astype(str).values.tolist()
        rt = Table(rr, repeatRows=1)
        rt.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.3,colors.grey),("BACKGROUND",(0,0),(-1,0),colors.lightgrey),("FONTSIZE",(0,0),(-1,-1),8)]))
        story.append(rt)
    doc.build(story)
    return output.getvalue()

st.title("Hispanic Heritage Historical Data Explorer")
st.write("Investigate genuine historical indicators for Hispanic countries and Latin America using polynomial models. "
         "Brazil is included as a Latin American comparison; it is not a Spanish-speaking Hispanic country.")
st.caption("Source: World Bank Indicators API. Historical relationships are descriptive, not proof of causation.")

with st.sidebar:
    st.header("Explore the data")
    country = st.selectbox("Country", list(COUNTRIES.keys()), index=list(COUNTRIES.keys()).index("Mexico"))
    indicator = st.selectbox("Indicator", list(INDICATORS.keys()))
    degree = st.slider("Polynomial degree", 3, 8, 3)
    sample_step = st.slider("Model marker sampling (years)", 1, 10, 1)
    horizon = st.slider("Forward display horizon (years)", 0, 50, 5)
    compare_countries = st.multiselect("Compare same indicator with", [c for c in COUNTRIES if c != country], default=["Colombia"] if country != "Colombia" else ["Mexico"])
    printer_mode = st.toggle("Printer-friendly view", value=False)

meta = INDICATORS[indicator]
current_year = datetime.now().year
start_year, end_year = current_year - 70, current_year - 1
st.subheader(f"{country}: {indicator}")
st.write(meta["definition"])
st.caption(f"Indicator code: {meta['code']} · Units: {meta['units']} · Requested years: {start_year}–{end_year}")
st.markdown(f"Source: [World Bank indicator page](https://data.worldbank.org/indicator/{meta['code']}) · "
            f"[API query](https://api.worldbank.org/v2/country/{COUNTRIES[country]}/indicator/{meta['code']}?date={start_year}:{end_year}&format=json)")

try:
    with st.spinner("Retrieving official World Bank data..."):
        source_data = fetch_world_bank(COUNTRIES[country], meta["code"], start_year, end_year)
except requests.exceptions.Timeout:
    st.error("The World Bank request timed out. Please try again later.")
    st.stop()
except requests.exceptions.RequestException as exc:
    st.error(f"Could not retrieve World Bank data: {exc}")
    st.stop()
except (ValueError, TypeError) as exc:
    st.error(f"Could not read the World Bank response: {exc}")
    st.stop()

if source_data.empty:
    st.warning("The World Bank returned no non-missing observations for this country and indicator in the requested window. No replacement data were fabricated.")
    st.stop()

if "working_data" not in st.session_state or st.session_state.get("working_key") != (country, indicator):
    st.session_state.working_data = source_data.copy()
    st.session_state.working_key = (country, indicator)
working = st.session_state.working_data.copy()
st.caption(f"Retrieved {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} · "
           f"Source coverage: {int(source_data.Year.min())}–{int(source_data.Year.max())} · "
           f"{len(source_data)} observations")
st.info("Edit the working table below. Changes are treated as a user-modified scenario and do not alter the original source data.")

left, right = st.columns([3, 1])
with left:
    edited = st.data_editor(working, num_rows="dynamic", use_container_width=True, key=f"editor_{country}_{indicator}")
with right:
    st.write("**Data controls**")
    if st.button("Reset to source", use_container_width=True):
        st.session_state.working_data = source_data.copy()
        st.rerun()
    st.write(f"Original observations: {len(source_data)}")
    st.write(f"Working rows: {len(edited)}")

edited_flag = not edited.reset_index(drop=True).equals(source_data.reset_index(drop=True))
try:
    if not {"Year", "Value"}.issubset(edited.columns):
        raise ValueError("The table must contain Year and Value columns.")
    cleaned = edited[["Year", "Value"]].copy()
    cleaned["Year"] = pd.to_numeric(cleaned["Year"], errors="coerce")
    cleaned["Value"] = pd.to_numeric(cleaned["Value"], errors="coerce")
    if cleaned.isna().any().any():
        raise ValueError("Every row must have a numeric year and value. Remove incomplete rows or correct them.")
    if not np.all(np.equal(cleaned["Year"], np.floor(cleaned["Year"]))):
        raise ValueError("Years must be integers.")
    cleaned["Year"] = cleaned["Year"].astype(int)
    if cleaned["Year"].duplicated().any():
        raise ValueError("Duplicate years are not allowed.")
    if (cleaned["Year"] < 1900).any() or (cleaned["Year"] > current_year).any():
        raise ValueError(f"Years must be between 1900 and {current_year}.")
    if indicator in ("Life expectancy at birth", "Crude birth rate", "Intentional homicide rate") and (cleaned["Value"] < 0).any():
        raise ValueError("Values for this indicator cannot be negative.")
    if indicator == "Life expectancy at birth" and (cleaned["Value"] > 125).any():
        raise ValueError("Life expectancy above 125 years is not accepted.")
    if indicator == "Unemployment rate (national estimate)" and ((cleaned["Value"] < 0) | (cleaned["Value"] > 100)).any():
        raise ValueError("Unemployment percentage must be between 0 and 100.")
    cleaned = cleaned.sort_values("Year").reset_index(drop=True)
    if cleaned.empty:
        raise ValueError("At least one valid row is required.")
except ValueError as exc:
    st.error(f"Working table validation: {exc}")
    st.stop()

if edited_flag:
    st.warning("User-modified scenario: model results and downloads use the edited working table.")
else:
    st.success("Working table matches the source observations.")

try:
    clean, base_year, coeff, predicted, residuals, r2, mae = fit_model(cleaned, degree)
except ValueError as exc:
    st.error(str(exc))
    st.stop()

def predict(years):
    return np.polyval(coeff, np.asarray(years, dtype=float) - base_year)

fit_min, fit_max = int(clean.Year.min()), int(clean.Year.max())
plot_end = max(fit_max, min(2100, fit_max + horizon))
smooth_years = np.linspace(fit_min, plot_end, max(300, (plot_end-fit_min+1)*2))
fig, ax = plt.subplots(figsize=(10, 5.2))
ax.scatter(clean.Year, clean.Value, label="Observed / working data", s=22)
hist_curve = smooth_years[smooth_years <= fit_max]
future_curve = smooth_years[smooth_years > fit_max]
if len(hist_curve):
    ax.plot(hist_curve, predict(hist_curve), label="Fitted model (historical range)", linewidth=2)
if len(future_curve):
    ax.plot(future_curve, predict(future_curve), "--", label="Extrapolated model", linewidth=2)
ax.axvline(fit_max, linestyle=":", linewidth=1.5, label="End of fitted range")
sample_years = np.arange(fit_min, plot_end + 1, sample_step)
ax.scatter(sample_years, predict(sample_years), marker="x", s=28, label=f"Sampled model markers ({sample_step}-year step)")
comparison_notes = []
for comp in compare_countries:
    try:
        comp_data = fetch_world_bank(COUNTRIES[comp], meta["code"], start_year, end_year)
        if not comp_data.empty:
            ax.plot(comp_data.Year, comp_data.Value, marker=".", markersize=3, linewidth=1.2, label=f"{comp} observed")
            comparison_notes.append(f"{comp}: {int(comp_data.Year.min())}–{int(comp_data.Year.max())}, {len(comp_data)} observations")
        else:
            comparison_notes.append(f"{comp}: no observations returned")
    except Exception as exc:
        comparison_notes.append(f"{comp}: retrieval failed ({exc})")
ax.set_xlabel("Year")
ax.set_ylabel(meta["units"])
ax.set_title(f"{country} — {indicator}")
ax.grid(True, alpha=0.25)
ax.legend(fontsize=8, loc="best")
fig.tight_layout()

summary = {
    "Country": country,
    "Indicator": indicator,
    "Indicator code": meta["code"],
    "Units": meta["units"],
    "Definition": meta["definition"],
    "Source": "World Bank Indicators API",
    "Source URL": f"https://data.worldbank.org/indicator/{meta['code']}",
    "Retrieved UTC": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
    "Actual fitted coverage": f"{fit_min}–{fit_max}",
    "Observation count": len(clean),
    "Data status": "User-modified scenario" if edited_flag else "Source data",
    "Polynomial degree": degree,
    "Base year": base_year,
    "Equation (t = year - base year)": equation_text(coeff, base_year),
    "R-squared (in-sample)": f"{r2:.8g}",
    "MAE (in-sample)": f"{mae:.8g} {meta['units']}",
    "Mean": f"{clean.Value.mean():.8g} {meta['units']}",
    "Comparison coverage": "; ".join(comparison_notes) if comparison_notes else "No comparison selected",
}
if printer_mode:
    st.subheader("Printer-friendly report")
    st.write(f"**{country} — {indicator}** ({meta['units']})")
    st.write(f"Coverage: {fit_min}–{fit_max} · n={len(clean)} · Degree {degree} · R²={r2:.5f} · MAE={mae:.5g}")
    st.write(f"Model uses centered time t = year − {base_year}.")
else:
    st.subheader("Historical data and polynomial model")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Observations", len(clean))
    c2.metric("R² (in-sample)", f"{r2:.5f}")
    c3.metric("MAE", f"{mae:.5g}")
    c4.metric("Mean", f"{clean.Value.mean():.5g}")
st.pyplot(fig, use_container_width=True)
plt.close(fig)
st.write(f"**Model equation:** p(t) = {equation_text(coeff, base_year)}, where t = year − {base_year}.")
st.caption(f"Fitted range: {fit_min}–{fit_max}. R² and MAE describe the in-sample fit; they do not prove forecast accuracy. Statistically validated prediction intervals are not provided.")
if comparison_notes:
    st.write("**Comparison coverage:** " + " · ".join(comparison_notes))

st.subheader("Single-year estimate")
single_year = st.number_input("Year to estimate (1950–2100)", min_value=1950, max_value=2100, value=min(max(fit_max, 1950), 2100), step=1)
if st.button("Calculate value"):
    value = float(predict([int(single_year)])[0])
    classification = "within fitted range" if fit_min <= single_year <= fit_max else "extrapolation outside fitted range"
    st.session_state.single_result = {"Year": int(single_year), "Estimate": value, "Units": meta["units"], "Classification": classification}
if "single_result" in st.session_state:
    sr = st.session_state.single_result
    st.write(f"**{sr['Year']}:** {sr['Estimate']:.8g} {sr['Units']} — {sr['Classification']}.")
    if sr["Year"] in set(clean.Year):
        observed = float(clean.loc[clean.Year == sr["Year"], "Value"].iloc[0])
        st.write(f"Observed working-data value for that year: {observed:.8g} {meta['units']}.")

st.subheader("Batch estimates")
batch_mode = st.radio("Batch input type", ["Year range", "Custom years"], horizontal=True)
batch_years = []
if batch_mode == "Year range":
    b1, b2, b3 = st.columns(3)
    with b1:
        batch_start = st.number_input("Start year", 1950, 2100, max(1950, min(fit_min, 2100)), step=1)
    with b2:
        batch_end = st.number_input("End year", 1950, 2100, max(1950, min(fit_max, 2100)), step=1)
    with b3:
        batch_step = st.number_input("Step (1–10 years)", 1, 10, 1, step=1)
    if batch_end <= batch_start:
        st.warning("End year must be greater than start year.")
    else:
        batch_years = list(range(int(batch_start), int(batch_end) + 1, int(batch_step)))
else:
    custom = st.text_input("Enter comma-separated years", f"{fit_min}, {fit_max}, {min(fit_max + 5, 2100)}")
    seen = set()
    invalid = []
    for token in custom.split(","):
        token = token.strip()
        if not token:
            continue
        try:
            num = float(token)
            if not num.is_integer() or not 1950 <= num <= 2100:
                invalid.append(token)
            elif int(num) not in seen:
                batch_years.append(int(num))
                seen.add(int(num))
        except ValueError:
            invalid.append(token)
    if invalid:
        st.warning("Ignored invalid/out-of-range entries: " + ", ".join(invalid))
st.write("**Valid years:** " + (", ".join(map(str, batch_years)) if batch_years else "None"))
if st.button("Calculate batch", disabled=not batch_years):
    rows = []
    for year in batch_years:
        cls = "within fitted range" if fit_min <= year <= fit_max else "extrapolation outside fitted range"
        rows.append({"Year": year, "Estimate": float(predict([year])[0]), "Units": meta["units"], "Classification": cls})
    st.session_state.batch_results = pd.DataFrame(rows)
if "batch_results" in st.session_state:
    results = st.session_state.batch_results
    st.dataframe(results, use_container_width=True)
else:
    results = pd.DataFrame()

st.subheader("Downloads")
st.download_button("Download working data (CSV)", data=make_csv(clean), file_name=f"{country}_{meta['code']}_working_data.csv", mime="text/csv")
xlsx_bytes = make_xlsx(summary, clean, results)
st.download_button("Download Excel workbook (.xlsx)", data=xlsx_bytes, file_name=f"{country}_{meta['code']}_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
pdf_bytes = make_pdf("Hispanic Heritage Historical Data Explorer", summary, clean, results, fig)
st.download_button("Download PDF report", data=pdf_bytes, file_name=f"{country}_{meta['code']}_report.pdf", mime="application/pdf")

st.subheader("Interpretation prompts")
st.write("1. What changed in the indicator over the observed period?")
st.write("2. How does the comparison country differ, and are the years/definitions comparable?")
st.write("3. Which historical or cultural context sources help explain the pattern without claiming that the indicator alone proves a cause?")
st.write("4. How do the degree, residuals, and extrapolation distance affect confidence in the model?")
st.caption("A higher-degree polynomial may fit historical data closely while behaving unrealistically outside that range. Do not treat this model as a validated forecast.")
