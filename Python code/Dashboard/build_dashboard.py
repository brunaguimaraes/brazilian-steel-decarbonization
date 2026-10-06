# -*- coding: utf-8 -*-
"""
Build the interactive results dashboard from the model output workbooks.

Input:  one results workbook (.xlsx) per scenario, as written by the steel
        model, plus the file dashboard_template.html (layout and chart code).
Output: one self-contained HTML file that opens in any browser, offline.

How to use:
  1. Put the results workbooks in RESULTS_DIR.
  2. List your scenarios (and, optionally, sensitivities) in the block
     "1. USER SETTINGS". The default values reproduce the article dashboard.
  3. Run this file. Requires: pandas, openpyxl.

Sheets read from each workbook:
  Production_route   Route x years (kt)                      required
  Emissions          Year, Emissions_tCO2 [, Cap_tCO2]       required
  Energy_mix         Year, Fossil/Renewable/Electricity %    required
  Fuel_use_PJ        Fuel x years (PJ)                       required
  Production_type    Type x years (kt)                       optional
  CO2_captured       Year, CO2_captured_<route>_tCO2         optional
"""

import json
import re
from pathlib import Path

import pandas as pd

# =============================================================================
# 1. USER SETTINGS
# =============================================================================
# --- Paths (built from the location of this script; edit if needed) ----------
#   Research Data_2/
#       Python code/Dashboard/   build_dashboard.py, dashboard_template.html
#       input/batch_dashboard/   results workbooks (.xlsx)
#       resultados/              steel_dashboard.html (output)
SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent.parent
RESULTS_DIR = BASE_DIR / "input" / "batch_dashboard"
TEMPLATE_HTML = SCRIPT_DIR / "dashboard_template.html"
OUTPUT_HTML = BASE_DIR / "resultados" / "steel_dashboard.html"

# --- Texts --------------------------------------------------------------------
TITLE = "Brazilian iron & steel: deep decarbonization pathways"
SUBTITLE = "Plant-level optimization results"     # period and counts are added

# --- Period shown (None = every year available in the workbooks) -------------
FIRST_YEAR = None
LAST_YEAR = 2050

# --- Main scenarios: (label, results file, color) -----------------------------
SCENARIOS = [
    ("Reference",              "resultados_modelo_V20_13_REFERENCE.xlsx",                 "#888888"),
    ("Emission Stabilization", "resultados_modelo_V20_13_NIS_V2.xlsx",                    "#0072B2"),
    ("Emission Reduction",     "resultados_modelo_V20_13_MIT_V2 -H2_27_GN_15_EL_31.xlsx", "#009E73"),
    ("Fossil Restriction",     "resultados_modelo_V20_13_FossilRestriction.xlsx",         "#D55E00"),
]
DEFAULT_SCENARIO = "Emission Reduction"      # scenario shown when the page opens

# --- Sensitivities (optional; use SENSITIVITIES = [] to leave them out) -------
# All sensitivities vary one input over the same base scenario.
SENSITIVITY_BASE = "Emission Reduction"           # label of one main scenario
SENSITIVITY_BASE_NOTE = "H2 27, GN 15, EL 31"     # short description, or ""
# (label, low-case file or None, high-case file or None)
SENSITIVITIES = [
    ("Hydrogen price",    "resultados_modelo_V20_13_MIT_V2 -H2_13_GN_15_EL_31.xlsx",
                          "resultados_modelo_V20_13_MIT_V2 -H2_41_GN_15_EL_31.xlsx"),
    ("Natural gas price", "resultados_modelo_V20_13_MIT_V2 -H2_27_GN_13_EL_31.xlsx",
                          "resultados_modelo_V20_13_MIT_V2 -H2_27_GN_18_EL_31.xlsx"),
    ("Electricity price", "resultados_modelo_V20_13_MIT_V2 -H2_27_GN_15_EL_27.xlsx",
                          "resultados_modelo_V20_13_MIT_V2 -H2_27_GN_15_EL_35.xlsx"),
    ("CCS capital cost",  None,
                          "resultados_modelo_V20_13_MIT_V2 -H2_27_GN_15_EL_31_HIGHCCS.xlsx"),
]

# =============================================================================
# 2. NAMES AND COLORS
# =============================================================================
# Route name in the model -> (label in the dashboard, color). The order here is
# the stacking order. Routes not listed are added at the end with spare colors.
ROUTES = {
    "BF-BOF_MC":  ("BF-BOF-MC",   "#4C5C68"),
    "BF-BOF_CC":  ("BF-BOF-CHAR", "#E69F00"),
    "BF-BOF-CCS": ("BF-BOF-CCS",  "#0072B2"),
    "DR-NG":      ("DR-NG",       "#CC79A7"),
    "DR-NG-CCS":  ("DR-NG-CCS",   "#882255"),
    "DR-H2":      ("DR-H2",       "#44AA99"),
    "EAF":        ("EAF",         "#117733"),
    "IBT":        ("IBT",         "#56B4E9"),
}
FUEL_COLOR = {
    "Coal": "#4C5C68", "Charcoal": "#E69F00", "Coke": "#9AA6AD",
    "Coke Oven Gas": "#D4B996", "Natural Gas": "#D55E00", "Hydrogen": "#44AA99",
    "Electricity": "#117733", "Diesel": "#7A3B00", "Scrap": "#56B4E9",
}
MIX_COLOR = {"Electricity": "#117733", "Fossil": "#4C5C68", "Renewable": "#0072B2"}
TYPE_COLOR = {"Existing": "#4C5C68", "Successor": "#E69F00", "Greenfield": "#0072B2"}
SPARE_COLORS = ["#999933", "#AA4499", "#332288", "#DDCC77", "#661100", "#6699CC"]

DATA_PLACEHOLDER = "/*__DATA__*/null"


# =============================================================================
# 3. READ ONE RESULTS WORKBOOK
# =============================================================================
def _find_file(results_dir, name):
    """Accept the file name as given, or with spaces replaced by underscores."""
    for candidate in (name, name.replace(" ", "_")):
        path = Path(results_dir) / candidate
        if path.exists():
            return path
    raise FileNotFoundError(f"Results file not found: '{name}' in {Path(results_dir).resolve()}")


def _year_columns(df):
    return [c for c in df.columns if isinstance(c, (int, float)) and float(c).is_integer()]


def _sheet(xls, name, path, required=True):
    if name in xls.sheet_names:
        return xls.parse(name)
    if required:
        raise KeyError(f"Sheet '{name}' not found in {path.name}")
    return None


def available_years(path):
    xls = pd.ExcelFile(path)
    return {int(c) for c in _year_columns(_sheet(xls, "Production_route", path))}


def read_scenario(path, years):
    """Return the data of one workbook in the structure used by the dashboard."""
    xls = pd.ExcelFile(path)
    out = {}

    def wide(sheet, key, decimals, required=True):
        df = _sheet(xls, sheet, path, required)
        if df is None:
            return {}
        df = df.set_index(key)
        return {str(k): [round(float(v), decimals) for v in df.loc[k, years]] for k in df.index
                if str(k).strip().lower() != "total"}      # skip total rows, if any

    out["route"] = wide("Production_route", "Route", 1)
    out["ptype"] = wide("Production_type", "Type", 1, required=False)
    out["fuel"] = wide("Fuel_use_PJ", "Fuel", 2)

    em = _sheet(xls, "Emissions", path).set_index("Year")
    has_cap = "Cap_tCO2" in em.columns
    out["emis"] = {str(y): {"e": round(float(em.loc[y, "Emissions_tCO2"]) / 1e6, 2),
                            "cap": round(float(em.loc[y, "Cap_tCO2"]) / 1e6, 2) if has_cap else None}
                   for y in years}

    mix = _sheet(xls, "Energy_mix", path).set_index("Year")
    out["mix"] = {str(y): {"f": round(float(mix.loc[y, "Fossil_share_%"]), 1),
                           "r": round(float(mix.loc[y, "Renewable_share_%"]), 1),
                           "e": round(float(mix.loc[y, "Electricity_share_%"]), 1)}
                  for y in years}

    out["co2cap"] = {}
    cap = _sheet(xls, "CO2_captured", path, required=False)
    if cap is not None:
        cap = cap.set_index("Year")
        for col in cap.columns:
            m = re.fullmatch(r"CO2_captured_(.+)_tCO2", str(col))
            if m and m.group(1).lower() != "total":
                out["co2cap"][m.group(1)] = [round(float(cap.loc[y, col]) / 1e6, 2) for y in years]
    return out


# =============================================================================
# 4. BUILD THE DATA BLOCK
# =============================================================================
def _slug(text):
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")


def _step_years(years):
    """First year, then every multiple of 5 at least 5 years later, then the last year."""
    steps = [years[0]] + [y for y in years if y % 5 == 0 and y >= years[0] + 5]
    if years[-1] not in steps:
        steps.append(years[-1])
    return steps


def build_data(results_dir=RESULTS_DIR):
    if not SCENARIOS:
        raise ValueError("SCENARIOS is empty: list at least one scenario.")
    labels = [s[0] for s in SCENARIOS]
    if DEFAULT_SCENARIO not in labels:
        raise ValueError(f"DEFAULT_SCENARIO '{DEFAULT_SCENARIO}' is not in SCENARIOS.")
    if SENSITIVITIES and SENSITIVITY_BASE not in labels:
        raise ValueError(f"SENSITIVITY_BASE '{SENSITIVITY_BASE}' is not in SCENARIOS.")

    # --- every file used, and the years they all share
    files = {lab: _find_file(results_dir, f) for lab, f, _ in SCENARIOS}
    sens_files = {}
    for lab, low, high in SENSITIVITIES:
        for case, f in (("low", low), ("high", high)):
            if f:
                sens_files[(lab, case)] = _find_file(results_dir, f)
    common = None
    for path in list(files.values()) + list(sens_files.values()):
        ys = available_years(path)
        common = ys if common is None else common & ys
    years = sorted(y for y in common
                   if (FIRST_YEAR is None or y >= FIRST_YEAR) and (LAST_YEAR is None or y <= LAST_YEAR))
    if not years:
        raise ValueError("No year is common to all workbooks within FIRST_YEAR and LAST_YEAR.")
    last = str(years[-1])

    # --- names and colors of routes (configured ones first, then any new route)
    spare = iter(SPARE_COLORS * 10)
    route_label = {r: lab for r, (lab, _) in ROUTES.items()}
    route_color = {lab: col for lab, col in ROUTES.values()}
    route_order = [lab for lab, _ in ROUTES.values()]
    fuel_color = dict(FUEL_COLOR)

    def relabel(scen):
        for block in ("route", "co2cap"):
            new = {}
            for r, v in scen[block].items():
                lab = route_label.get(r, r)
                if lab not in route_color:
                    route_color[lab] = next(spare)
                    route_order.append(lab)
                    print(f"  NOTE: route '{r}' is not in ROUTES; added with a spare color.")
                new[lab] = v
            scen[block] = new
        for f in scen["fuel"]:
            if f not in fuel_color:
                fuel_color[f] = next(spare)
        return scen

    data = {"years": years, "sensYears": years, "scenarios": {}, "scenLabel": {}, "scenColor": {},
            "sensRoutes": {}, "sensGroups": {}, "sensSummary": {}}
    keys = {}
    for lab, _, color in SCENARIOS:
        key = _slug(lab)
        keys[lab] = key
        data["scenarios"][key] = relabel(read_scenario(files[lab], years))
        data["scenLabel"][key] = lab
        data["scenColor"][key] = color
        print(f"  scenario     {lab:<26} <- {files[lab].name}")

    def add_sens(run_key, scen):
        data["sensRoutes"][run_key] = scen["route"]
        data["sensSummary"][run_key] = {
            "fossil": scen["mix"][last]["f"],
            "cap": round(sum(v[-1] for v in scen["co2cap"].values()), 2),
            "emis": scen["emis"][last]["e"],
        }

    if SENSITIVITIES:
        add_sens("base", data["scenarios"][keys[SENSITIVITY_BASE]])
        for lab, low, high in SENSITIVITIES:
            gkey = _slug(lab)
            group = {"label": lab, "low": None, "base": "base", "high": None}
            for case in ("low", "high"):
                if (lab, case) in sens_files:
                    run_key = f"{gkey}_{case}"
                    add_sens(run_key, relabel(read_scenario(sens_files[(lab, case)], years)))
                    group[case] = run_key
                    print(f"  sensitivity  {lab + ' (' + case + ')':<26} <- {sens_files[(lab, case)].name}")
            data["sensGroups"][gkey] = group

    data.update({"routeOrder": route_order, "routeColor": route_color, "fuelColor": fuel_color,
                 "mixColor": MIX_COLOR, "typeColor": TYPE_COLOR})

    n_s, n_g = len(SCENARIOS), len(SENSITIVITIES)
    subtitle = f"{SUBTITLE}, {years[0]}\u2013{years[-1]}. {n_s} scenario{'s' if n_s != 1 else ''}"
    subtitle += f" and {n_g} sensitivit{'ies' if n_g != 1 else 'y'}." if n_g else "."
    data["meta"] = {"title": TITLE, "subtitle": subtitle, "stepYears": _step_years(years),
                    "defaultScenario": keys[DEFAULT_SCENARIO],
                    "baseLabel": SENSITIVITY_BASE if SENSITIVITIES else "",
                    "baseNote": SENSITIVITY_BASE_NOTE if SENSITIVITIES else ""}
    return data


# =============================================================================
# 5. WRITE THE HTML
# =============================================================================
def build_dashboard(results_dir=RESULTS_DIR, template_html=TEMPLATE_HTML,
                    output_html=OUTPUT_HTML):
    template = Path(template_html).read_text(encoding="utf-8")
    if template.count(DATA_PLACEHOLDER) != 1:
        raise ValueError(f"{Path(template_html).name} is not a valid dashboard template.")
    data = build_data(results_dir)
    block = json.dumps(data, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
    Path(output_html).parent.mkdir(parents=True, exist_ok=True)
    Path(output_html).write_text(template.replace(DATA_PLACEHOLDER, block), encoding="utf-8")
    print(f"Dashboard saved: {output_html}")
    return data


# =============================================================================
# 6. RUN
# =============================================================================
if __name__ == "__main__":
    build_dashboard()
