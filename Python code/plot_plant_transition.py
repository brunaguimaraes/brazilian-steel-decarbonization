# -*- coding: utf-8 -*-
"""
Plant-level technology transition figure.

Reads one model results workbook (resultados_modelo_*.xlsx) and the existing
plants input file, and draws the plant transition grid:
  - Existing plants (P01..Pn): original route until Retrofitdate, then gray.
  - Successors: one row per successor unit, colored by the chosen route.
  - Greenfields: one row per route, number of active plants in each cell,
    color intensity scaled with the count.

Two ways of running (block "1. USER SETTINGS"):
  - Batch mode (BATCH_DIR = folder name): one figure for every results
    workbook found in the folder and in its subfolders.
  - Single file mode (BATCH_DIR = None): one figure for RESULTS_FILE.

Requires: pandas, openpyxl, matplotlib.
"""

from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Patch
from matplotlib.colors import to_rgb

# =============================================================================
# 1. USER SETTINGS
# =============================================================================
# Folder structure (paths are built from the location of this script):
#   Research Data_2/
#       Python code/   plot_plant_transition.py   (this file)
#       input/         Input_existing_plants.xlsx, batch_transition/
#       resultados/    figures_transition/        (figures are saved here)
BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "input"

# --- Batch mode: every results workbook inside BATCH_DIR (and its subfolders)
BATCH_DIR = INPUT_DIR / "batch_transition"      # None = single file mode
BATCH_PATTERN = "resultados_modelo*.xlsx"
OUTPUT_DIR = BASE_DIR / "resultados" / "figures_transition"

# --- Single file mode (used only when BATCH_DIR = None)
RESULTS_FILE = INPUT_DIR / "batch_transition" / "resultados_modelo_V20_11.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "Figure_plant_transition.png"

# --- Common settings
PLANTS_FILE = INPUT_DIR / "Input_existing_plants.xlsx"
PLANTS_SHEET = "Planilha1"

YEARS = [2023, 2025, 2030, 2035, 2040, 2045, 2050]   # columns of the grid
TITLE = None            # single file mode: e.g. "Emission Reduction"; None = no title
TITLE_FROM_NAME = False # batch mode: True = scenario name as figure title
SHOW_NOTE = True        # explanatory note below the grid
DPI = 300

# =============================================================================
# 2. STYLE (route names, labels and colors)
# =============================================================================
# Route names as written in the plants input file -> names used in the results
INPUT_ROUTE_MAP = {
    "BF-BOF CM": "BF-BOF_MC",
    "BF-BOF CV": "BF-BOF_CC",
    "EAF": "EAF",
}

# Display order, label and color of each route
ROUTE_STYLE = {          # colors identical to the dashboard (routeColor)
    "BF-BOF_MC":  ("BF-BOF-MC",   "#4C5C68"),
    "BF-BOF_CC":  ("BF-BOF-CHAR", "#E69F00"),
    "BF-BOF-CCS": ("BF-BOF-CCS",  "#0072B2"),
    "DR-NG":      ("DR-NG",       "#CC79A7"),
    "DR-NG-CCS":  ("DR-NG-CCS",   "#882255"),
    "DR-H2":      ("DR-H2",       "#44AA99"),
    "EAF":        ("EAF",         "#117733"),
    "IBT":        ("IBT",         "#56B4E9"),
}
GRAY = "#E8E6E0"
GRAY_LABEL = "Retired / substituted / not built"
MIN_INTENSITY = 0.40    # color intensity of a greenfield cell with 1 plant

# =============================================================================
# 3. DATA
# =============================================================================
def _clean_id(name):
    """Plant IDs use underscores in the results ('SBR00007 BOF' -> 'SBR00007_BOF')."""
    return str(name).strip().replace(" ", "_")


def load_data(results_file, plants_file, plants_sheet, years):
    """Return three tables (existing, successors, greenfields) indexed by row label."""
    plants = pd.read_excel(plants_file, sheet_name=plants_sheet)
    plants["PlantID"] = plants["Plantname"].map(_clean_id)
    plants["Label"] = [f"P{i + 1:02d}" for i in range(len(plants))]
    plants["RouteModel"] = plants["Route_detailed"].map(INPUT_ROUTE_MAP)
    if plants["RouteModel"].isna().any():
        bad = plants.loc[plants["RouteModel"].isna(), "Route_detailed"].unique()
        raise ValueError(f"Route not in INPUT_ROUTE_MAP: {list(bad)}")
    label_of = dict(zip(plants["PlantID"], plants["Label"]))

    # --- Existing plants: route until Retrofitdate (inclusive), then gray
    existing = []
    for _, p in plants.iterrows():
        cells = [p["RouteModel"] if p["Startyear"] <= y <= p["Retrofitdate"] else None
                 for y in years]
        existing.append((p["Label"], cells))

    # Consistency check against the production actually reported by the model
    prod = pd.read_excel(results_file, sheet_name="Production_long")
    prod = prod[(prod["Type"] == "Existing") & (prod["Year"].isin(years))]
    for _, p in plants.iterrows():
        for y in years:
            q = prod.loc[(prod["Source"] == p["PlantID"]) & (prod["Year"] == y),
                         "Production_kt"].sum()
            active = p["Startyear"] <= y <= p["Retrofitdate"]
            if q > 1 and not active:
                print(f"WARNING: {p['PlantID']} produces in {y} after Retrofitdate.")

    active = pd.read_excel(results_file, sheet_name="Active_status")
    active = active[(active["Year"].isin(years)) & (active["N_plants"] > 0)]

    # --- Successors: one row per successor unit actually built
    succ = active[active["Type"] == "Successor"]
    successors = []
    for unit, g in succ.groupby("Unit"):
        old_plant = unit.split("__")[0]
        route_by_year = dict(zip(g["Year"], g["Route"]))
        cells = [route_by_year.get(y) for y in years]
        order = int(label_of.get(old_plant, "P99")[1:])
        successors.append((order, f"Successor of {label_of.get(old_plant, old_plant)}", cells))
    successors = [(lab, cells) for _, lab, cells in sorted(successors)]

    # --- Greenfields: one row per route, number of plants in each year
    green = active[active["Type"] == "Greenfield"]
    counts = green.pivot_table(index="Route", columns="Year", values="N_plants",
                               aggfunc="sum").reindex(columns=years).fillna(0)
    greenfields = []
    for route in ROUTE_STYLE:
        if route in counts.index:
            greenfields.append((f"Greenfield {ROUTE_STYLE[route][0]}", route,
                                [int(round(n)) for n in counts.loc[route]]))
    unknown = set(counts.index) - set(ROUTE_STYLE)
    if unknown:
        raise ValueError(f"Route not in ROUTE_STYLE: {sorted(unknown)}")

    return existing, successors, greenfields


# =============================================================================
# 4. FIGURE
# =============================================================================
def _blend(color, intensity):
    """Mix a color with white (intensity 1 = full color)."""
    r, g, b = to_rgb(color)
    return tuple(1 - intensity * (1 - c) for c in (r, g, b))


def _cell(ax, x, y, color, text=None):
    w, h, gap = 1.0, 1.0, 0.04
    ax.add_patch(FancyBboxPatch((x + gap, y + gap + 0.03), w - 2 * gap, h - 2 * gap - 0.06,
                                boxstyle="round,pad=0,rounding_size=0.06",
                                facecolor=color, edgecolor="none"))
    if text is not None:
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
                color="white", fontsize=9, fontweight="bold")


def plot_plant_transition(results_file=RESULTS_FILE, plants_file=PLANTS_FILE,
                          plants_sheet=PLANTS_SHEET, output_file=OUTPUT_FILE,
                          years=YEARS, title=TITLE, show_note=SHOW_NOTE, dpi=DPI):
    existing, successors, greenfields = load_data(results_file, plants_file,
                                                  plants_sheet, years)
    n_years = len(years)
    max_count = max([max(c) for _, _, c in greenfields] + [1])
    used_routes = set()

    # Vertical layout (row 0 at the top); section headers take one row each
    rows = [("plant", lab, cells) for lab, cells in existing]
    if successors:
        rows.append(("header", "Successors (1 plant each)", None))
        rows += [("succ", lab, cells) for lab, cells in successors]
    if greenfields:
        rows.append(("header", "Greenfields (n plants)", None))
        rows += [("green", lab, (route, cnt)) for lab, route, cnt in greenfields]

    fig_h = 0.33 * len(rows) + 2.0
    fig, ax = plt.subplots(figsize=(8.2, fig_h))

    for j, y in enumerate(years):
        ax.text(j + 0.5, 0.35, str(y), ha="center", va="center", fontsize=10, color="#333333")

    for i, (kind, label, data) in enumerate(rows):
        top = -(i + 1)                      # y of the bottom edge of the row
        if kind == "header":
            ax.text(-0.15, top + 0.25, label, ha="right", va="center",
                    fontsize=8.5, style="italic", color="#555555")
            ax.plot([0, n_years], [top + 0.25, top + 0.25], color="#BBBBBB",
                    lw=0.8, ls=(0, (4, 3)))
            continue
        ax.text(-0.15, top + 0.5, label, ha="right", va="center", fontsize=9.5,
                style="normal" if kind == "plant" else "italic", color="#222222")
        for j in range(n_years):
            if kind == "green":
                route, cnt = data
                if cnt[j] > 0:
                    share = cnt[j] / max_count
                    color = _blend(ROUTE_STYLE[route][1],
                                   MIN_INTENSITY + (1 - MIN_INTENSITY) * share)
                    _cell(ax, j, top, color, str(cnt[j]))
                    used_routes.add(route)
                else:
                    _cell(ax, j, top, GRAY)
            else:
                route = data[j]
                if route is None:
                    _cell(ax, j, top, GRAY)
                else:
                    _cell(ax, j, top, ROUTE_STYLE[route][1])
                    used_routes.add(route)

    ax.set_xlim(-2.6, n_years + 0.1)
    ax.set_ylim(-len(rows) - 0.2, 0.8)
    ax.axis("off")

    handles = [Patch(facecolor=c, label=lab) for r, (lab, c) in ROUTE_STYLE.items()
               if r in used_routes]
    handles.append(Patch(facecolor=GRAY, label=GRAY_LABEL))
    fig.legend(handles=handles, loc="upper left", bbox_to_anchor=(0.02, 0.995),
               ncol=4, frameon=False, fontsize=9.5, handlelength=1.2,
               handleheight=1.2, columnspacing=1.4)
    if title:
        fig.suptitle(title, x=0.98, y=0.995, ha="right", va="top", fontsize=10.5,
                     fontweight="bold")
    if show_note:
        fig.text(0.02, 0.012,
                 "Numbers on greenfield rows indicate the number of plants of that route "
                 "active in the year; color intensity scales with the count.\n"
                 "Existing plants and successors are individual units (one plant per row); "
                 "greenfield rows aggregate multiple plants of the same route.",
                 fontsize=8, color="#666666", ha="left", va="bottom")

    top_margin = 1 - 0.75 / fig_h
    bottom_margin = (0.65 if show_note else 0.15) / fig_h
    fig.subplots_adjust(left=0.02, right=0.98, top=top_margin, bottom=bottom_margin)
    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(output_file, dpi=dpi, facecolor="white")
    print(f"Figure saved: {output_file}")
    return fig


# =============================================================================
# 5. BATCH
# =============================================================================
def _scenario_name(path, batch_dir):
    """Subfolder name if the file is inside a scenario subfolder, else file name."""
    if path.parent.resolve() != batch_dir.resolve():
        return path.parent.name
    return path.stem


def run_batch(batch_dir=BATCH_DIR, pattern=BATCH_PATTERN, output_dir=OUTPUT_DIR):
    batch_dir = Path(batch_dir)
    if not batch_dir.is_dir():
        raise FileNotFoundError(f"Batch folder not found: {batch_dir.resolve()}")
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    files = sorted(f for f in batch_dir.rglob(pattern) if not f.name.startswith("~$") and "REFERENCE" not in f.name.upper())
    if not files:
        print(f"No file matching '{pattern}' in {batch_dir.resolve()}")
        return

    summary = []
    for f in files:
        name = _scenario_name(f, batch_dir)
        try:
            fig = plot_plant_transition(
                results_file=f,
                output_file=out_dir / f"Figure_plant_transition_{name}.png",
                title=name if TITLE_FROM_NAME else None)
            plt.close(fig)
            summary.append((name, "OK"))
        except Exception as err:                      # keep going with the next file
            summary.append((name, f"ERROR: {err}"))

    print("\n" + "=" * 78 + "\nBATCH SUMMARY\n" + "=" * 78)
    for name, status in summary:
        print(f"  {name:<45} {status}")


# =============================================================================
# 6. RUN
# =============================================================================
if __name__ == "__main__":
    if BATCH_DIR:
        run_batch()
    else:
        plot_plant_transition()
        plt.show()
