
# Interactive LCA Explorer: Ref-Bead vs U@Bead



from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Tuple, Optional, List

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

# Excel parsing
import openpyxl


# Attempt to import OpenAI (new SDK style)
try:
    from openai import OpenAI

    _OPENAI_AVAILABLE = True
except Exception:
    OpenAI = None
    _OPENAI_AVAILABLE = False


# =============================================================================
# CONSTANTS
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "data"


ID_REF = "ref"
ID_MOF = "mof"

APP_TITLE = "Interactive LCA Explorer: Ref-Bead vs U@Bead"

DEFAULT_CUSTOM_GRIDS = {
    "QC Hydro": 0.002,
    "Canada Avg": 0.0737,
    "UK Grid": 0.225,
    "EU Avg": 0.25,
    "US Avg": 0.38,
    "China Grid": 0.58,
}

# Excel candidates 
EXCEL_CANDIDATES = [

    "All calculations.xlsx",
]

# =============================================================================
# PLOTLY DISPLAY + EXPORT CONFIG (HIGH QUALITY)
# =============================================================================
PLOTLY_CONFIG = {
    "displayModeBar": True,
    "displaylogo": False,
    "scrollZoom": True,
    "responsive": True,
    "toImageButtonOptions": {
        "format": "png",
        "filename": "figure",
        "scale": 5,  # high-res when user clicks the camera icon
    },
}


def apply_publication_style(fig: go.Figure, height: int = 520, title_size: int = 18) -> go.Figure:
    # Streamlit theme base is typically "light" or "dark"
    base = st.get_option("theme.base") or "light"
    is_dark = str(base).strip().lower() == "dark"

    template = "plotly_dark" if is_dark else "plotly_white"
    font_colour = "white" if is_dark else "black"

    fig.update_layout(
        template=template,
        height=height,
        font=dict(size=14, color=font_colour),
        title=dict(font=dict(size=title_size, color=font_colour)),
        legend=dict(font=dict(size=13, color=font_colour)),
        margin=dict(l=40, r=20, t=60, b=50),
    )

    try:
        fig.update_xaxes(
            tickfont=dict(size=13, color=font_colour),
            title_font=dict(size=14, color=font_colour),
        )
        fig.update_yaxes(
            tickfont=dict(size=13, color=font_colour),
            title_font=dict(size=14, color=font_colour),
        )
    except Exception:
        # Non-cartesian traces (eg Sankey)
        pass

    # Bar labels, etc.
    try:
        fig.update_traces(textfont=dict(size=13, color=font_colour))
    except Exception:
        pass

    return fig



# =============================================================================
# NORMALISATION
# =============================================================================
def norm_name(x: str) -> str:
    return str(x).strip().lower().replace("  ", " ")


# =============================================================================
# PNG EXPORT HELPERS
# =============================================================================
@st.cache_data(show_spinner=False)
def _plotly_json_to_png_bytes(fig_json: str, scale: int = 5) -> Optional[bytes]:
    """
    Converts a Plotly figure (as JSON string) to PNG bytes (requires kaleido).
    Export is publication-quality: scale=5 + fixed width/height.
    Cached to avoid regenerating on every rerun.
    """
    try:
        fig = go.Figure(json.loads(fig_json))

        # Force a sensible export size (prevents tiny/low-quality PNGs)
        export_width = 1600
        export_height = 900

        return pio.to_image(
            fig,
            format="png",
            scale=scale,
            width=export_width,
            height=export_height,
        )
    except Exception:
        return None


def plotly_png_bytes(fig: go.Figure, scale: int = 5) -> Optional[bytes]:
    try:
        return _plotly_json_to_png_bytes(fig.to_json(), scale=scale)
    except Exception:
        return None


def add_png_download_button(fig: go.Figure, filename: str, key: str, label: str = "Download PNG") -> None:
    """
    Adds a PNG download button for a Plotly figure.
    If kaleido is unavailable, shows a disabled button.
    """
    png_bytes = plotly_png_bytes(fig, scale=5)

    if png_bytes is None:
        st.download_button(
            f"{label} (requires kaleido)",
            data=b"",
            file_name=f"{filename}.png",
            mime="image/png",
            disabled=True,
            key=f"dl_{key}",
        )
        st.caption("PNG export requires the Plotly image export dependency (kaleido) to be available.")
    else:
        st.download_button(
            label,
            data=png_bytes,
            file_name=f"{filename}.png",
            mime="image/png",
            key=f"dl_{key}",
        )


# =============================================================================
# EXCEL LOADING 
# =============================================================================
def _find_excel_path() -> Optional[Path]:
    """
    Finds the Excel file automatically.
    Checks:
      1) Same folder as script
      2) /mnt/data
    """
    candidates: List[Path] = []
    base_dir = Path(__file__).resolve().parent

    for name in EXCEL_CANDIDATES:
        p = base_dir / name
        if p.exists():
            candidates.append(p)

    for name in EXCEL_CANDIDATES:
        p = Path("/mnt/data") / name
        if p.exists():
            candidates.append(p)

    return candidates[0] if candidates else None


def _ws_cell(ws, r: int, c: int):
    return ws.cell(row=r, column=c).value


def _find_row_with_first_cell(ws, text: str, col: int = 1, max_rows: int = 300) -> Optional[int]:
    """
    Finds a row where the cell in a given column exactly equals `text` (case-insensitive strip).
    """
    target = str(text).strip().lower()
    for r in range(1, min(ws.max_row, max_rows) + 1):
        v = _ws_cell(ws, r, col)
        if v is None:
            continue
        if isinstance(v, str) and v.strip().lower() == target:
            return r
    return None


def _read_table_down(ws, start_row: int, cols: List[int], stop_on_blank_col: int = 1, max_rows: int = 300) -> List[List[object]]:
    """
    Reads a simple table downwards from start_row until the stop_on_blank_col is blank.
    cols is the list of column indices to extract.
    """
    out = []
    for r in range(start_row, min(ws.max_row, start_row + max_rows) + 1):
        key = _ws_cell(ws, r, stop_on_blank_col)
        if key is None or (isinstance(key, str) and key.strip() == ""):
            break
        out.append([_ws_cell(ws, r, c) for c in cols])
    return out


def _safe_float(x, default: float = 0.0) -> float:
    try:
        if x is None or (isinstance(x, str) and x.strip() == ""):
            return float(default)
        return float(x)
    except Exception:
        return float(default)


def _is_number(x) -> bool:
    return isinstance(x, (int, float, np.integer, np.floating)) and not (isinstance(x, float) and np.isnan(x))


def _extract_unit_value(col_c, col_d) -> Tuple[str, float]:
    """
    Excel Table 1.2 is mixed:
      - Some rows are Unit (C), Value (D)
      - Some rows are Value (C), Unit (D)

    This function detects which is which and returns (unit, value).
    """
    c, d = col_c, col_d

    # Case A: C numeric, D text => Value(C), Unit(D)
    if _is_number(c) and (isinstance(d, str) or d is None):
        return (str(d).strip() if d is not None else "", float(c))

    # Case B: D numeric, C text => Unit(C), Value(D)
    if _is_number(d) and (isinstance(c, str) or c is None):
        return (str(c).strip() if c is not None else "", float(d))

    # Case C: Both numeric (rare) => take D as value
    if _is_number(c) and _is_number(d):
        return ("", float(d))

    # Case D: Numeric stored as string
    try:
        vv = float(d)
        return (str(c).strip() if c is not None else "", float(vv))
    except Exception:
        pass

    try:
        vv = float(c)
        return (str(d).strip() if d is not None else "", float(vv))
    except Exception:
        pass

    return (str(c).strip() if c is not None else "", float("nan"))


def load_tables_from_excel_v11(excel_path: Path) -> Optional[
    Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]
]:
    """
    

    Returns:
      ef_table_df   : Table 1.1 (inputs) + reagent_name column
      lab_table_df  : Table 1.2 (inputs), columns = Item name, Description, Unit, Value
      recipe_df     : Recipe table (inputs), columns = Parameter, Description, Unit, Value
      scale_df      : Scaling symbols (inputs, right block), columns = Parameter, Unit, Value, Notes
      perf_df       : route_id, capacity_mg_g (from Calculations q_ref/q_mof)
      lit_df        : Material, GWP_kgCO2_per_kg, Source, Type
    """
    try:
        wb_val = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return None

    if "inputs" not in wb_val.sheetnames:
        return None

    ws_in = wb_val["inputs"]

    # ---------------------------
    # Table 1.1 Emission Factors
    # ---------------------------
    ef_header_row = _find_row_with_first_cell(ws_in, "Item ", col=1, max_rows=80)
    if ef_header_row is None:
        ef_header_row = _find_row_with_first_cell(ws_in, "Item", col=1, max_rows=80)

    if ef_header_row is None:
        return None

    ef_rows = _read_table_down(
        ws_in,
        start_row=ef_header_row + 1,
        cols=[1, 2, 3, 4, 5, 6],
        stop_on_blank_col=1,
        max_rows=80,
    )
    ef_table_df = pd.DataFrame(ef_rows, columns=["Item", "Description", "Value", "Unit", "Ref", "Link"])

    # Map EF item codes to the reagent names used in Calculations
    item_to_reagent = {
        "EF_elec,CA": "Electricity (Canada)",
        "EF_chitosan": "Chitosan",
        "EF_PDChNF": "PDChNF",
        "EF_acetic": "Acetic Acid",
        "EF_ethanol": "Ethanol",
        "EF_formic": "Formic Acid",
        "EF_ZrCl4": "ZrCl4",
        "EF_2ATA": "2-ATA",
        # optional items in sheet (if present)
        "EF_ECH": "Epichlorohydrin (ECH)",
        "EF_NaOH": "Sodium hydroxide (NaOH)",
        "EF_H2SO4": "Sulfuric acid",
    }

    ef_table_df["reagent_name"] = ef_table_df["Item"].astype(str).map(item_to_reagent).fillna(ef_table_df["Item"].astype(str))
    ef_table_df["Value"] = pd.to_numeric(ef_table_df["Value"], errors="coerce")

    # ---------------------------
    # Table 1.2 Lab Data & Equipment (MIXED Unit/Value)
    # ---------------------------
    lab_header_row = _find_row_with_first_cell(ws_in, "Item name", col=1, max_rows=120)
    if lab_header_row is None:
        return None

    lab_rows = _read_table_down(
        ws_in,
        start_row=lab_header_row + 1,
        cols=[1, 2, 3, 4],
        stop_on_blank_col=1,
        max_rows=220,
    )

    lab_raw = pd.DataFrame(lab_rows, columns=["Item name", "Description", "colC", "colD"])
    units, values = [], []

    for _, rr in lab_raw.iterrows():
        u, v = _extract_unit_value(rr["colC"], rr["colD"])
        units.append(u)
        values.append(v)

    lab_table_df = pd.DataFrame(
        {
            "Item name": lab_raw["Item name"].astype(str).str.strip(),
            "Description": lab_raw["Description"],
            "Unit": units,
            "Value": pd.to_numeric(values, errors="coerce"),
        }
    )

    # ---------------------------
    # Recipe table (inputs rows 51..)
    # ---------------------------
    recipe_header = _find_row_with_first_cell(ws_in, "Parameter", col=1, max_rows=140)
    if recipe_header is None:
        recipe_start = _find_row_with_first_cell(ws_in, "m_p_chitosan", col=1, max_rows=200)
        if recipe_start is None:
            return None
        recipe_header = recipe_start - 1

    recipe_rows = _read_table_down(
        ws_in,
        start_row=recipe_header + 1,
        cols=[1, 2, 3, 4],
        stop_on_blank_col=1,
        max_rows=80,
    )
    recipe_df = pd.DataFrame(recipe_rows, columns=["Parameter", "Description", "Value", "Unit"])
    recipe_df["Value"] = pd.to_numeric(recipe_df["Value"], errors="coerce")

    # Convert conc_formic if unit is "%" and entered as 88 instead of 0.88
    mask_cf = recipe_df["Parameter"].astype(str).str.strip() == "conc_formic"
    if mask_cf.any():
        u = str(recipe_df.loc[mask_cf, "Unit"].iloc[0] or "").strip()
        v = recipe_df.loc[mask_cf, "Value"].iloc[0]
        if u == "%" and pd.notna(v) and float(v) > 1.0:
            recipe_df.loc[mask_cf, "Value"] = float(v) / 100.0
            recipe_df.loc[mask_cf, "Unit"] = "-"

    # Standardise to Unit then Value
    recipe_df = recipe_df[["Parameter", "Description", "Unit", "Value"]].copy()

    # ---------------------------
    # Scaling block (inputs sheet columns H-L)
    # ---------------------------
    # Columns: H Parameter | I Symbol | J Value | K Unit | L Notes
    scale_rows = []
    for r in range(1, min(ws_in.max_row, 250) + 1):
        param = _ws_cell(ws_in, r, 8)
        symbol = _ws_cell(ws_in, r, 9)
        value = _ws_cell(ws_in, r, 10)
        unit = _ws_cell(ws_in, r, 11)
        notes = _ws_cell(ws_in, r, 12)

        if symbol is None:
            continue
        if isinstance(symbol, str) and symbol.strip().lower() == "symbol":
            continue

        scale_rows.append([str(symbol).strip(), param, value, unit, notes])

    scale_df = pd.DataFrame(scale_rows, columns=["Parameter", "Description", "Value", "Unit", "Notes"])
    scale_df["Value"] = pd.to_numeric(scale_df["Value"], errors="coerce")

    # ---------------------------
    # Extra parameters present in inputs left table (C64/C65)
    # ---------------------------
    extra_keys = ["R_solv_recovery", "E_solv_rec_kWh_per_kg"]
    for k in extra_keys:
        found_r = _find_row_with_first_cell(ws_in, k, col=1, max_rows=220)
        if found_r is not None:
            c_val = _ws_cell(ws_in, found_r, 3)  # could be Value or Unit
            d_val = _ws_cell(ws_in, found_r, 4)  # could be Unit or Value
            unit, value = _extract_unit_value(c_val, d_val)
            desc = _ws_cell(ws_in, found_r, 2)

            row = {
                "Parameter": k,
                "Description": desc,
                "Value": value,
                "Unit": unit,
                "Notes": "From inputs table (C64/C65)",
            }

            if (scale_df["Parameter"].astype(str).str.strip() == k).sum() == 0:
                scale_df = pd.concat([scale_df, pd.DataFrame([row])], ignore_index=True)

    # ---------------------------
    # Performance (q_ref, q_mof) from Calculations sheet
    # ---------------------------
    perf_df = pd.DataFrame(
        [
            {"route_id": ID_REF, "capacity_mg_g": np.nan},
            {"route_id": ID_MOF, "capacity_mg_g": np.nan},
        ]
    )
    if "Calculations" in wb_val.sheetnames:
        ws_calc = wb_val["Calculations"]
        q_ref_row = _find_row_with_first_cell(ws_calc, "q_ref", col=1, max_rows=200)
        q_mof_row = _find_row_with_first_cell(ws_calc, "q_mof", col=1, max_rows=200)
        if q_ref_row is not None:
            perf_df.loc[perf_df["route_id"] == ID_REF, "capacity_mg_g"] = pd.to_numeric(_ws_cell(ws_calc, q_ref_row, 2), errors="coerce")
        if q_mof_row is not None:
            perf_df.loc[perf_df["route_id"] == ID_MOF, "capacity_mg_g"] = pd.to_numeric(_ws_cell(ws_calc, q_mof_row, 2), errors="coerce")

    # ---------------------------
    # Literature (Lit comparision)
    # ---------------------------
    lit_df = pd.DataFrame(columns=["Material", "GWP_kgCO2_per_kg", "Source", "Type"])
    if "Lit comparision" in wb_val.sheetnames:
        ws_lit = wb_val["Lit comparision"]
        rows = []
        for r in range(3, min(ws_lit.max_row, 250) + 1):
            mat = _ws_cell(ws_lit, r, 1)
            fu = _ws_cell(ws_lit, r, 2)
            gwp = _ws_cell(ws_lit, r, 3)
            ref = _ws_cell(ws_lit, r, 5)

            if mat is None:
                break

            gwp_val = pd.to_numeric(gwp, errors="coerce")
            typ = "Literature"
            if isinstance(ref, str) and "this work" in ref.lower():
                typ = "This work"

            rows.append(
                {
                    "Material": str(mat),
                    "GWP_kgCO2_per_kg": float(gwp_val) if pd.notna(gwp_val) else np.nan,
                    "Source": str(ref) if ref is not None else "",
                    "Type": typ,
                    "FU": str(fu) if fu is not None else "",
                }
            )

        df_lit_all = pd.DataFrame(rows)
        mask_fu1 = df_lit_all["FU"].astype(str).str.contains("1 kg", case=False, na=False)
        lit_df = df_lit_all.loc[mask_fu1, ["Material", "GWP_kgCO2_per_kg", "Source", "Type"]].copy()

    perf_df["capacity_mg_g"] = pd.to_numeric(perf_df["capacity_mg_g"], errors="coerce")
    perf_df = perf_df.dropna(subset=["capacity_mg_g"])

    return ef_table_df, lab_table_df, recipe_df, scale_df[["Parameter", "Description", "Unit", "Value", "Notes"]], perf_df, lit_df


# =============================================================================
# EMBEDDED DEFAULTS (ONLY USED IF EXCEL NOT FOUND)
# =============================================================================
def embedded_defaults_tables() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Minimal fallback.
    """
    ef_table_df = pd.DataFrame(
        [
            {"Item": "EF_elec,CA", "Description": "Canada grid electricity", "Value": 0.0737, "Unit": "kg CO2-eq/kWh", "Ref": "", "Link": "", "reagent_name": "Electricity (Canada)"},
            {"Item": "EF_chitosan", "Description": "Chitosan", "Value": 41.8, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "Chitosan"},
            {"Item": "EF_PDChNF", "Description": "PDChNF", "Value": 36.65, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "PDChNF"},
            {"Item": "EF_acetic", "Description": "Acetic Acid", "Value": 1.2, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "Acetic Acid"},
            {"Item": "EF_ethanol", "Description": "Ethanol", "Value": 1.24, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "Ethanol"},
            {"Item": "EF_formic", "Description": "Formic Acid", "Value": 2.51, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "Formic Acid"},
            {"Item": "EF_ZrCl4", "Description": "ZrCl4", "Value": 5.0, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "ZrCl4"},
            {"Item": "EF_2ATA", "Description": "2-ATA", "Value": 1.98, "Unit": "kg CO2-eq/kg", "Ref": "", "Link": "", "reagent_name": "2-ATA"},
        ]
    )

    
    lab_table_df = pd.DataFrame(
        [
            {"Item name": "m_bead_batch_Ref", "Description": "Ref bead per batch", "Unit": "g", "Value": 0.4},
            {"Item name": "m_bead_batch_MOF", "Description": "MOF bead per batch", "Unit": "g", "Value": 0.6},
            {"Item name": "P_MF", "Description": "Microfluidiser power", "Unit": "kW", "Value": 1.5},
            {"Item name": "t_MF_batch", "Description": "Microfluidiser time", "Unit": "h", "Value": 0.33},
            {"Item name": "P_hotplate_50", "Description": "Hotplate power", "Unit": "kW", "Value": 0.625},
            {"Item name": "t_mix_50", "Description": "Mix time", "Unit": "h", "Value": 6.0},
            {"Item name": "t_crosslink_50", "Description": "Crosslink time", "Unit": "h", "Value": 6.0},
            {"Item name": "P_stir_UiO", "Description": "MOF stir power", "Unit": "kW", "Value": 0.625},
            {"Item name": "t_Zr_step", "Description": "Zr step time", "Unit": "h", "Value": 12.0},
            {"Item name": "t_linker_step", "Description": "Linker step time", "Unit": "h", "Value": 12.0},
            {"Item name": "P_FD", "Description": "Freeze-dryer power", "Unit": "kW", "Value": 1.84},
            {"Item name": "t_FD_Ref_batch", "Description": "FD time ref", "Unit": "h", "Value": 8.0},
            {"Item name": "t_FD_MOF_batch", "Description": "FD time MOF", "Unit": "h", "Value": 8.0},
            {"Item name": "w_MOF_loading", "Description": "MOF loading fraction", "Unit": "-", "Value": 0.13},
            {"Item name": "V_crosslink_batch", "Description": "Crosslink volume", "Unit": "mL", "Value": 500.0},
            {"Item name": "V_Zr_solution_batch", "Description": "Zr solution volume", "Unit": "mL", "Value": 157.0},
            {"Item name": "V_linker_solution_batch", "Description": "Linker solution volume", "Unit": "mL", "Value": 50.0},
            {"Item name": "P_cent", "Description": "Centrifuge power", "Unit": "kW", "Value": 0.1},
            {"Item name": "t_cent", "Description": "Centrifuge time", "Unit": "h", "Value": 0.05},
            {"Item name": "V_cent_cap", "Description": "Centrifuge cap", "Unit": "mL", "Value": 1000.0},
            {"Item name": "V_dope", "Description": "Dope volume", "Unit": "mL", "Value": 72.0},
            {"Item name": "P_pump", "Description": "Pump power", "Unit": "kW", "Value": 0.005},
            {"Item name": "t_pump", "Description": "Pump time", "Unit": "h", "Value": 0.1},
            {"Item name": "SF_pump_def", "Description": "Pump scaling factor", "Unit": "-", "Value": 0.1},
            {"Item name": "P_stir_motor", "Description": "Stir motor power", "Unit": "kW", "Value": 0.01},
            {"Item name": "t_coag", "Description": "Coag time", "Unit": "h", "Value": 3.0},
            {"Item name": "V_bath", "Description": "Coag bath volume", "Unit": "mL", "Value": 500.0},
            {"Item name": "V_stir_cap", "Description": "Stirrer capacity", "Unit": "L", "Value": 20.0},
        ]
    )

    recipe_df = pd.DataFrame(
        [
            {"Parameter": "m_p_chitosan", "Description": "Chitosan solids", "Unit": "g", "Value": 2.16},
            {"Parameter": "m_p_nano", "Description": "PDChNF solids", "Unit": "g", "Value": 0.36},
            {"Parameter": "v_acetic", "Description": "Acetic acid volume", "Unit": "mL", "Value": 0.36},
            {"Parameter": "rho_acetic", "Description": "Acetic density", "Unit": "g/mL", "Value": 1.049},
            {"Parameter": "m_Zr_batch", "Description": "ZrCl4 per batch", "Unit": "g", "Value": 0.65},
            {"Parameter": "m_ATA_batch", "Description": "2-ATA per batch", "Unit": "g", "Value": 0.47},
            {"Parameter": "v_formic", "Description": "Formic solution volume", "Unit": "mL", "Value": 17.0},
            {"Parameter": "rho_formic", "Description": "Formic density", "Unit": "g/mL", "Value": 1.22},
            {"Parameter": "conc_formic", "Description": "Formic concentration", "Unit": "-", "Value": 0.88},
            {"Parameter": "v_eth", "Description": "Ethanol volume", "Unit": "mL", "Value": 50.0},
            {"Parameter": "rho_eth", "Description": "Ethanol density", "Unit": "g/mL", "Value": 0.789},
        ]
    )

    scale_df = pd.DataFrame(
        [
            {"Parameter": "V_MF_min", "Description": "Microfluidiser minimum volume", "Unit": "mL", "Value": 50.0, "Notes": ""},
            {"Parameter": "Q_MF", "Description": "Microfluidiser flow rate", "Unit": "mL/min", "Value": 120.0, "Notes": ""},
            {"Parameter": "n_passes", "Description": "Number of passes", "Unit": "-", "Value": 3.0, "Notes": ""},
            {"Parameter": "c_PDChNF_disp", "Description": "PDChNF dispersion concentration", "Unit": "g/mL", "Value": 0.01, "Notes": ""},
            {"Parameter": "V_acetic_soln_batch", "Description": "Acetic solution volume per dope batch", "Unit": "mL", "Value": 36.0, "Notes": ""},
            {"Parameter": "V_stir_cap", "Description": "Stirrer capacity", "Unit": "L", "Value": 20.0, "Notes": ""},
            {"Parameter": "V_FD_collector", "Description": "FD collector volume", "Unit": "L", "Value": 4.5, "Notes": ""},
            {"Parameter": "R_FD_24h", "Description": "FD removal rate", "Unit": "L/24h", "Value": 4.0, "Notes": ""},
            {"Parameter": "f_s_wet", "Description": "Wet solids fraction", "Unit": "-", "Value": 0.0338, "Notes": ""},
            {"Parameter": "R_solv_recovery", "Description": "Solvent recovery fraction (ethanol + formic)", "Unit": "-", "Value": 0.9, "Notes": "Luo et al. 2021"},
            {"Parameter": "E_solv_rec_kWh_per_kg", "Description": "Energy for solvent recovery per kg solvent recovered", "Unit": "kWh/kg", "Value": 2.11, "Notes": ""},
        ]
    )

    perf_df = pd.DataFrame(
        [
            {"route_id": ID_REF, "capacity_mg_g": 77.0},
            {"route_id": ID_MOF, "capacity_mg_g": 116.0},
        ]
    )

    lit_df = pd.DataFrame(
        [
            {"Material": "UiO-66-NH2 (solvothermal 1)", "GWP_kgCO2_per_kg": 353.0, "Source": "Luo et al. 2021", "Type": "Literature"},
            {"Material": "UiO-66-NH2 (solvothermal 2)", "GWP_kgCO2_per_kg": 180.0, "Source": "Luo et al. 2021", "Type": "Literature"},
            {"Material": "UiO-66-NH2 (aqueous)", "GWP_kgCO2_per_kg": 43.0, "Source": "Luo et al. 2021", "Type": "Literature"},
            {"Material": "Activated Carbon (Coal)", "GWP_kgCO2_per_kg": 18.28, "Source": "Gu et al. 2018", "Type": "Literature"},
        ]
    )

    return ef_table_df, lab_table_df, recipe_df, scale_df, perf_df, lit_df


# =============================================================================
# HEADER LOGOS
# =============================================================================
def render_header_logos() -> None:
    cov_path = ASSETS_DIR / "cov.png"
    ubc_path = ASSETS_DIR / "ubc.png"


    col_left, col_spacer, col_right = st.columns([1, 2, 1])
    with col_left:
        if cov_path.exists():
            st.image(str(cov_path))
    with col_right:
        if ubc_path.exists():
            st.image(str(ubc_path))


# =============================================================================
# SMALL HELPERS
# =============================================================================
def _val_from_table(df: pd.DataFrame, key_col: str, key: str, val_col: str = "Value", default: float = 0.0) -> float:
    if df is None or df.empty:
        return float(default)
    if key_col not in df.columns:
        return float(default)

    mask = df[key_col].astype(str).str.strip() == str(key).strip()
    if not mask.any():
        return float(default)

    v = df.loc[mask, val_col].iloc[0]
    try:
        return float(v)
    except Exception:
        return float(default)


def build_ef_df_from_table(ef_table_df: pd.DataFrame) -> pd.DataFrame:
    """
    Converts EF table to: reagent_name, GWP_kgCO2_per_kg
    Electricity EF is stored in kg CO2-eq/kWh, but we treat it as "per kg" for multiplication
    with kWh (so conceptually it's kg CO2-eq per kWh).
    """
    if ef_table_df is None or ef_table_df.empty:
        return pd.DataFrame(columns=["reagent_name", "GWP_kgCO2_per_kg"])

    df = ef_table_df.copy()
    if "reagent_name" not in df.columns:
        df["reagent_name"] = df.get("Item", "").astype(str)

    df["GWP_kgCO2_per_kg"] = pd.to_numeric(df["Value"], errors="coerce")
    out = df[["reagent_name", "GWP_kgCO2_per_kg"]].dropna(subset=["reagent_name"]).copy()
    out["reagent_name"] = out["reagent_name"].astype(str).str.strip()
    return out


# =============================================================================
# RECIPE + INVENTORY (FU1)
# =============================================================================
def compute_recipe_routes(
    lab_table_df: pd.DataFrame,
    recipe_table_df: pd.DataFrame,
    mof_loading_fraction: float,
) -> Tuple[pd.DataFrame, float]:
    """
    Builds the routes_df inventory per 1 kg bead output (FU1).


    - w_MOF_loading is stored as a fraction (e.g. 0.13 = 13 wt%)
    - polymer support fraction = 1 - w_MOF_loading
    - Formic/Ethanol base inventories are computed without recovery, then recovery is applied as (1-R)
    """

    # batch yields (g per batch)
    m_ref_g = _val_from_table(lab_table_df, "Item name", "m_bead_batch_Ref", default=0.4)
    m_mof_g = _val_from_table(lab_table_df, "Item name", "m_bead_batch_MOF", default=0.6)

    # w_MOF_loading stored as fraction (0-1). Allow user entry as 13 -> convert
    w_mof = float(mof_loading_fraction)
    if w_mof > 1.0:
        w_mof = w_mof / 100.0
    w_mof = max(0.0, min(w_mof, 0.999999))
    polymer_fraction_support = max(0.0, 1.0 - w_mof)

    # polymer recipe
    m_p_chitosan_g = _val_from_table(recipe_table_df, "Parameter", "m_p_chitosan", default=2.16)
    m_p_nano_g = _val_from_table(recipe_table_df, "Parameter", "m_p_nano", default=0.36)
    pol_total = max(m_p_chitosan_g + m_p_nano_g, 1e-12)

    frac_chi = m_p_chitosan_g / pol_total
    frac_nano = m_p_nano_g / pol_total

    # acetic
    v_acetic_ml = _val_from_table(recipe_table_df, "Parameter", "v_acetic", default=0.36)
    rho_acetic = _val_from_table(recipe_table_df, "Parameter", "rho_acetic", default=1.049)
    m_acetic_g = float(v_acetic_ml) * float(rho_acetic)

    # Excel logic: (Acetic mass in dope / total polymer solids) scaled to 1 kg bead output
    ref_acetic_kg = (m_acetic_g / pol_total) * 1.0  # kg per kg bead
    ref_chitosan_kg = frac_chi * 1.0
    ref_pdchnf_kg = frac_nano * 1.0

    # MOF inventories (base masses without recovery applied here)
    m_zr_g = _val_from_table(recipe_table_df, "Parameter", "m_Zr_batch", default=0.65)
    m_ata_g = _val_from_table(recipe_table_df, "Parameter", "m_ATA_batch", default=0.47)

    v_formic_ml = _val_from_table(recipe_table_df, "Parameter", "v_formic", default=17.0)
    rho_formic = _val_from_table(recipe_table_df, "Parameter", "rho_formic", default=1.22)
    conc_formic = _val_from_table(recipe_table_df, "Parameter", "conc_formic", default=0.88)

    v_eth_ml = _val_from_table(recipe_table_df, "Parameter", "v_eth", default=50.0)
    rho_eth = _val_from_table(recipe_table_df, "Parameter", "rho_eth", default=0.789)

    m_mof_kg = max(float(m_mof_g) / 1000.0, 1e-12)

    zr_per_kg = (float(m_zr_g) / 1000.0) / m_mof_kg
    ata_per_kg = (float(m_ata_g) / 1000.0) / m_mof_kg

    m_formic_solution_g = float(v_formic_ml) * float(rho_formic)
    m_formic_pure_g = m_formic_solution_g * float(conc_formic)
    formic_base_per_kg = (m_formic_pure_g / 1000.0) / m_mof_kg

    m_eth_g = float(v_eth_ml) * float(rho_eth)
    eth_base_per_kg = (m_eth_g / 1000.0) / m_mof_kg

    # support masses for MOF bead
    mof_chitosan_kg = ref_chitosan_kg * polymer_fraction_support
    mof_pdchnf_kg = ref_pdchnf_kg * polymer_fraction_support
    mof_acetic_kg = ref_acetic_kg * polymer_fraction_support

    routes_df = pd.DataFrame(
        [
            # Ref
            {"route_id": ID_REF, "route_name": "Ref-Bead", "reagent_name": "Chitosan", "mass_kg_per_fu": ref_chitosan_kg, "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_REF, "route_name": "Ref-Bead", "reagent_name": "PDChNF", "mass_kg_per_fu": ref_pdchnf_kg, "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_REF, "route_name": "Ref-Bead", "reagent_name": "Acetic Acid", "mass_kg_per_fu": float(ref_acetic_kg), "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            # MOF bead: support
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "Chitosan", "mass_kg_per_fu": mof_chitosan_kg, "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "PDChNF", "mass_kg_per_fu": mof_pdchnf_kg, "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "Acetic Acid", "mass_kg_per_fu": mof_acetic_kg, "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            # MOF reagents (base, no recovery applied yet)
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "ZrCl4", "mass_kg_per_fu": float(zr_per_kg), "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "2-ATA", "mass_kg_per_fu": float(ata_per_kg), "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "Formic Acid", "mass_kg_per_fu": float(formic_base_per_kg), "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
            {"route_id": ID_MOF, "route_name": "U@Bead", "reagent_name": "Ethanol", "mass_kg_per_fu": float(eth_base_per_kg), "electricity_kwh_per_fu": 0.0, "electricity_source": "Electricity (Canada)"},
        ]
    )

    return routes_df, polymer_fraction_support


# =============================================================================
# SCALING ENGINE (ELECTRICITY)
# =============================================================================
@dataclass
class ScalingOutputs:
    polymer_fraction_support: float
    scaling_factors: Dict[str, float]
    baseline_step_df: pd.DataFrame
    scaled_step_df: pd.DataFrame
    route_elec_kwh_per_kg_baseline: Dict[str, float]
    route_elec_kwh_per_kg_scaled: Dict[str, float]


def compute_scaling_from_input_tables(
    lab_table_df: pd.DataFrame,
    recipe_table_df: pd.DataFrame,
    scale_table_df: pd.DataFrame,
    polymer_fraction_support: float,
    solvent_recovery_frac: float,
    solvent_recovery_energy_kwh_per_kg: float,
    mof_solvent_base_mass_kg_per_kgbead: float,
) -> ScalingOutputs:
    """
    

    REF baseline (raw, no utilisation): kWh/kg = sum(P*t)/m_ref_kg for each step.
    REF scaled: multiply each step by its scaling factor:
      SF_MF, SF_mix, SF_cent, SF_pump_def, SF_coag, SF_crosslink, SF_FD_ref

    MOF baseline / scaled:
      Support electricity = REF_total * polymer support fraction
      + Zr stirring + Linker stirring + 2nd FD
      + Solvent recovery electricity

    Excel uses:
      +((fresh_formic + fresh_ethanol) * R/(1-R)) * E

    Since fresh = base*(1-R), this simplifies exactly to:
      Solvent recovery electricity = (base_formic + base_ethanol) * R * E
    """

    # outputs per batch
    m_ref_g = _val_from_table(lab_table_df, "Item name", "m_bead_batch_Ref", default=0.4)
    m_mof_g = _val_from_table(lab_table_df, "Item name", "m_bead_batch_MOF", default=0.6)

    # equipment powers and times
    P_MF = _val_from_table(lab_table_df, "Item name", "P_MF", default=1.5)
    t_MF_batch = _val_from_table(lab_table_df, "Item name", "t_MF_batch", default=0.33)

    P_hotplate = _val_from_table(lab_table_df, "Item name", "P_hotplate_50", default=0.625)
    t_mix = _val_from_table(lab_table_df, "Item name", "t_mix_50", default=6.0)
    t_cross = _val_from_table(lab_table_df, "Item name", "t_crosslink_50", default=6.0)

    P_stir_uio = _val_from_table(lab_table_df, "Item name", "P_stir_UiO", default=0.625)
    t_zr = _val_from_table(lab_table_df, "Item name", "t_Zr_step", default=12.0)
    t_linker = _val_from_table(lab_table_df, "Item name", "t_linker_step", default=12.0)

    P_FD = _val_from_table(lab_table_df, "Item name", "P_FD", default=1.84)
    t_FD_ref = _val_from_table(lab_table_df, "Item name", "t_FD_Ref_batch", default=8.0)
    t_FD_mof = _val_from_table(lab_table_df, "Item name", "t_FD_MOF_batch", default=8.0)

    # volumes
    V_crosslink = _val_from_table(lab_table_df, "Item name", "V_crosslink_batch", default=500.0)
    V_zr_sol = _val_from_table(lab_table_df, "Item name", "V_Zr_solution_batch", default=157.0)
    V_linker_sol = _val_from_table(lab_table_df, "Item name", "V_linker_solution_batch", default=50.0)

    V_dope = _val_from_table(lab_table_df, "Item name", "V_dope", default=72.0)

    # centrifuge + pump + coag
    P_cent = _val_from_table(lab_table_df, "Item name", "P_cent", default=0.1)
    t_cent = _val_from_table(lab_table_df, "Item name", "t_cent", default=0.05)
    V_cent_cap = _val_from_table(lab_table_df, "Item name", "V_cent_cap", default=1000.0)

    P_pump = _val_from_table(lab_table_df, "Item name", "P_pump", default=0.005)
    t_pump = _val_from_table(lab_table_df, "Item name", "t_pump", default=0.1)
    SF_pump_def = _val_from_table(lab_table_df, "Item name", "SF_pump_def", default=0.1)

    P_stir_motor = _val_from_table(lab_table_df, "Item name", "P_stir_motor", default=0.01)
    t_coag = _val_from_table(lab_table_df, "Item name", "t_coag", default=3.0)
    V_bath = _val_from_table(lab_table_df, "Item name", "V_bath", default=500.0)

    # scaling parameters from scale table
    V_MF_min = _val_from_table(scale_table_df, "Parameter", "V_MF_min", default=50.0)
    Q_MF = _val_from_table(scale_table_df, "Parameter", "Q_MF", default=120.0)
    n_passes = _val_from_table(scale_table_df, "Parameter", "n_passes", default=3.0)
    c_PDChNF_disp = _val_from_table(scale_table_df, "Parameter", "c_PDChNF_disp", default=0.01)

    # IMPORTANT: Excel uses V_acetic_soln_batch (36 mL) explicitly in dope volume for scaling
    V_acetic_soln_batch = _val_from_table(scale_table_df, "Parameter", "V_acetic_soln_batch", default=36.0)

    V_stir_cap_L = _val_from_table(scale_table_df, "Parameter", "V_stir_cap", default=20.0)
    V_FD_collector = _val_from_table(scale_table_df, "Parameter", "V_FD_collector", default=4.5)
    R_FD_24h = _val_from_table(scale_table_df, "Parameter", "R_FD_24h", default=4.0)
    f_s_wet = _val_from_table(scale_table_df, "Parameter", "f_s_wet", default=0.0338)

    # derived volumes for MF scaling (Excel logic)
    m_p_nano = _val_from_table(recipe_table_df, "Parameter", "m_p_nano", default=0.36)
    V_PDChNF_disp_batch = float(m_p_nano) / max(float(c_PDChNF_disp), 1e-12)
    V_dope_batch = float(V_PDChNF_disp_batch) + float(V_acetic_soln_batch)

    V_MF_eff = max(float(V_PDChNF_disp_batch), float(V_MF_min))
    t_MF_eff_h = (float(n_passes) * float(V_MF_eff) / max(float(Q_MF), 1e-9)) / 60.0
    SF_MF = float(t_MF_eff_h) / max(float(t_MF_batch), 1e-12)

    V_stir_cap_mL = float(V_stir_cap_L) * 1000.0
    SF_mix = float(V_dope_batch) / max(float(V_stir_cap_mL), 1e-12)
    SF_crosslink = float(V_crosslink) / max(float(V_stir_cap_mL), 1e-12)
    SF_Zr = float(V_zr_sol) / max(float(V_stir_cap_mL), 1e-12)
    SF_linker = float(V_linker_sol) / max(float(V_stir_cap_mL), 1e-12)

    # Freeze-drying scaling (Excel logic)
    t_FD_batch = float(t_FD_ref)
    V_FD_rate_cap_8h = float(R_FD_24h) * (float(t_FD_batch) / 24.0)
    V_FD_lim_8h = min(float(V_FD_collector), float(V_FD_rate_cap_8h))

    m_ref_kg = max(float(m_ref_g) / 1000.0, 1e-12)
    m_mof_kg = max(float(m_mof_g) / 1000.0, 1e-12)

    V_water_ref_L = ((m_ref_kg / max(float(f_s_wet), 1e-12)) - m_ref_kg)
    V_water_mof_L = ((m_mof_kg / max(float(f_s_wet), 1e-12)) - m_mof_kg)

    SF_FD_ref = float(V_water_ref_L) / max(float(V_FD_lim_8h), 1e-12)
    SF_FD_mof = float(V_water_mof_L) / max(float(V_FD_lim_8h), 1e-12)

    # Centrifuge scaling and coag scaling
    SF_cent = float(V_dope) / max(float(V_cent_cap), 1e-12)
    SF_coag = (float(V_bath) / 1000.0) / max(float(V_stir_cap_L), 1e-12)
    SF_pump = float(SF_pump_def)

    scaling_factors = {
        "SF_MF": SF_MF,
        "SF_mix": SF_mix,
        "SF_crosslink": SF_crosslink,
        "SF_cent": SF_cent,
        "SF_pump_def": SF_pump,
        "SF_coag": SF_coag,
        "SF_FD_ref": SF_FD_ref,
        "SF_Zr": SF_Zr,
        "SF_linker": SF_linker,
        "SF_FD_mof": SF_FD_mof,
    }

    def kwh_per_kg(power_kw: float, time_h: float, sf: float, mass_g_out: float) -> float:
        mass_kg_out = max(float(mass_g_out) / 1000.0, 1e-12)
        return (float(power_kw) * float(time_h) * float(sf)) / mass_kg_out

    # REF baseline and scaled step list (Excel order and names)
    ref_steps_base = [
        ("Microfluidizer", kwh_per_kg(P_MF, t_MF_batch, 1.0, m_ref_g)),
        ("Mixing (50°C)", kwh_per_kg(P_hotplate, t_mix, 1.0, m_ref_g)),
        ("Centrifugation", kwh_per_kg(P_cent, t_cent, 1.0, m_ref_g)),
        ("Syringe Pump", kwh_per_kg(P_pump, t_pump, 1.0, m_ref_g)),
        ("Coagulation Stirring", kwh_per_kg(P_stir_motor, t_coag, 1.0, m_ref_g)),
        ("Crosslinking", kwh_per_kg(P_hotplate, t_cross, 1.0, m_ref_g)),
        ("Freeze Drying", kwh_per_kg(P_FD, t_FD_ref, 1.0, m_ref_g)),
    ]
    ref_steps_scaled = [
        ("Microfluidizer", kwh_per_kg(P_MF, t_MF_batch, SF_MF, m_ref_g)),
        ("Mixing (50°C)", kwh_per_kg(P_hotplate, t_mix, SF_mix, m_ref_g)),
        ("Centrifugation", kwh_per_kg(P_cent, t_cent, SF_cent, m_ref_g)),
        ("Syringe Pump", kwh_per_kg(P_pump, t_pump, SF_pump, m_ref_g)),
        ("Coagulation Stirring", kwh_per_kg(P_stir_motor, t_coag, SF_coag, m_ref_g)),
        ("Crosslinking", kwh_per_kg(P_hotplate, t_cross, SF_crosslink, m_ref_g)),
        ("Freeze Drying", kwh_per_kg(P_FD, t_FD_ref, SF_FD_ref, m_ref_g)),
    ]

    ref_total_base = float(sum(v for _, v in ref_steps_base))
    ref_total_scaled = float(sum(v for _, v in ref_steps_scaled))

    # MOF-specific electricity (scaled and baseline steps)
    mof_specific_base = [
        ("Elec: Zr Stirring", kwh_per_kg(P_stir_uio, t_zr, 1.0, m_mof_g)),
        ("Elec: Linker Stirring", kwh_per_kg(P_stir_uio, t_linker, 1.0, m_mof_g)),
        ("Elec: 2nd FD", kwh_per_kg(P_FD, t_FD_mof, 1.0, m_mof_g)),
    ]
    mof_specific_scaled = [
        ("Elec: Zr Stirring", kwh_per_kg(P_stir_uio, t_zr, SF_Zr, m_mof_g)),
        ("Elec: Linker Stirring", kwh_per_kg(P_stir_uio, t_linker, SF_linker, m_mof_g)),
        ("Elec: 2nd FD", kwh_per_kg(P_FD, t_FD_mof, SF_FD_mof, m_mof_g)),
    ]

    # Support electricity (Excel: Total Ref Elec * polymer support fraction)
    support_elec_base = polymer_fraction_support * ref_total_base
    support_elec_scaled = polymer_fraction_support * ref_total_scaled

    # Solvent recovery electricity (Excel equivalent)
    R = max(0.0, min(float(solvent_recovery_frac), 0.999999))
    E = max(0.0, float(solvent_recovery_energy_kwh_per_kg))

    # Excel-equivalent energy = base_solvent_mass * R * E
    if R > 0 and E > 0:
        solv_rec_kwh_per_kg = float(mof_solvent_base_mass_kg_per_kgbead) * R * E
    else:
        solv_rec_kwh_per_kg = 0.0

    # Build step dataframes
    baseline_step_rows = [{"route_id": ID_REF, "Bead": "Ref-Bead", "Step": s, "kWh_per_kg": float(v)} for s, v in ref_steps_base]
    scaled_step_rows = [{"route_id": ID_REF, "Bead": "Ref-Bead", "Step": s, "kWh_per_kg": float(v)} for s, v in ref_steps_scaled]

    baseline_step_rows.append({"route_id": ID_MOF, "Bead": "U@Bead", "Step": "Support Electricity", "kWh_per_kg": float(support_elec_base)})
    scaled_step_rows.append({"route_id": ID_MOF, "Bead": "U@Bead", "Step": "Support Electricity", "kWh_per_kg": float(support_elec_scaled)})

    for s, v in mof_specific_base:
        baseline_step_rows.append({"route_id": ID_MOF, "Bead": "U@Bead", "Step": s, "kWh_per_kg": float(v)})
    for s, v in mof_specific_scaled:
        scaled_step_rows.append({"route_id": ID_MOF, "Bead": "U@Bead", "Step": s, "kWh_per_kg": float(v)})

    if solv_rec_kwh_per_kg > 0:
        baseline_step_rows.append({"route_id": ID_MOF, "Bead": "U@Bead", "Step": "Solvent recovery (ethanol + formic)", "kWh_per_kg": float(solv_rec_kwh_per_kg)})
        scaled_step_rows.append({"route_id": ID_MOF, "Bead": "U@Bead", "Step": "Solvent recovery (ethanol + formic)", "kWh_per_kg": float(solv_rec_kwh_per_kg)})

    baseline_step_df = pd.DataFrame(baseline_step_rows)
    scaled_step_df = pd.DataFrame(scaled_step_rows)

    # Route totals
    mof_total_base = support_elec_base + float(sum(v for _, v in mof_specific_base)) + float(solv_rec_kwh_per_kg)
    mof_total_scaled = support_elec_scaled + float(sum(v for _, v in mof_specific_scaled)) + float(solv_rec_kwh_per_kg)

    route_elec_kwh_per_kg_baseline = {ID_REF: ref_total_base, ID_MOF: mof_total_base}
    route_elec_kwh_per_kg_scaled = {ID_REF: ref_total_scaled, ID_MOF: mof_total_scaled}

    return ScalingOutputs(
        polymer_fraction_support=polymer_fraction_support,
        scaling_factors=scaling_factors,
        baseline_step_df=baseline_step_df,
        scaled_step_df=scaled_step_df,
        route_elec_kwh_per_kg_baseline=route_elec_kwh_per_kg_baseline,
        route_elec_kwh_per_kg_scaled=route_elec_kwh_per_kg_scaled,
    )


# =============================================================================
# IMPACT ENGINE (FU1)
# =============================================================================
def calculate_impacts(
    route_id: str,
    ef_df: pd.DataFrame,
    routes_df: pd.DataFrame,
    efficiency_factor: float = 1.0,
    recycling_rate_pct: float = 0.0,  # solvent recovery (%)
    yield_rate: float = 100.0,
    transport_pct: float = 0.0,
) -> Tuple[Optional[dict], Optional[pd.DataFrame]]:
    """
    Calculates GWP per kg bead (FU1):
    - Electricity (kWh/kg) * EF_elec
    - Reagents mass inventory * EF_reagent
    - Solvent recovery: implemented as lower fresh solvent mass (1-R) in inventory
      AND recovery electricity already included in electricity_kwh_per_fu (scaling engine)
    """
    required_cols = {
        "route_id",
        "route_name",
        "electricity_kwh_per_fu",
        "electricity_source",
        "reagent_name",
        "mass_kg_per_fu",
    }
    missing = required_cols - set(routes_df.columns)
    if missing:
        st.error(f"routes_df missing columns: {sorted(list(missing))}")
        return None, None

    route_data = routes_df[routes_df["route_id"] == route_id].copy()
    if route_data.empty:
        return None, None

    yield_multiplier = 1.0 / max((yield_rate / 100.0), 1e-12)
    R = max(0.0, min(float(recycling_rate_pct) / 100.0, 0.999999))

    # Electricity
    base_elec_kwh = float(route_data.iloc[0]["electricity_kwh_per_fu"])
    elec_kwh = base_elec_kwh * float(efficiency_factor) * yield_multiplier

    elec_source = str(route_data.iloc[0]["electricity_source"])
    ef_elec_row = ef_df[ef_df["reagent_name"] == elec_source]
    ef_elec = float(ef_elec_row["GWP_kgCO2_per_kg"].iloc[0]) if not ef_elec_row.empty else 0.0
    gwp_elec = elec_kwh * ef_elec

    contributions = [{"Component": "Electricity", "Category": "Electricity", "Mass (kg)": 0.0, "GWP": gwp_elec}]
    total_reagent_gwp = 0.0

    solvent_set = {"Ethanol", "Formic Acid"}  # Excel applies recovery to these

    for _, row in route_data.iterrows():
        reagent = str(row["reagent_name"]).strip()
        base_mass_no_recovery = float(row["mass_kg_per_fu"]) if pd.notna(row["mass_kg_per_fu"]) else 0.0
        base_mass_no_recovery = base_mass_no_recovery * yield_multiplier

        if reagent in solvent_set:
            # Excel: fresh solvent = base*(1-R)
            effective_mass = base_mass_no_recovery * (1.0 - R)
        else:
            effective_mass = base_mass_no_recovery

        ef_row = ef_df[ef_df["reagent_name"] == reagent]
        ef_val = float(ef_row["GWP_kgCO2_per_kg"].iloc[0]) if not ef_row.empty else 0.0

        gwp_val = effective_mass * ef_val
        total_reagent_gwp += gwp_val

        if reagent in ["Chitosan", "PDChNF", "Acetic Acid"]:
            cat = "GWP Support"
        elif reagent in ["ZrCl4", "2-ATA", "Ethanol", "Formic Acid"]:
            cat = "GWP MOF Reagents"
        else:
            cat = "Other"

        contributions.append({"Component": reagent, "Category": cat, "Mass (kg)": effective_mass, "GWP": gwp_val})

    raw_total_gwp = gwp_elec + total_reagent_gwp
    transport_gwp = raw_total_gwp * (float(transport_pct) / 100.0)

    if transport_gwp > 0:
        contributions.append({"Component": "Transport", "Category": "Transport", "Mass (kg)": 0.0, "GWP": transport_gwp})

    final_total_gwp = raw_total_gwp + transport_gwp

    results = {
        "id": route_id,
        "name": str(route_data.iloc[0]["route_name"]),
        "Total GWP": final_total_gwp,
        "Electricity GWP": gwp_elec,
        "Non-Electric GWP": total_reagent_gwp + transport_gwp,
        "Electricity kWh": elec_kwh,
        "Electricity EF Used": ef_elec,
    }

    return results, pd.DataFrame(contributions)


# =============================================================================
# SANKEY (VALUES ON-NODES, WITH IMPROVED READABILITY)
# =============================================================================
def plot_sankey_materials_processes(
    results_list: List[dict],
    contrib_df_all: pd.DataFrame,
    step_df: pd.DataFrame,
    route_id: str,
) -> go.Figure:
    target_res = next((r for r in results_list if r["id"] == route_id), None)
    if target_res is None:
        return go.Figure()

    bead_name = target_res["name"]
    total_gwp = float(target_res["Total GWP"])
    elec_gwp_total = float(target_res["Electricity GWP"])

    df_b = contrib_df_all[contrib_df_all["Bead"] == bead_name].copy()
    if df_b.empty:
        return go.Figure()

    df_mat = df_b[(df_b["Component"] != "Electricity") & (df_b["Component"] != "Transport")].copy()
    df_mat = df_mat[df_mat["GWP"] > 0]

    transport_gwp = 0.0
    df_transport = df_b[df_b["Component"] == "Transport"]
    if not df_transport.empty:
        transport_gwp = float(df_transport["GWP"].sum())

    df_steps = step_df[step_df["route_id"] == route_id].copy()
    df_steps = df_steps[df_steps["kWh_per_kg"] > 0]

    step_elec_rows = []
    if not df_steps.empty and elec_gwp_total > 0:
        kwh_vals = df_steps["kWh_per_kg"].to_numpy(dtype=float)
        kwh_sum = float(np.sum(kwh_vals))
        if kwh_sum > 0:
            for step_name, kwh_val in zip(df_steps["Step"].tolist(), kwh_vals.tolist()):
                frac = float(kwh_val) / kwh_sum
                step_elec_rows.append({"Step": step_name, "GWP": elec_gwp_total * frac})

    df_elec_steps = pd.DataFrame(step_elec_rows)

    node_labels: List[str] = []
    node_colors: List[str] = []

    def _add_node(lbl: str, color: str) -> None:
        if lbl not in node_labels:
            node_labels.append(lbl)
            node_colors.append(color)

    for comp in df_mat["Component"].tolist():
        _add_node(f"Material: {comp}", "#90EE90")

    if not df_elec_steps.empty:
        for step_name in df_elec_steps["Step"].tolist():
            _add_node(f"Process: {step_name}", "#FFD700")

    if transport_gwp > 0:
        _add_node("Transport", "#A9A9A9")

    _add_node("Materials supply", "#B6E3B6")
    _add_node("Electricity supply", "#FFE48A")
    _add_node(f"{bead_name} synthesis", "#87CEFA")
    _add_node("Total GWP", "#FF6347")

    idx = {lbl: i for i, lbl in enumerate(node_labels)}

    sources: List[int] = []
    targets: List[int] = []
    values: List[float] = []

    mat_sum = 0.0
    for _, row in df_mat.iterrows():
        v = float(row["GWP"])
        if v <= 0:
            continue
        src = idx.get(f"Material: {row['Component']}")
        tgt = idx.get("Materials supply")
        if src is None or tgt is None:
            continue
        sources.append(src)
        targets.append(tgt)
        values.append(v)
        mat_sum += v

    elec_sum = 0.0
    if not df_elec_steps.empty:
        for _, row in df_elec_steps.iterrows():
            v = float(row["GWP"])
            if v <= 0:
                continue
            src = idx.get(f"Process: {row['Step']}")
            tgt = idx.get("Electricity supply")
            if src is None or tgt is None:
                continue
            sources.append(src)
            targets.append(tgt)
            values.append(v)
            elec_sum += v

    if transport_gwp > 0:
        sources.append(idx["Transport"])
        targets.append(idx[f"{bead_name} synthesis"])
        values.append(float(transport_gwp))

    if mat_sum > 0:
        sources.append(idx["Materials supply"])
        targets.append(idx[f"{bead_name} synthesis"])
        values.append(mat_sum)

    if elec_sum > 0:
        sources.append(idx["Electricity supply"])
        targets.append(idx[f"{bead_name} synthesis"])
        values.append(elec_sum)

    to_total = mat_sum + elec_sum + float(transport_gwp)
    if to_total <= 0:
        to_total = total_gwp

    sources.append(idx[f"{bead_name} synthesis"])
    targets.append(idx["Total GWP"])
    values.append(to_total)
        # ---------------------------------------------------------------------
    # Add numeric values to node labels (GWP contributions, kg CO2-eq per kg bead)
    # ---------------------------------------------------------------------
    node_value: Dict[str, float] = {}

    # Material nodes
    for _, row in df_mat.iterrows():
        lbl = f"Material: {row['Component']}"
        node_value[lbl] = node_value.get(lbl, 0.0) + float(row["GWP"])

    # Process-step nodes (electricity split)
    if not df_elec_steps.empty:
        for _, row in df_elec_steps.iterrows():
            lbl = f"Process: {row['Step']}"
            node_value[lbl] = node_value.get(lbl, 0.0) + float(row["GWP"])

    # Aggregate nodes
    node_value["Materials supply"] = float(mat_sum)
    node_value["Electricity supply"] = float(elec_sum)
    if transport_gwp > 0:
        node_value["Transport"] = float(transport_gwp)

    node_value[f"{bead_name} synthesis"] = float(to_total)
    node_value["Total GWP"] = float(to_total)

    def _fmt(v: float) -> str:
        v = float(v)
        if abs(v) >= 10:
            return f"{v:.1f}"
        if abs(v) >= 1:
            return f"{v:.2f}"
        return f"{v:.3f}"

    node_labels_display = [
        f"{lbl}: {_fmt(node_value.get(lbl, 0.0))}" for lbl in node_labels
    ]


        fig = go.Figure(
        data=[
            go.Sankey(
                textfont=dict(size=14, color="black"),  # <-- key: forces solid black text
                node=dict(
                    pad=18,
                    thickness=22,
                    line=dict(color="black", width=0.8),
                    label=node_labels_display,          # <-- use labels with numbers
                    color=node_colors,
                ),
                link=dict(source=sources, target=targets, value=values),
            )
        ]
    )

    fig.update_layout(
        title_text=f"Sankey (materials + processes): {bead_name}",
        template="plotly_white",          # <-- lock to white theme
        paper_bgcolor="white",            # <-- avoid dark-mode styling bleed-through
        font=dict(size=14, color="black"),
        height=520,
        margin=dict(l=10, r=10, t=60, b=10),
    )

    return fig


# =============================================================================
# SYSTEM BOUNDARY GRAPHVIZ
# =============================================================================
def render_system_boundary_graphviz() -> None:
    dot = """
    digraph {
        rankdir=LR;
        bgcolor="transparent";
        node [shape=box, style="filled,rounded", fontname="Sans-Serif", fontsize=10];
        edge [fontname="Sans-Serif", fontsize=9, color="#666666"];

        subgraph cluster_upstream {
            label = "Upstream (Excluded)";
            style = "dashed";
            color = "#808080";
            fontcolor = "#808080";
            node [fillcolor="#f9f9f9", color="#808080", fontcolor="#808080"];
            Raw [label="Raw materials\\n(mining, chemicals, waste feedstocks)"];
            Trans [label="Transport\\n(to laboratory)"];
        }

        subgraph cluster_gate {
            label = "System Boundary (Included: Gate-to-Gate)";
            style = "solid";
            color = "#2E8B57";
            penwidth = 2;
            fontcolor = "#2E8B57";
            node [fillcolor="#E8F5E9", color="#2E8B57", fontcolor="black"];

            Inputs [label="Inputs:\\nElectricity, water,\\nreagents, solvents"];

            subgraph cluster_ref {
                label = "Ref-Bead process";
                style = "dotted";
                color = "#2E8B57";

                MF   [label="Microfluidizer\\n(PDChNF dispersion)"];
                Mix  [label="Mixing at 50C\\n(dope preparation)"];
                Cent [label="Centrifugation\\n(degassing)"];
                Pump [label="Syringe pump\\n(extrusion)"];
                Coag [label="Coagulation bath\\nstirring"];
                Cross [label="Crosslinking\\n& washing"];
                FD1  [label="Freeze Drying (1)"];
            }

            subgraph cluster_mof {
                label = "MOF add-on process (U@Bead)";
                style = "dotted";
                color = "#2E8B57";

                ZrStep     [label="Zr precursor step\\n(stirring at 50C)"];
                LinkerStep [label="Linker step (2-ATA)\\n(stirring at 50C)"];
                FD2        [label="Freeze Drying (2)"];
                SolvRec    [label="Solvent recovery\\n(ethanol + formic)"];
            }

            RefProduct [label="Ref-Bead (dry)"];
            FinalProduct [label="Final dry bead\\n(lab gate)"];
        }

        subgraph cluster_downstream {
            label = "Downstream (Excluded)";
            style = "dashed";
            color = "#808080";
            fontcolor = "#808080";
            node [fillcolor="#f9f9f9", color="#808080", fontcolor="#808080"];
            Use [label="Use phase\\n(copper removal)"];
            EOL [label="End of life\\n(regeneration/disposal)"];
        }

        Raw -> Trans;
        Trans -> Inputs;

        Inputs -> MF;
        MF -> Mix;
        Mix -> Cent;
        Cent -> Pump;
        Pump -> Coag;
        Coag -> Cross;
        Cross -> FD1;

        FD1 -> RefProduct [label="Ref-Bead", style="dashed"];
        RefProduct -> FinalProduct [label="Ref-Bead", style="dashed"];

        Inputs -> ZrStep;
        RefProduct -> ZrStep [label="Support"];

        ZrStep -> LinkerStep;
        LinkerStep -> FD2;
        FD2 -> SolvRec;
        SolvRec -> FinalProduct [label="U@Bead"];

        FinalProduct -> Use;
        Use -> EOL;
    }
    """
    try:
        st.graphviz_chart(dot, use_container_width=True)
        st.caption("Gate-to-gate system boundary with explicit unit operations used in the electricity inventory")
    except Exception as e:
        st.error(f"Graphviz rendering failed: {e}")


# =============================================================================
# AI HELPER
# =============================================================================
def get_ai_insight(context_data: List[dict], user_question: str, model_name: str) -> str:
    if not _OPENAI_AVAILABLE:
        return "Error: OpenAI library not installed. Add `openai` to requirements."

    api_key = st.secrets.get("openai_api_key2")
    if not api_key:
        return "Error: API Key 'openai_api_key2' not found in Streamlit secrets."

    client = OpenAI(api_key=api_key)

    context_str = "Current LCA Results:\n"
    for res in context_data:
        context_str += (
            f"- {res['name']}: "
            f"Total GWP={res['Total GWP']:.3e}, "
            f"Elec GWP={res['Electricity GWP']:.3e}, "
            f"Elec%={res.get('Electricity %', 0.0):.1f}%\n"
        )

    system_prompt = f"""
You are an expert in Life Cycle Assessment (LCA) for lab-to-pilot scaling of materials synthesis.
Use the provided data to answer the user's question.

Context Data:
{context_str}

Guidelines:
- Be concise and scientific.
- Explain why impacts are high at lab scale (low utilisation, high fixed duty cycles, solvents).
- Link hotspots to specific unit operations and chemicals.
- Provide actionable optimisation suggestions (solvent recovery, batching, heat integration, better utilisation, grid mix).
""".strip()

    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_question.strip()},
            ],
            temperature=0.3,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI Error: {e}"


# =============================================================================
# SESSION STATE INITIALISATION
# =============================================================================
def reset_to_defaults_auto() -> None:
    """
    Loads defaults from Excel if available (no upload), otherwise embedded defaults.
    """
    excel_path = _find_excel_path()
    used_excel = False

    if excel_path is not None:
        loaded = load_tables_from_excel_v11(excel_path)
        if loaded is not None:
            ef_table_df, lab_table_df, recipe_table_df, scale_table_df, perf_df, lit_df = loaded
            used_excel = True
        else:
            ef_table_df, lab_table_df, recipe_table_df, scale_table_df, perf_df, lit_df = embedded_defaults_tables()
    else:
        ef_table_df, lab_table_df, recipe_table_df, scale_table_df, perf_df, lit_df = embedded_defaults_tables()

    st.session_state["used_excel_defaults"] = bool(used_excel)
    st.session_state["excel_path_found"] = str(excel_path) if excel_path else ""

    st.session_state["ef_table_df"] = ef_table_df
    st.session_state["lab_table_df"] = lab_table_df
    st.session_state["recipe_table_df"] = recipe_table_df
    st.session_state["scale_table_df"] = scale_table_df

    st.session_state["perf_df"] = perf_df
    st.session_state["lit_df"] = lit_df

    st.session_state["custom_grids"] = dict(DEFAULT_CUSTOM_GRIDS)


def ensure_session_state() -> None:
    if "ef_table_df" not in st.session_state:
        reset_to_defaults_auto()


# =============================================================================
# MAIN APP
# =============================================================================
def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    ensure_session_state()

    render_header_logos()

    st.title(APP_TITLE)
    st.markdown("Compare Ref-Bead (polymer) vs U@Bead (MOF-functionalised) .")

    # Sidebar
    with st.sidebar:
        st.header("Control Panel")

        if st.session_state.get("excel_path_found"):
            if st.session_state.get("used_excel_defaults"):
                st.success(f"Loaded defaults from Excel: {st.session_state.get('excel_path_found')}")
            else:
                st.warning(f"Excel found but parsing failed; using embedded defaults. File: {st.session_state.get('excel_path_found')}")
        else:
            st.info("No Excel file found. Using embedded defaults.")

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("Reset (auto-load Excel if found)", key="btn_reset_auto"):
                reset_to_defaults_auto()
                st.rerun()
        with col_r2:
            if st.button("Reset to embedded defaults", key="btn_reset_embedded"):
                ef_table_df, lab_table_df, recipe_table_df, scale_table_df, perf_df, lit_df = embedded_defaults_tables()
                st.session_state["used_excel_defaults"] = False
                st.session_state["excel_path_found"] = ""
                st.session_state["ef_table_df"] = ef_table_df
                st.session_state["lab_table_df"] = lab_table_df
                st.session_state["recipe_table_df"] = recipe_table_df
                st.session_state["scale_table_df"] = scale_table_df
                st.session_state["perf_df"] = perf_df
                st.session_state["lit_df"] = lit_df
                st.session_state["custom_grids"] = dict(DEFAULT_CUSTOM_GRIDS)
                st.rerun()

        st.divider()
        st.subheader("Scenario parameters")

        # grid EF slider driven by "Electricity (Canada)" row
        ef_table_df = st.session_state["ef_table_df"].copy()
        elec_mask = ef_table_df["reagent_name"].astype(str).str.strip().eq("Electricity (Canada)")
        if elec_mask.any():
            elec_idx = ef_table_df[elec_mask].index[0]
            current_grid_ef = float(pd.to_numeric(ef_table_df.loc[elec_idx, "Value"], errors="coerce"))
        else:
            current_grid_ef = 0.0737

        new_grid_ef = st.slider(
            "Grid carbon intensity (kg CO2/kWh)",
            min_value=0.0,
            max_value=1.0,
            value=float(current_grid_ef),
            step=0.01,
        )
        if float(new_grid_ef) != float(current_grid_ef) and elec_mask.any():
            st.session_state["ef_table_df"].loc[elec_idx, "Value"] = float(new_grid_ef)

        eff_factor = st.slider(
            "Efficiency multiplier",
            0.1,
            1.0,
            1.0,
            help="Lower values approximate better industrial energy efficiency.",
        )

        # Default solvent recovery from scale table if available (R_solv_recovery is fraction)
        scale_df_sidebar = st.session_state["scale_table_df"]
        default_R = _val_from_table(scale_df_sidebar, "Parameter", "R_solv_recovery", default=0.0)
        default_R = max(0.0, min(default_R, 0.95))
        default_R_pct = int(round(100.0 * default_R))
        recycle_rate = st.slider("Solvent recovery (%)", 0, 95, default_R_pct)

        yield_rate = st.slider("Global yield (%)", 10, 100, 100)
        transport_overhead = st.slider("Transport overhead (%)", 0, 50, 0)

        st.divider()
        st.subheader("Input tables (editable)")

        with st.expander("Table 1.1 Emission factors (EF)", expanded=False):
            st.session_state["ef_table_df"] = st.data_editor(
                st.session_state["ef_table_df"],
                key="ed_ef_table",
                num_rows="dynamic",
                use_container_width=True,
            )

        with st.expander("Table 1.2 Lab data and equipment", expanded=False):
            st.session_state["lab_table_df"] = st.data_editor(
                st.session_state["lab_table_df"],
                key="ed_lab_table",
                num_rows="dynamic",
                use_container_width=True,
            )

        with st.expander("Recipe table", expanded=False):
            st.session_state["recipe_table_df"] = st.data_editor(
                st.session_state["recipe_table_df"],
                key="ed_recipe_table",
                num_rows="dynamic",
                use_container_width=True,
            )

        with st.expander("Scaling parameters (Excel symbols + recovery params)", expanded=False):
            st.session_state["scale_table_df"] = st.data_editor(
                st.session_state["scale_table_df"],
                key="ed_scale_table",
                num_rows="dynamic",
                use_container_width=True,
            )

        with st.expander("Performance data (FU2 capacity)", expanded=False):
            st.session_state["perf_df"] = st.data_editor(
                st.session_state["perf_df"],
                key="ed_perf",
                num_rows="dynamic",
                use_container_width=True,
            )

        with st.expander("Literature data", expanded=False):
            st.session_state["lit_df"] = st.data_editor(
                st.session_state["lit_df"],
                key="ed_lit",
                num_rows="dynamic",
                use_container_width=True,
            )

        st.divider()
        sAI_MODEL_NAME = "gpt-4o-mini"  # or whichever fixed model is intended


    # Build calculation dataframes
    ef_df = build_ef_df_from_table(st.session_state["ef_table_df"])
    lab_table_df = st.session_state["lab_table_df"]
    recipe_table_df = st.session_state["recipe_table_df"]
    scale_table_df = st.session_state["scale_table_df"]
    perf_df = st.session_state["perf_df"]
    lit_df = st.session_state["lit_df"]

    # MOF loading fraction from lab table
    w_mof_loading = _val_from_table(lab_table_df, "Item name", "w_MOF_loading", default=0.13)
    if w_mof_loading > 1.0:
        w_mof_loading = w_mof_loading / 100.0
    w_mof_loading = max(0.0, min(float(w_mof_loading), 0.999999))

    routes_df, polymer_fraction_support = compute_recipe_routes(
        lab_table_df=lab_table_df,
        recipe_table_df=recipe_table_df,
        mof_loading_fraction=w_mof_loading,
    )

    # Solvent recovery energy parameter (kWh/kg recovered)
    E_solv_rec = _val_from_table(scale_table_df, "Parameter", "E_solv_rec_kWh_per_kg", default=0.0)

    # Total base solvent mass per kg bead (ethanol + formic), without recovery applied
    mof_solvent_base_mass = 0.0
    if not routes_df.empty:
        mof_solvent_base_mass = float(
            routes_df[(routes_df["route_id"] == ID_MOF) & (routes_df["reagent_name"].isin(["Ethanol", "Formic Acid"]))]["mass_kg_per_fu"].sum()
        )

    # Scaling engine (baseline and scaled)
    scaling = compute_scaling_from_input_tables(
        lab_table_df=lab_table_df,
        recipe_table_df=recipe_table_df,
        scale_table_df=scale_table_df,
        polymer_fraction_support=polymer_fraction_support,
        solvent_recovery_frac=float(recycle_rate) / 100.0,
        solvent_recovery_energy_kwh_per_kg=float(E_solv_rec),
        mof_solvent_base_mass_kg_per_kgbead=float(mof_solvent_base_mass),
    )

    # Apply electricity kWh/kg to routes
    routes_base = routes_df.copy()
    routes_scaled = routes_df.copy()

    for rid, kwh in scaling.route_elec_kwh_per_kg_baseline.items():
        routes_base.loc[routes_base["route_id"] == rid, "electricity_kwh_per_fu"] = float(kwh)
    for rid, kwh in scaling.route_elec_kwh_per_kg_scaled.items():
        routes_scaled.loc[routes_scaled["route_id"] == rid, "electricity_kwh_per_fu"] = float(kwh)

    unique_routes = [ID_REF, ID_MOF]

    def run_batch(routes_df_use: pd.DataFrame) -> Tuple[List[dict], List[pd.DataFrame]]:
        res_list, dfs_list = [], []
        for rid in unique_routes:
            res, df = calculate_impacts(
                rid,
                ef_df,
                routes_df_use,
                efficiency_factor=eff_factor,
                recycling_rate_pct=recycle_rate,
                yield_rate=yield_rate,
                transport_pct=transport_overhead,
            )
            if res is not None and df is not None:
                res["Electricity %"] = (res["Electricity GWP"] / res["Total GWP"]) * 100.0 if res["Total GWP"] > 0 else 0.0
                res_list.append(res)
                dfs_list.append(df)
        return res_list, dfs_list

    base_results_list, base_dfs_list = run_batch(routes_base)
    scaled_results_list, scaled_dfs_list = run_batch(routes_scaled)

    # Performance map
    perf_map = {}
    if not perf_df.empty and "route_id" in perf_df.columns and "capacity_mg_g" in perf_df.columns:
        for _, row in perf_df.iterrows():
            rid = str(row["route_id"])
            cap = pd.to_numeric(row["capacity_mg_g"], errors="coerce")
            if pd.notna(cap):
                perf_map[rid] = float(cap)

    def make_summary(res_list: List[dict]) -> pd.DataFrame:
        rows = []
        for r in res_list:
            cap = perf_map.get(r["id"], 0.001)
            rows.append(
                {
                    "Bead": r["name"],
                    "Total GWP": r["Total GWP"],
                    "Non-Electric GWP": r["Non-Electric GWP"],
                    "GWP per g Cu": (r["Total GWP"] / cap) if cap else np.nan,
                    "Electricity %": r.get("Electricity %", 0.0),
                }
            )
        return pd.DataFrame(rows)

    base_sum_df = make_summary(base_results_list)
    scaled_sum_df = make_summary(scaled_results_list)

    tab_scale, tab1, tab2, tab3, tab4, tab5 = st.tabs(
        [
            "Scaling (Excel-aligned)",
            "LCA results",
            "Sensitivity and scaling",
            "Inventory and flows",
            "Literature comparison",
            "AI insights",
        ]
    )

    # -------------------------------------------------------------------------
    # TAB: SCALING
    # -------------------------------------------------------------------------
    with tab_scale:
        st.header("Scaling")

        st.markdown(
            """
This tab computes electricity intensities from table inputs.

Baseline scenario:
- Full lab duty cycles are allocated to the small batch output (no utilisation correction).

Scaled scenario:
- Applies utilisation and capacity scaling factors derived from the scaling symbols (SF_MF, SF_mix, SF_FD_ref, etc.).
- Includes solvent recovery electricity for ethanol + formic using the Excel-equivalent formula.
"""
        )

        col_s1, col_s2 = st.columns(2)

        with col_s1:
            st.subheader("Computed scaling factors")
            sf_df = pd.DataFrame([{"Scaling factor": k, "Value": float(v)} for k, v in scaling.scaling_factors.items()])
            st.dataframe(sf_df.style.format({"Value": "{:.6f}"}), use_container_width=True, hide_index=True)
            st.info(f"Polymer support fraction in U@Bead = {scaling.polymer_fraction_support:.3f} (from w_MOF_loading).")

        with col_s2:
            st.subheader("Electricity intensity (kWh/kg)")
            elec_summary = pd.DataFrame(
                {
                    "Scenario": ["Baseline", "Baseline", "Scaled", "Scaled"],
                    "Bead": ["Ref-Bead", "U@Bead", "Ref-Bead", "U@Bead"],
                    "Electricity kWh/kg": [
                        scaling.route_elec_kwh_per_kg_baseline.get(ID_REF, np.nan),
                        scaling.route_elec_kwh_per_kg_baseline.get(ID_MOF, np.nan),
                        scaling.route_elec_kwh_per_kg_scaled.get(ID_REF, np.nan),
                        scaling.route_elec_kwh_per_kg_scaled.get(ID_MOF, np.nan),
                    ],
                }
            )
            st.dataframe(elec_summary.style.format({"Electricity kWh/kg": "{:.3f}"}), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Step-level electricity breakdown (kWh/kg bead)")

        col_sb1, col_sb2 = st.columns(2)
        with col_sb1:
            st.markdown("**Baseline (lab duty cycles)**")
            st.dataframe(scaling.baseline_step_df.style.format({"kWh_per_kg": "{:.3f}"}), use_container_width=True, hide_index=True)
            fig_steps_base = px.bar(
                scaling.baseline_step_df,
                x="Step",
                y="kWh_per_kg",
                color="Bead",
                barmode="group",
                title="Baseline electricity by step (kWh/kg)",
                text_auto=".2f",
            )
            fig_steps_base.update_layout(xaxis_tickangle=-25)
            fig_steps_base = apply_publication_style(fig_steps_base, height=520)
            st.plotly_chart(fig_steps_base, use_container_width=True, key="scale_steps_base", config=PLOTLY_CONFIG)

        with col_sb2:
            st.markdown("**Scaled (Excel-aligned utilisation)**")
            st.dataframe(scaling.scaled_step_df.style.format({"kWh_per_kg": "{:.3f}"}), use_container_width=True, hide_index=True)
            fig_steps_scaled = px.bar(
                scaling.scaled_step_df,
                x="Step",
                y="kWh_per_kg",
                color="Bead",
                barmode="group",
                title="Scaled electricity by step (kWh/kg)",
                text_auto=".2f",
            )
            fig_steps_scaled.update_layout(xaxis_tickangle=-25)
            fig_steps_scaled = apply_publication_style(fig_steps_scaled, height=520)
            st.plotly_chart(fig_steps_scaled, use_container_width=True, key="scale_steps_scaled", config=PLOTLY_CONFIG)

        st.divider()
        st.subheader("Figure 3: Electricity breakdown by unit operation (utilisation-scaled kWh kg−1)")

        df_fig3 = scaling.scaled_step_df.copy()

        def _step_group(step: str) -> str:
            s = str(step)
            if "Microfluidizer" in s:
                return "Microfluidisation"
            if "Mixing" in s or "Crosslinking" in s:
                return "Mixing and crosslinking"
            if "Freeze" in s:
                return "Freeze-drying"
            if "Zr" in s or "Linker" in s:
                return "MOF-step stirring"
            if "Support Electricity" in s:
                return "Support electricity"
            if "Solvent recovery" in s:
                return "Solvent recovery"
            return "Other unit ops"

        df_fig3["Unit operation"] = df_fig3["Step"].map(_step_group)
        df_fig3_grouped = df_fig3.groupby(["Bead", "Unit operation"], as_index=False)["kWh_per_kg"].sum()

        fig3 = px.bar(
            df_fig3_grouped,
            x="Bead",
            y="kWh_per_kg",
            color="Unit operation",
            barmode="stack",
            title="Figure 3: Unit-operation electricity breakdown (scaled)",
            text_auto=".2f",
        )
        fig3 = apply_publication_style(fig3, height=520)
        st.plotly_chart(fig3, use_container_width=True, key="fig3_unit_ops", config=PLOTLY_CONFIG)
        add_png_download_button(fig3, filename="Figure_3_unit_operation_electricity_scaled", key="fig3_unit_ops")

    # -------------------------------------------------------------------------
    # TAB: RESULTS
    # -------------------------------------------------------------------------
    with tab1:
        st.header("LCA results (GWP)")

        fmt = {"Total GWP": "{:.3f}", "Non-Electric GWP": "{:.3f}", "GWP per g Cu": "{:.3f}", "Electricity %": "{:.1f}"}

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.subheader("Baseline")
            st.dataframe(base_sum_df.style.format(fmt), hide_index=True, use_container_width=True)
            fig_base = px.bar(
                base_sum_df,
                x="Bead",
                y="Total GWP",
                color="Bead",
                title="Total GWP (baseline) (kg CO2-eq per kg bead)",
                text_auto=".2f",
            )
            fig_base = apply_publication_style(fig_base, height=520)
            st.plotly_chart(fig_base, use_container_width=True, key="res_base_total", config=PLOTLY_CONFIG)

        with col_r2:
            st.subheader("Scaled")
            st.dataframe(scaled_sum_df.style.format(fmt), hide_index=True, use_container_width=True)
            fig_scaled = px.bar(
                scaled_sum_df,
                x="Bead",
                y="Total GWP",
                color="Bead",
                title="Total GWP (scaled) (kg CO2-eq per kg bead)",
                text_auto=".2f",
            )
            fig_scaled = apply_publication_style(fig_scaled, height=520)
            st.plotly_chart(fig_scaled, use_container_width=True, key="res_scaled_total", config=PLOTLY_CONFIG)

        st.divider()
        st.subheader("Performance normalised (FU2)")
        st.markdown("Impact per gram of copper removed (using adsorption capacity values from the Excel calculations sheet).")

        col_fu1, col_fu2 = st.columns(2)
        with col_fu1:
            fig_fu2_base = px.bar(
                base_sum_df,
                x="Bead",
                y="GWP per g Cu",
                color="Bead",
                title="GWP per g Cu removed (baseline)",
                text_auto=".3f",
            )
            fig_fu2_base = apply_publication_style(fig_fu2_base, height=520)
            st.plotly_chart(fig_fu2_base, use_container_width=True, key="fu2_base", config=PLOTLY_CONFIG)

        with col_fu2:
            st.markdown("**Figure 1: Utilisation-scaled GWP100 per g Cu removed (FU2)**")
            fig1 = px.bar(
                scaled_sum_df,
                x="Bead",
                y="GWP per g Cu",
                color="Bead",
                title="Figure 1: Utilisation-scaled GWP100 per g Cu removed (FU2)",
                text_auto=".3f",
            )
            fig1 = apply_publication_style(fig1, height=520)
            st.plotly_chart(fig1, use_container_width=True, key="fig1_fu2_scaled", config=PLOTLY_CONFIG)
            add_png_download_button(fig1, filename="Figure_1_FU2_GWP_per_g_Cu_scaled", key="fig1_fu2_scaled")

        st.divider()
        st.subheader("Figure 2: Stacked FU1 split (electricity, polymer support, MOF growth inputs)")

        # Build scaled contribution table for Figure 2
        all_contribs_scaled = []
        for i, df in enumerate(scaled_dfs_list):
            d = df.copy()
            d["Bead"] = scaled_results_list[i]["name"]
            all_contribs_scaled.append(d)
        df_all_scaled = pd.concat(all_contribs_scaled) if all_contribs_scaled else pd.DataFrame()

        polymer_components = {"Chitosan", "PDChNF", "Acetic Acid"}
        mof_components = {"Ethanol", "Formic Acid", "ZrCl4", "2-ATA"}

        fig2_rows = []
        if not df_all_scaled.empty:
            for bead in df_all_scaled["Bead"].unique():
                dfb = df_all_scaled[df_all_scaled["Bead"] == bead].copy()
                gwp_elec = float(dfb.loc[dfb["Component"] == "Electricity", "GWP"].sum())
                gwp_poly = float(dfb.loc[dfb["Component"].isin(polymer_components), "GWP"].sum())
                gwp_mof = float(dfb.loc[dfb["Component"].isin(mof_components), "GWP"].sum())
                gwp_transport = float(dfb.loc[dfb["Component"] == "Transport", "GWP"].sum())

                raw_total = gwp_elec + gwp_poly + gwp_mof
                if raw_total > 0 and gwp_transport > 0:
                    gwp_elec += gwp_transport * (gwp_elec / raw_total)
                    gwp_poly += gwp_transport * (gwp_poly / raw_total)
                    gwp_mof += gwp_transport * (gwp_mof / raw_total)

                fig2_rows.extend(
                    [
                        {"Bead": bead, "Split": "Electricity", "GWP": gwp_elec},
                        {"Bead": bead, "Split": "Polymer support", "GWP": gwp_poly},
                        {"Bead": bead, "Split": "MOF growth inputs", "GWP": gwp_mof},
                    ]
                )

        df_fig2 = pd.DataFrame(fig2_rows)

        if not df_fig2.empty:
            fig2 = px.bar(
                df_fig2,
                x="Bead",
                y="GWP",
                color="Split",
                barmode="stack",
                title="Figure 2: Utilisation-scaled GWP100 per kg bead (FU1), split by contribution class",
                text_auto=".2f",
            )
            fig2 = apply_publication_style(fig2, height=520)
            st.plotly_chart(fig2, use_container_width=True, key="fig2_fu1_stacked_scaled", config=PLOTLY_CONFIG)
            add_png_download_button(fig2, filename="Figure_2_FU1_stacked_split_scaled", key="fig2_fu1_stacked_scaled")
            st.caption("Transport overhead (if non-zero) is allocated proportionally across the three contribution classes.")
        else:
            st.warning("Figure 2 could not be generated (scaled inventory table is empty).")

        st.divider()
        st.subheader("System boundary")
        render_system_boundary_graphviz()

    # -------------------------------------------------------------------------
    # TAB: SENSITIVITY AND SCALING
    # -------------------------------------------------------------------------
    with tab2:
        st.header("Sensitivity and scaling")

        st.subheader("1) Grid intensity sensitivity (custom grids)")
        sens_rows_base = []
        sens_rows_scaled = []

        for g_name, g_val in st.session_state["custom_grids"].items():
            temp_ef = ef_df.copy()
            temp_ef.loc[temp_ef["reagent_name"].str.contains("Electricity", case=False, na=False), "GWP_kgCO2_per_kg"] = float(g_val)

            for rid in unique_routes:
                res_b, _ = calculate_impacts(
                    rid,
                    temp_ef,
                    routes_base,
                    efficiency_factor=eff_factor,
                    recycling_rate_pct=recycle_rate,
                    yield_rate=yield_rate,
                    transport_pct=transport_overhead,
                )
                res_s, _ = calculate_impacts(
                    rid,
                    temp_ef,
                    routes_scaled,
                    efficiency_factor=eff_factor,
                    recycling_rate_pct=recycle_rate,
                    yield_rate=yield_rate,
                    transport_pct=transport_overhead,
                )
                if res_b:
                    sens_rows_base.append({"Grid": g_name, "Grid Value": g_val, "Bead": res_b["name"], "Total GWP": res_b["Total GWP"]})
                if res_s:
                    sens_rows_scaled.append({"Grid": g_name, "Grid Value": g_val, "Bead": res_s["name"], "Total GWP": res_s["Total GWP"]})

        df_sens_base = pd.DataFrame(sens_rows_base).sort_values("Grid Value")
        df_sens_scaled = pd.DataFrame(sens_rows_scaled).sort_values("Grid Value")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("**Baseline**")
            fig_sens_base = px.line(
                df_sens_base,
                x="Grid",
                y="Total GWP",
                color="Bead",
                markers=True,
                title="Total GWP vs grid intensity (baseline)",
                hover_data=["Grid Value"],
            )
            fig_sens_base = apply_publication_style(fig_sens_base, height=520)
            st.plotly_chart(fig_sens_base, use_container_width=True, key="sens_grid_base", config=PLOTLY_CONFIG)

        with col_g2:
            st.markdown("**Scaled**")
            fig_sens_scaled = px.line(
                df_sens_scaled,
                x="Grid",
                y="Total GWP",
                color="Bead",
                markers=True,
                title="Total GWP vs grid intensity (scaled)",
                hover_data=["Grid Value"],
            )
            fig_sens_scaled = apply_publication_style(fig_sens_scaled, height=520)
            st.plotly_chart(fig_sens_scaled, use_container_width=True, key="sens_grid_scaled", config=PLOTLY_CONFIG)

        st.divider()
        st.subheader("2) Batch size projection (electricity-driven illustration)")

        LAB_BATCH_REF_KG = max(_val_from_table(lab_table_df, "Item name", "m_bead_batch_Ref", default=0.4) / 1000.0, 1e-12)
        LAB_BATCH_MOF_KG = max(_val_from_table(lab_table_df, "Item name", "m_bead_batch_MOF", default=0.6) / 1000.0, 1e-12)

        batch_sizes = np.logspace(np.log10(min(LAB_BATCH_REF_KG, LAB_BATCH_MOF_KG)), np.log10(100.0), num=40)

        scale_rows_base = []
        scale_rows_scaled = []

        for rid in unique_routes:
            base_res, _ = calculate_impacts(
                rid, ef_df, routes_base, efficiency_factor=1.0, recycling_rate_pct=recycle_rate, yield_rate=100.0, transport_pct=transport_overhead
            )
            scaled_res, _ = calculate_impacts(
                rid, ef_df, routes_scaled, efficiency_factor=1.0, recycling_rate_pct=recycle_rate, yield_rate=100.0, transport_pct=transport_overhead
            )
            if base_res is None or scaled_res is None:
                continue

            base_elec_intensity = float(base_res["Electricity kWh"])
            scaled_elec_intensity = float(scaled_res["Electricity kWh"])

            lab_batch = LAB_BATCH_REF_KG if rid == ID_REF else LAB_BATCH_MOF_KG

            for b in batch_sizes:
                sf = float(lab_batch) / float(b)

                new_elec_b = base_elec_intensity * sf
                gwp_b = new_elec_b * float(base_res["Electricity EF Used"]) + float(base_res["Non-Electric GWP"])
                scale_rows_base.append({"Batch size (kg)": b, "Bead": base_res["name"], "Projected GWP": gwp_b})

                new_elec_s = scaled_elec_intensity * sf
                gwp_s = new_elec_s * float(scaled_res["Electricity EF Used"]) + float(scaled_res["Non-Electric GWP"])
                scale_rows_scaled.append({"Batch size (kg)": b, "Bead": scaled_res["name"], "Projected GWP": gwp_s})

        df_proj_base = pd.DataFrame(scale_rows_base)
        df_proj_scaled = pd.DataFrame(scale_rows_scaled)

        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.markdown("**Projection using baseline electricity**")
            fig_proj_base = px.line(
                df_proj_base,
                x="Batch size (kg)",
                y="Projected GWP",
                color="Bead",
                log_x=True,
                log_y=True,
                markers=True,
                title="Projected GWP vs batch size (baseline, log-log)",
            )
            fig_proj_base = apply_publication_style(fig_proj_base, height=520)
            st.plotly_chart(fig_proj_base, use_container_width=True, key="proj_base", config=PLOTLY_CONFIG)

        with col_p2:
            st.markdown("**Projection using scaled electricity**")
            fig_proj_scaled = px.line(
                df_proj_scaled,
                x="Batch size (kg)",
                y="Projected GWP",
                color="Bead",
                log_x=True,
                log_y=True,
                markers=True,
                title="Projected GWP vs batch size (scaled, log-log)",
            )
            fig_proj_scaled = apply_publication_style(fig_proj_scaled, height=520)
            st.plotly_chart(fig_proj_scaled, use_container_width=True, key="proj_scaled", config=PLOTLY_CONFIG)

    # -------------------------------------------------------------------------
    # TAB: INVENTORY AND FLOWS
    # -------------------------------------------------------------------------
    with tab3:
        st.header("Inventory and flows")

        all_contribs_base = []
        for i, df in enumerate(base_dfs_list):
            d = df.copy()
            d["Bead"] = base_results_list[i]["name"]
            all_contribs_base.append(d)
        df_all_base = pd.concat(all_contribs_base) if all_contribs_base else pd.DataFrame()

        all_contribs_scaled2 = []
        for i, df in enumerate(scaled_dfs_list):
            d = df.copy()
            d["Bead"] = scaled_results_list[i]["name"]
            all_contribs_scaled2.append(d)
        df_all_scaled2 = pd.concat(all_contribs_scaled2) if all_contribs_scaled2 else pd.DataFrame()

        if not df_all_base.empty and not df_all_scaled2.empty:
            st.subheader("A) Total impact breakdown")
            col_b1, col_b2 = st.columns(2)

            with col_b1:
                st.markdown("**Baseline**")
                fig_break_base = px.bar(
                    df_all_base,
                    x="Bead",
                    y="GWP",
                    color="Component",
                    barmode="group",
                    title="GWP breakdown (baseline)",
                )
                fig_break_base = apply_publication_style(fig_break_base, height=520)
                st.plotly_chart(fig_break_base, use_container_width=True, key="inv_break_base", config=PLOTLY_CONFIG)

            with col_b2:
                st.markdown("**Scaled**")
                fig_break_scaled = px.bar(
                    df_all_scaled2,
                    x="Bead",
                    y="GWP",
                    color="Component",
                    barmode="group",
                    title="GWP breakdown (scaled)",
                )
                fig_break_scaled = apply_publication_style(fig_break_scaled, height=520)
                st.plotly_chart(fig_break_scaled, use_container_width=True, key="inv_break_scaled", config=PLOTLY_CONFIG)

            st.divider()
            st.subheader("B) Mass inventory per kg bead (scaled)")
            df_mass_scaled = df_all_scaled2[df_all_scaled2["Category"] != "Electricity"].copy()
            fig_mass = px.bar(
                df_mass_scaled,
                x="Component",
                y="Mass (kg)",
                color="Component",
                facet_col="Bead",
                title="Mass input per kg bead (scaled)",
            )
            fig_mass.update_yaxes(matches=None, showticklabels=True)
            fig_mass = apply_publication_style(fig_mass, height=520)
            st.plotly_chart(fig_mass, use_container_width=True, key="inv_mass_scaled", config=PLOTLY_CONFIG)

            st.divider()
            st.subheader("C) Sankey diagrams (materials + processes)")

            col_sk1, col_sk2 = st.columns(2)
            with col_sk1:
                st.markdown("**Ref-Bead (scaled)**")
                fig_sank_ref_scaled = plot_sankey_materials_processes(
                    scaled_results_list,
                    contrib_df_all=df_all_scaled2,
                    step_df=scaling.scaled_step_df,
                    route_id=ID_REF,
                )
                st.plotly_chart(fig_sank_ref_scaled, use_container_width=True, key="sankey_ref_scaled", config=PLOTLY_CONFIG)

            with col_sk2:
                st.markdown("**U@Bead (scaled)**")
                fig_sank_mof_scaled = plot_sankey_materials_processes(
                    scaled_results_list,
                    contrib_df_all=df_all_scaled2,
                    step_df=scaling.scaled_step_df,
                    route_id=ID_MOF,
                )
                st.plotly_chart(fig_sank_mof_scaled, use_container_width=True, key="sankey_mof_scaled", config=PLOTLY_CONFIG)

            st.divider()
            st.subheader("D) Sankey diagrams (baseline)")

            col_sk3, col_sk4 = st.columns(2)
            with col_sk3:
                st.markdown("**Ref-Bead (baseline)**")
                fig_sank_ref_base = plot_sankey_materials_processes(
                    base_results_list,
                    contrib_df_all=df_all_base,
                    step_df=scaling.baseline_step_df,
                    route_id=ID_REF,
                )
                st.plotly_chart(fig_sank_ref_base, use_container_width=True, key="sankey_ref_base", config=PLOTLY_CONFIG)

            with col_sk4:
                st.markdown("**U@Bead (baseline)**")
                fig_sank_mof_base = plot_sankey_materials_processes(
                    base_results_list,
                    contrib_df_all=df_all_base,
                    step_df=scaling.baseline_step_df,
                    route_id=ID_MOF,
                )
                st.plotly_chart(fig_sank_mof_base, use_container_width=True, key="sankey_mof_base", config=PLOTLY_CONFIG)

        else:
            st.warning("Inventory tables are empty. Check your routes and EF tables.")

    # -------------------------------------------------------------------------
    # TAB: LITERATURE
    # -------------------------------------------------------------------------
    with tab4:
        st.header("Literature comparison")

        current_data_scaled = []
        for r in scaled_results_list:
            current_data_scaled.append(
                {
                    "Material": f"{r['name']} (scaled, this work)",
                    "GWP_kgCO2_per_kg": r["Total GWP"],
                    "Source": "This work (scaled)",
                    "Type": "This work",
                }
            )

        lit_combined_scaled = pd.concat([lit_df, pd.DataFrame(current_data_scaled)], ignore_index=True) if not lit_df.empty else pd.DataFrame(current_data_scaled)

        st.subheader("Figure 4: Benchmark comparison (utilisation-scaled FU1 vs literature)")

        def _classify_scale_form(material: str) -> str:
            m = str(material).lower()
            if "bead" in m or "u@bead" in m or "ref-bead" in m:
                return "Structured bead-form composite"
            if "activated carbon" in m or "biochar" in m or "commercial" in m:
                return "Industrial-scale powder"
            if "uio" in m or "mof" in m:
                return "MOF powder (literature)"
            return "Other"

        df_fig4 = lit_combined_scaled.copy()
        df_fig4["Scale/Form"] = df_fig4["Material"].map(_classify_scale_form)

        fig4 = px.bar(
            df_fig4,
            x="Material",
            y="GWP_kgCO2_per_kg",
            color="Scale/Form",
            title="Figure 4: Benchmark comparison of FU1 (scaled) against selected literature values",
            hover_data=["Source", "Type", "Scale/Form"],
            log_y=False,
        )
        fig4.update_layout(xaxis_tickangle=-45)
        fig4 = apply_publication_style(fig4, height=560)
        st.plotly_chart(fig4, use_container_width=True, key="fig4_benchmark_scaled", config=PLOTLY_CONFIG)
        add_png_download_button(fig4, filename="Figure_4_benchmark_comparison_scaled", key="fig4_benchmark_scaled")

        st.divider()
        col_l1, col_l2 = st.columns(2)

        with col_l1:
            st.subheader("Baseline")
            current_data_base = []
            for r in base_results_list:
                current_data_base.append(
                    {
                        "Material": f"{r['name']} (baseline, this work)",
                        "GWP_kgCO2_per_kg": r["Total GWP"],
                        "Source": "This work (baseline)",
                        "Type": "This work",
                    }
                )
            lit_combined_base = pd.concat([lit_df, pd.DataFrame(current_data_base)], ignore_index=True) if not lit_df.empty else pd.DataFrame(current_data_base)
            fig_lit_base = px.bar(
                lit_combined_base,
                x="Material",
                y="GWP_kgCO2_per_kg",
                color="Source",
                title="GWP comparison with literature (baseline)",
                text="Source",
            )
            fig_lit_base.update_layout(xaxis_tickangle=-45)
            fig_lit_base = apply_publication_style(fig_lit_base, height=560)
            st.plotly_chart(fig_lit_base, use_container_width=True, key="lit_base", config=PLOTLY_CONFIG)

        with col_l2:
            st.subheader("Scaled")
            fig_lit_scaled = px.bar(
                lit_combined_scaled,
                x="Material",
                y="GWP_kgCO2_per_kg",
                color="Source",
                title="GWP comparison with literature (scaled)",
                text="Source",
            )
            fig_lit_scaled.update_layout(xaxis_tickangle=-45)
            fig_lit_scaled = apply_publication_style(fig_lit_scaled, height=560)
            st.plotly_chart(fig_lit_scaled, use_container_width=True, key="lit_scaled", config=PLOTLY_CONFIG)

    # -------------------------------------------------------------------------
    # TAB: AI INSIGHTS
    # -------------------------------------------------------------------------
    with tab5:
        st.header("AI insights")
        st.caption("AI interprets the current scaled results. Requires OpenAI SDK and the secret key `openai_api_key2`.")

        route_name_map = {r["id"]: r["name"] for r in scaled_results_list}
        focus_options = ["All routes"] + [route_name_map.get(rid, rid) for rid in unique_routes]

        focus_choice = st.radio("Route focus (optional)", options=focus_options, index=0, key="ai_focus")

        if focus_choice == "All routes":
            ai_context_results = scaled_results_list
        else:
            chosen_id = None
            for rid, nm in route_name_map.items():
                if nm == focus_choice:
                    chosen_id = rid
                    break
            ai_context_results = [r for r in scaled_results_list if r["id"] == chosen_id] if chosen_id else scaled_results_list

        sample_questions = [
            "What dominates the total GWP and why?",
            "How do the scaling factors reduce electricity impacts?",
            "Which step is the biggest electricity hotspot in the scaled scenario?",
            "What are the most effective levers to reduce impacts (solvents vs electricity)?",
        ]

        col_q1, col_q2 = st.columns(2)
        if "ai_custom_q" not in st.session_state:
            st.session_state["ai_custom_q"] = ""

        for i, q in enumerate(sample_questions):
            with (col_q1 if i % 2 == 0 else col_q2):
                if st.button(q, key=f"btn_ai_sample_{i}"):
                    st.session_state["ai_custom_q"] = q

        user_q = st.text_area("Ask a question:", key="ai_custom_q", height=140)

        if st.button("Analyse with AI", key="btn_ai_run"):
            if user_q.strip():
                with st.spinner("Analysing..."):
                    ans = get_ai_insight(ai_context_results, user_q, model_name=ai_model)
                st.markdown("### AI analysis")
                st.info(ans)
            else:
                st.warning("Please enter a question or click a sample prompt.")


if __name__ == "__main__":
    main()






