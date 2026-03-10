import streamlit as st
import pandas as pd
import numpy as np
import json
from io import BytesIO
from pathlib import Path
from datetime import date

st.title("Generar Reporte Consolidado")
st.markdown("Sube el Excel rellenado por los operarios y genera el reporte de KPIs.")

from metodologia import render_download as _render_metodologia
_render_metodologia("reporte_consolidado")


def find_col(columns, candidates):
    cols_lower = {str(c).lower().strip(): c for c in columns}
    for cand in candidates:
        cand_lower = cand.lower().strip()
        if cand_lower in cols_lower:
            return cols_lower[cand_lower]
        for k, v in cols_lower.items():
            if cand_lower in k or k in cand_lower:
                return v
    return None


# =============================================================================
# PERSISTENT STORAGE
# =============================================================================
_DATA_DIR = Path(__file__).resolve().parent.parent / "data"
CENTROS_FILE = _DATA_DIR / "centros_trabajo.json"
CONSOL_HISTORY_FILE = _DATA_DIR / "consolidado_historico.json"


def load_centros_trabajo():
    if CENTROS_FILE.exists():
        with open(CENTROS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def load_consol_history():
    if CONSOL_HISTORY_FILE.exists():
        with open(CONSOL_HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_consol_history(entries):
    CONSOL_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CONSOL_HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


# =============================================================================
# 1. CENTER SELECTOR + FILE UPLOADS — persisted in session_state
# =============================================================================
SS = st.session_state

centros = load_centros_trabajo()
if centros:
    centro_sel = st.selectbox(
        "Centro de trabajo",
        options=centros,
        key="rpt_centro",
        help="Selecciona el centro al que corresponde esta auditoría",
    )
else:
    st.info("No hay centros registrados. Regístralos en la pestaña de Absentismo.")
    centro_sel = None

st.sidebar.header("Archivos — Reporte")
audit_up = st.sidebar.file_uploader("Excel de auditoría rellenado", type=["xlsx", "xls"], key="rpt_audit_up")
stock_up = st.sidebar.file_uploader("Stock original (opcional, para % cobertura)", type=["xlsx", "xls"], key="rpt_stock_up")
values_up = st.sidebar.file_uploader("Excel de Valores Unitarios (opcional, para pérdidas)", type=["xlsx", "xls"], key="rpt_values_up")

if audit_up is not None:
    SS["rpt_audit_bytes"] = audit_up.getvalue()
if stock_up is not None:
    SS["rpt_stock_bytes"] = stock_up.getvalue()
if values_up is not None:
    SS["rpt_values_bytes"] = values_up.getvalue()

if "rpt_audit_bytes" in SS:
    st.sidebar.success("Auditoría cargada")
if "rpt_stock_bytes" in SS:
    st.sidebar.success("Stock original cargado")
if "rpt_values_bytes" in SS:
    st.sidebar.success("Valores cargados")

if "rpt_audit_bytes" not in SS:
    st.info("Sube el Excel de auditoría rellenado en la barra lateral.")
    st.stop()


# =============================================================================
# 2. LOAD
# =============================================================================
@st.cache_data
def load_all_sheets_bytes(data):
    xls = pd.ExcelFile(BytesIO(data))
    sheets = {}
    for name in xls.sheet_names:
        sheets[name] = pd.read_excel(xls, sheet_name=name)
    return sheets


@st.cache_data
def load_excel_bytes(data):
    return pd.read_excel(BytesIO(data))


all_sheets = load_all_sheets_bytes(SS["rpt_audit_bytes"])
st.sidebar.caption(f"Pestañas: {', '.join(all_sheets.keys())}")

df_stock_orig = None
if "rpt_stock_bytes" in SS:
    df_stock_orig = load_excel_bytes(SS["rpt_stock_bytes"])

# External values file (for loss calculation when audit Excel doesn't have unit values)
df_values_ext = None
if "rpt_values_bytes" in SS:
    _xls_v = pd.ExcelFile(BytesIO(SS["rpt_values_bytes"]))
    _frames_v = [pd.read_excel(_xls_v, sheet_name=s) for s in _xls_v.sheet_names]
    df_values_ext = pd.concat(_frames_v, ignore_index=True)


# =============================================================================
# 3. CLASSIFY
# =============================================================================
def classify_sheet(name):
    n = name.lower().strip()
    if "control" in n or "diferenc" in n:
        return "CONTROL DIFERENCIADO"
    elif "valios" in n:
        return "MATERIAL VALIOSO"
    elif "aleatori" in n:
        return "ALEATORIO"
    return None


sheet_map = {}
for name, df in all_sheets.items():
    stype = classify_sheet(name)
    if stype:
        sheet_map[stype] = df

if not sheet_map:
    st.error("No se detectaron pestañas de tipo aleatorio/control/valioso.")
    st.stop()


# =============================================================================
# 4. PROCESS
# =============================================================================
def process_sheet(df, is_control=False, df_values_ext=None):
    col_ubic = find_col(df.columns, ["Ubicacion", "Ubicación"])
    col_stock = find_col(df.columns, ["Stock", "Cantidad"])
    col_fisica = find_col(df.columns, ["Cant. Física", "Cant. Fisica"])
    col_valor_total = find_col(df.columns, ["Valor total", "Valor Total"])
    col_valor_unit = find_col(df.columns, ["Valor unitario", "Valor Unitario", "Precio"])
    col_fallo = find_col(df.columns, ["Fallo en el proceso", "Fallo en proceso"])

    col_obs_inv = find_col(df.columns, ["Obs. Inventario", "Observaciones Inventario", "Observaciones"])
    col_obs_proc = find_col(df.columns, ["Obs. Proceso", "Observaciones Proceso"])

    if not col_ubic or not col_stock:
        return None, None

    df = df.copy()
    df["_stock"] = pd.to_numeric(df[col_stock], errors="coerce").fillna(0)
    df["_fisica"] = pd.to_numeric(df[col_fisica], errors="coerce") if col_fisica else np.nan
    df["_valor_total"] = pd.to_numeric(df[col_valor_total], errors="coerce").fillna(0) if col_valor_total else 0
    df["_valor_unit"] = pd.to_numeric(df[col_valor_unit], errors="coerce").fillna(0) if col_valor_unit else 0

    # Enrich unit values from external values file if available
    if df_values_ext is not None:
        col_mat = find_col(df.columns, ["Ref. Material", "Material"])
        col_lote = find_col(df.columns, ["Nº Lote", "Lote"])
        ext_mat = find_col(df_values_ext.columns, ["Material", "Ref. Material", "Referencia"])
        ext_lote = find_col(df_values_ext.columns, ["Batch", "Lote", "Nº Lote"])
        ext_valor = find_col(df_values_ext.columns, ["Valor unitario", "Valor Unitario", "Unit Value", "Precio"])
        if ext_valor and col_mat:
            need_fill = (df["_valor_unit"] == 0) | df["_valor_unit"].isna()
            if need_fill.any():
                vdf = df_values_ext.copy()
                vdf["_ext_val"] = pd.to_numeric(vdf[ext_valor], errors="coerce")
                vdf = vdf.dropna(subset=["_ext_val"])
                if ext_mat:
                    vdf["_ext_mat"] = vdf[ext_mat].astype(str).str.strip()
                if ext_lote:
                    vdf["_ext_lote"] = vdf[ext_lote].astype(str).str.strip()
                df["_mat_k"] = df[col_mat].astype(str).str.strip() if col_mat else ""
                df["_lote_k"] = df[col_lote].astype(str).str.strip() if col_lote else ""
                # 3-level cascade: mat+lote → lote → mat
                # Use map-based approach to avoid index alignment issues
                if ext_mat and ext_lote and col_lote:
                    vd1 = vdf.drop_duplicates(subset=["_ext_mat", "_ext_lote"], keep="first")
                    lookup1 = dict(zip(
                        vd1["_ext_mat"] + "||" + vd1["_ext_lote"],
                        vd1["_ext_val"]
                    ))
                    keys1 = df["_mat_k"] + "||" + df["_lote_k"]
                    mapped1 = keys1.map(lookup1)
                    fill_mask1 = need_fill & mapped1.notna()
                    df.loc[fill_mask1, "_valor_unit"] = mapped1[fill_mask1]
                    need_fill = (df["_valor_unit"] == 0) | df["_valor_unit"].isna()
                if ext_lote and col_lote and need_fill.any():
                    vd2 = vdf.drop_duplicates(subset=["_ext_lote"], keep="first")
                    lookup2 = dict(zip(vd2["_ext_lote"], vd2["_ext_val"]))
                    mapped2 = df["_lote_k"].map(lookup2)
                    fill_mask2 = need_fill & mapped2.notna()
                    df.loc[fill_mask2, "_valor_unit"] = mapped2[fill_mask2]
                    need_fill = (df["_valor_unit"] == 0) | df["_valor_unit"].isna()
                if ext_mat and need_fill.any():
                    vd3 = vdf.drop_duplicates(subset=["_ext_mat"], keep="first")
                    lookup3 = dict(zip(vd3["_ext_mat"], vd3["_ext_val"]))
                    mapped3 = df["_mat_k"].map(lookup3)
                    fill_mask3 = need_fill & mapped3.notna()
                    df.loc[fill_mask3, "_valor_unit"] = mapped3[fill_mask3]
                df.drop(columns=["_mat_k", "_lote_k"], errors="ignore", inplace=True)

    df["_is_error"] = False
    mask_filled = df["_fisica"].notna()
    if mask_filled.any():
        df.loc[mask_filled, "_is_error"] = df.loc[mask_filled, "_stock"] != df.loc[mask_filled, "_fisica"]

    # Erroneous units: absolute discrepancy for lines with error
    df["_uds_erroneas"] = np.where(mask_filled & df["_is_error"], (df["_fisica"] - df["_stock"]).abs(), 0)

    # Units: stock vs physical (only where physical is filled)
    df["_uds_inventariadas"] = np.where(mask_filled, df["_fisica"], np.nan)
    df["_uds_stock"] = np.where(mask_filled, df["_stock"], np.nan)
    df["_descuadre"] = np.where(mask_filled, df["_fisica"] - df["_stock"], np.nan)
    # Negative discrepancy = loss (physical < stock)
    df["_descuadre_neg"] = np.where(
        mask_filled & (df["_fisica"] < df["_stock"]),
        df["_stock"] - df["_fisica"],  # positive amount of missing units
        0,
    )
    # Monetary loss = missing units × unit value
    df["_perdida_monetaria"] = df["_descuadre_neg"] * df["_valor_unit"]

    # Fallo: ANY non-empty value counts (except "no", "n", empty, nan)
    df["_fallo_proceso"] = False
    if is_control and col_fallo:
        fallo_raw = df[col_fallo].fillna("").astype(str).str.strip()
        fallo_lower = fallo_raw.str.lower()
        df["_fallo_proceso"] = (fallo_raw != "") & (~fallo_lower.isin(["nan", "no", "n"]))

    df["_obs_inv"] = ""
    if col_obs_inv:
        df["_obs_inv"] = df[col_obs_inv].fillna("").astype(str)
    df["_obs_proc"] = ""
    if col_obs_proc:
        df["_obs_proc"] = df[col_obs_proc].fillna("").astype(str)

    agg = {
        "lotes_auditados": ("_stock", "size"),
        "lotes_erroneos": ("_is_error", "sum"),
        "valor_total": ("_valor_total", "sum"),
        "uds_inventariadas": ("_uds_inventariadas", lambda x: x.dropna().sum()),
        "uds_stock": ("_uds_stock", lambda x: x.dropna().sum()),
        "uds_erroneas": ("_uds_erroneas", "sum"),
        "uds_descuadre_abs": ("_descuadre", lambda x: x.dropna().abs().sum()),
        "perdida_monetaria": ("_perdida_monetaria", "sum"),
        "obs_inventario": ("_obs_inv", lambda x: "; ".join(dict.fromkeys(o.strip() for o in x if o and o != "nan" and o.strip()))),
    }
    if is_control:
        agg["fallos_proceso"] = ("_fallo_proceso", "sum")
        agg["lotes_rev_proceso"] = ("_fallo_proceso", "size")
        agg["obs_proceso"] = ("_obs_proc", lambda x: "; ".join(dict.fromkeys(o.strip() for o in x if o and o != "nan" and o.strip())))

    grouped = df.groupby(col_ubic, sort=True).agg(**agg).reset_index()
    grouped.rename(columns={col_ubic: "Ubicación"}, inplace=True)
    grouped["lotes_erroneos"] = grouped["lotes_erroneos"].astype(int)
    grouped["Fiabilidad"] = 1 - (grouped["lotes_erroneos"] / grouped["lotes_auditados"])
    grouped["Fiabilidad_uds"] = np.where(
        grouped["uds_inventariadas"] > 0,
        1 - (grouped["uds_descuadre_abs"] / grouped["uds_inventariadas"]),
        1.0,
    )

    if is_control:
        grouped["fallos_proceso"] = grouped["fallos_proceso"].astype(int)
        grouped["Cumplimiento proceso"] = 1 - (grouped["fallos_proceso"] / grouped["lotes_rev_proceso"])

    col_mat = find_col(df.columns, ["Ref. Material", "Material"])
    t_lotes = int(grouped["lotes_auditados"].sum())
    t_err = int(grouped["lotes_erroneos"].sum())
    t_fproc = int(grouped["fallos_proceso"].sum()) if is_control and "fallos_proceso" in grouped.columns else 0
    t_lproc = int(grouped["lotes_rev_proceso"].sum()) if is_control and "lotes_rev_proceso" in grouped.columns else 0

    t_uds_inv = float(grouped["uds_inventariadas"].sum())
    t_uds_err = float(grouped["uds_erroneas"].sum())
    t_uds_desc = float(grouped["uds_descuadre_abs"].sum())
    t_perdida = float(grouped["perdida_monetaria"].sum())

    stats = {
        "ubicaciones": len(grouped),
        "lotes_total": t_lotes,
        "lotes_error": t_err,
        "fiabilidad_global": 1 - (t_err / t_lotes) if t_lotes > 0 else 1.0,
        "valor_auditado": float(grouped["valor_total"].sum()),
        "referencias_unicas": int(df[col_mat].nunique()) if col_mat else 0,
        "fallos_proceso": t_fproc,
        "lotes_proceso": t_lproc,
        "cumplimiento_proceso": 1 - (t_fproc / t_lproc) if t_lproc > 0 else 1.0,
        "uds_inventariadas": t_uds_inv,
        "uds_erroneas": t_uds_err,
        "uds_descuadre": t_uds_desc,
        "fiabilidad_uds": 1 - (t_uds_desc / t_uds_inv) if t_uds_inv > 0 else 1.0,
        "perdida_monetaria": t_perdida,
    }
    return grouped, stats


# =============================================================================
# 4b. VALIDATION — unjustified discrepancies
# =============================================================================
def validate_sheet(df, stype):
    """Find discrepancies without justification."""
    is_control = (stype == "CONTROL DIFERENCIADO")
    warnings = []

    col_ubic = find_col(df.columns, ["Ubicacion", "Ubicación"])
    col_mat = find_col(df.columns, ["Ref. Material", "Material"])
    col_lote = find_col(df.columns, ["Nº Lote", "Lote"])
    col_stock = find_col(df.columns, ["Stock", "Cantidad"])
    col_fisica = find_col(df.columns, ["Cant. Física", "Cant. Fisica"])
    col_fallo = find_col(df.columns, ["Fallo en el proceso", "Fallo en proceso"])
    col_obs_inv = find_col(df.columns, ["Obs. Inventario", "Observaciones Inventario", "Observaciones"])
    col_obs_proc = find_col(df.columns, ["Obs. Proceso", "Observaciones Proceso"])

    if not col_ubic or not col_stock:
        return warnings

    for _, row in df.iterrows():
        ubic = str(row[col_ubic]) if col_ubic else ""
        mat = str(row[col_mat]) if col_mat else ""
        lote = str(row[col_lote]) if col_lote else ""

        # --- Descuadre without observation (ALL segregations) ---
        if col_fisica:
            fis = pd.to_numeric(row[col_fisica], errors="coerce")
            stk = pd.to_numeric(row[col_stock], errors="coerce")
            if not pd.isna(fis) and not pd.isna(stk) and fis != stk:
                obs = ""
                if col_obs_inv:
                    obs = str(row[col_obs_inv]).strip()
                    if obs.lower() in ("nan", ""):
                        obs = ""
                if not obs:
                    warnings.append({
                        "Sección": stype,
                        "Tipo": "Descuadre sin justificar",
                        "Ubicación": ubic,
                        "Material": mat,
                        "Lote": lote,
                        "Descuadre": fis - stk,
                    })

        # --- Fallo en proceso without Obs. Proceso (control only) ---
        if is_control and col_fallo:
            fallo_val = str(row[col_fallo]).strip()
            fallo_lower = fallo_val.lower()
            if fallo_val and fallo_lower not in ("nan", "", "no", "n"):
                obs_proc = ""
                if col_obs_proc:
                    obs_proc = str(row[col_obs_proc]).strip()
                    if obs_proc.lower() in ("nan", ""):
                        obs_proc = ""
                if not obs_proc:
                    warnings.append({
                        "Sección": stype,
                        "Tipo": "Fallo en proceso sin justificar",
                        "Ubicación": ubic,
                        "Material": mat,
                        "Lote": lote,
                        "Descuadre": "",
                    })

    return warnings


def extract_loss_detail(df, stype, df_values_ext=None):
    """Extract lines with negative discrepancy (physical < stock) and their monetary loss."""
    col_ubic = find_col(df.columns, ["Ubicacion", "Ubicación"])
    col_mat = find_col(df.columns, ["Ref. Material", "Material"])
    col_lote = find_col(df.columns, ["Nº Lote", "Lote"])
    col_desc = find_col(df.columns, ["Descripción", "Descripcion"])
    col_stock = find_col(df.columns, ["Stock", "Cantidad"])
    col_fisica = find_col(df.columns, ["Cant. Física", "Cant. Fisica"])
    col_valor_unit = find_col(df.columns, ["Valor unitario", "Valor Unitario", "Precio"])

    if not col_ubic or not col_stock or not col_fisica:
        return []

    # Build external value lookup if available
    _ext_lookup = {}
    if df_values_ext is not None:
        ext_mat = find_col(df_values_ext.columns, ["Material", "Ref. Material", "Referencia"])
        ext_valor = find_col(df_values_ext.columns, ["Valor unitario", "Valor Unitario", "Unit Value", "Precio"])
        if ext_mat and ext_valor:
            for _, vr in df_values_ext.iterrows():
                mk = str(vr[ext_mat]).strip()
                vv = pd.to_numeric(vr[ext_valor], errors="coerce")
                if not pd.isna(vv) and mk not in _ext_lookup:
                    _ext_lookup[mk] = vv

    losses = []
    for _, row in df.iterrows():
        fis = pd.to_numeric(row[col_fisica], errors="coerce") if col_fisica else np.nan
        stk = pd.to_numeric(row[col_stock], errors="coerce") if col_stock else np.nan
        if pd.isna(fis) or pd.isna(stk):
            continue
        if fis < stk:
            uds_perdidas = stk - fis
            val_unit = pd.to_numeric(row[col_valor_unit], errors="coerce") if col_valor_unit else np.nan
            # Fallback to external values
            if pd.isna(val_unit) and col_mat and _ext_lookup:
                mk = str(row[col_mat]).strip()
                val_unit = _ext_lookup.get(mk, np.nan)
            perdida = uds_perdidas * val_unit if not pd.isna(val_unit) else np.nan
            losses.append({
                "Sección": stype,
                "Ubicación": str(row[col_ubic]) if col_ubic else "",
                "Material": str(row[col_mat]) if col_mat else "",
                "Descripción": str(row[col_desc]) if col_desc else "",
                "Lote": str(row[col_lote]) if col_lote else "",
                "Stock SAP": stk,
                "Cant. Física": fis,
                "Uds. perdidas": uds_perdidas,
                "Valor unitario": val_unit if not pd.isna(val_unit) else None,
                "Pérdida (€)": perdida if not pd.isna(perdida) else None,
            })
    return losses


# =============================================================================
# 5. RUN PROCESSING + VALIDATION
# =============================================================================
results = {}
all_stats = {}
for stype, df in sheet_map.items():
    kpis, stats = process_sheet(df, is_control=(stype == "CONTROL DIFERENCIADO"), df_values_ext=df_values_ext)
    if kpis is not None:
        results[stype] = kpis
        all_stats[stype] = stats

all_warnings = []
all_losses = []
for stype, df in sheet_map.items():
    all_warnings.extend(validate_sheet(df, stype))
    all_losses.extend(extract_loss_detail(df, stype, df_values_ext=df_values_ext))

stock_stats = {}
if df_stock_orig is not None:
    cu = find_col(df_stock_orig.columns, ["Ubicacion", "Ubicación"])
    cm = find_col(df_stock_orig.columns, ["Ref. Material", "Material"])
    if cu:
        stock_stats["total_ubicaciones"] = df_stock_orig[cu].nunique()
        stock_stats["total_lotes"] = len(df_stock_orig)
        stock_stats["total_referencias"] = df_stock_orig[cm].nunique() if cm else 0


# =============================================================================
# 6. DISPLAY
# =============================================================================
st.header("Resumen General")

total_lotes = sum(s["lotes_total"] for s in all_stats.values())
total_err = sum(s["lotes_error"] for s in all_stats.values())
total_ubics = sum(s["ubicaciones"] for s in all_stats.values())
total_fiab = 1 - (total_err / total_lotes) if total_lotes > 0 else 1.0
total_uds_inv = sum(s["uds_inventariadas"] for s in all_stats.values())
total_uds_err = sum(s["uds_erroneas"] for s in all_stats.values())
total_uds_desc = sum(s["uds_descuadre"] for s in all_stats.values())
total_fiab_uds = 1 - (total_uds_desc / total_uds_inv) if total_uds_inv > 0 else 1.0
total_perdida = sum(s["perdida_monetaria"] for s in all_stats.values())

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Ubicaciones auditadas", total_ubics)
m2.metric("Lotes auditados", f"{total_lotes:,}")
m3.metric("Uds. inventariadas", f"{total_uds_inv:,.0f}")
m4.metric("Fiabilidad por lotes", f"{total_fiab:.2%}")
m5.metric("Fiabilidad por uds.", f"{total_fiab_uds:.2%}",
          help="1 − (Uds. con descuadre / Uds. inventariadas)")

m6, m7, m8, m9 = st.columns(4)
m6.metric("Lotes erróneos", total_err)
m7.metric("Uds. erróneas", f"{total_uds_err:,.0f}",
          help="Unidades con descuadre (valor absoluto)")
if total_perdida > 0:
    m8.metric("Pérdida monetaria", f"{total_perdida:,.2f} €",
              help="Valor de las uds. que faltan (descuadre negativo × valor unitario)")
else:
    m8.metric("Pérdida monetaria", "0,00 €")
m9.metric("", "")

if stock_stats:
    st.subheader("Cobertura vs Stock total")
    cv1, cv2, cv3 = st.columns(3)
    pu = total_ubics / stock_stats["total_ubicaciones"] * 100 if stock_stats.get("total_ubicaciones") else 0
    pl = total_lotes / stock_stats["total_lotes"] * 100 if stock_stats.get("total_lotes") else 0
    cv1.metric("% Ubicaciones", f"{pu:.1f}%", delta=f"{total_ubics} de {stock_stats['total_ubicaciones']}")
    cv2.metric("% Lotes", f"{pl:.1f}%", delta=f"{total_lotes:,} de {stock_stats['total_lotes']:,}")
    if stock_stats.get("total_referencias"):
        tr = sum(s["referencias_unicas"] for s in all_stats.values())
        pr = tr / stock_stats["total_referencias"] * 100
        cv3.metric("% Referencias", f"{pr:.1f}%", delta=f"{tr} de {stock_stats['total_referencias']}")

# --- Loss detail ---
if all_losses:
    st.divider()
    st.header("Detalle de pérdidas (descuadre negativo)")
    df_losses = pd.DataFrame(all_losses)
    total_uds_perd = df_losses["Uds. perdidas"].sum()
    total_val_perd = df_losses["Pérdida (€)"].sum()
    lp1, lp2, lp3 = st.columns(3)
    lp1.metric("Líneas con pérdida", len(df_losses))
    lp2.metric("Uds. perdidas totales", f"{total_uds_perd:,.0f}")
    lp3.metric("Pérdida total", f"{total_val_perd:,.2f} €" if not pd.isna(total_val_perd) else "Sin valor unit.")

    # Format for display
    df_loss_disp = df_losses.copy()
    df_loss_disp["Uds. perdidas"] = df_loss_disp["Uds. perdidas"].map(lambda x: f"{x:,.0f}")
    df_loss_disp["Valor unitario"] = df_loss_disp["Valor unitario"].map(
        lambda x: f"{x:,.2f} €" if pd.notna(x) else "—")
    df_loss_disp["Pérdida (€)"] = df_loss_disp["Pérdida (€)"].map(
        lambda x: f"{x:,.2f} €" if pd.notna(x) else "—")
    st.dataframe(df_loss_disp, use_container_width=True, hide_index=True)

st.divider()

section_order = ["ALEATORIO", "CONTROL DIFERENCIADO", "MATERIAL VALIOSO"]

for stype in section_order:
    if stype not in results:
        continue
    kpis = results[stype]
    stats = all_stats[stype]

    st.subheader(stype)

    sc1, sc2, sc3, sc4, sc5, sc6 = st.columns(6)
    sc1.metric("Ubicaciones", stats["ubicaciones"])
    sc2.metric("Lotes auditados", f"{stats['lotes_total']:,}")
    sc3.metric("Uds. inventariadas", f"{stats['uds_inventariadas']:,.0f}")
    sc4.metric("Uds. erróneas", f"{stats['uds_erroneas']:,.0f}")
    sc5.metric("Fiabilidad (lotes)", f"{stats['fiabilidad_global']:.2%}")
    sc6.metric("Fiabilidad (uds.)", f"{stats['fiabilidad_uds']:.2%}")

    if stats["perdida_monetaria"] > 0:
        st.warning(f"Pérdida monetaria por descuadre negativo: **{stats['perdida_monetaria']:,.2f} €**")

    if stype == "CONTROL DIFERENCIADO":
        pc1, pc2, pc3 = st.columns(3)
        pc1.metric("Fallos proceso", stats["fallos_proceso"])
        pc2.metric("Lotes revisados (proceso)", stats["lotes_proceso"])
        pc3.metric("Cumplimiento proceso", f"{stats['cumplimiento_proceso']:.2%}")
        st.caption("**Fiabilidad** = mercancía. **Cumplimiento proceso** = pegatinas/etiquetas. Son independientes.")

    if stype == "MATERIAL VALIOSO" and stats["valor_auditado"] > 0:
        vc1, vc2 = st.columns(2)
        vc1.metric("Valor auditado", f"{stats['valor_auditado']:,.2f}")
        vc2.metric("Referencias únicas", stats["referencias_unicas"])

    # Table
    disp = pd.DataFrame()
    disp["Ubicación"] = kpis["Ubicación"]
    disp["Lotes auditados"] = kpis["lotes_auditados"]
    disp["Lotes erróneos"] = kpis["lotes_erroneos"]
    disp["Fiabilidad (lotes)"] = kpis["Fiabilidad"].map(lambda x: f"{x:.2%}" if x < 1 else "100%")
    disp["Uds. inventariadas"] = kpis["uds_inventariadas"].map(lambda x: f"{x:,.0f}")
    disp["Uds. erróneas"] = kpis["uds_erroneas"].map(lambda x: f"{x:,.0f}" if x > 0 else "")
    disp["Fiabilidad (uds.)"] = kpis["Fiabilidad_uds"].map(lambda x: f"{x:.2%}" if x < 1 else "100%")
    if kpis["perdida_monetaria"].sum() > 0:
        disp["Pérdida (€)"] = kpis["perdida_monetaria"].map(lambda x: f"{x:,.2f}" if x > 0 else "")

    if stype == "CONTROL DIFERENCIADO" and "fallos_proceso" in kpis.columns:
        disp["Fallos proceso"] = kpis["fallos_proceso"]
        disp["Cumplim. proceso"] = kpis["Cumplimiento proceso"].map(lambda x: f"{x:.2%}" if x < 1 else "100%")

    # Tipo de error: combine unique obs
    if stype == "CONTROL DIFERENCIADO" and "obs_proceso" in kpis.columns:
        def _combine_obs(r):
            parts = []
            inv = str(r.get("obs_inventario", "")).strip()
            if inv and inv.lower() != "nan":
                parts.append(inv)
            proc = str(r.get("obs_proceso", "")).strip()
            if proc and proc.lower() != "nan":
                parts.append(proc)
            return "; ".join(parts)
        disp["Tipo de error"] = kpis.apply(_combine_obs, axis=1)
    else:
        disp["Tipo de error"] = kpis["obs_inventario"]

    st.dataframe(disp, use_container_width=True, hide_index=True)
    st.divider()


# =============================================================================
# 7. VALIDATION WARNINGS
# =============================================================================
if all_warnings:
    st.header("Validación")
    df_warn = pd.DataFrame(all_warnings)

    n_desc = len(df_warn[df_warn["Tipo"] == "Descuadre sin justificar"])
    n_fallo = len(df_warn[df_warn["Tipo"] == "Fallo en proceso sin justificar"])

    if n_desc > 0:
        st.error(f"**{n_desc}** descuadre(s) sin observación que lo justifique")
    if n_fallo > 0:
        st.error(f"**{n_fallo}** fallo(s) en proceso sin observación que lo justifique")

    st.dataframe(
        df_warn[["Sección", "Tipo", "Ubicación", "Material", "Lote", "Descuadre"]],
        use_container_width=True, hide_index=True,
    )
else:
    st.header("Validación")
    st.success("Todas las incidencias tienen observación justificativa.")


# =============================================================================
# 8. EXCEL REPORT
# =============================================================================
st.header("Descargar Reporte")


def build_excel(results, all_stats, stock_stats, all_warnings, all_losses=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        ws_name = "KIPs Consolidado Aud.Interna"
        pd.DataFrame().to_excel(writer, sheet_name=ws_name, index=False)
        wb = writer.book
        ws = writer.sheets[ws_name]

        # Formats
        title_f = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F3864", "bottom": 2})
        sub_f = wb.add_format({"bold": True, "font_size": 11, "font_color": "#1F3864", "italic": True})
        hdr_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
        t_lbl = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#D6DCE4", "border": 1})
        t_num = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#D6DCE4", "border": 1, "num_format": "#,##0", "align": "center"})
        t_pct = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#D6DCE4", "border": 1, "num_format": "0.00%", "align": "center"})
        sec_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#E2EFDA", "border": 1})
        d_f = wb.add_format({"font_size": 10, "border": 1})
        d_c = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0"})
        p_ok = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#006100", "bg_color": "#C6EFCE"})
        p_w = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%"})
        p_bad = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#9C0006", "bg_color": "#FFC7CE"})
        obs_f = wb.add_format({"font_size": 9, "border": 1, "text_wrap": True})
        s_title = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#FFC000", "font_color": "#1F3864", "border": 1, "align": "center"})
        s_hdr = wb.add_format({"bold": True, "font_size": 9, "bg_color": "#FFC000", "font_color": "#1F3864", "border": 1, "text_wrap": True, "align": "center"})
        cov_h = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#4472C4", "font_color": "white", "border": 1, "align": "center"})
        cov_p = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.0%"})
        money_f = wb.add_format({"font_size": 10, "border": 1, "num_format": "#,##0.00", "align": "center"})

        def wpct(r, c, v):
            f = p_ok if v >= 1.0 else (p_bad if v < 0.9 else p_w)
            ws.write_number(r, c, v, f)

        widths = [24, 18, 14, 16, 14, 14, 16, 16, 16, 50, 3, 24, 18, 22, 50]
        for i, w in enumerate(widths):
            ws.set_column(i, i, w)

        tu = sum(s["ubicaciones"] for s in all_stats.values())
        tl = sum(s["lotes_total"] for s in all_stats.values())
        te = sum(s["lotes_error"] for s in all_stats.values())
        tf = 1 - (te / tl) if tl > 0 else 1.0
        t_uds = sum(s["uds_inventariadas"] for s in all_stats.values())
        t_uds_d = sum(s["uds_descuadre"] for s in all_stats.values())
        tf_uds = 1 - (t_uds_d / t_uds) if t_uds > 0 else 1.0
        t_perd = sum(s["perdida_monetaria"] for s in all_stats.values())

        loss_f = wb.add_format({"bold": True, "font_size": 11, "border": 1, "num_format": "#,##0.00 €",
                                "align": "center", "font_color": "#9C0006", "bg_color": "#FFC7CE"})
        loss_hdr = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#9C0006", "font_color": "white",
                                  "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})

        row = 0
        ws.merge_range(row, 0, row, 7, f"CONSOLIDADO AUDITORÍA INTERNA — {date.today().strftime('%d/%m/%Y')}", title_f)
        ws.set_row(row, 30)
        row += 2

        t_uds_e = sum(s["uds_erroneas"] for s in all_stats.values())

        for c, h in enumerate(["", "Ubicaciones\nauditadas", "Lotes\nauditados", "Uds.\ninventariadas",
                                "Lotes\nerróneos", "Uds.\nerróneas", "Fiabilidad\n(lotes)", "Fiabilidad\n(uds.)", "Pérdida\nmonetaria"]):
            fmt = loss_hdr if c == 8 else hdr_f
            ws.write(row, c, h, fmt)
        ws.set_row(row, 35)
        row += 1
        ws.write(row, 0, "TOTAL", t_lbl); ws.write(row, 1, tu, t_num); ws.write(row, 2, tl, t_num)
        ws.write(row, 3, t_uds, t_num); ws.write(row, 4, te, t_num); ws.write(row, 5, t_uds_e, t_num)
        ws.write(row, 6, tf, t_pct); ws.write(row, 7, tf_uds, t_pct)
        ws.write(row, 8, t_perd, loss_f if t_perd > 0 else money_f)
        row += 2

        if stock_stats:
            ws.merge_range(row, 0, row, 5, "COBERTURA vs STOCK TOTAL", sub_f)
            row += 1
            for c, h in enumerate(["", "Auditado", "Total almacén", "% Cobertura", "", ""]):
                ws.write(row, c, h, cov_h)
            row += 1
            ws.write(row, 0, "Ubicaciones", d_f); ws.write(row, 1, tu, d_c)
            ws.write(row, 2, stock_stats.get("total_ubicaciones", 0), d_c)
            ws.write(row, 3, tu / stock_stats["total_ubicaciones"] if stock_stats.get("total_ubicaciones") else 0, cov_p)
            row += 1
            ws.write(row, 0, "Lotes", d_f); ws.write(row, 1, tl, d_c)
            ws.write(row, 2, stock_stats.get("total_lotes", 0), d_c)
            ws.write(row, 3, tl / stock_stats["total_lotes"] if stock_stats.get("total_lotes") else 0, cov_p)
            row += 1
            if stock_stats.get("total_referencias"):
                trefs = sum(s["referencias_unicas"] for s in all_stats.values())
                ws.write(row, 0, "Referencias", d_f); ws.write(row, 1, trefs, d_c)
                ws.write(row, 2, stock_stats["total_referencias"], d_c)
                ws.write(row, 3, trefs / stock_stats["total_referencias"], cov_p)
                row += 1
            row += 1

        for c, h in enumerate(["", "Ubicación", "Lotes\nauditados", "Uds.\ninventariadas",
                                "Lotes\nerróneos", "Uds.\nerróneas", "Fiabilidad\n(lotes)", "Fiabilidad\n(uds.)",
                                "Pérdida (€)", "Tipo de error"]):
            fmt = loss_hdr if c == 8 else hdr_f
            ws.write(row, c, h, fmt)
        ws.set_row(row, 35)
        row += 1

        loss_cell = wb.add_format({"font_size": 10, "border": 1, "num_format": "#,##0.00",
                                   "align": "center", "font_color": "#9C0006"})

        ctrl_start = None
        for stype in ["ALEATORIO", "CONTROL DIFERENCIADO", "MATERIAL VALIOSO"]:
            if stype not in results:
                continue
            kpis = results[stype]
            stats = all_stats[stype]
            for i, (_, r) in enumerate(kpis.iterrows()):
                if i == 0:
                    ws.write(row, 0, stype, sec_f)
                    if stype == "CONTROL DIFERENCIADO":
                        ctrl_start = row
                else:
                    ws.write(row, 0, "", d_f)
                ws.write(row, 1, r["Ubicación"], d_f)
                ws.write(row, 2, int(r["lotes_auditados"]), d_c)
                ws.write(row, 3, int(r["uds_inventariadas"]), d_c)
                ws.write(row, 4, int(r["lotes_erroneos"]), d_c)
                uds_err_val = int(r.get("uds_erroneas", 0))
                ws.write(row, 5, uds_err_val, d_c) if uds_err_val > 0 else ws.write(row, 5, "", d_f)
                wpct(row, 6, r["Fiabilidad"])
                wpct(row, 7, r["Fiabilidad_uds"])
                perd_val = float(r.get("perdida_monetaria", 0))
                if perd_val > 0:
                    ws.write(row, 8, perd_val, loss_cell)
                else:
                    ws.write(row, 8, "", d_f)
                obs_text = r.get("obs_inventario", "")
                if isinstance(obs_text, str) and obs_text.strip().lower() == "nan":
                    obs_text = ""
                ws.write(row, 9, obs_text, obs_f)
                row += 1
            ws.write(row, 0, f"Subtotal {stype}", t_lbl)
            ws.write(row, 1, f"{stats['ubicaciones']} ubic.", t_lbl)
            ws.write(row, 2, stats["lotes_total"], t_num)
            ws.write(row, 3, int(stats["uds_inventariadas"]), t_num)
            ws.write(row, 4, stats["lotes_error"], t_num)
            ws.write(row, 5, int(stats["uds_erroneas"]), t_num)
            ws.write(row, 6, stats["fiabilidad_global"], t_pct)
            ws.write(row, 7, stats["fiabilidad_uds"], t_pct)
            ws.write(row, 8, stats["perdida_monetaria"], loss_f if stats["perdida_monetaria"] > 0 else t_num)
            ws.write(row, 9, "", t_lbl)
            row += 1

        # === SIDE TABLE: Control Diferenciado — Process compliance ===
        if ctrl_start is not None and "CONTROL DIFERENCIADO" in results:
            ck = results["CONTROL DIFERENCIADO"]
            cs = all_stats["CONTROL DIFERENCIADO"]
            sc = 11  # start column (after main table cols 0-9 + gap)
            for ci in range(sc, sc + 5):
                ws.set_column(ci, ci, 22)
            hr = max(ctrl_start - 2, 0)
            ws.merge_range(hr, sc, hr, sc + 4, "Control Diferenciado — Cumplimiento del proceso", s_title)
            sr = hr + 1
            for c, h in enumerate(["Ubicación", "Lotes\nauditados", "Lotes\nerróneos", "Cumplimiento\nproceso", "Tipo de error"]):
                ws.write(sr, sc + c, h, s_hdr)
            ws.set_row(sr, 35)
            for i, (_, r) in enumerate(ck.iterrows()):
                dr = sr + 1 + i
                ws.write(dr, sc, r["Ubicación"], d_f)
                ws.write(dr, sc + 1, int(r["lotes_auditados"]), d_c)
                ws.write(dr, sc + 2, int(r["lotes_erroneos"]), d_c)
                if "Cumplimiento proceso" in r.index:
                    wpct(dr, sc + 3, r["Cumplimiento proceso"])
                obs_proc = r.get("obs_proceso", "")
                if isinstance(obs_proc, str) and obs_proc.strip().lower() == "nan":
                    obs_proc = ""
                ws.write(dr, sc + 4, obs_proc, obs_f)
            tr = sr + 1 + len(ck)
            ws.write(tr, sc, "TOTAL", t_lbl)
            ws.write(tr, sc + 1, cs["lotes_total"], t_num)
            ws.write(tr, sc + 2, cs["lotes_error"], t_num)
            ws.write(tr, sc + 3, cs["cumplimiento_proceso"], t_pct)
            ws.write(tr, sc + 4, "", t_lbl)

        if "MATERIAL VALIOSO" in all_stats:
            vs = all_stats["MATERIAL VALIOSO"]
            row += 1
            ws.merge_range(row, 0, row, 9, "RESUMEN MATERIAL VALIOSO", sub_f)
            row += 1
            ws.write(row, 0, "Valor total auditado", d_f); ws.write(row, 1, vs["valor_auditado"], money_f)
            ws.write(row, 2, "Referencias únicas", d_f); ws.write(row, 3, vs["referencias_unicas"], d_c)
            ws.write(row, 4, "Pérdida monetaria", d_f)
            ws.write(row, 5, vs["perdida_monetaria"], loss_f if vs["perdida_monetaria"] > 0 else money_f)
            row += 1
            ws.write(row, 0, "Ubicaciones", d_f); ws.write(row, 1, vs["ubicaciones"], d_c)
            ws.write(row, 2, "Fiabilidad (lotes)", d_f); wpct(row, 3, vs["fiabilidad_global"])
            ws.write(row, 4, "Fiabilidad (uds.)", d_f); wpct(row, 5, vs["fiabilidad_uds"])

        ws.print_area(0, 0, row + 1, 15)
        ws.set_landscape()
        ws.set_paper(9)
        ws.fit_to_pages(1, 0)

        # --- Validation sheet ---
        if all_warnings:
            vws_name = "Validación"
            df_w = pd.DataFrame(all_warnings)
            df_w.to_excel(writer, sheet_name=vws_name, index=False)
            vws = writer.sheets[vws_name]

            v_hdr = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#C00000", "font_color": "white", "border": 1, "text_wrap": True, "align": "center"})
            v_cell = wb.add_format({"font_size": 10, "border": 1})
            v_warn = wb.add_format({"font_size": 10, "border": 1, "bg_color": "#FFC7CE"})

            for ci, cn in enumerate(df_w.columns):
                vws.write(0, ci, cn, v_hdr)
            for ri in range(len(df_w)):
                for ci in range(len(df_w.columns)):
                    val = df_w.iloc[ri, ci]
                    fmt = v_warn if df_w.columns[ci] == "Tipo" else v_cell
                    if pd.isna(val) or val == "":
                        vws.write_blank(ri + 1, ci, "", v_cell)
                    else:
                        vws.write(ri + 1, ci, val, fmt)

            for ci, cn in enumerate(df_w.columns):
                mx = max(df_w[cn].astype(str).map(len).max(), len(str(cn)))
                vws.set_column(ci, ci, min(mx + 3, 40))
            vws.freeze_panes(1, 0)

        # --- Loss detail sheet ---
        if all_losses:
            lws_name = "Detalle Pérdidas"
            df_l = pd.DataFrame(all_losses)
            df_l.to_excel(writer, sheet_name=lws_name, index=False)
            lws = writer.sheets[lws_name]

            l_hdr = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#9C0006",
                                   "font_color": "white", "border": 1, "text_wrap": True, "align": "center"})
            l_cell = wb.add_format({"font_size": 10, "border": 1})
            l_money = wb.add_format({"font_size": 10, "border": 1, "num_format": "#,##0.00 €",
                                     "font_color": "#9C0006"})
            l_num = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0"})

            for ci, cn in enumerate(df_l.columns):
                lws.write(0, ci, cn, l_hdr)
            for ri in range(len(df_l)):
                for ci in range(len(df_l.columns)):
                    val = df_l.iloc[ri, ci]
                    cn = df_l.columns[ci]
                    if pd.isna(val) or val is None:
                        lws.write_blank(ri + 1, ci, "", l_cell)
                    elif cn in ("Pérdida (€)", "Valor unitario"):
                        lws.write_number(ri + 1, ci, float(val), l_money)
                    elif cn in ("Stock SAP", "Cant. Física", "Uds. perdidas"):
                        lws.write_number(ri + 1, ci, float(val), l_num)
                    else:
                        lws.write(ri + 1, ci, val, l_cell)

            # Total row
            total_row = len(df_l) + 1
            lws.write(total_row, 0, "TOTAL", t_lbl)
            for ci in range(1, len(df_l.columns)):
                cn = df_l.columns[ci]
                if cn == "Uds. perdidas":
                    lws.write_number(total_row, ci, float(df_l[cn].sum()), t_num)
                elif cn == "Pérdida (€)":
                    total_p = df_l[cn].sum()
                    lws.write_number(total_row, ci, float(total_p) if not pd.isna(total_p) else 0, loss_f)
                else:
                    lws.write(total_row, ci, "", t_lbl)

            for ci, cn in enumerate(df_l.columns):
                mx = max(df_l[cn].astype(str).map(len).max(), len(str(cn)))
                lws.set_column(ci, ci, min(mx + 3, 40))
            lws.freeze_panes(1, 0)
            lws.set_landscape()

    return output.getvalue()


if results:
    st.download_button(
        "Descargar Reporte Consolidado (Excel)",
        build_excel(results, all_stats, stock_stats, all_warnings, all_losses),
        file_name=f"CONSOLIDADO_AUDITORIA_INTERNA_{date.today().strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )


# =============================================================================
# 9. SAVE TO HISTORY
# =============================================================================
if results and not SS.get("consol_saved_to_history", False):
    st.divider()
    st.markdown(
        "**Guardar en historial** — Para mantener un registro de auditorías realizadas."
    )
    if st.button("Guardar este reporte en el historial", type="secondary", use_container_width=True):
        hist = load_consol_history()
        key = f"{date.today().isoformat()}_{centro_sel or 'sin_centro'}"
        entry = {
            "key": key,
            "fecha": date.today().isoformat(),
            "centro": centro_sel,
            "total": {
                "ubicaciones": sum(s["ubicaciones"] for s in all_stats.values()),
                "lotes_total": sum(s["lotes_total"] for s in all_stats.values()),
                "lotes_error": sum(s["lotes_error"] for s in all_stats.values()),
                "uds_inventariadas": sum(s["uds_inventariadas"] for s in all_stats.values()),
                "uds_erroneas": sum(s["uds_erroneas"] for s in all_stats.values()),
                "fiabilidad_lotes": total_fiab,
                "fiabilidad_uds": total_fiab_uds,
                "perdida_monetaria": total_perdida,
            },
            "secciones": {stype: stats for stype, stats in all_stats.items()},
        }
        if stock_stats:
            entry["cobertura"] = stock_stats
        # Dedup by key
        hist = [h for h in hist if h.get("key") != key]
        hist.append(entry)
        hist.sort(key=lambda h: h["key"])
        save_consol_history(hist)
        SS["consol_saved_to_history"] = True
        st.rerun()
elif results and SS.get("consol_saved_to_history"):
    st.success("Reporte guardado en el historial.")


# =============================================================================
# 10. HISTORICAL DATA
# =============================================================================
st.divider()
st.header("Historial de auditorías")

consol_hist = load_consol_history()

if not consol_hist:
    st.info("No hay auditorías guardadas en el historial.")
else:
    # Trend table
    hist_rows = []
    for h in consol_hist:
        t = h.get("total", {})
        hist_rows.append({
            "Fecha": h.get("fecha", ""),
            "Centro": h.get("centro", "—"),
            "Ubicaciones": t.get("ubicaciones", 0),
            "Lotes": t.get("lotes_total", 0),
            "Lotes err.": t.get("lotes_error", 0),
            "Uds. inv.": f"{t.get('uds_inventariadas', 0):,.0f}",
            "Uds. err.": f"{t.get('uds_erroneas', 0):,.0f}",
            "Fiab. lotes": f"{t.get('fiabilidad_lotes', 1):.2%}",
            "Fiab. uds.": f"{t.get('fiabilidad_uds', 1):.2%}",
            "Pérdida (€)": f"{t.get('perdida_monetaria', 0):,.2f}",
        })
    st.dataframe(pd.DataFrame(hist_rows), use_container_width=True, hide_index=True)

    with st.expander("Gestionar historial"):
        for i, h in enumerate(consol_hist):
            c1, c2 = st.columns([5, 1])
            c1.write(f"**{h.get('fecha', '')}** — {h.get('centro', 'Sin centro')} "
                     f"({h.get('total', {}).get('ubicaciones', 0)} ubic.)")
            if c2.button("Eliminar", key=f"del_consol_{i}", type="secondary"):
                hist_current = load_consol_history()
                hist_current = [x for x in hist_current if x.get("key") != h["key"]]
                save_consol_history(hist_current)
                st.rerun()
        if st.button("Limpiar todo el historial", type="secondary"):
            save_consol_history([])
            st.rerun()
