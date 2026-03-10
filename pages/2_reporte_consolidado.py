import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
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
# 1. FILE UPLOADS — persisted in session_state
# =============================================================================
SS = st.session_state

st.sidebar.header("Archivos — Reporte")
audit_up = st.sidebar.file_uploader("Excel de auditoría rellenado", type=["xlsx", "xls"], key="rpt_audit_up")
stock_up = st.sidebar.file_uploader("Stock original (opcional, para % cobertura)", type=["xlsx", "xls"], key="rpt_stock_up")

if audit_up is not None:
    SS["rpt_audit_bytes"] = audit_up.getvalue()
if stock_up is not None:
    SS["rpt_stock_bytes"] = stock_up.getvalue()

if "rpt_audit_bytes" in SS:
    st.sidebar.success("Auditoría cargada")
if "rpt_stock_bytes" in SS:
    st.sidebar.success("Stock original cargado")

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
def process_sheet(df, is_control=False):
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

    df["_is_error"] = False
    mask_filled = df["_fisica"].notna()
    if mask_filled.any():
        df.loc[mask_filled, "_is_error"] = df.loc[mask_filled, "_stock"] != df.loc[mask_filled, "_fisica"]

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


# =============================================================================
# 5. RUN PROCESSING + VALIDATION
# =============================================================================
results = {}
all_stats = {}
for stype, df in sheet_map.items():
    kpis, stats = process_sheet(df, is_control=(stype == "CONTROL DIFERENCIADO"))
    if kpis is not None:
        results[stype] = kpis
        all_stats[stype] = stats

all_warnings = []
for stype, df in sheet_map.items():
    all_warnings.extend(validate_sheet(df, stype))

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
total_uds_desc = sum(s["uds_descuadre"] for s in all_stats.values())
total_fiab_uds = 1 - (total_uds_desc / total_uds_inv) if total_uds_inv > 0 else 1.0
total_perdida = sum(s["perdida_monetaria"] for s in all_stats.values())

m1, m2, m3, m4 = st.columns(4)
m1.metric("Ubicaciones auditadas", total_ubics)
m2.metric("Lotes auditados", f"{total_lotes:,}")
m3.metric("Uds. inventariadas", f"{total_uds_inv:,.0f}")
m4.metric("Fiabilidad por lotes", f"{total_fiab:.2%}")

m5, m6, m7, m8 = st.columns(4)
m5.metric("Lotes erróneos", total_err)
m6.metric("Fiabilidad por uds.", f"{total_fiab_uds:.2%}",
          help="1 − (Uds. con descuadre / Uds. inventariadas)")
if total_perdida > 0:
    m7.metric("Pérdida monetaria", f"{total_perdida:,.2f} €",
              help="Valor de las uds. que faltan (descuadre negativo × valor unitario)")
else:
    m7.metric("Pérdida monetaria", "0,00 €")
m8.metric("", "")

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

st.divider()

section_order = ["ALEATORIO", "CONTROL DIFERENCIADO", "MATERIAL VALIOSO"]

for stype in section_order:
    if stype not in results:
        continue
    kpis = results[stype]
    stats = all_stats[stype]

    st.subheader(stype)

    sc1, sc2, sc3, sc4, sc5 = st.columns(5)
    sc1.metric("Ubicaciones", stats["ubicaciones"])
    sc2.metric("Lotes auditados", f"{stats['lotes_total']:,}")
    sc3.metric("Uds. inventariadas", f"{stats['uds_inventariadas']:,.0f}")
    sc4.metric("Fiabilidad (lotes)", f"{stats['fiabilidad_global']:.2%}")
    sc5.metric("Fiabilidad (uds.)", f"{stats['fiabilidad_uds']:.2%}")

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


def build_excel(results, all_stats, stock_stats, all_warnings):
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

        widths = [24, 18, 14, 16, 14, 16, 16, 16, 50, 3, 24, 18, 22, 50]
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

        for c, h in enumerate(["", "Ubicaciones\nauditadas", "Lotes\nauditados", "Uds.\ninventariadas",
                                "Lotes\nerróneos", "Fiabilidad\n(lotes)", "Fiabilidad\n(uds.)", "Pérdida\nmonetaria"]):
            fmt = loss_hdr if c == 7 else hdr_f
            ws.write(row, c, h, fmt)
        ws.set_row(row, 35)
        row += 1
        ws.write(row, 0, "TOTAL", t_lbl); ws.write(row, 1, tu, t_num); ws.write(row, 2, tl, t_num)
        ws.write(row, 3, t_uds, t_num); ws.write(row, 4, te, t_num)
        ws.write(row, 5, tf, t_pct); ws.write(row, 6, tf_uds, t_pct)
        ws.write(row, 7, t_perd, loss_f if t_perd > 0 else money_f)
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
                                "Lotes\nerróneos", "Fiabilidad\n(lotes)", "Fiabilidad\n(uds.)",
                                "Pérdida (€)", "Tipo de error"]):
            fmt = loss_hdr if c == 7 else hdr_f
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
                wpct(row, 5, r["Fiabilidad"])
                wpct(row, 6, r["Fiabilidad_uds"])
                perd_val = float(r.get("perdida_monetaria", 0))
                if perd_val > 0:
                    ws.write(row, 7, perd_val, loss_cell)
                else:
                    ws.write(row, 7, "", d_f)
                obs_text = r.get("obs_inventario", "")
                if isinstance(obs_text, str) and obs_text.strip().lower() == "nan":
                    obs_text = ""
                ws.write(row, 8, obs_text, obs_f)
                row += 1
            ws.write(row, 0, f"Subtotal {stype}", t_lbl)
            ws.write(row, 1, f"{stats['ubicaciones']} ubic.", t_lbl)
            ws.write(row, 2, stats["lotes_total"], t_num)
            ws.write(row, 3, int(stats["uds_inventariadas"]), t_num)
            ws.write(row, 4, stats["lotes_error"], t_num)
            ws.write(row, 5, stats["fiabilidad_global"], t_pct)
            ws.write(row, 6, stats["fiabilidad_uds"], t_pct)
            ws.write(row, 7, stats["perdida_monetaria"], loss_f if stats["perdida_monetaria"] > 0 else t_num)
            ws.write(row, 8, "", t_lbl)
            row += 1

        # === SIDE TABLE: Control Diferenciado — Process compliance ===
        if ctrl_start is not None and "CONTROL DIFERENCIADO" in results:
            ck = results["CONTROL DIFERENCIADO"]
            cs = all_stats["CONTROL DIFERENCIADO"]
            sc = 10  # start column (after main table cols 0-8 + gap)
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
            ws.merge_range(row, 0, row, 8, "RESUMEN MATERIAL VALIOSO", sub_f)
            row += 1
            ws.write(row, 0, "Valor total auditado", d_f); ws.write(row, 1, vs["valor_auditado"], money_f)
            ws.write(row, 2, "Referencias únicas", d_f); ws.write(row, 3, vs["referencias_unicas"], d_c)
            ws.write(row, 4, "Pérdida monetaria", d_f)
            ws.write(row, 5, vs["perdida_monetaria"], loss_f if vs["perdida_monetaria"] > 0 else money_f)
            row += 1
            ws.write(row, 0, "Ubicaciones", d_f); ws.write(row, 1, vs["ubicaciones"], d_c)
            ws.write(row, 2, "Fiabilidad (lotes)", d_f); wpct(row, 3, vs["fiabilidad_global"])
            ws.write(row, 4, "Fiabilidad (uds.)", d_f); wpct(row, 5, vs["fiabilidad_uds"])

        ws.print_area(0, 0, row + 1, 14)
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

    return output.getvalue()


if results:
    st.download_button(
        "Descargar Reporte Consolidado (Excel)",
        build_excel(results, all_stats, stock_stats, all_warnings),
        file_name=f"CONSOLIDADO_AUDITORIA_INTERNA_{date.today().strftime('%d-%m-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )
