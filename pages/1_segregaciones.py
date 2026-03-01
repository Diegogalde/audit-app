import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date
from xlsxwriter.utility import xl_col_to_name

st.title("Generar Segregaciones")
st.markdown("Genera muestras de auditoría a partir del extracto de stock.")


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
# 1. SEGREGATION TYPE SELECTOR
# =============================================================================
tipos_seg = st.multiselect(
    "Segregaciones a generar",
    ["Aleatorio", "Control Diferenciado", "Material Valioso"],
    default=["Aleatorio", "Control Diferenciado", "Material Valioso"],
    key="tipos_seg",
)

if not tipos_seg:
    st.warning("Selecciona al menos un tipo de segregación.")
    st.stop()


# =============================================================================
# 2. FILE UPLOADS — stored in session_state for persistence
# =============================================================================
SS = st.session_state

needs_values = "Aleatorio" in tipos_seg or "Material Valioso" in tipos_seg
needs_control = "Control Diferenciado" in tipos_seg

st.sidebar.header("Archivos — Segregaciones")

stock_up = st.sidebar.file_uploader("Excel de Stock", type=["xlsx", "xls"], key="seg_stock_up")
values_up = st.sidebar.file_uploader("Excel de Valores Unitarios", type=["xlsx", "xls"], key="seg_values_up")
control_up = st.sidebar.file_uploader("Excel de Control Diferenciado", type=["xlsx", "xls"], key="seg_control_up")

# Persist in session_state
if stock_up is not None:
    SS["seg_stock_bytes"] = stock_up.getvalue()
if values_up is not None:
    SS["seg_values_bytes"] = values_up.getvalue()
if control_up is not None:
    SS["seg_control_bytes"] = control_up.getvalue()

has_stock = "seg_stock_bytes" in SS
has_values = "seg_values_bytes" in SS
has_control = "seg_control_bytes" in SS

if has_stock:
    st.sidebar.success("Stock cargado")
if has_values:
    st.sidebar.success("Valores cargados")
if has_control:
    st.sidebar.success("Control cargado")

# Dynamic file requirements
missing = []
if not has_stock:
    missing.append("Stock")
if needs_values and not has_values:
    missing.append("Valores Unitarios")
if needs_control and not has_control:
    missing.append("Control Diferenciado")

if missing:
    st.info(f"Sube los archivos necesarios: **{', '.join(missing)}**")
    st.stop()


# =============================================================================
# 3. LOAD DATA
# =============================================================================
@st.cache_data
def load_excel_bytes(data):
    return pd.read_excel(BytesIO(data))


@st.cache_data
def load_all_sheets_bytes(data):
    xls = pd.ExcelFile(BytesIO(data))
    frames = []
    for name in xls.sheet_names:
        frames.append(pd.read_excel(xls, sheet_name=name))
    return pd.concat(frames, ignore_index=True)


df_stock = load_excel_bytes(SS["seg_stock_bytes"])
df_values_all = load_all_sheets_bytes(SS["seg_values_bytes"]) if has_values else None
df_control = load_excel_bytes(SS["seg_control_bytes"]) if has_control else None


# =============================================================================
# 4. AUTO-DETECT COLUMNS
# =============================================================================
COL_UBIC = find_col(df_stock.columns, ["Ubicacion", "Ubicación"])
COL_MAT = find_col(df_stock.columns, ["Ref. Material", "Ref.Material", "Material"])
COL_LOTE = find_col(df_stock.columns, ["Nº Lote", "N° Lote", "Lote", "Batch"])
COL_CANT = find_col(df_stock.columns, ["Stock", "Cantidad", "Qty"])
COL_DESC = find_col(df_stock.columns, ["Descripción", "Descripcion"])
COL_SERIE = find_col(df_stock.columns, ["Nº Serie", "N° Serie"])
COL_UB = find_col(df_stock.columns, ["Unidad Base", "Unidad base"])
COL_CENTRO = find_col(df_stock.columns, ["Ref. centro", "Ref centro", "Centro"])
COL_ALMACEN = find_col(df_stock.columns, ["Ref. Almacén", "Ref. Almacen", "Almacén"])
COL_SOK = find_col(df_stock.columns, ["Stock OK"])
COL_SBL = find_col(df_stock.columns, ["Stock Bloqueado"])
COL_TBLOQ = find_col(df_stock.columns, ["Tipo Bloqueo", "Tipo bloqueo"])

VAL_MAT = VAL_LOTE = VAL_VALOR = None
if df_values_all is not None:
    VAL_MAT = find_col(df_values_all.columns, ["Material", "Ref. Material", "Referencia"])
    VAL_LOTE = find_col(df_values_all.columns, ["Batch", "Lote", "Nº Lote"])
    VAL_VALOR = find_col(df_values_all.columns, ["Valor unitario", "Valor Unitario", "Unit Value", "Precio"])

CTRL_MAT = None
if df_control is not None:
    CTRL_MAT = find_col(df_control.columns, ["Material", "Ref. Material", "Referencia", "Código"])

missing_cols = []
if not COL_UBIC: missing_cols.append("Ubicación")
if not COL_MAT: missing_cols.append("Material")
if not COL_LOTE: missing_cols.append("Lote")
if not COL_CANT: missing_cols.append("Stock/Cantidad")
if needs_values and not VAL_VALOR: missing_cols.append("Valor unitario")
if needs_control and not CTRL_MAT: missing_cols.append("Material (control)")
if missing_cols:
    st.error(f"No se detectaron: {', '.join(missing_cols)}")
    st.stop()


# =============================================================================
# 5. VALUE MERGE
# =============================================================================
@st.cache_data
def merge_values(_df_stock, _df_values, col_mat, col_lote, col_cant, val_mat, val_lote, val_valor):
    df_work = _df_stock.copy()
    df_work["_mat_key"] = df_work[col_mat].astype(str).str.strip()
    df_work["_lote_key"] = df_work[col_lote].astype(str).str.strip()

    df_val = _df_values.copy()
    if val_mat:
        df_val["_mat_key"] = df_val[val_mat].astype(str).str.strip()
    if val_lote:
        df_val["_lote_key"] = df_val[val_lote].astype(str).str.strip()
    df_val["_valor_unitario"] = pd.to_numeric(df_val[val_valor], errors="coerce")

    merged = df_work.copy()
    merged["_valor_unitario"] = np.nan

    if val_mat and val_lote:
        vd1 = df_val.dropna(subset=["_valor_unitario"]).drop_duplicates(subset=["_mat_key", "_lote_key"], keep="first")
        m1 = merged.merge(vd1[["_mat_key", "_lote_key", "_valor_unitario"]].rename(columns={"_valor_unitario": "_v1"}), on=["_mat_key", "_lote_key"], how="left")
        merged["_valor_unitario"] = m1["_v1"]

    mask = merged["_valor_unitario"].isna()
    if val_lote and mask.any():
        vd2 = df_val.dropna(subset=["_valor_unitario"]).drop_duplicates(subset=["_lote_key"], keep="first")
        m2 = merged.loc[mask].merge(vd2[["_lote_key", "_valor_unitario"]].rename(columns={"_valor_unitario": "_v2"}), on="_lote_key", how="left")
        merged.loc[mask, "_valor_unitario"] = m2["_v2"].values

    mask = merged["_valor_unitario"].isna()
    if val_mat and mask.any():
        vd3 = df_val.dropna(subset=["_valor_unitario"]).drop_duplicates(subset=["_mat_key"], keep="first")
        m3 = merged.loc[mask].merge(vd3[["_mat_key", "_valor_unitario"]].rename(columns={"_valor_unitario": "_v3"}), on="_mat_key", how="left")
        merged.loc[mask, "_valor_unitario"] = m3["_v3"].values

    merged["_cantidad"] = pd.to_numeric(merged[col_cant], errors="coerce").fillna(0)
    merged["Valor_Total"] = merged["_cantidad"] * merged["_valor_unitario"].fillna(0)
    return merged


if df_values_all is not None:
    merged = merge_values(df_stock, df_values_all, COL_MAT, COL_LOTE, COL_CANT, VAL_MAT, VAL_LOTE, VAL_VALOR)
else:
    merged = df_stock.copy()
    merged["_mat_key"] = merged[COL_MAT].astype(str).str.strip()
    merged["_lote_key"] = merged[COL_LOTE].astype(str).str.strip()
    merged["_cantidad"] = pd.to_numeric(merged[COL_CANT], errors="coerce").fillna(0)
    merged["_valor_unitario"] = np.nan
    merged["Valor_Total"] = 0

if df_control is not None:
    ctrl_set = set(df_control[CTRL_MAT].astype(str).str.strip().unique())
    merged["_es_control"] = merged["_mat_key"].isin(ctrl_set)
else:
    merged["_es_control"] = False


# =============================================================================
# 6. LINES WITHOUT VALUE (only when values loaded)
# =============================================================================
if df_values_all is not None:
    sin_valor = merged[merged["_valor_unitario"].isna()]
    n_sin = len(sin_valor)
    n_total = len(merged)
    pct_sin = n_sin / n_total * 100 if n_total > 0 else 0

    st.header("Cruce de valores unitarios")

    if n_sin == 0:
        st.success(f"**100%** de las líneas tienen valor unitario ({n_total:,} líneas)")
    else:
        st.warning(f"**{n_sin:,}** líneas sin valor unitario (**{pct_sin:.1f}%** de {n_total:,})")

        refs_sin = sin_valor[[COL_MAT, COL_LOTE]].drop_duplicates()
        with st.expander(f"Ver {len(refs_sin):,} refs/lotes sin valor"):
            st.dataframe(refs_sin, use_container_width=True, height=250)

        def refs_to_xl(df):
            buf = BytesIO()
            df.to_excel(buf, index=False, sheet_name="Sin valor")
            return buf.getvalue()

        st.download_button(
            f"Descargar {len(refs_sin):,} refs sin valor",
            refs_to_xl(refs_sin), file_name="refs_sin_valor.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if st.checkbox("Eliminar líneas sin valor unitario", key="excl_sin_val"):
            merged = merged[merged["_valor_unitario"].notna()].copy()
            st.info(f"Eliminadas {n_sin:,} líneas. Quedan **{len(merged):,}**.")


# =============================================================================
# 7. PREVIEW
# =============================================================================
with st.expander("Vista previa de datos"):
    tab_names = ["Stock"]
    tab_dfs = [df_stock]
    if df_values_all is not None:
        tab_names.append("Valores")
        tab_dfs.append(df_values_all)
    if df_control is not None:
        tab_names.append("Control")
        tab_dfs.append(df_control)
    tabs_prev = st.tabs(tab_names)
    for tab, df_prev in zip(tabs_prev, tab_dfs):
        with tab:
            st.dataframe(df_prev.head(10), use_container_width=True)


# =============================================================================
# 8. PARAMETERS
# =============================================================================
st.header("Configuración por segregación")
seed = st.sidebar.number_input("Semilla aleatoria", value=42, min_value=0, step=1, key="seg_seed")

# Defaults
n_ubic_alea = 10; min_lin_alea = 4
n_ubic_ctrl = 10; min_lin_ctrl = 4
pct_valor = 80; max_ubic_val = 30; min_lin_val = 4

param_cols = st.columns(len(tipos_seg))
for i, tipo in enumerate(tipos_seg):
    with param_cols[i]:
        if tipo == "Aleatorio":
            st.subheader("Aleatorio")
            n_ubic_alea = st.slider("Nº de ubicaciones", 3, 30, 10, key="n_u_a")
            min_lin_alea = st.slider("Mín. líneas/ubicación", 1, 20, 4, key="ml_a")
        elif tipo == "Control Diferenciado":
            st.subheader("Control Diferenciado")
            n_ubic_ctrl = st.slider("Nº de ubicaciones", 3, 30, 10, key="n_u_c")
            min_lin_ctrl = st.slider("Mín. líneas/ubicación", 1, 20, 4, key="ml_c")
        elif tipo == "Material Valioso":
            st.subheader("Material Valioso")
            pct_valor = st.slider("% del valor a cubrir", 10, 100, 80, step=5, key="pv")
            max_ubic_val = st.slider("Máx. ubicaciones", 5, 100, 30, key="mu_v")
            min_lin_val = st.slider("Mín. líneas/ubicación", 1, 20, 4, key="ml_v")


# =============================================================================
# 9. GENERATE
# =============================================================================
if st.button("Generar Segregaciones", type="primary", use_container_width=True):
    rng = np.random.default_rng(seed)

    def get_eligible(df, min_lines, filter_col=None):
        if filter_col:
            counts = df[df[filter_col]].groupby(COL_UBIC).size()
        else:
            counts = df.groupby(COL_UBIC).size()
        return counts[counts >= min_lines].index.tolist()

    def sample_locs(eligible, n):
        n = min(n, len(eligible))
        if n == 0:
            return []
        idx = rng.choice(len(eligible), size=n, replace=False)
        return [eligible[i] for i in idx]

    used = set()
    valor_total_almacen = merged["Valor_Total"].sum()

    top_val = []
    samp_ctrl = []
    samp_alea = []

    # 1) MATERIAL VALIOSO
    if "Material Valioso" in tipos_seg:
        elig_v = get_eligible(merged, min_lin_val)
        vbl = merged[merged[COL_UBIC].isin(elig_v)].groupby(COL_UBIC)["Valor_Total"].sum().sort_values(ascending=False)
        cs = vbl.cumsum()
        tgt = valor_total_almacen * (pct_valor / 100.0)
        top_val = cs[cs <= tgt].index.tolist()
        if len(top_val) < len(vbl):
            rem = [l for l in vbl.index if l not in top_val]
            if rem:
                top_val.append(rem[0])
        if len(top_val) > max_ubic_val:
            cap = vbl.head(max_ubic_val).index.tolist()
            cap_v = vbl.loc[cap].sum()
            pr = cap_v / valor_total_almacen * 100 if valor_total_almacen > 0 else 0
            st.error(f"Se necesitan **{len(top_val)}** ubicaciones para el {pct_valor}%. Con {max_ubic_val} se cubre el **{pr:.1f}%**.")
            st.stop()
        used.update(top_val)

    # 2) CONTROL DIFERENCIADO
    if "Control Diferenciado" in tipos_seg:
        elig_c = [l for l in get_eligible(merged, min_lin_ctrl, "_es_control") if l not in used]
        samp_ctrl = sample_locs(elig_c, n_ubic_ctrl)
        used.update(samp_ctrl)

    # 3) ALEATORIO
    if "Aleatorio" in tipos_seg:
        elig_g = [l for l in get_eligible(merged, min_lin_alea) if l not in used]
        samp_alea = sample_locs(elig_g, n_ubic_alea)
        used.update(samp_alea)

    today_str = date.today().strftime("%d-%m-%Y")

    def build_audit_df(seg_df, include_value=False, is_control=False):
        out = pd.DataFrame()
        out["Fecha"] = [today_str] * len(seg_df)
        out["Ref. centro"] = seg_df[COL_CENTRO].values if COL_CENTRO else ""
        out["Ref. Almacén"] = seg_df[COL_ALMACEN].values if COL_ALMACEN else ""
        out["Ubicacion"] = seg_df[COL_UBIC].values
        out["Ref. Material"] = seg_df[COL_MAT].values
        out["Descripción"] = seg_df[COL_DESC].values if COL_DESC else ""
        out["Nº Lote"] = seg_df[COL_LOTE].values
        if include_value:
            out["Valor unitario"] = seg_df["_valor_unitario"].values
            out["Valor total"] = seg_df["Valor_Total"].values
        out["Nº Serie"] = seg_df[COL_SERIE].values if COL_SERIE else ""
        out["Stock"] = seg_df[COL_CANT].values
        out["Cant. Física"] = np.nan
        out["Descuadre"] = np.nan
        out["Unidad Base"] = seg_df[COL_UB].values if COL_UB else ""
        out["Stock OK"] = seg_df[COL_SOK].values if COL_SOK else ""
        out["Stock Bloqueado"] = seg_df[COL_SBL].values if COL_SBL else 0
        out["Tipo Bloqueo"] = seg_df[COL_TBLOQ].values if COL_TBLOQ else ""
        if is_control:
            out["Fallo en el proceso"] = np.nan
            out["Obs. Inventario"] = np.nan
            out["Obs. Proceso"] = np.nan
        else:
            out["Observaciones Inventario"] = np.nan
        return out.sort_values("Ubicacion").reset_index(drop=True)

    # Build only selected segregations
    seg_alea_fmt = seg_ctrl_fmt = seg_val_fmt = pd.DataFrame()

    if "Aleatorio" in tipos_seg:
        seg_alea = merged[merged[COL_UBIC].isin(samp_alea)].copy()
        seg_alea_fmt = build_audit_df(seg_alea, include_value=True)
    if "Control Diferenciado" in tipos_seg:
        seg_ctrl = merged[merged[COL_UBIC].isin(samp_ctrl)].copy()
        seg_ctrl_fmt = build_audit_df(seg_ctrl, is_control=True)
    if "Material Valioso" in tipos_seg:
        seg_val = merged[merged[COL_UBIC].isin(top_val)].copy()
        seg_val_fmt = build_audit_df(seg_val, include_value=True)

    # === DISPLAY ===
    st.header("Resultados")
    st.markdown(f"**{len(used)} ubicaciones** en total (sin repeticiones)")

    tab_labels = []
    tab_items = []

    if "Aleatorio" in tipos_seg:
        tab_labels.append(f"Aleatorio ({len(samp_alea)})")
        tab_items.append(("alea", samp_alea, seg_alea_fmt, None))
    if "Control Diferenciado" in tipos_seg:
        tab_labels.append(f"Control Diferenciado ({len(samp_ctrl)})")
        seg_ctrl_raw = merged[merged[COL_UBIC].isin(samp_ctrl)] if samp_ctrl else pd.DataFrame()
        tab_items.append(("ctrl", samp_ctrl, seg_ctrl_fmt, seg_ctrl_raw))
    if "Material Valioso" in tipos_seg:
        tab_labels.append(f"Material Valioso ({len(top_val)})")
        seg_val_raw = merged[merged[COL_UBIC].isin(top_val)] if top_val else pd.DataFrame()
        tab_items.append(("val", top_val, seg_val_fmt, seg_val_raw))

    tabs = st.tabs(tab_labels)
    for tab, (kind, locs, fmt_df, raw_df) in zip(tabs, tab_items):
        with tab:
            if kind == "alea":
                st.markdown(f"**{len(locs)}** ubic. · **{len(fmt_df):,}** líneas")
            elif kind == "ctrl":
                nc = raw_df["_es_control"].sum() if raw_df is not None and not raw_df.empty else 0
                st.markdown(f"**{len(locs)}** ubic. · **{len(fmt_df):,}** líneas · Control: **{nc:,}**")
                st.caption("Columnas operarios: **Cant. Física**, **Fallo en el proceso**, **Obs. Inventario**, **Obs. Proceso**")
            elif kind == "val":
                val_cubierto = raw_df["Valor_Total"].sum() if raw_df is not None and not raw_df.empty else 0
                pct_cub = (val_cubierto / valor_total_almacen * 100) if valor_total_almacen > 0 else 0
                st.markdown(f"**{len(locs)}** ubic. · **{len(fmt_df):,}** líneas · Valor: **{val_cubierto:,.2f}** ({pct_cub:.1f}%)")

            if not fmt_df.empty:
                st.dataframe(fmt_df, use_container_width=True, height=500)
            else:
                st.warning("Sin datos")

    # === DOWNLOAD ===
    st.header("Descargar")

    def to_formatted_excel(sheets, edit_map):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            for sname, df in sheets.items():
                safe = sname[:31]
                df.to_excel(writer, sheet_name=safe, index=False)
                wb = writer.book
                ws = writer.sheets[safe]

                hdr_blue = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
                hdr_yellow = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#FFD966", "font_color": "#1F3864", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
                cell_normal = wb.add_format({"font_size": 9, "border": 1, "valign": "vcenter"})
                cell_yellow = wb.add_format({"font_size": 10, "border": 1, "valign": "vcenter", "bg_color": "#FFF2CC"})
                cell_yellow_txt = wb.add_format({"font_size": 10, "border": 1, "valign": "vcenter", "bg_color": "#FFF2CC", "text_wrap": True})
                cell_money = wb.add_format({"font_size": 9, "border": 1, "num_format": "#,##0.00"})

                editable = edit_map.get(sname, [])

                ws.set_row(0, 30)
                for ci, cn in enumerate(df.columns):
                    ws.write(0, ci, cn, hdr_yellow if cn in editable else hdr_blue)

                # Detect columns for Descuadre formula
                _col_list = list(df.columns)
                _stock_lt = xl_col_to_name(_col_list.index("Stock")) if "Stock" in _col_list else None
                _fisica_lt = xl_col_to_name(_col_list.index("Cant. Física")) if "Cant. Física" in _col_list else None

                for ri in range(len(df)):
                    for ci, cn in enumerate(df.columns):
                        val = df.iloc[ri, ci]
                        is_edit = cn in editable
                        # Descuadre → Excel formula
                        if cn == "Descuadre" and _stock_lt and _fisica_lt:
                            er = ri + 2
                            formula = f'=IF({_fisica_lt}{er}="","",' \
                                      f'{_fisica_lt}{er}-{_stock_lt}{er})'
                            ws.write_formula(ri + 1, ci, formula, cell_yellow)
                        elif pd.isna(val):
                            ws.write_blank(ri + 1, ci, "", cell_yellow if is_edit else cell_normal)
                        elif is_edit:
                            if isinstance(val, (int, float)):
                                ws.write_number(ri + 1, ci, val, cell_yellow)
                            else:
                                ws.write(ri + 1, ci, val, cell_yellow_txt)
                        elif cn in ("Valor unitario", "Valor total"):
                            ws.write_number(ri + 1, ci, float(val) if not pd.isna(val) else 0, cell_money)
                        else:
                            ws.write(ri + 1, ci, val, cell_normal)

                for ci, cn in enumerate(df.columns):
                    mx = max(df[cn].astype(str).map(len).max(), len(str(cn)))
                    ws.set_column(ci, ci, min(mx + 3, 40))
                ws.freeze_panes(1, 0)
        return output.getvalue()

    edit_map = {
        "material  control diferenc": ["Cant. Física", "Descuadre", "Fallo en el proceso", "Obs. Inventario", "Obs. Proceso"],
        "material Valioso": ["Cant. Física", "Descuadre", "Observaciones Inventario"],
        "material aleatorio": ["Cant. Física", "Descuadre", "Observaciones Inventario"],
    }

    combined = {}
    if "Control Diferenciado" in tipos_seg and not seg_ctrl_fmt.empty:
        combined["material  control diferenc"] = seg_ctrl_fmt
    if "Material Valioso" in tipos_seg and not seg_val_fmt.empty:
        combined["material Valioso"] = seg_val_fmt
    if "Aleatorio" in tipos_seg and not seg_alea_fmt.empty:
        combined["material aleatorio"] = seg_alea_fmt

    if combined:
        st.download_button(
            "Descargar Auditoría completa (Excel)",
            to_formatted_excel(combined, edit_map),
            file_name=f"AUDITORIA_INTERNA_{date.today().strftime('%B_%Y').upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )
