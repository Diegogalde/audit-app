"""
5_historiales.py
================
Gestión centralizada de historiales de todas las secciones de la app.
Permite filtrar por centro, subir entradas y eliminar individualmente.
"""
import re
import streamlit as st
import pandas as pd
import json
from io import BytesIO
from pathlib import Path
from datetime import date

st.title("Historiales")
st.markdown("Gestiona los historiales de todas las secciones de la aplicación.")

CENTROS_DISPONIBLES = ["Noain", "Post-Venta", "Export-OTC", "Arazuri"]

MONTH_NAMES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}
_MONTH_NAME_TO_NUM = {v.upper(): k for k, v in MONTH_NAMES.items()}

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"
SEG_HISTORY_FILE = _DATA_DIR / "audit_history.json"
CONSOL_HISTORY_FILE = _DATA_DIR / "consolidado_historico.json"
ABS_HISTORY_FILE = _DATA_DIR / "absentismo_historico.json"
RECEP_HISTORY_FILE = _DATA_DIR / "recepciones_historico.json"


# =============================================================================
# Load / Save helpers
# =============================================================================
def _load_json(path):
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def _save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =============================================================================
# Parse helpers
# =============================================================================
def _find_col(columns, candidates):
    cols_lower = {str(c).lower().strip(): c for c in columns}
    for cand in candidates:
        cl = cand.lower().strip()
        if cl in cols_lower:
            return cols_lower[cl]
        for k, v in cols_lower.items():
            if cl in k or k in cl:
                return v
    return None


def _parse_seg_excel(data_bytes):
    """Parse a segregation Excel to extract ubicaciones per type."""
    try:
        xls = pd.ExcelFile(BytesIO(data_bytes))
    except Exception:
        return None

    val_ubics = []
    ctrl_ubics = []
    alea_ubics = []

    for name in xls.sheet_names:
        nl = name.lower().strip()
        if nl == "instrucciones":
            continue
        df = pd.read_excel(xls, sheet_name=name)
        col_ubic = _find_col(df.columns, ["Ubicacion", "Ubicación"])
        if not col_ubic:
            continue
        ubics = [str(u) for u in df[col_ubic].dropna().unique()]
        if "control" in nl or "diferenc" in nl:
            ctrl_ubics = ubics
        elif "valios" in nl:
            val_ubics = ubics
        elif "aleatori" in nl:
            alea_ubics = ubics

    if not val_ubics and not ctrl_ubics and not alea_ubics:
        return None

    fecha = date.today().isoformat()
    for name in xls.sheet_names:
        if name.lower().strip() == "instrucciones":
            continue
        df = pd.read_excel(xls, sheet_name=name)
        col_fecha = _find_col(df.columns, ["Fecha"])
        if col_fecha and not df[col_fecha].dropna().empty:
            raw = str(df[col_fecha].dropna().iloc[0]).strip()
            if len(raw) >= 10:
                if "-" in raw and raw[2] == "-":
                    parts = raw[:10].split("-")
                    if len(parts) == 3:
                        fecha = f"{parts[2]}-{parts[1]}-{parts[0]}"
                elif "-" in raw and raw[4] == "-":
                    fecha = raw[:10]
            break

    return {
        "fecha": fecha,
        "centro": None,
        "valioso_ubicaciones": val_ubics,
        "control_ubicaciones": ctrl_ubics,
        "aleatorio_ubicaciones": alea_ubics,
    }


def _parse_consol_report(data_bytes):
    """Parse an exported consolidado Excel back into a history entry."""
    try:
        df = pd.read_excel(BytesIO(data_bytes), sheet_name=0, header=None)
    except Exception:
        return None
    title = str(df.iloc[0, 0]) if not pd.isna(df.iloc[0, 0]) else ""
    m = re.search(r"(\d{2}/\d{2}/\d{4})", title)
    if not m:
        return None
    fecha_str = m.group(1)
    parts = fecha_str.split("/")
    fecha_iso = f"{parts[2]}-{parts[1]}-{parts[0]}"
    tu = tl = te = t_uds = t_uds_e = 0
    tf = tf_uds = tf_pl = tf_pu = 1.0
    t_perd = t_pot = 0.0
    for ri in range(2, min(10, len(df))):
        v1 = df.iloc[ri, 1]
        if isinstance(v1, (int, float)) and not pd.isna(v1) and v1 > 0:
            tu = int(v1) if not pd.isna(df.iloc[ri, 1]) else 0
            tl = int(df.iloc[ri, 2]) if not pd.isna(df.iloc[ri, 2]) else 0
            t_uds = int(df.iloc[ri, 3]) if not pd.isna(df.iloc[ri, 3]) else 0
            te = int(df.iloc[ri, 4]) if not pd.isna(df.iloc[ri, 4]) else 0
            t_uds_e = int(df.iloc[ri, 5]) if not pd.isna(df.iloc[ri, 5]) else 0
            tf = float(df.iloc[ri, 6]) if not pd.isna(df.iloc[ri, 6]) else 1.0
            tf_uds = float(df.iloc[ri, 7]) if not pd.isna(df.iloc[ri, 7]) else 1.0
            tf_pl = float(df.iloc[ri, 8]) if not pd.isna(df.iloc[ri, 8]) else 1.0
            tf_pu = float(df.iloc[ri, 9]) if not pd.isna(df.iloc[ri, 9]) else 1.0
            t_perd = float(df.iloc[ri, 11]) if df.shape[1] > 11 and not pd.isna(df.iloc[ri, 11]) else 0.0
            t_pot = float(df.iloc[ri, 12]) if df.shape[1] > 12 and not pd.isna(df.iloc[ri, 12]) else 0.0
            break
    key = f"{fecha_iso}_imported"
    return {
        "key": key,
        "fecha": fecha_iso,
        "centro": None,
        "total": {
            "ubicaciones": tu, "lotes_total": tl, "lotes_error": te,
            "uds_inventariadas": float(t_uds), "uds_erroneas": float(t_uds_e),
            "fiabilidad_lotes": tf, "fiabilidad_uds": tf_uds,
            "perdida_monetaria": t_perd,
            "fiab_proc_lotes": tf_pl, "fiab_proc_uds": tf_pu,
            "potencial_perdida": t_pot,
        },
        "secciones": {},
    }


def _parse_abs_report(data_bytes):
    """Parse an exported absentismo Excel report back into a history entry."""
    try:
        df = pd.read_excel(BytesIO(data_bytes), sheet_name="Resumen", header=None)
    except Exception:
        return None

    title = str(df.iloc[0, 0]) if not pd.isna(df.iloc[0, 0]) else ""
    m = re.search(r"—\s*(\w+)\s+(\d{4})", title)
    if not m:
        return None
    month_name, year_str = m.group(1).upper(), m.group(2)
    month_num = _MONTH_NAME_TO_NUM.get(month_name)
    if not month_num:
        return None
    year = int(year_str)

    header_row = df.iloc[2]
    centro_names = []
    for ci in range(1, len(header_row)):
        val = str(header_row.iloc[ci]).strip() if not pd.isna(header_row.iloc[ci]) else ""
        if val and val.upper() != "TOTAL":
            centro_names.append((ci, val))

    if not centro_names:
        return None

    metric_keys = [
        "plantilla", "plantilla_efectiva", "dias_laborables", "dias_trabajados",
        "dias_vacaciones", "dias_baja", "dias_ap", "dias_permiso", "dias_excedencia",
        "total_ausencias_con_vac", "total_ausencias_sin_vac",
    ]

    centros = []
    for ci, cname in centro_names:
        centro_data = {"centro": cname}
        for ri, key in enumerate(metric_keys):
            val = df.iloc[3 + ri, ci]
            centro_data[key] = float(val) if not pd.isna(val) else 0
        pct_row_con = 3 + len(metric_keys) + 1
        pct_row_sin = pct_row_con + 1
        pct_con_val = df.iloc[pct_row_con, ci] if pct_row_con < len(df) else 0
        pct_sin_val = df.iloc[pct_row_sin, ci] if pct_row_sin < len(df) else 0
        pct_con = float(pct_con_val) if not pd.isna(pct_con_val) else 0
        pct_sin = float(pct_sin_val) if not pd.isna(pct_sin_val) else 0
        if pct_con > 1:
            pct_con = pct_con / 100
        if pct_sin > 1:
            pct_sin = pct_sin / 100
        centro_data["pct_con_vac"] = round(pct_con * 100, 2)
        centro_data["pct_sin_vac"] = round(pct_sin * 100, 2)
        centros.append(centro_data)

    total_plantilla = sum(c.get("plantilla", 0) for c in centros)
    total_pct_con = sum(c.get("pct_con_vac", 0) * c.get("plantilla", 0) for c in centros)
    total_pct_sin = sum(c.get("pct_sin_vac", 0) * c.get("plantilla", 0) for c in centros)
    if total_plantilla > 0:
        total_pct_con /= total_plantilla
        total_pct_sin /= total_plantilla

    key = f"{year}-{month_num:02d}"
    return {
        "key": key,
        "year": year,
        "month": month_num,
        "centros": centros,
        "total_plantilla": total_plantilla,
        "total_pct_con": round(total_pct_con, 2),
        "total_pct_sin": round(total_pct_sin, 2),
    }


def _parse_recep_excel(data_bytes):
    """Parse a recepciones BIP Excel to extract summary info."""
    try:
        df = pd.read_excel(BytesIO(data_bytes), sheet_name=0)
    except Exception:
        return None

    col_fecha = _find_col(df.columns, ["Fecha Recepcion", "Fecha recepción", "Fecha"])
    col_proveedor = _find_col(df.columns, ["Proveedor", "Nombre Proveedor"])
    col_pedido = _find_col(df.columns, ["Pedido", "Nº Pedido", "Num Pedido"])

    if col_fecha is None and col_proveedor is None:
        return None

    n_lineas = len(df)
    n_proveedores = df[col_proveedor].nunique() if col_proveedor else 0
    n_pedidos = df[col_pedido].nunique() if col_pedido else 0

    fecha_min = fecha_max = None
    if col_fecha:
        fechas = pd.to_datetime(df[col_fecha], errors="coerce").dropna()
        if not fechas.empty:
            fecha_min = fechas.min().strftime("%Y-%m-%d")
            fecha_max = fechas.max().strftime("%Y-%m-%d")

    return {
        "key": f"{fecha_min or date.today().isoformat()}_recep",
        "fecha_desde": fecha_min or date.today().isoformat(),
        "fecha_hasta": fecha_max or date.today().isoformat(),
        "centro": None,
        "n_lineas": n_lineas,
        "n_proveedores": n_proveedores,
        "n_pedidos": n_pedidos,
    }


# =============================================================================
# CENTRO FILTER
# =============================================================================
centro_filter = st.selectbox(
    "Filtrar por centro",
    options=["Todos"] + CENTROS_DISPONIBLES,
    key="hist_centro_filter",
)


def _matches_centro(entry, centro):
    if centro == "Todos":
        return True
    return entry.get("centro") == centro or not entry.get("centro")


# =============================================================================
# CATEGORY SELECTOR
# =============================================================================
categoria = st.radio(
    "Sección",
    ["Inventarios", "Absentismo", "Recepciones"],
    horizontal=True,
    key="hist_categoria",
)

# =============================================================================
# INVENTARIOS — sub-tabs: Segregaciones + Reportes Consolidados
# =============================================================================
if categoria == "Inventarios":
    tab_seg, tab_consol = st.tabs(["Segregaciones", "Reportes Consolidados"])

    # --- SEGREGACIONES ---
    with tab_seg:
        seg_all = _load_json(SEG_HISTORY_FILE)
        seg_filtered = [e for e in seg_all if _matches_centro(e, centro_filter)]

        st.subheader(f"Segregaciones ({len(seg_filtered)} de {len(seg_all)})")

        if seg_filtered:
            seg_rows = []
            for e in seg_filtered:
                seg_rows.append({
                    "Fecha": e.get("fecha", ""),
                    "Centro": e.get("centro", "—"),
                    "Valioso": len(e.get("valioso_ubicaciones", [])),
                    "Control": len(e.get("control_ubicaciones", [])),
                    "Aleatorio": len(e.get("aleatorio_ubicaciones", [])),
                })
            st.dataframe(pd.DataFrame(seg_rows), use_container_width=True, hide_index=True)
        else:
            st.info("No hay segregaciones en el historial.")

        # Upload
        st.divider()
        st.markdown("**Subir segregación**")
        st.caption("Sube un Excel de auditoría (generado por esta app) para añadir sus ubicaciones al historial.")
        seg_upload = st.file_uploader("Excel de auditoría", type=["xlsx", "xls"], key="hist_seg_upload")
        if seg_upload:
            parsed = _parse_seg_excel(seg_upload.getvalue())
            if parsed:
                n_v = len(parsed.get("valioso_ubicaciones", []))
                n_c = len(parsed.get("control_ubicaciones", []))
                n_a = len(parsed.get("aleatorio_ubicaciones", []))
                st.success(f"Detectado: **{parsed['fecha']}** — Valioso: {n_v}, Control: {n_c}, Aleatorio: {n_a}")
                parsed_centro = st.selectbox("Centro para esta entrada", CENTROS_DISPONIBLES, key="seg_upload_centro")
                parsed["centro"] = parsed_centro
                if st.button("Guardar en historial", key="btn_save_seg", type="primary"):
                    _h = _load_json(SEG_HISTORY_FILE)
                    _h.append(parsed)
                    _save_json(SEG_HISTORY_FILE, _h)
                    st.success("Guardado")
                    st.rerun()
            else:
                st.error("No se pudo leer el Excel. Asegúrate de que tiene pestañas de segregación.")

        # Individual delete
        if seg_filtered:
            st.divider()
            st.markdown("**Eliminar entradas**")
            for i, entry in enumerate(seg_filtered):
                orig_idx = seg_all.index(entry)
                c1, c2 = st.columns([4, 1])
                n_v = len(entry.get("valioso_ubicaciones", []))
                n_c = len(entry.get("control_ubicaciones", []))
                n_a = len(entry.get("aleatorio_ubicaciones", []))
                c_label = entry.get("centro", "Sin centro")
                c1.text(f"{entry.get('fecha', '?')} [{c_label}] — V:{n_v} C:{n_c} A:{n_a}")
                if c2.button("Eliminar", key=f"del_seg_{i}", type="secondary"):
                    _h = _load_json(SEG_HISTORY_FILE)
                    if orig_idx < len(_h):
                        _h.pop(orig_idx)
                    _save_json(SEG_HISTORY_FILE, _h)
                    st.rerun()

            st.divider()
            if st.button("Limpiar todo el historial de segregaciones", key="seg_clear_all", type="secondary"):
                _save_json(SEG_HISTORY_FILE, [])
                st.rerun()

    # --- REPORTES CONSOLIDADOS ---
    with tab_consol:
        consol_all = _load_json(CONSOL_HISTORY_FILE)
        consol_filtered = [e for e in consol_all if _matches_centro(e, centro_filter)]

        st.subheader(f"Reportes Consolidados ({len(consol_filtered)} de {len(consol_all)})")

        if consol_filtered:
            consol_rows = []
            for h in consol_filtered:
                t = h.get("total", {})
                consol_rows.append({
                    "Fecha": h.get("fecha", ""),
                    "Centro": h.get("centro", "—"),
                    "Ubicaciones": t.get("ubicaciones", 0),
                    "Lotes": t.get("lotes_total", 0),
                    "Lotes err.": t.get("lotes_error", 0),
                    "Fiab. lotes": f"{t.get('fiabilidad_lotes', 1):.2%}",
                    "Fiab. uds.": f"{t.get('fiabilidad_uds', 1):.2%}",
                    "Pérdida (€)": f"{t.get('perdida_monetaria', 0):,.2f}",
                    "Pot. pérdida (€)": f"{t.get('potencial_perdida', 0):,.2f}",
                })
            st.dataframe(pd.DataFrame(consol_rows), use_container_width=True, hide_index=True)
        else:
            st.info("No hay reportes consolidados en el historial.")

        # Upload
        st.divider()
        st.markdown("**Subir reporte pasado**")
        st.caption("Sube un Excel de reporte consolidado (generado por esta app) para añadirlo al historial.")
        consol_upload = st.file_uploader("Excel de reporte", type=["xlsx", "xls"], key="hist_consol_upload")
        if consol_upload:
            parsed = _parse_consol_report(consol_upload.getvalue())
            if parsed:
                st.success(f"Detectado: **{parsed['fecha']}** — "
                           f"{parsed['total'].get('ubicaciones', 0)} ubic.")
                parsed_centro = st.selectbox("Centro para esta entrada", CENTROS_DISPONIBLES, key="consol_upload_centro")
                parsed["centro"] = parsed_centro
                parsed["key"] = f"{parsed['fecha']}_{parsed_centro}"
                if st.button("Guardar en historial", key="btn_save_consol", type="primary"):
                    _h = _load_json(CONSOL_HISTORY_FILE)
                    _h = [x for x in _h if x.get("key") != parsed["key"]]
                    _h.append(parsed)
                    _h.sort(key=lambda x: x["key"])
                    _save_json(CONSOL_HISTORY_FILE, _h)
                    st.success("Guardado")
                    st.rerun()
            else:
                st.error("No se pudo leer el reporte. Asegúrate de que tiene el formato correcto.")

        # Individual delete
        if consol_filtered:
            st.divider()
            st.markdown("**Eliminar reportes**")
            for i, h in enumerate(consol_filtered):
                c1, c2 = st.columns([4, 1])
                t = h.get("total", {})
                c1.text(f"{h.get('fecha', '')} — {h.get('centro', 'Sin centro')} — "
                        f"{t.get('ubicaciones', 0)} ubic. — Fiab: {t.get('fiabilidad_lotes', 1):.2%}")
                if c2.button("Eliminar", key=f"del_consol_{i}", type="secondary"):
                    _h = _load_json(CONSOL_HISTORY_FILE)
                    _h = [x for x in _h if x.get("key") != h["key"]]
                    _save_json(CONSOL_HISTORY_FILE, _h)
                    st.rerun()

            st.divider()
            if st.button("Limpiar todo el historial de reportes", key="consol_clear_all", type="secondary"):
                _save_json(CONSOL_HISTORY_FILE, [])
                st.rerun()


# =============================================================================
# ABSENTISMO
# =============================================================================
elif categoria == "Absentismo":
    abs_all = _load_json(ABS_HISTORY_FILE)

    st.subheader(f"Historial de Absentismo ({len(abs_all)} mes(es))")

    if abs_all:
        abs_rows = []
        for h in abs_all:
            month_label = f"{MONTH_NAMES.get(h.get('month', 0), '?')} {h.get('year', '')}"
            abs_rows.append({
                "Mes": month_label,
                "Plantilla": h.get("total_plantilla", 0),
                "% Con Vac.": f"{h.get('total_pct_con', 0):.2f}%",
                "% Sin Vac.": f"{h.get('total_pct_sin', 0):.2f}%",
                "Centros": len(h.get("centros", [])),
            })
        st.dataframe(pd.DataFrame(abs_rows), use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos de absentismo en el historial.")

    # Upload
    st.divider()
    st.markdown("**Subir reporte de absentismo**")
    st.caption("Sube un Excel de reporte de absentismo (generado por esta app) para añadirlo al historial.")
    abs_upload = st.file_uploader("Excel de absentismo", type=["xlsx", "xls"], key="hist_abs_upload")
    if abs_upload:
        parsed = _parse_abs_report(abs_upload.getvalue())
        if parsed:
            month_label = f"{MONTH_NAMES[parsed['month']]} {parsed['year']}"
            st.success(f"Detectado: **{month_label}** — {len(parsed['centros'])} centro(s)")
            if st.button(f"Guardar {month_label} en historial", key="btn_save_abs", type="primary"):
                _h = _load_json(ABS_HISTORY_FILE)
                _h = [x for x in _h if x.get("key") != parsed["key"]]
                _h.append(parsed)
                _h.sort(key=lambda x: x["key"])
                _save_json(ABS_HISTORY_FILE, _h)
                st.success(f"**{month_label}** guardado en historial")
                st.rerun()
        else:
            st.error("No se pudo leer el reporte. Asegúrate de que tiene la hoja 'Resumen' con el formato correcto.")

    # Individual delete
    if abs_all:
        st.divider()
        st.markdown("**Eliminar meses**")
        for h in abs_all:
            c1, c2 = st.columns([4, 1])
            month_label = f"{MONTH_NAMES.get(h.get('month', 0), '?')} {h.get('year', '')}"
            c1.text(f"{month_label} — Plantilla: {h.get('total_plantilla', 0)} — "
                    f"Con vac: {h.get('total_pct_con', 0):.2f}% — Sin vac: {h.get('total_pct_sin', 0):.2f}%")
            if c2.button("Eliminar", key=f"del_abs_{h.get('key', '')}", type="secondary"):
                _h = _load_json(ABS_HISTORY_FILE)
                _h = [x for x in _h if x.get("key") != h["key"]]
                _save_json(ABS_HISTORY_FILE, _h)
                st.rerun()

        st.divider()
        if st.button("Limpiar todo el historial de absentismo", key="abs_clear_all", type="secondary"):
            _save_json(ABS_HISTORY_FILE, [])
            st.rerun()


# =============================================================================
# RECEPCIONES
# =============================================================================
elif categoria == "Recepciones":
    recep_all = _load_json(RECEP_HISTORY_FILE)
    recep_filtered = [e for e in recep_all if _matches_centro(e, centro_filter)]

    st.subheader(f"Historial de Recepciones ({len(recep_filtered)} de {len(recep_all)})")

    if recep_filtered:
        recep_rows = []
        for h in recep_filtered:
            recep_rows.append({
                "Desde": h.get("fecha_desde", ""),
                "Hasta": h.get("fecha_hasta", ""),
                "Centro": h.get("centro", "—"),
                "Líneas": h.get("n_lineas", 0),
                "Proveedores": h.get("n_proveedores", 0),
                "Pedidos": h.get("n_pedidos", 0),
            })
        st.dataframe(pd.DataFrame(recep_rows), use_container_width=True, hide_index=True)
    else:
        st.info("No hay recepciones en el historial.")

    # Upload
    st.divider()
    st.markdown("**Subir extracto de recepciones**")
    st.caption("Sube un Excel de recepciones BIP para añadirlo al historial.")
    recep_upload = st.file_uploader("Excel de recepciones", type=["xlsx", "xls"], key="hist_recep_upload")
    if recep_upload:
        parsed = _parse_recep_excel(recep_upload.getvalue())
        if parsed:
            st.success(f"Detectado: **{parsed['fecha_desde']}** a **{parsed['fecha_hasta']}** — "
                       f"{parsed['n_lineas']} líneas, {parsed['n_proveedores']} proveedores")
            parsed_centro = st.selectbox("Centro para esta entrada", CENTROS_DISPONIBLES, key="recep_upload_centro")
            parsed["centro"] = parsed_centro
            parsed["key"] = f"{parsed['fecha_desde']}_{parsed_centro}_recep"
            if st.button("Guardar en historial", key="btn_save_recep", type="primary"):
                _h = _load_json(RECEP_HISTORY_FILE)
                _h = [x for x in _h if x.get("key") != parsed["key"]]
                _h.append(parsed)
                _h.sort(key=lambda x: x["key"])
                _save_json(RECEP_HISTORY_FILE, _h)
                st.success("Guardado")
                st.rerun()
        else:
            st.error("No se pudo leer el Excel. Asegúrate de que tiene columnas de recepciones (Fecha, Proveedor, etc.).")

    # Individual delete
    if recep_filtered:
        st.divider()
        st.markdown("**Eliminar entradas**")
        for i, h in enumerate(recep_filtered):
            c1, c2 = st.columns([4, 1])
            c1.text(f"{h.get('fecha_desde', '')} → {h.get('fecha_hasta', '')} — "
                    f"{h.get('centro', 'Sin centro')} — {h.get('n_lineas', 0)} líneas")
            if c2.button("Eliminar", key=f"del_recep_{i}", type="secondary"):
                _h = _load_json(RECEP_HISTORY_FILE)
                _h = [x for x in _h if x.get("key") != h["key"]]
                _save_json(RECEP_HISTORY_FILE, _h)
                st.rerun()

        st.divider()
        if st.button("Limpiar todo el historial de recepciones", key="recep_clear_all", type="secondary"):
            _save_json(RECEP_HISTORY_FILE, [])
            st.rerun()
