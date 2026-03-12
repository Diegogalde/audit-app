"""
5_historiales.py
================
Gestión centralizada de historiales: segregaciones, reportes consolidados.
Permite filtrar por centro, subir entradas y eliminar individualmente.
"""
import streamlit as st
import pandas as pd
import json
from io import BytesIO
from pathlib import Path
from datetime import date

st.title("Historiales")
st.markdown("Gestiona los historiales de segregaciones y reportes consolidados.")

CENTROS_DISPONIBLES = ["Noain", "Post-Venta", "Export-OTC", "Arazuri"]

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"
SEG_HISTORY_FILE = _DATA_DIR / "audit_history.json"
CONSOL_HISTORY_FILE = _DATA_DIR / "consolidado_historico.json"


# =============================================================================
# Load / Save helpers
# =============================================================================
def load_seg_history():
    if SEG_HISTORY_FILE.exists():
        with open(SEG_HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_seg_history(entries):
    SEG_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(SEG_HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


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
# Parse helpers for uploads
# =============================================================================
def _parse_seg_excel(data_bytes):
    """Parse a segregation Excel to extract ubicaciones per type."""
    try:
        xls = pd.ExcelFile(BytesIO(data_bytes))
    except Exception:
        return None

    val_ubics = []
    ctrl_ubics = []
    alea_ubics = []

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

    # Try to get date from data
    fecha = date.today().isoformat()
    for name in xls.sheet_names:
        if name.lower().strip() == "instrucciones":
            continue
        df = pd.read_excel(xls, sheet_name=name)
        col_fecha = _find_col(df.columns, ["Fecha"])
        if col_fecha and not df[col_fecha].dropna().empty:
            raw = str(df[col_fecha].dropna().iloc[0]).strip()
            # Try dd-mm-yyyy or yyyy-mm-dd
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
    import re as _re
    try:
        df = pd.read_excel(BytesIO(data_bytes), sheet_name=0, header=None)
    except Exception:
        return None
    title = str(df.iloc[0, 0]) if not pd.isna(df.iloc[0, 0]) else ""
    m = _re.search(r"(\d{2}/\d{2}/\d{4})", title)
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
# TAB 1: SEGREGACIONES
# =============================================================================
tab_seg, tab_consol = st.tabs(["Segregaciones", "Reportes Consolidados"])

with tab_seg:
    seg_all = load_seg_history()
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

            # Let user assign centro
            parsed_centro = st.selectbox("Centro para esta entrada", CENTROS_DISPONIBLES, key="seg_upload_centro")
            parsed["centro"] = parsed_centro

            if st.button("Guardar en historial", key="btn_save_seg", type="primary"):
                _h = load_seg_history()
                _h.append(parsed)
                save_seg_history(_h)
                st.success("Guardado")
                st.rerun()
        else:
            st.error("No se pudo leer el Excel. Asegúrate de que tiene pestañas de segregación.")

    # Individual delete
    if seg_filtered:
        st.divider()
        st.markdown("**Eliminar entradas**")
        for i, entry in enumerate(seg_filtered):
            # Find original index
            orig_idx = seg_all.index(entry)
            c1, c2 = st.columns([4, 1])
            n_v = len(entry.get("valioso_ubicaciones", []))
            n_c = len(entry.get("control_ubicaciones", []))
            n_a = len(entry.get("aleatorio_ubicaciones", []))
            c_label = entry.get("centro", "Sin centro")
            c1.text(f"{entry.get('fecha', '?')} [{c_label}] — V:{n_v} C:{n_c} A:{n_a}")
            if c2.button("Eliminar", key=f"del_seg_{i}", type="secondary"):
                _h = load_seg_history()
                if orig_idx < len(_h):
                    _h.pop(orig_idx)
                save_seg_history(_h)
                st.rerun()

        st.divider()
        if st.button("Limpiar todo el historial de segregaciones", key="seg_clear_all", type="secondary"):
            save_seg_history([])
            st.rerun()


# =============================================================================
# TAB 2: REPORTES CONSOLIDADOS
# =============================================================================
with tab_consol:
    consol_all = load_consol_history()
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
            # Update key with centro
            parsed["key"] = f"{parsed['fecha']}_{parsed_centro}"

            if st.button("Guardar en historial", key="btn_save_consol", type="primary"):
                _h = load_consol_history()
                _h = [x for x in _h if x.get("key") != parsed["key"]]
                _h.append(parsed)
                _h.sort(key=lambda x: x["key"])
                save_consol_history(_h)
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
                _h = load_consol_history()
                _h = [x for x in _h if x.get("key") != h["key"]]
                save_consol_history(_h)
                st.rerun()

        st.divider()
        if st.button("Limpiar todo el historial de reportes", key="consol_clear_all", type="secondary"):
            save_consol_history([])
            st.rerun()
