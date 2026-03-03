"""
4_absentismo.py
===============
Análisis de absentismo por centro de trabajo.
Sube un Excel mensual por centro (cuadrante de horas) y obtiene KPIs,
validaciones y un reporte consolidado descargable.
"""
import re
import json
import streamlit as st
import pandas as pd
import numpy as np
import calendar
import plotly.graph_objects as go
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path

st.title("Análisis de Absentismo")
st.markdown("Sube los cuadrantes mensuales de horas (un Excel por centro) y genera el análisis.")

from metodologia import render_download as _render_metodologia
_render_metodologia("absentismo")

SS = st.session_state

MONTH_NAMES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# =============================================================================
# 1. HOLIDAYS — Navarra (national + regional)
# =============================================================================
def _easter_date(year):
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def get_default_holidays(year):
    """National Spanish + Navarra regional holidays."""
    easter = _easter_date(year)
    good_friday = easter - timedelta(days=2)
    holy_thursday = easter - timedelta(days=3)
    easter_monday = easter + timedelta(days=1)
    return {
        date(year, 1, 1): "Año Nuevo",
        date(year, 1, 6): "Epifanía",
        holy_thursday: "Jueves Santo",
        good_friday: "Viernes Santo",
        easter_monday: "Lunes de Pascua",
        date(year, 5, 1): "Día del Trabajo",
        date(year, 8, 15): "Asunción",
        date(year, 10, 12): "Día de la Hispanidad",
        date(year, 11, 1): "Todos los Santos",
        date(year, 12, 3): "San Francisco Javier (Navarra)",
        date(year, 12, 6): "Día de la Constitución",
        date(year, 12, 8): "Inmaculada Concepción",
        date(year, 12, 25): "Navidad",
    }


# =============================================================================
# 2. CUSTOM CALENDAR — persistent JSON
# =============================================================================
CALENDAR_FILE = Path(__file__).resolve().parent.parent / "data" / "calendario_laboral.json"


def load_custom_calendar():
    if CALENDAR_FILE.exists():
        with open(CALENDAR_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def save_custom_calendar(data):
    CALENDAR_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CALENDAR_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def generate_calendar_template(year):
    """Generate an Excel template with all days of the year."""
    rows = []
    holidays = get_default_holidays(year)
    n_days_year = (date(year, 12, 31) - date(year, 1, 1)).days + 1
    day_names_es = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    for i in range(n_days_year):
        d = date(year, 1, 1) + timedelta(days=i)
        is_weekend = d.weekday() >= 5
        is_holiday = d in holidays
        rows.append({
            "Fecha": d.strftime("%Y-%m-%d"),
            "Día": day_names_es[d.weekday()],
            "No laborable": "Sí" if (is_weekend or is_holiday) else "",
            "Motivo": holidays.get(d, "Fin de semana" if is_weekend else ""),
        })
    df = pd.DataFrame(rows)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Calendario")
        wb = writer.book
        ws = writer.sheets["Calendario"]
        hdr_f = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1})
        weekend_f = wb.add_format({"bg_color": "#D6DCE4", "border": 1})
        holiday_f = wb.add_format({"bg_color": "#FFC7CE", "border": 1})
        normal_f = wb.add_format({"border": 1})
        for ci, cn in enumerate(df.columns):
            ws.write(0, ci, cn, hdr_f)
        for ri in range(len(df)):
            row = df.iloc[ri]
            fmt = normal_f
            if row["Motivo"] == "Fin de semana":
                fmt = weekend_f
            elif row["No laborable"] == "Sí":
                fmt = holiday_f
            for ci in range(len(df.columns)):
                ws.write(ri + 1, ci, row.iloc[ci], fmt)
        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 12)
        ws.set_column(2, 2, 14)
        ws.set_column(3, 3, 35)
        ws.freeze_panes(1, 0)
    return buf.getvalue()


def parse_calendar_upload(data_bytes):
    """Parse an uploaded calendar Excel and return dict {year: {month: set_of_non_working_day_numbers}}."""
    df = pd.read_excel(BytesIO(data_bytes), sheet_name=0)
    col_fecha = df.columns[0]
    col_no_lab = df.columns[2] if len(df.columns) > 2 else None
    if col_no_lab is None:
        return None
    result = {}  # {"YYYY-MM-DD": True/False}
    for _, row in df.iterrows():
        fecha_str = str(row[col_fecha]).strip()[:10]
        no_lab = str(row[col_no_lab]).strip().lower()
        is_non_working = no_lab in ("sí", "si", "yes", "1", "true", "x")
        result[fecha_str] = is_non_working
    return result


def get_non_working_days(year, month, custom_cal=None):
    """Return set of day numbers (1-31) that are non-working in this month."""
    non_working = set()
    n_days = calendar.monthrange(year, month)[1]
    holidays = get_default_holidays(year)
    for day in range(1, n_days + 1):
        d = date(year, month, day)
        key = d.strftime("%Y-%m-%d")
        if custom_cal and key in custom_cal:
            if custom_cal[key]:
                non_working.add(day)
        else:
            if d.weekday() >= 5 or d in holidays:
                non_working.add(day)
    return non_working


def working_days_in_month(year, month, custom_cal=None):
    n_days = calendar.monthrange(year, month)[1]
    non_working = get_non_working_days(year, month, custom_cal)
    return n_days - len(non_working)


# =============================================================================
# 3. KNOWN CODES
# =============================================================================
KNOWN_CODES = {
    "V": "Vacaciones",
    "B": "Baja",
    "AP": "Asuntos Propios",
    "P": "Permiso",
    "PR": "Permiso Retribuido",
    "E": "Excedencia",
}

ABSENCE_CODES = {"V", "B", "AP", "P", "PR", "E"}

# For row detection — skip these as "employee names"
SKIP_NAMES = {
    "nan", "", "total", "totales", "suma", "nombre", "empleado", "trabajador",
    "departamento", "centro", "mes", "año", "seccion", "sección", "turno",
    "puesto", "categoria", "categoría", "observaciones", "notas", "dias",
    "horas", "plantilla",
}


def classify_cell(val):
    if pd.isna(val):
        return "empty", True
    s = str(val).strip()
    if s == "":
        return "empty", True
    try:
        float(s.replace(",", "."))
        return "worked", False
    except ValueError:
        pass
    upper = s.upper()
    if upper in KNOWN_CODES:
        return upper, False
    return f"unknown:{s}", True


# =============================================================================
# 4. PARSE EXCEL
# =============================================================================
def parse_cuadrante(data_bytes, filename="", non_working_days=None):
    """
    Parse a monthly hours grid Excel.
    non_working_days: set of day numbers to ignore for empty-cell warnings.
    """
    df_raw = pd.read_excel(BytesIO(data_bytes), header=None)
    warnings = []

    # --- Find the row with day numbers (1, 2, 3, ...) ---
    day_row_idx = None
    day_cols = {}

    for ri in range(min(15, len(df_raw))):
        row_vals = df_raw.iloc[ri].tolist()
        found_days = {}
        for ci, v in enumerate(row_vals):
            try:
                n = int(float(str(v).strip()))
                if 1 <= n <= 31:
                    found_days[n] = ci
            except (ValueError, TypeError):
                pass
        if len(found_days) >= 15:
            day_row_idx = ri
            day_cols = found_days
            break

    if day_row_idx is None:
        return None, [f"No se encontró fila de días (1-31) en '{filename}'"]

    # --- Detect employee name column ---
    min_day_col = min(day_cols.values())
    name_col = None
    for ci in range(min_day_col - 1, -1, -1):
        sample = df_raw.iloc[day_row_idx + 1:day_row_idx + 6, ci].dropna()
        if len(sample) > 0:
            has_text = any(
                isinstance(v, str) or (not pd.isna(v) and not str(v).strip().replace(".", "").isdigit())
                for v in sample
            )
            if has_text:
                name_col = ci
                break
    if name_col is None:
        name_col = 0

    # --- Detect month/year from header ---
    detected_month = None
    detected_year = None
    month_names_es = {
        "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
        "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
        "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
    }
    for ri in range(day_row_idx):
        for ci in range(min(10, len(df_raw.columns))):
            v = str(df_raw.iloc[ri, ci]).lower().strip()
            for mname, mnum in month_names_es.items():
                if mname in v:
                    detected_month = mnum
                    years = re.findall(r"20\d{2}", v)
                    if years:
                        detected_year = int(years[0])
                    break
    if detected_year is None:
        for ri in range(day_row_idx):
            for ci in range(min(10, len(df_raw.columns))):
                v = str(df_raw.iloc[ri, ci]).strip()
                years = re.findall(r"20\d{2}", v)
                if years:
                    detected_year = int(years[0])
                    break
            if detected_year:
                break

    # --- Non-working days for empty-cell filtering ---
    nw = non_working_days if non_working_days else set()

    # --- Parse employee rows ---
    max_day = max(day_cols.keys())
    employees = []
    unknown_codes = set()
    empty_cells = []

    for ri in range(day_row_idx + 1, len(df_raw)):
        name_val = df_raw.iloc[ri, name_col]
        if pd.isna(name_val) or str(name_val).strip() == "":
            row_data = df_raw.iloc[ri, min_day_col:max(day_cols.values()) + 1]
            if row_data.isna().all() or (row_data.astype(str).str.strip() == "").all():
                continue
            name_val = f"Empleado fila {ri + 1}"

        name = str(name_val).strip()
        # Skip header/title rows
        if name.lower() in SKIP_NAMES:
            continue
        # Skip if name looks like a date or number only
        if name.replace(".", "").replace(",", "").replace("-", "").replace("/", "").isdigit():
            continue

        emp = {
            "nombre": name,
            "worked": 0,
            "V": 0, "B": 0, "AP": 0, "P": 0, "PR": 0, "E": 0,
            "empty": 0, "unknown": 0,
            "unknown_codes": [],
        }

        for day_num, col_idx in day_cols.items():
            if day_num > max_day:
                continue
            cell_val = df_raw.iloc[ri, col_idx]
            cat, _ = classify_cell(cell_val)

            if cat == "worked":
                emp["worked"] += 1
            elif cat in ABSENCE_CODES:
                emp[cat] += 1
            elif cat == "empty":
                # Only flag if it's a working day
                if day_num not in nw:
                    emp["empty"] += 1
                    empty_cells.append({
                        "empleado": name, "dia": day_num,
                        "fila": ri + 1, "columna": col_idx + 1,
                    })
                # else: weekend/holiday → ignore silently
            elif cat.startswith("unknown:"):
                emp["unknown"] += 1
                code = cat.split(":", 1)[1]
                emp["unknown_codes"].append(code)
                unknown_codes.add(code)

        # Only real employees: need at least 1 working day or absence
        total_entries = emp["worked"] + sum(emp[c] for c in ABSENCE_CODES)
        if total_entries > 0:
            employees.append(emp)

    if not employees:
        return None, [f"No se encontraron empleados en '{filename}'"]

    if empty_cells:
        warnings.append({
            "tipo": "Celdas vacías",
            "mensaje": f"{len(empty_cells)} celda(s) vacía(s) en días laborables",
            "detalle": empty_cells,
        })
    if unknown_codes:
        warnings.append({
            "tipo": "Códigos desconocidos",
            "mensaje": f"Códigos no reconocidos: {', '.join(sorted(unknown_codes))}",
            "detalle": list(unknown_codes),
        })

    return {
        "employees": employees,
        "n_employees": len(employees),
        "detected_month": detected_month,
        "detected_year": detected_year,
        "max_day": max_day,
        "filename": filename,
    }, warnings


# =============================================================================
# 5. CALCULATE KPIs
# =============================================================================
def calculate_kpis(parsed, year, month, custom_cal=None):
    emps = parsed["employees"]
    n = len(emps)
    w_days = working_days_in_month(year, month, custom_cal)

    total_worked = sum(e["worked"] for e in emps)
    total_V = sum(e["V"] for e in emps)
    total_B = sum(e["B"] for e in emps)
    total_AP = sum(e["AP"] for e in emps)
    total_P = sum(e["P"] + e["PR"] for e in emps)  # PR counts as Permiso
    total_E = sum(e["E"] for e in emps)

    dias_teoricos = n * w_days

    ausencias_con_vac = total_V + total_B + total_AP + total_P + total_E
    pct_absent_con = (ausencias_con_vac / dias_teoricos * 100) if dias_teoricos > 0 else 0

    ausencias_sin_vac = total_B + total_AP + total_P + total_E
    pct_absent_sin = (ausencias_sin_vac / dias_teoricos * 100) if dias_teoricos > 0 else 0

    return {
        "plantilla": n,
        "dias_laborables": w_days,
        "dias_teoricos": dias_teoricos,
        "dias_trabajados": total_worked,
        "dias_vacaciones": total_V,
        "dias_baja": total_B,
        "dias_ap": total_AP,
        "dias_permiso": total_P,
        "dias_excedencia": total_E,
        "total_ausencias_con_vac": ausencias_con_vac,
        "total_ausencias_sin_vac": ausencias_sin_vac,
        "pct_absentismo_con_vac": round(pct_absent_con, 2),
        "pct_absentismo_sin_vac": round(pct_absent_sin, 2),
        "year": year,
        "month": month,
    }


# =============================================================================
# 6. HISTORY — persistent JSON
# =============================================================================
HISTORY_FILE = Path(__file__).resolve().parent.parent / "data" / "absentismo_historico.json"


def load_abs_history():
    if HISTORY_FILE.exists():
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_abs_history(entries):
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(entries, f, ensure_ascii=False, indent=2)


# =============================================================================
# 7. SIDEBAR — FILES + CALENDAR
# =============================================================================
st.sidebar.header("Archivos — Absentismo")

uploaded_files = st.sidebar.file_uploader(
    "Cuadrantes mensuales (uno por centro)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="abs_files",
)

if uploaded_files:
    SS["abs_file_data"] = [(f.name, f.getvalue()) for f in uploaded_files]
    st.sidebar.success(f"{len(uploaded_files)} fichero(s) cargado(s)")

# --- Custom calendar ---
st.sidebar.divider()
st.sidebar.subheader("Calendario laboral")

custom_cal = load_custom_calendar()
if custom_cal:
    st.sidebar.success("Calendario personalizado cargado")

cal_year_tmpl = st.sidebar.number_input("Año plantilla", 2020, 2030, date.today().year, key="cal_year_tmpl")
st.sidebar.download_button(
    "Descargar plantilla calendario",
    generate_calendar_template(cal_year_tmpl),
    file_name=f"calendario_laboral_{cal_year_tmpl}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

cal_upload = st.sidebar.file_uploader("Subir calendario cumplimentado", type=["xlsx", "xls"], key="cal_up")
if cal_upload:
    parsed_cal = parse_calendar_upload(cal_upload.getvalue())
    if parsed_cal:
        save_custom_calendar(parsed_cal)
        custom_cal = parsed_cal
        st.sidebar.success("Calendario guardado")
    else:
        st.sidebar.error("No se pudo leer el calendario")

if custom_cal and st.sidebar.button("Borrar calendario personalizado", type="secondary"):
    if CALENDAR_FILE.exists():
        CALENDAR_FILE.unlink()
    custom_cal = None
    st.rerun()

# --- Guard ---
if "abs_file_data" not in SS or not SS["abs_file_data"]:
    st.info("Sube los cuadrantes mensuales de horas en la barra lateral (un Excel por centro).")
    st.stop()


# =============================================================================
# 8. MONTH / YEAR + CONFIG
# =============================================================================
first_parsed, _ = parse_cuadrante(SS["abs_file_data"][0][1], SS["abs_file_data"][0][0])
auto_month = first_parsed["detected_month"] if first_parsed and first_parsed["detected_month"] else date.today().month
auto_year = first_parsed["detected_year"] if first_parsed and first_parsed["detected_year"] else date.today().year

st.header("Configuración")
c1, c2 = st.columns(2)
with c1:
    sel_month = st.selectbox("Mes", list(range(1, 13)), index=auto_month - 1,
                             format_func=lambda m: MONTH_NAMES[m], key="abs_month")
with c2:
    sel_year = st.number_input("Año", 2020, 2030, auto_year, key="abs_year")

w_days = working_days_in_month(sel_year, sel_month, custom_cal)
nw_days = get_non_working_days(sel_year, sel_month, custom_cal)

# Show holidays for this month
holidays_default = get_default_holidays(sel_year)
month_holidays = {d: n for d, n in holidays_default.items() if d.month == sel_month}
cal_source = "calendario personalizado" if custom_cal else "Navarra (nacional + foral)"
st.caption(
    f"**{MONTH_NAMES[sel_month]} {sel_year}**: {w_days} días laborables — {cal_source}"
    + (f" · Festivos: {', '.join(f'{d.day}-{n}' for d, n in sorted(month_holidays.items()))}" if month_holidays and not custom_cal else "")
)


# =============================================================================
# 9. PROCESS
# =============================================================================
if st.button("Analizar Absentismo", type="primary", use_container_width=True):
    all_results = []
    all_warnings = {}
    all_employee_details = {}

    for fname, fdata in SS["abs_file_data"]:
        centro_name = fname.rsplit(".", 1)[0]
        parsed, warns = parse_cuadrante(fdata, fname, non_working_days=nw_days)
        if parsed is None:
            st.error(f"Error procesando **{fname}**: {'; '.join(w if isinstance(w, str) else w.get('mensaje', '') for w in warns)}")
            continue
        kpis = calculate_kpis(parsed, sel_year, sel_month, custom_cal)
        kpis["centro"] = centro_name
        all_results.append(kpis)
        all_warnings[centro_name] = warns
        all_employee_details[centro_name] = parsed["employees"]

    if all_results:
        SS["abs_results"] = {
            "kpis": all_results,
            "warnings": all_warnings,
            "employees": all_employee_details,
            "month": sel_month,
            "year": sel_year,
        }


# =============================================================================
# 10. DISPLAY RESULTS
# =============================================================================
if "abs_results" in SS:
    res = SS["abs_results"]
    kpis_list = res["kpis"]
    all_warnings = res["warnings"]
    all_employees = res["employees"]
    r_month = res["month"]
    r_year = res["year"]

    # --- Warnings (only unknown codes in UI — empty cells just count) ---
    for centro, warns in all_warnings.items():
        for w in warns:
            if w["tipo"] == "Códigos desconocidos":
                st.error(f"**{centro}**: {w['mensaje']}")
            elif w["tipo"] == "Celdas vacías":
                st.warning(f"**{centro}**: {w['mensaje']}")

    # --- KPI Cards ---
    st.header(f"Resumen — {MONTH_NAMES[r_month]} {r_year}")

    total_plantilla = sum(k["plantilla"] for k in kpis_list)
    total_dias_teo = sum(k["dias_teoricos"] for k in kpis_list)
    total_trabajados = sum(k["dias_trabajados"] for k in kpis_list)
    total_ausencias_con = sum(k["total_ausencias_con_vac"] for k in kpis_list)
    total_ausencias_sin = sum(k["total_ausencias_sin_vac"] for k in kpis_list)
    pct_con = round(total_ausencias_con / total_dias_teo * 100, 2) if total_dias_teo > 0 else 0
    pct_sin = round(total_ausencias_sin / total_dias_teo * 100, 2) if total_dias_teo > 0 else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Plantilla total", total_plantilla)
    m2.metric("Días trabajados", f"{total_trabajados:,}")
    m3.metric("Absentismo (con vac.)", f"{pct_con:.2f}%")
    m4.metric("Absentismo (sin vac.)", f"{pct_sin:.2f}%")

    # --- Tabs ---
    if len(kpis_list) > 1:
        tab_names = [k["centro"] for k in kpis_list] + ["Consolidado"]
    else:
        tab_names = [kpis_list[0]["centro"]]

    tabs = st.tabs(tab_names)

    def show_kpi_detail(k, employees=None):
        c1, c2 = st.columns(2)
        c1.metric("Plantilla", k["plantilla"])
        c2.metric("Días laborables", k["dias_laborables"])

        st.subheader("Desglose de días")
        df_detail = pd.DataFrame({
            "Concepto": ["Días trabajados", "Vacaciones (V)", "Bajas (B)",
                         "Asuntos Propios (AP)", "Permisos (P/PR)", "Excedencias (E)"],
            "Días": [k["dias_trabajados"], k["dias_vacaciones"], k["dias_baja"],
                     k["dias_ap"], k["dias_permiso"], k["dias_excedencia"]],
        })
        st.dataframe(df_detail, use_container_width=True, hide_index=True)

        a1, a2 = st.columns(2)
        a1.metric("% Absentismo CON vacaciones", f"{k['pct_absentismo_con_vac']:.2f}%",
                   help="(V + B + AP + P/PR + E) / Días teóricos × 100")
        a2.metric("% Absentismo SIN vacaciones", f"{k['pct_absentismo_sin_vac']:.2f}%",
                   help="(B + AP + P/PR + E) / Días teóricos × 100")

        if employees:
            st.subheader("Detalle por empleado")
            rows = []
            for e in employees:
                rows.append({
                    "Empleado": e["nombre"],
                    "Trabajados": e["worked"],
                    "Vacaciones": e["V"],
                    "Bajas": e["B"],
                    "A. Propios": e["AP"],
                    "Permisos": e["P"] + e["PR"],
                    "Excedencias": e["E"],
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True, height=400)

    for i, tab in enumerate(tabs):
        with tab:
            if i < len(kpis_list):
                show_kpi_detail(kpis_list[i], all_employees.get(kpis_list[i]["centro"], []))
            else:
                consol = {
                    "plantilla": total_plantilla,
                    "dias_laborables": kpis_list[0]["dias_laborables"],
                    "dias_trabajados": total_trabajados,
                    "dias_vacaciones": sum(k["dias_vacaciones"] for k in kpis_list),
                    "dias_baja": sum(k["dias_baja"] for k in kpis_list),
                    "dias_ap": sum(k["dias_ap"] for k in kpis_list),
                    "dias_permiso": sum(k["dias_permiso"] for k in kpis_list),
                    "dias_excedencia": sum(k["dias_excedencia"] for k in kpis_list),
                    "total_ausencias_con_vac": total_ausencias_con,
                    "total_ausencias_sin_vac": total_ausencias_sin,
                    "pct_absentismo_con_vac": pct_con,
                    "pct_absentismo_sin_vac": pct_sin,
                }
                show_kpi_detail(consol)

    # =================================================================
    # 11. DOWNLOAD EXCEL
    # =================================================================
    st.header("Descargar")

    def build_absentismo_excel(kpis_list, all_employees, r_month, r_year):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book
            title_f = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864", "bottom": 2})
            hdr_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
            lbl_f = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#D6DCE4"})
            num_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0"})
            pct_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%"})
            pct_green = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#006100", "bg_color": "#C6EFCE"})
            pct_red = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#9C0006", "bg_color": "#FFC7CE"})
            legend_f = wb.add_format({"font_size": 9, "italic": True, "text_wrap": True})

            def wpct(ws, r, c, v):
                fmt = pct_green if v < 0.05 else (pct_red if v > 0.10 else pct_f)
                ws.write_number(r, c, v, fmt)

            sn = "Resumen"
            pd.DataFrame().to_excel(writer, sheet_name=sn, index=False)
            ws = writer.sheets[sn]
            ws.set_column(0, 0, 30)
            for ci in range(1, len(kpis_list) + 2):
                ws.set_column(ci, ci, 18)

            row = 0
            ws.merge_range(row, 0, row, len(kpis_list) + 1,
                           f"ANÁLISIS DE ABSENTISMO — {MONTH_NAMES[r_month]} {r_year}", title_f)
            row += 2

            headers = [""] + [k["centro"] for k in kpis_list]
            if len(kpis_list) > 1:
                headers.append("TOTAL")
            for ci, h in enumerate(headers):
                ws.write(row, ci, h, hdr_f)
            ws.set_row(row, 25)
            row += 1

            metrics = [
                ("Plantilla", "plantilla"),
                ("Días laborables", "dias_laborables"),
                ("Días trabajados", "dias_trabajados"),
                ("Vacaciones (días)", "dias_vacaciones"),
                ("Bajas (días)", "dias_baja"),
                ("Asuntos Propios (días)", "dias_ap"),
                ("Permisos (días)", "dias_permiso"),
                ("Excedencias (días)", "dias_excedencia"),
                ("Total ausencias (con vac.)", "total_ausencias_con_vac"),
                ("Total ausencias (sin vac.)", "total_ausencias_sin_vac"),
            ]

            for label, key in metrics:
                ws.write(row, 0, label, lbl_f)
                for ci, k in enumerate(kpis_list):
                    ws.write(row, ci + 1, k[key], num_f)
                if len(kpis_list) > 1:
                    ws.write(row, len(kpis_list) + 1, sum(k[key] for k in kpis_list), num_f)
                row += 1

            row += 1
            ws.write(row, 0, "% Absentismo CON vacaciones", lbl_f)
            for ci, k in enumerate(kpis_list):
                wpct(ws, row, ci + 1, k["pct_absentismo_con_vac"] / 100)
            if len(kpis_list) > 1:
                t_teo = sum(k["dias_teoricos"] for k in kpis_list)
                t_con = sum(k["total_ausencias_con_vac"] for k in kpis_list)
                wpct(ws, row, len(kpis_list) + 1, t_con / t_teo if t_teo else 0)
            row += 1

            ws.write(row, 0, "% Absentismo SIN vacaciones", lbl_f)
            for ci, k in enumerate(kpis_list):
                wpct(ws, row, ci + 1, k["pct_absentismo_sin_vac"] / 100)
            if len(kpis_list) > 1:
                t_sin = sum(k["total_ausencias_sin_vac"] for k in kpis_list)
                wpct(ws, row, len(kpis_list) + 1, t_sin / t_teo if t_teo else 0)

            # Legend
            row += 2
            ws.merge_range(row, 0, row, len(kpis_list) + 1,
                           "Leyenda: Verde = < 5% absentismo · Sin color = 5%-10% · Rojo = > 10%", legend_f)

            ws.print_area(0, 0, row + 1, len(kpis_list) + 1)
            ws.set_landscape()
            ws.fit_to_pages(1, 0)

            # --- Employee detail sheets ---
            for k in kpis_list:
                centro = k["centro"]
                emps = all_employees.get(centro, [])
                if not emps:
                    continue
                rows_data = []
                for e in emps:
                    rows_data.append({
                        "Empleado": e["nombre"],
                        "Trabajados": e["worked"],
                        "Vacaciones": e["V"],
                        "Bajas": e["B"],
                        "A. Propios": e["AP"],
                        "Permisos": e["P"] + e["PR"],
                        "Excedencias": e["E"],
                    })
                df_emp = pd.DataFrame(rows_data)
                safe_name = centro[:31]
                df_emp.to_excel(writer, sheet_name=safe_name, index=False)
                ws_emp = writer.sheets[safe_name]
                for ci, cn in enumerate(df_emp.columns):
                    ws_emp.write(0, ci, cn, hdr_f)
                    mx = max(len(cn), int(df_emp[cn].astype(str).str.len().max()) if len(df_emp) else 0)
                    ws_emp.set_column(ci, ci, min(mx + 3, 30))
                ws_emp.freeze_panes(1, 0)

        return output.getvalue()

    excel_data = build_absentismo_excel(kpis_list, all_employees, r_month, r_year)
    st.download_button(
        "Descargar Análisis de Absentismo (Excel)",
        excel_data,
        file_name=f"ABSENTISMO_{MONTH_NAMES[r_month].upper()}_{r_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )

    # =================================================================
    # 12. SAVE TO HISTORY
    # =================================================================
    st.divider()
    if st.button("Guardar mes en historial", type="secondary", use_container_width=True,
                  help="Guarda estos resultados para el análisis histórico"):
        hist = load_abs_history()
        key = f"{r_year}-{r_month:02d}"
        # Replace if same month already exists
        hist = [h for h in hist if h.get("key") != key]
        hist.append({
            "key": key,
            "year": r_year,
            "month": r_month,
            "centros": [
                {
                    "centro": k["centro"],
                    "plantilla": k["plantilla"],
                    "dias_laborables": k["dias_laborables"],
                    "dias_trabajados": k["dias_trabajados"],
                    "dias_vacaciones": k["dias_vacaciones"],
                    "dias_baja": k["dias_baja"],
                    "dias_ap": k["dias_ap"],
                    "dias_permiso": k["dias_permiso"],
                    "dias_excedencia": k["dias_excedencia"],
                    "pct_con": k["pct_absentismo_con_vac"],
                    "pct_sin": k["pct_absentismo_sin_vac"],
                }
                for k in kpis_list
            ],
            "total_pct_con": pct_con,
            "total_pct_sin": pct_sin,
            "total_plantilla": total_plantilla,
        })
        hist.sort(key=lambda h: h["key"])
        save_abs_history(hist)
        st.success(f"Guardado **{MONTH_NAMES[r_month]} {r_year}** en historial")


# =============================================================================
# 13. HISTORICAL ANALYSIS
# =============================================================================
st.divider()
st.header("Histórico de Absentismo")

hist = load_abs_history()
if not hist:
    st.info("No hay datos históricos. Procesa meses y guárdalos con el botón 'Guardar mes en historial'.")
else:
    st.caption(f"**{len(hist)}** mes(es) guardado(s)")

    # Build trend data
    trend_rows = []
    for h in hist:
        label = f"{MONTH_NAMES[h['month']]} {h['year']}"
        trend_rows.append({
            "Mes": label,
            "key": h["key"],
            "Plantilla": h["total_plantilla"],
            "% Con Vac.": h["total_pct_con"],
            "% Sin Vac.": h["total_pct_sin"],
        })
    df_trend = pd.DataFrame(trend_rows)

    # Chart: absenteeism trend
    if len(df_trend) >= 2:
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df_trend["Mes"], y=df_trend["% Con Vac."],
            name="Con vacaciones", mode="lines+markers",
            line=dict(color="#FF5722", width=2),
            marker=dict(size=8),
        ))
        fig.add_trace(go.Scatter(
            x=df_trend["Mes"], y=df_trend["% Sin Vac."],
            name="Sin vacaciones", mode="lines+markers",
            line=dict(color="#2196F3", width=2),
            marker=dict(size=8),
        ))
        fig.update_layout(
            title="Evolución del absentismo",
            yaxis_title="% Absentismo",
            xaxis_title="",
            height=400,
            legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
        )
        st.plotly_chart(fig, use_container_width=True)

    # Table
    st.dataframe(df_trend[["Mes", "Plantilla", "% Con Vac.", "% Sin Vac."]],
                 use_container_width=True, hide_index=True)

    # Per-center breakdown
    if len(hist) >= 2:
        centros_set = set()
        for h in hist:
            for c in h["centros"]:
                centros_set.add(c["centro"])

        if len(centros_set) > 1:
            st.subheader("Evolución por centro")
            fig2 = go.Figure()
            colors = ["#2196F3", "#FF5722", "#4CAF50", "#FF9800", "#9C27B0", "#00BCD4"]
            for idx, centro in enumerate(sorted(centros_set)):
                x_vals, y_vals = [], []
                for h in hist:
                    for c in h["centros"]:
                        if c["centro"] == centro:
                            x_vals.append(f"{MONTH_NAMES[h['month']]} {h['year']}")
                            y_vals.append(c["pct_sin"])
                fig2.add_trace(go.Scatter(
                    x=x_vals, y=y_vals, name=centro,
                    mode="lines+markers",
                    line=dict(color=colors[idx % len(colors)], width=2),
                ))
            fig2.update_layout(
                title="Absentismo sin vacaciones por centro",
                yaxis_title="% Absentismo",
                height=400,
                legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
            )
            st.plotly_chart(fig2, use_container_width=True)

    # Clear history
    with st.expander("Gestionar historial"):
        for h in hist:
            st.text(f"{MONTH_NAMES[h['month']]} {h['year']} — Plantilla: {h['total_plantilla']} — "
                    f"Abs. con vac: {h['total_pct_con']:.2f}% — sin vac: {h['total_pct_sin']:.2f}%")
        if st.button("Limpiar todo el historial", type="secondary"):
            save_abs_history([])
            st.rerun()
