"""
4_absentismo.py
===============
Análisis de absentismo por centro de trabajo.
Sube un Excel mensual por centro (cuadrante de horas) y obtiene KPIs,
validaciones y un reporte consolidado descargable.
"""
import streamlit as st
import pandas as pd
import numpy as np
import calendar
from datetime import date, timedelta
from io import BytesIO

st.title("Análisis de Absentismo")
st.markdown("Sube los cuadrantes mensuales de horas (un Excel por centro) y genera el análisis.")

from metodologia import render_download as _render_metodologia
_render_metodologia("absentismo")

SS = st.session_state

# =============================================================================
# 1. SPANISH HOLIDAYS
# =============================================================================
def _easter_date(year):
    """Computus algorithm for Easter Sunday."""
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


def get_spanish_holidays(year):
    """National Spanish holidays for a given year."""
    easter = _easter_date(year)
    good_friday = easter - timedelta(days=2)
    return {
        date(year, 1, 1): "Año Nuevo",
        date(year, 1, 6): "Epifanía",
        good_friday: "Viernes Santo",
        date(year, 5, 1): "Día del Trabajo",
        date(year, 8, 15): "Asunción",
        date(year, 10, 12): "Día de la Hispanidad",
        date(year, 11, 1): "Todos los Santos",
        date(year, 12, 6): "Día de la Constitución",
        date(year, 12, 8): "Inmaculada Concepción",
        date(year, 12, 25): "Navidad",
    }


def working_days_in_month(year, month, extra_holidays=None):
    """Count working days (Mon-Fri excluding national holidays)."""
    holidays = get_spanish_holidays(year)
    if extra_holidays:
        for d in extra_holidays:
            holidays[d] = "Festivo local"
    n_days = calendar.monthrange(year, month)[1]
    count = 0
    for day in range(1, n_days + 1):
        d = date(year, month, day)
        if d.weekday() < 5 and d not in holidays:
            count += 1
    return count


# =============================================================================
# 2. KNOWN CODES
# =============================================================================
KNOWN_CODES = {
    "V": "Vacaciones",
    "B": "Baja",
    "AP": "Asuntos Propios",
    "P": "Permiso",
    "E": "Excedencia",
}

ABSENCE_CODES = {"V", "B", "AP", "P", "E"}


def classify_cell(val):
    """
    Classify a cell value.
    Returns: (category, is_warning)
    Categories: 'worked', 'V', 'B', 'AP', 'P', 'E', 'empty', 'unknown'
    """
    if pd.isna(val):
        return "empty", True
    s = str(val).strip()
    if s == "":
        return "empty", True
    # Try numeric
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
# 3. PARSE EXCEL
# =============================================================================
def parse_cuadrante(data_bytes, filename=""):
    """
    Parse a monthly hours grid Excel.
    Returns: dict with parsed data, warnings, and metadata.
    """
    df_raw = pd.read_excel(BytesIO(data_bytes), header=None)
    warnings = []

    # --- Find the row with day numbers (1, 2, 3, ...) ---
    day_row_idx = None
    day_cols = {}  # {day_number: column_index}

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
        # Need at least 15 consecutive-ish days to consider this the header
        if len(found_days) >= 15:
            day_row_idx = ri
            day_cols = found_days
            break

    if day_row_idx is None:
        return None, [f"No se encontró fila de días (1-31) en '{filename}'"]

    # --- Detect employee name column (first text column before day columns) ---
    min_day_col = min(day_cols.values())
    name_col = None
    for ci in range(min_day_col - 1, -1, -1):
        # Check if this column has text values below the header
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

    # --- Extract month/year from header area ---
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
                    # Try to find year nearby
                    import re
                    years = re.findall(r"20\d{2}", v)
                    if years:
                        detected_year = int(years[0])
                    break

    # Try to find year from header if not found yet
    if detected_year is None:
        for ri in range(day_row_idx):
            for ci in range(min(10, len(df_raw.columns))):
                v = str(df_raw.iloc[ri, ci]).strip()
                import re
                years = re.findall(r"20\d{2}", v)
                if years:
                    detected_year = int(years[0])
                    break
            if detected_year:
                break

    # --- Parse employee rows ---
    max_day = max(day_cols.keys())
    employees = []
    unknown_codes = set()
    empty_cells = []

    for ri in range(day_row_idx + 1, len(df_raw)):
        name_val = df_raw.iloc[ri, name_col]
        if pd.isna(name_val) or str(name_val).strip() == "":
            # Check if entire row is empty → stop
            row_data = df_raw.iloc[ri, min_day_col:max(day_cols.values()) + 1]
            if row_data.isna().all():
                continue
            name_val = f"Empleado fila {ri + 1}"
        name = str(name_val).strip()
        if name.lower() in ("nan", "", "total", "totales", "suma"):
            continue

        emp = {
            "nombre": name,
            "worked": 0,
            "V": 0, "B": 0, "AP": 0, "P": 0, "E": 0,
            "empty": 0, "unknown": 0,
            "unknown_codes": [],
        }

        for day_num, col_idx in day_cols.items():
            if day_num > max_day:
                continue
            cell_val = df_raw.iloc[ri, col_idx]
            cat, is_warn = classify_cell(cell_val)

            if cat == "worked":
                emp["worked"] += 1
            elif cat in ABSENCE_CODES:
                emp[cat] += 1
            elif cat == "empty":
                emp["empty"] += 1
                empty_cells.append({
                    "empleado": name, "dia": day_num,
                    "fila": ri + 1, "columna": col_idx + 1,
                })
            elif cat.startswith("unknown:"):
                emp["unknown"] += 1
                code = cat.split(":", 1)[1]
                emp["unknown_codes"].append(code)
                unknown_codes.add(code)

        # Only add if they have some data
        total_entries = emp["worked"] + emp["V"] + emp["B"] + emp["AP"] + emp["P"] + emp["E"]
        if total_entries > 0 or emp["empty"] > 0:
            employees.append(emp)

    if not employees:
        return None, [f"No se encontraron empleados en '{filename}'"]

    # Build warnings
    if empty_cells:
        warnings.append({
            "tipo": "Celdas vacías",
            "mensaje": f"{len(empty_cells)} celda(s) vacía(s) detectadas",
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
# 4. CALCULATE KPIs
# =============================================================================
def calculate_kpis(parsed, year, month):
    """Calculate absenteeism KPIs from parsed data."""
    emps = parsed["employees"]
    n = len(emps)
    w_days = working_days_in_month(year, month)

    total_worked = sum(e["worked"] for e in emps)
    total_V = sum(e["V"] for e in emps)
    total_B = sum(e["B"] for e in emps)
    total_AP = sum(e["AP"] for e in emps)
    total_P = sum(e["P"] for e in emps)
    total_E = sum(e["E"] for e in emps)

    dias_teoricos = n * w_days
    horas_teoricas = dias_teoricos * 8

    # Absentismo CON vacaciones
    ausencias_con_vac = total_V + total_B + total_AP + total_P + total_E
    pct_absent_con = (ausencias_con_vac / dias_teoricos * 100) if dias_teoricos > 0 else 0

    # Absentismo SIN vacaciones
    ausencias_sin_vac = total_B + total_AP + total_P + total_E
    pct_absent_sin = (ausencias_sin_vac / dias_teoricos * 100) if dias_teoricos > 0 else 0

    return {
        "plantilla": n,
        "dias_laborables": w_days,
        "dias_teoricos": dias_teoricos,
        "horas_teoricas": horas_teoricas,
        "dias_trabajados": total_worked,
        "horas_trabajadas": total_worked * 8,
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
# 5. SIDEBAR — FILE UPLOADS
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

if "abs_file_data" not in SS or not SS["abs_file_data"]:
    st.info("Sube los cuadrantes mensuales de horas en la barra lateral (un Excel por centro).")
    st.stop()


# =============================================================================
# 6. MONTH / YEAR SELECTION
# =============================================================================
MONTH_NAMES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# Try to auto-detect from first file
first_parsed, _ = parse_cuadrante(SS["abs_file_data"][0][1], SS["abs_file_data"][0][0])
auto_month = first_parsed["detected_month"] if first_parsed and first_parsed["detected_month"] else date.today().month
auto_year = first_parsed["detected_year"] if first_parsed and first_parsed["detected_year"] else date.today().year

st.header("Configuración")
c1, c2 = st.columns(2)
with c1:
    sel_month = st.selectbox(
        "Mes",
        options=list(range(1, 13)),
        index=auto_month - 1,
        format_func=lambda m: MONTH_NAMES[m],
        key="abs_month",
    )
with c2:
    sel_year = st.number_input("Año", min_value=2020, max_value=2030, value=auto_year, key="abs_year")

w_days = working_days_in_month(sel_year, sel_month)
holidays = get_spanish_holidays(sel_year)
month_holidays = {d: n for d, n in holidays.items() if d.month == sel_month}

st.caption(
    f"**{MONTH_NAMES[sel_month]} {sel_year}**: {w_days} días laborables"
    + (f" (festivos: {', '.join(f'{d.day} - {n}' for d, n in sorted(month_holidays.items()))})" if month_holidays else "")
)


# =============================================================================
# 7. PROCESS
# =============================================================================
if st.button("Analizar Absentismo", type="primary", use_container_width=True):
    all_results = []
    all_warnings = {}
    all_employee_details = {}

    for fname, fdata in SS["abs_file_data"]:
        centro_name = fname.rsplit(".", 1)[0]  # Remove extension
        parsed, warns = parse_cuadrante(fdata, fname)

        if parsed is None:
            st.error(f"Error procesando **{fname}**: {'; '.join(w if isinstance(w, str) else w.get('mensaje', '') for w in warns)}")
            continue

        kpis = calculate_kpis(parsed, sel_year, sel_month)
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
# 8. DISPLAY RESULTS
# =============================================================================
if "abs_results" in SS:
    res = SS["abs_results"]
    kpis_list = res["kpis"]
    all_warnings = res["warnings"]
    all_employees = res["employees"]
    r_month = res["month"]
    r_year = res["year"]

    # --- Warnings ---
    has_warnings = False
    for centro, warns in all_warnings.items():
        for w in warns:
            has_warnings = True
            if w["tipo"] == "Celdas vacías":
                st.warning(f"**{centro}**: {w['mensaje']}")
                with st.expander(f"Detalle celdas vacías — {centro}"):
                    df_empty = pd.DataFrame(w["detalle"])
                    st.dataframe(df_empty, use_container_width=True, hide_index=True)
            elif w["tipo"] == "Códigos desconocidos":
                st.error(f"**{centro}**: {w['mensaje']}")

    if not has_warnings:
        st.success("Todos los cuadrantes procesados sin incidencias.")

    # --- KPI Cards ---
    st.header(f"Resumen — {MONTH_NAMES[r_month]} {r_year}")

    # Consolidated totals
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

    # --- Per-center tabs ---
    if len(kpis_list) > 1:
        tab_names = [k["centro"] for k in kpis_list] + ["Consolidado"]
    else:
        tab_names = [kpis_list[0]["centro"]]

    tabs = st.tabs(tab_names)

    def show_kpi_detail(k, employees=None):
        """Display detailed KPIs for a center or consolidated."""
        c1, c2, c3 = st.columns(3)
        c1.metric("Plantilla", k["plantilla"])
        c2.metric("Días laborables", k["dias_laborables"])
        c3.metric("Horas teóricas", f"{k['horas_teoricas']:,}")

        st.subheader("Desglose de días")
        data = {
            "Concepto": [
                "Días trabajados",
                "Vacaciones (V)",
                "Bajas (B)",
                "Asuntos Propios (AP)",
                "Permisos (P)",
                "Excedencias (E)",
            ],
            "Días": [
                k["dias_trabajados"], k["dias_vacaciones"],
                k["dias_baja"], k["dias_ap"],
                k["dias_permiso"], k["dias_excedencia"],
            ],
            "Horas (×8)": [
                k["dias_trabajados"] * 8, k["dias_vacaciones"] * 8,
                k["dias_baja"] * 8, k["dias_ap"] * 8,
                k["dias_permiso"] * 8, k["dias_excedencia"] * 8,
            ],
        }
        df_detail = pd.DataFrame(data)
        st.dataframe(df_detail, use_container_width=True, hide_index=True)

        a1, a2 = st.columns(2)
        a1.metric("% Absentismo CON vacaciones", f"{k['pct_absentismo_con_vac']:.2f}%",
                   help="(V + B + AP + P + E) / Días teóricos × 100")
        a2.metric("% Absentismo SIN vacaciones", f"{k['pct_absentismo_sin_vac']:.2f}%",
                   help="(B + AP + P + E) / Días teóricos × 100")

        # Employee detail table
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
                    "Permisos": e["P"],
                    "Excedencias": e["E"],
                    "Vacías": e["empty"],
                })
            df_emp = pd.DataFrame(rows)
            st.dataframe(df_emp, use_container_width=True, hide_index=True, height=400)

    for i, tab in enumerate(tabs):
        with tab:
            if i < len(kpis_list):
                k = kpis_list[i]
                emps = all_employees.get(k["centro"], [])
                show_kpi_detail(k, emps)
            else:
                # Consolidated
                consol = {
                    "plantilla": total_plantilla,
                    "dias_laborables": kpis_list[0]["dias_laborables"],
                    "horas_teoricas": sum(k["horas_teoricas"] for k in kpis_list),
                    "dias_trabajados": total_trabajados,
                    "horas_trabajadas": sum(k["horas_trabajadas"] for k in kpis_list),
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
    # 9. DOWNLOAD EXCEL
    # =================================================================
    st.header("Descargar")

    def build_absentismo_excel(kpis_list, all_employees, r_month, r_year):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book

            # Formats
            title_f = wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F3864", "bottom": 2})
            hdr_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
            lbl_f = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#D6DCE4"})
            num_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0"})
            pct_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%"})
            pct_green = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#006100", "bg_color": "#C6EFCE"})
            pct_red = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#9C0006", "bg_color": "#FFC7CE"})
            sub_f = wb.add_format({"bold": True, "font_size": 11, "font_color": "#1F3864", "italic": True})
            cell_f = wb.add_format({"font_size": 10, "border": 1})
            cell_c = wb.add_format({"font_size": 10, "border": 1, "align": "center"})

            def wpct(ws, r, c, v):
                fmt = pct_green if v < 0.05 else (pct_red if v > 0.10 else pct_f)
                ws.write_number(r, c, v, fmt)

            # --- Sheet 1: Resumen por centro ---
            sn = "Resumen"
            pd.DataFrame().to_excel(writer, sheet_name=sn, index=False)
            ws = writer.sheets[sn]

            ws.set_column(0, 0, 28)
            for ci in range(1, len(kpis_list) + 2):
                ws.set_column(ci, ci, 18)

            row = 0
            ws.merge_range(row, 0, row, len(kpis_list) + 1,
                           f"ANÁLISIS DE ABSENTISMO — {MONTH_NAMES[r_month]} {r_year}", title_f)
            row += 2

            # Headers: concepto | centro1 | centro2 | ... | TOTAL
            headers = [""] + [k["centro"] for k in kpis_list]
            if len(kpis_list) > 1:
                headers.append("TOTAL")
            for ci, h in enumerate(headers):
                ws.write(row, ci, h, hdr_f)
            ws.set_row(row, 25)
            row += 1

            # Rows of metrics
            metrics = [
                ("Plantilla", "plantilla", num_f),
                ("Días laborables", "dias_laborables", num_f),
                ("Horas teóricas", "horas_teoricas", num_f),
                ("Días trabajados", "dias_trabajados", num_f),
                ("Horas trabajadas", "horas_trabajadas", num_f),
                ("Vacaciones (días)", "dias_vacaciones", num_f),
                ("Bajas (días)", "dias_baja", num_f),
                ("Asuntos Propios (días)", "dias_ap", num_f),
                ("Permisos (días)", "dias_permiso", num_f),
                ("Excedencias (días)", "dias_excedencia", num_f),
                ("Total ausencias (con vac.)", "total_ausencias_con_vac", num_f),
                ("Total ausencias (sin vac.)", "total_ausencias_sin_vac", num_f),
            ]

            for label, key, fmt in metrics:
                ws.write(row, 0, label, lbl_f)
                for ci, k in enumerate(kpis_list):
                    ws.write(row, ci + 1, k[key], fmt)
                if len(kpis_list) > 1:
                    total = sum(k[key] for k in kpis_list)
                    ws.write(row, len(kpis_list) + 1, total, fmt)
                row += 1

            # Percentage rows
            row += 1
            ws.write(row, 0, "% Absentismo CON vacaciones", lbl_f)
            for ci, k in enumerate(kpis_list):
                wpct(ws, row, ci + 1, k["pct_absentismo_con_vac"] / 100)
            if len(kpis_list) > 1:
                total_teo = sum(k["dias_teoricos"] for k in kpis_list)
                total_aus_con = sum(k["total_ausencias_con_vac"] for k in kpis_list)
                wpct(ws, row, len(kpis_list) + 1, total_aus_con / total_teo if total_teo else 0)
            row += 1

            ws.write(row, 0, "% Absentismo SIN vacaciones", lbl_f)
            for ci, k in enumerate(kpis_list):
                wpct(ws, row, ci + 1, k["pct_absentismo_sin_vac"] / 100)
            if len(kpis_list) > 1:
                total_aus_sin = sum(k["total_ausencias_sin_vac"] for k in kpis_list)
                wpct(ws, row, len(kpis_list) + 1, total_aus_sin / total_teo if total_teo else 0)

            ws.print_area(0, 0, row + 1, len(kpis_list) + 1)
            ws.set_landscape()
            ws.fit_to_pages(1, 0)

            # --- Sheet per center: employee detail ---
            for k in kpis_list:
                centro = k["centro"]
                safe_name = centro[:31]
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
                        "Permisos": e["P"],
                        "Excedencias": e["E"],
                        "Celdas vacías": e["empty"],
                    })
                df_emp = pd.DataFrame(rows_data)
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
        use_container_width=True,
        type="primary",
    )
