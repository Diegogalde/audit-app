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

CENTROS_DISPONIBLES = ["Noain", "Post-Venta", "Export-OTC", "Arazuri"]
abs_centro_sel = st.selectbox(
    "Centro de trabajo",
    options=CENTROS_DISPONIBLES,
    key="abs_centro",
    help="Selecciona el centro para este análisis",
)

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
# 2. CUSTOM CALENDAR — persistent JSON (supports per-center calendars)
# =============================================================================
CALENDAR_FILE = Path(__file__).resolve().parent.parent / "data" / "calendario_laboral.json"
EMPLOYEES_FILE = Path(__file__).resolve().parent.parent / "data" / "plantilla_empleados.json"
CENTROS_FILE = Path(__file__).resolve().parent.parent / "data" / "centros_trabajo.json"


def _load_raw_calendar():
    """Load raw calendar JSON, migrating old flat format if needed."""
    if not CALENDAR_FILE.exists():
        return {"__default__": None, "centros": {}}
    with open(CALENDAR_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    # Migrate old flat format: {"2026-01-01": true, ...} → new structure
    if isinstance(data, dict) and "__default__" not in data and "centros" not in data:
        return {"__default__": data, "centros": {}}
    return data


def load_custom_calendar():
    """Load the default custom calendar (backward compatible)."""
    raw = _load_raw_calendar()
    return raw.get("__default__")


def load_center_calendars():
    """Load per-center calendars: {centro_name: {date_str: bool}}."""
    raw = _load_raw_calendar()
    return raw.get("centros", {})


def get_calendar_for_centro(centro_name, default_cal=None):
    """Get the calendar for a specific center. Falls back to default if none set."""
    center_cals = load_center_calendars()
    # Try exact match, then case-insensitive partial match
    for key, cal in center_cals.items():
        if key.upper() == centro_name.upper() or key.upper() in centro_name.upper() or centro_name.upper() in key.upper():
            return cal
    return default_cal


def save_custom_calendar(data):
    """Save the default calendar (backward compatible)."""
    raw = _load_raw_calendar()
    raw["__default__"] = data
    CALENDAR_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CALENDAR_FILE, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)


def save_center_calendar(centro_name, data):
    """Save a calendar for a specific center."""
    raw = _load_raw_calendar()
    if "centros" not in raw:
        raw["centros"] = {}
    raw["centros"][centro_name] = data
    CALENDAR_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CALENDAR_FILE, "w", encoding="utf-8") as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)


def delete_center_calendar(centro_name):
    """Remove a center-specific calendar."""
    raw = _load_raw_calendar()
    if "centros" in raw and centro_name in raw["centros"]:
        del raw["centros"][centro_name]
        with open(CALENDAR_FILE, "w", encoding="utf-8") as f:
            json.dump(raw, f, ensure_ascii=False, indent=2)


def load_centros_trabajo():
    """Load the list of registered work centers."""
    if CENTROS_FILE.exists():
        with open(CENTROS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_centros_trabajo(centros):
    """Save the list of registered work centers."""
    CENTROS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(CENTROS_FILE, "w", encoding="utf-8") as f:
        json.dump(centros, f, ensure_ascii=False, indent=2)


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


# =============================================================================
# 3. EMPLOYEE LIST — optional override (persistent JSON)
# =============================================================================
def load_employee_list():
    """Load saved employee list: {centro_name: [name1, name2, ...]}"""
    if EMPLOYEES_FILE.exists():
        with open(EMPLOYEES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def save_employee_list(data):
    EMPLOYEES_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(EMPLOYEES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def generate_employee_template():
    """Generate an Excel template for employee lists per center."""
    df = pd.DataFrame({
        "Centro": ["ejemplo_centro1", "ejemplo_centro1", "ejemplo_centro2"],
        "Empleado": ["GARCIA LOPEZ, JUAN", "MARTINEZ RUIZ, ANA", "PEREZ GIL, CARLOS"],
    })
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Empleados")
        wb = writer.book
        ws = writer.sheets["Empleados"]
        hdr_f = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white", "border": 1})
        normal_f = wb.add_format({"border": 1})
        for ci, cn in enumerate(df.columns):
            ws.write(0, ci, cn, hdr_f)
        for ri in range(len(df)):
            for ci in range(len(df.columns)):
                ws.write(ri + 1, ci, df.iloc[ri, ci], normal_f)
        ws.set_column(0, 0, 25)
        ws.set_column(1, 1, 35)
        ws.freeze_panes(1, 0)
    return buf.getvalue()


def generate_cuadrante_template(year, month):
    """Generate a blank monthly schedule template matching the format the parser expects."""
    max_day = calendar.monthrange(year, month)[1]
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        ws = wb.add_worksheet("Cuadrante")

        # Formats
        title_f = wb.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter", "bg_color": "#1F3864", "font_color": "white"})
        hdr_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "align": "center", "valign": "vcenter"})
        hdr_we_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#D6DCE4", "font_color": "#1F3864", "border": 1, "align": "center", "valign": "vcenter"})
        name_f = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#D9E2F3", "valign": "vcenter"})
        cell_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "valign": "vcenter"})
        legend_title_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1})
        legend_f = wb.add_format({"font_size": 9, "border": 1, "text_wrap": True})

        month_name = MONTH_NAMES[month]
        day_abbr = ["L", "M", "X", "J", "V", "S", "D"]

        # Row 0: title
        ws.merge_range(0, 0, 0, max_day, f"CUADRANTE — {month_name.upper()} {year}", title_f)
        ws.set_row(0, 30)

        # Row 1: header — "Empleado" + day numbers
        ws.write(1, 0, "Empleado", hdr_f)
        for d in range(1, max_day + 1):
            wd = date(year, month, d).weekday()
            fmt = hdr_we_f if wd >= 5 else hdr_f
            ws.write(1, d, d, fmt)
        ws.set_row(1, 22)

        # Row 2: day-of-week letters
        ws.write(2, 0, "", hdr_f)
        for d in range(1, max_day + 1):
            wd = date(year, month, d).weekday()
            fmt = hdr_we_f if wd >= 5 else hdr_f
            ws.write(2, d, day_abbr[wd], fmt)

        # Rows 3-22: 20 blank employee rows
        for ri in range(20):
            ws.write(3 + ri, 0, "", name_f)
            for d in range(1, max_day + 1):
                ws.write(3 + ri, d, "", cell_f)

        # Column widths
        ws.set_column(0, 0, 30)
        ws.set_column(1, max_day, 4.5)

        # Legend below
        legend_row = 25
        ws.merge_range(legend_row, 0, legend_row, 5, "CÓDIGOS DE AUSENCIA", legend_title_f)
        codes = [
            ("Número (8, 7.5, etc.)", "Jornada trabajada (horas)"),
            ("V", "Vacaciones"),
            ("B", "Baja (IT / enfermedad)"),
            ("AP", "Asuntos propios"),
            ("P", "Permiso no retribuido"),
            ("PR", "Permiso retribuido"),
            ("E", "Excedencia"),
        ]
        for i, (code, desc) in enumerate(codes):
            ws.write(legend_row + 1 + i, 0, code, legend_f)
            ws.merge_range(legend_row + 1 + i, 1, legend_row + 1 + i, 5, desc, legend_f)

        ws.freeze_panes(3, 1)
        ws.print_area(0, 0, 22, max_day)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)

    return buf.getvalue()


def parse_employee_upload(data_bytes):
    """Parse uploaded employee Excel → {centro: [name1, name2, ...]}"""
    df = pd.read_excel(BytesIO(data_bytes), sheet_name=0)
    if len(df.columns) < 2:
        return None
    result = {}
    for _, row in df.iterrows():
        centro = str(row.iloc[0]).strip()
        nombre = str(row.iloc[1]).strip()
        if centro and nombre and centro.lower() != "nan" and nombre.lower() != "nan":
            result.setdefault(centro, []).append(nombre.upper())
    return result if result else None


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
    "horas", "plantilla", "resumen", "cuadrante", "horario", "calendario",
    "jornada", "personal", "media", "promedio", "acumulado", "periodo",
    "fecha", "firma", "responsable", "supervisor", "encargado", "gerente",
    "director", "jefe", "coordinador", "subtotal", "nº", "num",
}

# Patterns that indicate a title/section row rather than an employee name
_TITLE_PATTERNS = re.compile(
    r"(?i)^(enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
    r"septiembre|octubre|noviembre|diciembre|"
    r"cuadrante|horario|turno|secci[oó]n|departamento|"
    r"centro|almac[eé]n|planta|nave|total\b|"
    r"20\d{2}|horas\s|d[ií]as\s|plantilla\b|"
    r"observ|notas?$|resumen|periodo)"
)


def _is_title_row(name):
    """Detect if a row name is likely a title/section header, not an employee."""
    low = name.lower().strip()
    if low in SKIP_NAMES:
        return True
    # Employee names always start with a letter (A-Z, Ñ, Á, etc.)
    # This catches ∑, Σ, symbols, numbers, and any other non-name row
    if name and not name[0].isalpha():
        return True
    # Matches known title patterns (months, departments, etc.)
    if _TITLE_PATTERNS.search(low):
        return True
    return False


def _normalize_name(name):
    """Normalize an employee name for duplicate comparison."""
    n = name.upper().strip()
    n = re.sub(r'\s+', ' ', n)
    return n


def detect_duplicates(all_employee_details):
    """
    Detect employees appearing in multiple centers.
    Returns dict {normalized_name: [(centro, emp_data), ...]} for names in 2+ centers.
    """
    name_map = {}
    for centro, employees in all_employee_details.items():
        for emp in employees:
            norm = _normalize_name(emp["nombre"])
            name_map.setdefault(norm, []).append((centro, emp))
    duplicates = {}
    for name, entries in name_map.items():
        centers = set(c for c, _ in entries)
        if len(centers) > 1:
            duplicates[name] = entries
    return duplicates


def merge_employee_records(all_employee_details):
    """
    Merge employees from all centers, removing duplicates.
    For duplicate employees, keeps the record with more total days.
    """
    seen = {}
    for centro, employees in all_employee_details.items():
        for emp in employees:
            norm = _normalize_name(emp["nombre"])
            total = emp["worked"] + sum(emp[c] for c in ABSENCE_CODES)
            if norm not in seen or total > seen[norm][1]:
                seen[norm] = (emp, total)
    merged = [emp for emp, _ in seen.values()]
    merged.sort(key=lambda e: e["nombre"])
    return merged


def classify_cell(val):
    if pd.isna(val):
        return "empty", True
    s = str(val).strip()
    if s == "":
        return "empty", True
    try:
        num = float(s.replace(",", "."))
        if num > 0:
            return "worked", False
        # 0 hours = registered but not working (e.g., listed in center but works elsewhere)
        return "zero", False
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

    for ri in range(day_row_idx + 1, len(df_raw)):
        name_val = df_raw.iloc[ri, name_col]
        if pd.isna(name_val) or str(name_val).strip() == "":
            row_data = df_raw.iloc[ri, min_day_col:max(day_cols.values()) + 1]
            if row_data.isna().all() or (row_data.astype(str).str.strip() == "").all():
                continue
            # Check if unnamed row looks like a summary (values > 1 suggest sums, not daily entries)
            numeric_vals = pd.to_numeric(row_data, errors="coerce").dropna()
            if len(numeric_vals) > 0 and (numeric_vals > 1).sum() > len(numeric_vals) * 0.5:
                continue  # likely a sum/total row
            name_val = f"Empleado fila {ri + 1}"

        name = str(name_val).strip()
        # Skip header/title/section rows
        if _is_title_row(name):
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
            elif cat == "zero":
                pass  # 0 hours: registered but not working — don't count
            elif cat in ABSENCE_CODES:
                emp[cat] += 1
            elif cat == "empty":
                if day_num not in nw:
                    emp["empty"] += 1
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
    plantilla_efectiva = round(total_worked / w_days, 2) if w_days > 0 else 0

    ausencias_con_vac = total_V + total_B + total_AP + total_P + total_E
    pct_absent_con = (ausencias_con_vac / dias_teoricos * 100) if dias_teoricos > 0 else 0

    ausencias_sin_vac = total_B + total_AP + total_P + total_E
    pct_absent_sin = (ausencias_sin_vac / dias_teoricos * 100) if dias_teoricos > 0 else 0

    return {
        "plantilla": n,
        "plantilla_efectiva": plantilla_efectiva,
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


# Reverse mapping: month name → number
_MONTH_NAME_TO_NUM = {v.upper(): k for k, v in MONTH_NAMES.items()}


def parse_report_excel(data_bytes):
    """Parse an exported absentismo Excel report back into a history entry.

    Reads the 'Resumen' sheet and extracts KPIs per center.
    Returns a history-compatible dict or None on failure.
    """
    try:
        df = pd.read_excel(BytesIO(data_bytes), sheet_name="Resumen", header=None)
    except Exception:
        return None

    # Row 0: title — extract month and year
    title = str(df.iloc[0, 0]) if not pd.isna(df.iloc[0, 0]) else ""
    # Pattern: "ANÁLISIS DE ABSENTISMO — MES AÑO"
    m = re.search(r"—\s*(\w+)\s+(\d{4})", title)
    if not m:
        return None
    month_name, year_str = m.group(1).upper(), m.group(2)
    month_num = _MONTH_NAME_TO_NUM.get(month_name)
    if not month_num:
        return None
    year = int(year_str)

    # Row 2: headers — center names (skip col 0 which is empty, skip "TOTAL")
    header_row = df.iloc[2]
    centro_names = []
    for ci in range(1, len(header_row)):
        val = str(header_row.iloc[ci]).strip() if not pd.isna(header_row.iloc[ci]) else ""
        if val and val.upper() != "TOTAL":
            centro_names.append((ci, val))

    if not centro_names:
        return None

    # Metric rows start at row 3, in the same order as the export
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
        # Percentages are in rows after a blank row (row 3+11=14 is blank, 15 and 16)
        pct_row_con = 3 + len(metric_keys) + 1  # skip blank row
        pct_row_sin = pct_row_con + 1
        pct_con_val = df.iloc[pct_row_con, ci] if pct_row_con < len(df) else 0
        pct_sin_val = df.iloc[pct_row_sin, ci] if pct_row_sin < len(df) else 0
        # Values may be decimal (0.0523) or percentage (5.23)
        pct_con_v = float(pct_con_val) if not pd.isna(pct_con_val) else 0
        pct_sin_v = float(pct_sin_val) if not pd.isna(pct_sin_val) else 0
        # If stored as decimal fraction (< 1), convert to percentage
        if pct_con_v < 1:
            pct_con_v = round(pct_con_v * 100, 2)
        if pct_sin_v < 1:
            pct_sin_v = round(pct_sin_v * 100, 2)
        centro_data["pct_con"] = pct_con_v
        centro_data["pct_sin"] = pct_sin_v
        centros.append(centro_data)

    total_plantilla = sum(c["plantilla"] for c in centros)
    total_dias_teo = sum(c["plantilla"] * c["dias_laborables"] for c in centros)
    total_aus_con = sum(c["total_ausencias_con_vac"] for c in centros)
    total_aus_sin = sum(c["total_ausencias_sin_vac"] for c in centros)

    return {
        "key": f"{year}-{month_num:02d}",
        "year": year,
        "month": month_num,
        "centros": centros,
        "total_pct_con": round(total_aus_con / total_dias_teo * 100, 2) if total_dias_teo > 0 else 0,
        "total_pct_sin": round(total_aus_sin / total_dias_teo * 100, 2) if total_dias_teo > 0 else 0,
        "total_plantilla": total_plantilla,
    }


# =============================================================================
# 7. SIDEBAR — FILES + CONFIG (organized in collapsible sections)
# =============================================================================
st.sidebar.header("Absentismo")

# ── 1. SUBIDA DE CUADRANTES (siempre visible — es lo principal) ──
uploaded_files = st.sidebar.file_uploader(
    "Cuadrantes mensuales (uno por centro)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    key="abs_files",
)

if uploaded_files:
    SS["abs_file_data"] = [(f.name, f.getvalue()) for f in uploaded_files]
    st.sidebar.success(f"{len(uploaded_files)} fichero(s) cargado(s)")

# ── 2. PLANTILLA CUADRANTE EN BLANCO ──
with st.sidebar.expander("Plantilla cuadrante en blanco"):
    st.caption("Descarga una plantilla vacía del mes para cumplimentar.")
    _tc1, _tc2 = st.columns(2)
    tmpl_month = _tc1.selectbox("Mes", range(1, 13), format_func=lambda m: MONTH_NAMES[m], index=date.today().month - 1, key="tmpl_cuad_month")
    tmpl_year = _tc2.number_input("Año", 2020, 2030, date.today().year, key="tmpl_cuad_year")
    st.download_button(
        "Descargar cuadrante en blanco",
        generate_cuadrante_template(tmpl_year, tmpl_month),
        file_name=f"cuadrante_{MONTH_NAMES[tmpl_month].lower()}_{tmpl_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── 3. CALENDARIO LABORAL + CENTROS DE TRABAJO ──
custom_cal = load_custom_calendar()
center_cals = load_center_calendars()

_cal_status = ""
if custom_cal:
    _cal_status += "Default "
if center_cals:
    _cal_status += f"+ {len(center_cals)} centro(s)"
_cal_label = f"Calendario laboral ({_cal_status.strip()})" if _cal_status.strip() else "Calendario laboral"

with st.sidebar.expander(_cal_label):
    if custom_cal:
        st.success("Calendario por defecto cargado")
    if center_cals:
        st.info(f"Calendarios por centro: {', '.join(sorted(center_cals.keys()))}")

    # Download template
    cal_year_tmpl = st.number_input("Año plantilla", 2020, 2030, date.today().year, key="cal_year_tmpl")
    st.download_button(
        "Descargar plantilla calendario",
        generate_calendar_template(cal_year_tmpl),
        file_name=f"calendario_laboral_{cal_year_tmpl}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.divider()

    _cal_mode = st.radio(
        "Subir calendario para:",
        ["Por defecto (todos)", "Centro específico"],
        key="cal_upload_mode",
        horizontal=True,
    )

    if _cal_mode == "Centro específico":
        centros_registrados = load_centros_trabajo()
        _cal_centro_opts = centros_registrados if centros_registrados else []
        _cal_centro_name = st.text_input(
            "Nombre del centro",
            placeholder="Ej: Almacén Pamplona",
            key="cal_centro_name",
            help="Debe coincidir con el nombre del fichero del cuadrante (sin extensión)." + (
                f" Centros registrados: {', '.join(_cal_centro_opts)}" if _cal_centro_opts else ""
            ),
        )
        cal_upload = st.file_uploader("Subir calendario del centro", type=["xlsx", "xls"], key="cal_up_centro")
        if cal_upload and _cal_centro_name.strip():
            parsed_cal = parse_calendar_upload(cal_upload.getvalue())
            if parsed_cal:
                save_center_calendar(_cal_centro_name.strip(), parsed_cal)
                center_cals = load_center_calendars()
                st.success(f"Calendario de **{_cal_centro_name.strip()}** guardado")
            else:
                st.error("No se pudo leer el calendario")
        elif cal_upload and not _cal_centro_name.strip():
            st.warning("Indica el nombre del centro antes de subir")
    else:
        cal_upload = st.file_uploader("Subir calendario cumplimentado", type=["xlsx", "xls"], key="cal_up")
        if cal_upload:
            parsed_cal = parse_calendar_upload(cal_upload.getvalue())
            if parsed_cal:
                save_custom_calendar(parsed_cal)
                custom_cal = parsed_cal
                st.success("Calendario por defecto guardado")
            else:
                st.error("No se pudo leer el calendario")

    # Delete calendars
    if custom_cal or center_cals:
        st.divider()
        st.caption("**Eliminar calendarios**")
        if custom_cal:
            if st.button("Borrar calendario por defecto", key="del_cal_default", type="secondary", use_container_width=True):
                raw = _load_raw_calendar()
                raw["__default__"] = None
                with open(CALENDAR_FILE, "w", encoding="utf-8") as f:
                    json.dump(raw, f, ensure_ascii=False, indent=2)
                st.rerun()
        for cname in sorted(center_cals.keys()):
            if st.button(f"Borrar: {cname}", key=f"del_cal_{cname}", type="secondary", use_container_width=True):
                delete_center_calendar(cname)
                st.rerun()

# ── 4. CENTROS DE TRABAJO (removed — now hardcoded at page top) ──

# ── 5. PLANTILLA DE EMPLEADOS ──
employee_override = load_employee_list()
_emp_label = "Plantilla de empleados"
if employee_override:
    total_emp = sum(len(v) for v in employee_override.values())
    _emp_label += f" ({total_emp} emp.)"

with st.sidebar.expander(_emp_label):
    st.caption("Opcional: sube una lista con los empleados reales de cada centro para evitar que se cuenten filas de resumen o títulos.")
    if employee_override:
        st.success(f"{total_emp} empleados en {len(employee_override)} centro(s)")

    st.download_button(
        "Descargar plantilla empleados",
        generate_employee_template(),
        file_name="plantilla_empleados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    emp_upload = st.file_uploader("Subir plantilla cumplimentada", type=["xlsx", "xls"], key="emp_up")
    if emp_upload:
        parsed_emp = parse_employee_upload(emp_upload.getvalue())
        if parsed_emp:
            save_employee_list(parsed_emp)
            employee_override = parsed_emp
            total_emp = sum(len(v) for v in parsed_emp.values())
            st.success(f"Guardado: {total_emp} empleados en {len(parsed_emp)} centro(s)")
        else:
            st.error("No se pudo leer la plantilla. Verifica que tenga columnas Centro y Empleado.")

    if employee_override:
        if st.button("Borrar lista de empleados", type="secondary", use_container_width=True):
            if EMPLOYEES_FILE.exists():
                EMPLOYEES_FILE.unlink()
            employee_override = None
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
cal_info = f"**{MONTH_NAMES[sel_month]} {sel_year}**: {w_days} días laborables — {cal_source}"
if month_holidays and not custom_cal:
    cal_info += f" · Festivos: {', '.join(f'{d.day}-{n}' for d, n in sorted(month_holidays.items()))}"
if center_cals:
    cal_info += f" · Calendarios específicos: {', '.join(sorted(center_cals.keys()))}"
st.caption(cal_info)

# =============================================================================
# 9. PROCESS
# =============================================================================
if st.button("Analizar Absentismo", type="primary", use_container_width=True):
    all_results = []
    all_warnings = {}
    all_employee_details = {}

    for fname, fdata in SS["abs_file_data"]:
        centro_name = fname.rsplit(".", 1)[0]
        # Use center-specific calendar if available, else default
        centro_cal = get_calendar_for_centro(centro_name, custom_cal)
        centro_nw_days = get_non_working_days(sel_year, sel_month, centro_cal)
        parsed, warns = parse_cuadrante(fdata, fname, non_working_days=centro_nw_days)
        if parsed is None:
            st.error(f"Error procesando **{fname}**: {'; '.join(w if isinstance(w, str) else w.get('mensaje', '') for w in warns)}")
            continue
        # Filter employees by override list if available
        if employee_override:
            allowed = None
            # Try exact match first, then case-insensitive partial match
            for key in employee_override:
                if key.upper() == centro_name.upper() or key.upper() in centro_name.upper() or centro_name.upper() in key.upper():
                    allowed = {n.upper() for n in employee_override[key]}
                    break
            if allowed:
                parsed["employees"] = [e for e in parsed["employees"] if e["nombre"].upper() in allowed]
        kpis = calculate_kpis(parsed, sel_year, sel_month, centro_cal)
        kpis["centro"] = centro_name
        all_results.append(kpis)
        all_warnings[centro_name] = warns
        all_employee_details[centro_name] = parsed["employees"]

    if all_results:
        # Detect duplicates across centers
        duplicates = detect_duplicates(all_employee_details)

        SS["abs_results"] = {
            "kpis": all_results,
            "warnings": all_warnings,
            "employees": all_employee_details,
            "month": sel_month,
            "year": sel_year,
            "duplicates": duplicates,
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

    # --- Warnings (only unknown codes) ---
    for centro, warns in all_warnings.items():
        for w in warns:
            if w["tipo"] == "Códigos desconocidos":
                st.error(f"**{centro}**: {w['mensaje']}")

    # --- KPI Cards ---
    st.header(f"Resumen — {MONTH_NAMES[r_month]} {r_year}")

    total_plantilla = sum(k["plantilla"] for k in kpis_list)
    total_dias_teo = sum(k["dias_teoricos"] for k in kpis_list)
    total_trabajados = sum(k["dias_trabajados"] for k in kpis_list)
    total_ausencias_con = sum(k["total_ausencias_con_vac"] for k in kpis_list)
    total_ausencias_sin = sum(k["total_ausencias_sin_vac"] for k in kpis_list)
    pct_con = round(total_ausencias_con / total_dias_teo * 100, 2) if total_dias_teo > 0 else 0
    pct_sin = round(total_ausencias_sin / total_dias_teo * 100, 2) if total_dias_teo > 0 else 0

    # Sum individual plantilla efectiva (each center may have different working days)
    total_plantilla_efectiva = round(sum(k["plantilla_efectiva"] for k in kpis_list), 2)

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Plantilla total", total_plantilla)
    m2.metric("Plantilla efectiva", f"{total_plantilla_efectiva:.1f}")
    m3.metric("Días trabajados", f"{total_trabajados:,}")
    m4.metric("Absentismo (con vac.)", f"{pct_con:.2f}%")
    m5.metric("Absentismo (sin vac.)", f"{pct_sin:.2f}%")

    # --- Duplicate detection & interactive merge ---
    duplicates = res.get("duplicates", {})
    merged_info = res.get("merged_centers", [])

    # Show success message if merge was just applied
    if merged_info:
        removed_names = res.get("merged_duplicates_removed", [])
        st.divider()
        st.success(
            f"**Centros fusionados**: {', '.join(merged_info)} — "
            f"Se eliminaron **{len(removed_names)} empleado(s) repetido(s)**."
        )
        if removed_names:
            with st.expander("Empleados deduplicados"):
                for norm_name in removed_names:
                    orig_dups = res.get("original_duplicates", {})
                    if norm_name in orig_dups:
                        display = orig_dups[norm_name][0][1]["nombre"]
                    else:
                        display = norm_name
                    st.markdown(f"- {display}")

    # Show duplicates and merge UI (only when there are 2+ separate centers)
    if duplicates and len(kpis_list) > 1:
        st.divider()
        st.warning(f"**{len(duplicates)} empleado(s) repetido(s)** detectados en varios centros")
        with st.expander(f"Ver empleados repetidos ({len(duplicates)})", expanded=True):
            dup_rows = []
            for name, entries in sorted(duplicates.items()):
                centros_list = [c for c, _ in entries]
                dup_rows.append({
                    "Empleado": entries[0][1]["nombre"],
                    "Aparece en centros": ", ".join(centros_list),
                    "Nº centros": len(centros_list),
                })
            st.dataframe(pd.DataFrame(dup_rows), use_container_width=True, hide_index=True)

        # --- Merge UI ---
        st.subheader("Fusionar centros")
        st.caption("Selecciona los centros que comparten plantilla para unificarlos y eliminar repetidos.")
        available_centers = [k["centro"] for k in kpis_list]
        selected_to_merge = st.multiselect(
            "Centros a fusionar:",
            available_centers,
            key="abs_merge_selection",
        )

        if len(selected_to_merge) >= 2:
            # Preview which duplicates would be removed
            merge_dups = {}
            for name, entries in duplicates.items():
                relevant_centers = set(c for c, _ in entries if c in selected_to_merge)
                if len(relevant_centers) > 1:
                    merge_dups[name] = [(c, e) for c, e in entries if c in selected_to_merge]

            if merge_dups:
                # === VALIDATE: compare day distributions ===
                _COMPARE_KEYS = ["worked", "V", "B", "AP", "P", "PR", "E"]
                _COMPARE_LABELS = ["Trabajados", "Vacaciones", "Bajas", "A.Propios", "Permiso", "P.Retrib.", "Excedencia"]
                matching = {}   # name → True if all centers match
                mismatch_details = []  # rows for mismatch table

                for norm_name, entries in merge_dups.items():
                    display_name = entries[0][1]["nombre"]
                    first_centro, first_emp = entries[0]
                    all_match = True
                    for other_centro, other_emp in entries[1:]:
                        for k in _COMPARE_KEYS:
                            if first_emp.get(k, 0) != other_emp.get(k, 0):
                                all_match = False
                                break
                    matching[norm_name] = all_match

                    if not all_match:
                        for centro, emp in entries:
                            mismatch_details.append({
                                "Empleado": display_name,
                                "Centro": centro,
                                **{lab: emp.get(k, 0) for k, lab in zip(_COMPARE_KEYS, _COMPARE_LABELS)},
                            })

                n_ok = sum(1 for v in matching.values() if v)
                n_bad = sum(1 for v in matching.values() if not v)

                if n_bad > 0:
                    st.error(
                        f"**{n_bad} empleado(s) con distribución de días DIFERENTE** entre centros. "
                        f"Revisa antes de fusionar — al fusionar se mantiene solo un registro y se pierde el otro."
                    )
                    df_mismatch = pd.DataFrame(mismatch_details)
                    st.dataframe(df_mismatch, use_container_width=True, hide_index=True)

                    # Show which are OK
                    if n_ok > 0:
                        ok_names = [merge_dups[n][0][1]["nombre"] for n, v in matching.items() if v]
                        st.success(f"**{n_ok} empleado(s) coinciden** perfectamente: {', '.join(ok_names)}")

                    st.warning("Puedes fusionar de todos modos, pero los registros que no coincidan perderán datos del centro descartado.")
                    _allow_merge = st.checkbox(
                        "Entiendo el riesgo, fusionar igualmente",
                        key="force_merge_mismatch",
                    )
                else:
                    ok_names = [merge_dups[n][0][1]["nombre"] for n in merge_dups]
                    st.success(
                        f"**{n_ok} empleado(s) coinciden** perfectamente entre centros: {', '.join(ok_names)}. "
                        f"Fusión segura."
                    )
                    _allow_merge = True
            else:
                st.info("No hay empleados repetidos entre los centros seleccionados.")
                _allow_merge = True

            if _allow_merge and st.button("Fusionar seleccionados", type="primary", use_container_width=True):
                merged_name = " + ".join(selected_to_merge)
                merge_emp_details = {c: all_employees[c] for c in selected_to_merge if c in all_employees}
                merged_employees = merge_employee_records(merge_emp_details)
                merged_parsed = {"employees": merged_employees}
                merged_kpis = calculate_kpis(merged_parsed, r_year, r_month, get_calendar_for_centro(merged_name, custom_cal))
                merged_kpis["centro"] = merged_name

                new_kpis = [k for k in kpis_list if k["centro"] not in selected_to_merge] + [merged_kpis]
                new_employees = {c: emps for c, emps in all_employees.items() if c not in selected_to_merge}
                new_employees[merged_name] = merged_employees

                # Recalculate duplicates for new configuration
                new_duplicates = detect_duplicates(new_employees)

                SS["abs_results"] = {
                    "kpis": new_kpis,
                    "warnings": res["warnings"],
                    "employees": new_employees,
                    "month": r_month,
                    "year": r_year,
                    "duplicates": new_duplicates,
                    "original_duplicates": duplicates,
                    "merged_centers": selected_to_merge,
                    "merged_duplicates_removed": list(merge_dups.keys()),
                }
                st.rerun()

        elif len(selected_to_merge) == 1:
            st.caption("Selecciona al menos 2 centros para fusionar.")

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
                # Show range if centers have different working days
                all_wdays = set(k["dias_laborables"] for k in kpis_list)
                avg_wdays = round(sum(k["dias_laborables"] for k in kpis_list) / len(kpis_list)) if kpis_list else 0
                consol = {
                    "plantilla": total_plantilla,
                    "dias_laborables": avg_wdays,
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
            title_f = wb.add_format({"bold": True, "font_size": 14, "bg_color": "#1F3864", "font_color": "white", "align": "center", "valign": "vcenter"})
            hdr_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
            lbl_f = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#D6DCE4"})
            num_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0"})
            # Yellow highlight for plantilla efectiva
            lbl_yellow = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#FFD966", "font_color": "#1F3864"})
            num_yellow = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0.0", "bg_color": "#FFF2CC"})
            # Red highlight for total ausencias
            lbl_red = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#FFC7CE", "font_color": "#9C0006"})
            num_red = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0", "bg_color": "#FFC7CE", "font_color": "#9C0006"})
            # Percentage formats
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
            ncols = len(kpis_list) + (1 if len(kpis_list) > 1 else 0)
            ws.merge_range(row, 0, row, ncols,
                           f"ANÁLISIS DE ABSENTISMO — {MONTH_NAMES[r_month].upper()} {r_year}", title_f)
            ws.set_row(row, 30)
            row += 2

            headers = [""] + [k["centro"] for k in kpis_list]
            if len(kpis_list) > 1:
                headers.append("TOTAL")
            for ci, h in enumerate(headers):
                ws.write(row, ci, h, hdr_f)
            ws.set_row(row, 25)
            row += 1

            # Metrics with per-row formatting
            metrics = [
                ("Plantilla", "plantilla", lbl_f, num_f),
                ("Plantilla efectiva", "plantilla_efectiva", lbl_yellow, num_yellow),
                ("Días laborables", "dias_laborables", lbl_f, num_f),
                ("Días trabajados", "dias_trabajados", lbl_f, num_f),
                ("Vacaciones (días)", "dias_vacaciones", lbl_f, num_f),
                ("Bajas (días)", "dias_baja", lbl_f, num_f),
                ("Asuntos Propios (días)", "dias_ap", lbl_f, num_f),
                ("Permisos (días)", "dias_permiso", lbl_f, num_f),
                ("Excedencias (días)", "dias_excedencia", lbl_f, num_f),
                ("Total ausencias (con vac.)", "total_ausencias_con_vac", lbl_red, num_red),
                ("Total ausencias (sin vac.)", "total_ausencias_sin_vac", lbl_red, num_red),
            ]

            for label, key, lfmt, nfmt in metrics:
                ws.write(row, 0, label, lfmt)
                for ci, k in enumerate(kpis_list):
                    ws.write(row, ci + 1, k[key], nfmt)
                if len(kpis_list) > 1:
                    total_val = sum(k[key] for k in kpis_list)
                    if key == "plantilla_efectiva":
                        # Sum individual center effective headcounts
                        total_val = round(sum(k["plantilla_efectiva"] for k in kpis_list), 1)
                    elif key == "dias_laborables":
                        # Show average when centers differ
                        total_val = round(total_val / len(kpis_list))
                    ws.write(row, len(kpis_list) + 1, total_val, nfmt)
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
            ws.merge_range(row, 0, row, ncols,
                           "Leyenda: Verde = < 5% absentismo · Sin color = 5%-10% · Rojo = > 10%", legend_f)
            row += 1
            ws.merge_range(row, 0, row, ncols,
                           "Plantilla efectiva = Días trabajados ÷ Días laborables. "
                           "Indica el nº medio de personas que realmente han trabajado cada día del mes.", legend_f)

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
                    "plantilla_efectiva": k["plantilla_efectiva"],
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

    # --- Download historical Excel ---
    def build_historical_excel(hist_data):
        """Build a comprehensive Excel with historical analysis, charts, and per-month detail tabs."""
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            wb = writer.book

            # === Formats ===
            title_f = wb.add_format({"bold": True, "font_size": 16, "bg_color": "#1F3864", "font_color": "white", "align": "center", "valign": "vcenter"})
            subtitle_f = wb.add_format({"bold": True, "font_size": 12, "bg_color": "#2E75B6", "font_color": "white", "align": "center", "valign": "vcenter"})
            hdr_f = wb.add_format({"bold": True, "font_size": 10, "bg_color": "#1F3864", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"})
            lbl_f = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#D6DCE4"})
            num_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0"})
            num_dec_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0.0"})
            pct_f = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%"})
            pct_green = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#006100", "bg_color": "#C6EFCE"})
            pct_yellow = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#9C6500", "bg_color": "#FFEB9C"})
            pct_red = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "0.00%", "font_color": "#9C0006", "bg_color": "#FFC7CE"})
            lbl_yellow = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#FFD966", "font_color": "#1F3864"})
            num_yellow = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0.0", "bg_color": "#FFF2CC"})
            lbl_red = wb.add_format({"bold": True, "font_size": 10, "border": 1, "bg_color": "#FFC7CE", "font_color": "#9C0006"})
            num_red = wb.add_format({"font_size": 10, "border": 1, "align": "center", "num_format": "#,##0", "bg_color": "#FFC7CE", "font_color": "#9C0006"})
            legend_f = wb.add_format({"font_size": 9, "italic": True, "text_wrap": True})

            def wpct(ws, r, c, v):
                fmt = pct_green if v < 0.05 else (pct_red if v > 0.10 else pct_yellow)
                ws.write_number(r, c, v, fmt)

            # ===============================================================
            # SHEET 1: Resumen Histórico (trend table + chart)
            # ===============================================================
            sn = "Resumen Histórico"
            pd.DataFrame().to_excel(writer, sheet_name=sn, index=False)
            ws = writer.sheets[sn]

            # Collect all unique centers across history
            all_centros_hist = []
            for h in hist_data:
                for c in h["centros"]:
                    if c["centro"] not in all_centros_hist:
                        all_centros_hist.append(c["centro"])

            # Title
            row = 0
            ncols = 4 + len(all_centros_hist) * 2  # Mes, Plantilla, %Con, %Sin + per-center %con/%sin
            ws.merge_range(row, 0, row, max(ncols, 5),
                           "REPORTE HISTÓRICO DE ABSENTISMO", title_f)
            ws.set_row(row, 35)
            row += 2

            # --- Trend table ---
            # Headers
            trend_headers = ["Mes", "Plantilla"]
            if len(all_centros_hist) > 1:
                for cn in all_centros_hist:
                    short = cn[:15] if len(cn) > 15 else cn
                    trend_headers.append(f"% Con Vac.\n{short}")
                    trend_headers.append(f"% Sin Vac.\n{short}")
            trend_headers.append("% Con Vac. TOTAL")
            trend_headers.append("% Sin Vac. TOTAL")

            for ci, h_txt in enumerate(trend_headers):
                ws.write(row, ci, h_txt, hdr_f)
            ws.set_row(row, 30)
            row += 1

            data_start_row = row
            for h in hist_data:
                label = f"{MONTH_NAMES[h['month']]} {h['year']}"
                ws.write(row, 0, label, lbl_f)
                ws.write(row, 1, h["total_plantilla"], num_f)
                ci = 2
                if len(all_centros_hist) > 1:
                    for cn in all_centros_hist:
                        centro_data = next((c for c in h["centros"] if c["centro"] == cn), None)
                        if centro_data:
                            wpct(ws, row, ci, centro_data["pct_con"] / 100)
                            wpct(ws, row, ci + 1, centro_data["pct_sin"] / 100)
                        else:
                            ws.write(row, ci, "", num_f)
                            ws.write(row, ci + 1, "", num_f)
                        ci += 2
                wpct(ws, row, ci, h["total_pct_con"] / 100)
                wpct(ws, row, ci + 1, h["total_pct_sin"] / 100)
                row += 1
            data_end_row = row - 1

            # Column widths
            ws.set_column(0, 0, 20)
            ws.set_column(1, 1, 12)
            for ci in range(2, len(trend_headers)):
                ws.set_column(ci, ci, 16)

            # --- Chart: Absenteeism trend ---
            if len(hist_data) >= 2:
                row += 1
                chart = wb.add_chart({"type": "line"})
                chart.set_title({"name": "Evolución del Absentismo"})
                chart.set_y_axis({"name": "% Absentismo", "num_format": "0.00%"})
                chart.set_x_axis({"name": ""})
                chart.set_size({"width": 800, "height": 420})
                chart.set_legend({"position": "bottom"})

                # Total con vacaciones
                total_con_col = len(trend_headers) - 2
                total_sin_col = len(trend_headers) - 1
                chart.add_series({
                    "name": "Con vacaciones (Total)",
                    "categories": [sn, data_start_row, 0, data_end_row, 0],
                    "values": [sn, data_start_row, total_con_col, data_end_row, total_con_col],
                    "line": {"color": "#FF5722", "width": 2.5},
                    "marker": {"type": "circle", "size": 6, "fill": {"color": "#FF5722"}},
                })
                chart.add_series({
                    "name": "Sin vacaciones (Total)",
                    "categories": [sn, data_start_row, 0, data_end_row, 0],
                    "values": [sn, data_start_row, total_sin_col, data_end_row, total_sin_col],
                    "line": {"color": "#2196F3", "width": 2.5},
                    "marker": {"type": "circle", "size": 6, "fill": {"color": "#2196F3"}},
                })

                # Per-center series (sin vac only, to avoid clutter)
                colors_chart = ["#4CAF50", "#FF9800", "#9C27B0", "#00BCD4", "#795548", "#607D8B"]
                if len(all_centros_hist) > 1:
                    for idx, cn in enumerate(all_centros_hist):
                        sin_col = 2 + idx * 2 + 1  # %Sin Vac column for this center
                        chart.add_series({
                            "name": f"{cn} (sin vac.)",
                            "categories": [sn, data_start_row, 0, data_end_row, 0],
                            "values": [sn, data_start_row, sin_col, data_end_row, sin_col],
                            "line": {"color": colors_chart[idx % len(colors_chart)], "width": 1.5, "dash_type": "dash"},
                            "marker": {"type": "diamond", "size": 4},
                        })

                ws.insert_chart(row, 0, chart)
                row += 22  # chart height in rows

            # Legend
            row += 1
            ws.merge_range(row, 0, row, max(ncols, 5),
                           "Verde = < 5% · Amarillo = 5%-10% · Rojo = > 10%", legend_f)
            row += 1
            ws.merge_range(row, 0, row, max(ncols, 5),
                           "% Con Vac. = (V+B+AP+P+E) / Días teóricos × 100 · "
                           "% Sin Vac. = (B+AP+P+E) / Días teóricos × 100", legend_f)

            ws.set_landscape()
            ws.fit_to_pages(1, 0)

            # ===============================================================
            # SHEETS 2+: One per month with full detail
            # ===============================================================
            for h in hist_data:
                month_label = f"{MONTH_NAMES[h['month']]} {h['year']}"
                safe_sn = month_label[:31]
                pd.DataFrame().to_excel(writer, sheet_name=safe_sn, index=False)
                ws_m = writer.sheets[safe_sn]

                centros = h["centros"]
                ncols_m = len(centros) + (1 if len(centros) > 1 else 0)

                # Title
                row_m = 0
                ws_m.merge_range(row_m, 0, row_m, ncols_m,
                                 f"ANÁLISIS DE ABSENTISMO — {month_label.upper()}", title_f)
                ws_m.set_row(row_m, 30)
                row_m += 2

                # Headers
                hdrs = [""] + [c["centro"] for c in centros]
                if len(centros) > 1:
                    hdrs.append("TOTAL")
                for ci, h_txt in enumerate(hdrs):
                    ws_m.write(row_m, ci, h_txt, hdr_f)
                ws_m.set_row(row_m, 25)
                row_m += 1

                # Metric rows
                metrics_hist = [
                    ("Plantilla", "plantilla", lbl_f, num_f),
                    ("Plantilla efectiva", "plantilla_efectiva", lbl_yellow, num_yellow),
                    ("Días laborables", "dias_laborables", lbl_f, num_f),
                    ("Días trabajados", "dias_trabajados", lbl_f, num_f),
                    ("Vacaciones (días)", "dias_vacaciones", lbl_f, num_f),
                    ("Bajas (días)", "dias_baja", lbl_f, num_f),
                    ("Asuntos Propios (días)", "dias_ap", lbl_f, num_f),
                    ("Permisos (días)", "dias_permiso", lbl_f, num_f),
                    ("Excedencias (días)", "dias_excedencia", lbl_f, num_f),
                    ("Total ausencias (con vac.)", "total_ausencias_con_vac", lbl_red, num_red),
                    ("Total ausencias (sin vac.)", "total_ausencias_sin_vac", lbl_red, num_red),
                ]

                for label, key, lfmt, nfmt in metrics_hist:
                    ws_m.write(row_m, 0, label, lfmt)
                    for ci, c in enumerate(centros):
                        val = c.get(key, 0)
                        ws_m.write(row_m, ci + 1, val, nfmt)
                    if len(centros) > 1:
                        if key == "plantilla_efectiva":
                            t_worked = sum(c.get("dias_trabajados", 0) for c in centros)
                            t_wdays = centros[0].get("dias_laborables", 1) if centros else 1
                            total_val = round(t_worked / t_wdays, 1) if t_wdays > 0 else 0
                        elif key == "dias_laborables":
                            total_val = centros[0].get(key, 0) if centros else 0
                        else:
                            total_val = sum(c.get(key, 0) for c in centros)
                        ws_m.write(row_m, len(centros) + 1, total_val, nfmt)
                    row_m += 1

                # Percentages
                row_m += 1
                ws_m.write(row_m, 0, "% Absentismo CON vacaciones", lbl_f)
                for ci, c in enumerate(centros):
                    wpct(ws_m, row_m, ci + 1, c.get("pct_con", 0) / 100)
                if len(centros) > 1:
                    wpct(ws_m, row_m, len(centros) + 1, h.get("total_pct_con", 0) / 100)
                row_m += 1

                ws_m.write(row_m, 0, "% Absentismo SIN vacaciones", lbl_f)
                for ci, c in enumerate(centros):
                    wpct(ws_m, row_m, ci + 1, c.get("pct_sin", 0) / 100)
                if len(centros) > 1:
                    wpct(ws_m, row_m, len(centros) + 1, h.get("total_pct_sin", 0) / 100)

                # Mini chart per month (bar chart comparing centers)
                if len(centros) > 1:
                    row_m += 2
                    # Write data for chart
                    chart_data_row = row_m
                    ws_m.write(row_m, 0, "", lbl_f)
                    for ci, c in enumerate(centros):
                        ws_m.write(row_m, ci + 1, c["centro"], hdr_f)
                    row_m += 1
                    ws_m.write(row_m, 0, "% Con Vac.", lbl_f)
                    for ci, c in enumerate(centros):
                        ws_m.write_number(row_m, ci + 1, c.get("pct_con", 0) / 100, pct_f)
                    row_m += 1
                    ws_m.write(row_m, 0, "% Sin Vac.", lbl_f)
                    for ci, c in enumerate(centros):
                        ws_m.write_number(row_m, ci + 1, c.get("pct_sin", 0) / 100, pct_f)

                    chart_m = wb.add_chart({"type": "column"})
                    chart_m.set_title({"name": f"Absentismo por centro — {month_label}"})
                    chart_m.set_y_axis({"name": "%", "num_format": "0.00%"})
                    chart_m.set_size({"width": 600, "height": 350})
                    chart_m.set_legend({"position": "bottom"})
                    chart_m.add_series({
                        "name": "Con vacaciones",
                        "categories": [safe_sn, chart_data_row, 1, chart_data_row, len(centros)],
                        "values": [safe_sn, chart_data_row + 1, 1, chart_data_row + 1, len(centros)],
                        "fill": {"color": "#FF5722"},
                    })
                    chart_m.add_series({
                        "name": "Sin vacaciones",
                        "categories": [safe_sn, chart_data_row, 1, chart_data_row, len(centros)],
                        "values": [safe_sn, chart_data_row + 2, 1, chart_data_row + 2, len(centros)],
                        "fill": {"color": "#2196F3"},
                    })
                    row_m += 2
                    ws_m.insert_chart(row_m, 0, chart_m)
                    row_m += 18

                # Legend
                row_m += 2
                ws_m.merge_range(row_m, 0, row_m, ncols_m,
                                 "Verde = < 5% · Amarillo = 5%-10% · Rojo = > 10%", legend_f)
                row_m += 1
                ws_m.merge_range(row_m, 0, row_m, ncols_m,
                                 "Plantilla efectiva = Días trabajados / Días laborables", legend_f)

                # Column widths
                ws_m.set_column(0, 0, 30)
                for ci in range(1, ncols_m + 1):
                    ws_m.set_column(ci, ci, 18)
                ws_m.set_landscape()
                ws_m.fit_to_pages(1, 0)

        return output.getvalue()

    hist_excel = build_historical_excel(hist)
    st.download_button(
        "Descargar Reporte Histórico (Excel)",
        hist_excel,
        file_name="HISTORICO_ABSENTISMO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    # --- Manage history ---
    with st.expander("Gestionar historial"):
        # Upload a past report Excel
        st.markdown("**Subir reporte pasado**")
        st.caption("Sube un Excel de reporte de absentismo (generado por esta app) para añadirlo al historial.")
        hist_upload = st.file_uploader(
            "Excel de reporte", type=["xlsx", "xls"], key="hist_report_upload",
        )
        if hist_upload:
            parsed_entry = parse_report_excel(hist_upload.getvalue())
            if parsed_entry:
                month_label = f"{MONTH_NAMES[parsed_entry['month']]} {parsed_entry['year']}"
                st.success(f"Detectado: **{month_label}** — {len(parsed_entry['centros'])} centro(s)")
                if st.button(f"Guardar {month_label} en historial", type="primary"):
                    hist_current = load_abs_history()
                    hist_current = [h for h in hist_current if h.get("key") != parsed_entry["key"]]
                    hist_current.append(parsed_entry)
                    hist_current.sort(key=lambda h: h["key"])
                    save_abs_history(hist_current)
                    st.success(f"**{month_label}** guardado en historial")
                    st.rerun()
            else:
                st.error("No se pudo leer el reporte. Asegúrate de que tiene la hoja 'Resumen' con el formato correcto.")

        # Individual month deletion
        st.divider()
        st.markdown("**Eliminar meses individuales**")
        for h in hist:
            col_info, col_btn = st.columns([4, 1])
            with col_info:
                st.text(
                    f"{MONTH_NAMES[h['month']]} {h['year']} — "
                    f"Plantilla: {h['total_plantilla']} — "
                    f"Con vac: {h['total_pct_con']:.2f}% — Sin vac: {h['total_pct_sin']:.2f}%"
                )
            with col_btn:
                if st.button("Eliminar", key=f"del_hist_{h['key']}", type="secondary"):
                    hist_current = load_abs_history()
                    hist_current = [x for x in hist_current if x.get("key") != h["key"]]
                    save_abs_history(hist_current)
                    st.rerun()

        st.divider()
        if st.button("Limpiar todo el historial", type="secondary"):
            save_abs_history([])
            st.rerun()
