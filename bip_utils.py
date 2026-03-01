"""
bip_utils.py
============
Analisis de Recepciones BIP — funciones de normalizacion, tokenizacion
de proveedores, calculo de KPIs mensuales, Pareto y desglose mensual.

Portado del notebook ``Copy_of_BIP_1_PUNTO_MEJORADO_super.ipynb``
con las siguientes mejoras:
- aliases, force-split y paqueteria se cargan desde un Excel externo
  (3 hojas: Aliases, Force Split, Paqueteria)
- funciones puras que no dependen de variables globales
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter, defaultdict
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import numpy as np
import pandas as pd


# =====================================================================
# NORMALIZACION DE TEXTO
# =====================================================================

def _norm(s: str) -> str:
    """Quita acentos, pasa a minusculas, elimina sufijos legales."""
    if pd.isna(s):
        return ""
    t = (
        unicodedata.normalize("NFKD", str(s))
        .encode("ascii", "ignore")
        .decode()
        .lower()
    )
    # permitimos + / , ; & y 'y'
    t = re.sub(r"[^a-z0-9\-\+\&\/,; y]+", " ", t)
    t = re.sub(
        r"\b(s\.?l\.?|s\.?a\.?u?|s\.?a\.?|gmbh|llc|ltd|inc|srl|spa|sas|bv|ag|co|corp)\b",
        "",
        t,
    )
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _typo_fix(key: str) -> str:
    """Correcciones globales de typos sobre clave normalizada."""
    key = re.sub(r"\btechnink\b", "technik", key)
    return key


def _make_key(s: str) -> str:
    """Clave de busqueda: _norm → quitar guiones → colapsar espacios."""
    return re.sub(r"\s+", " ", _typo_fix(_norm(s).replace("-", " "))).strip()


def _clean_label(v: str) -> str:
    """Etiqueta final presentable: si contiene nordex → Nordex, si no → Title Case."""
    base = _norm(v)
    if "nordex" in base:
        return "Nordex"
    return base.title()


# =====================================================================
# CARGA DEL FICHERO EXTERNO DE ALIASES
# =====================================================================

def load_aliases(path: str | Path) -> tuple[dict, dict, set]:
    """
    Carga el Excel de aliases con 3 hojas.

    Hoja **Aliases**  : columnas ``Variante``, ``Nombre Canonico``
    Hoja **Force Split**: columnas ``Texto Original``, ``Proveedor 1`` .. ``Proveedor 4``
    Hoja **Paqueteria** : columnas ``Proveedor``, ``Tipo``

    Returns
    -------
    aliases : dict   {clave_normalizada: NombreCanonicoLimpio}
    force_split : dict  {clave_normalizada: [parte1, parte2, ...]}
    paqueteria : set   {NombreCanonicoLimpio, ...}
    """
    path = Path(path)

    # --- Aliases ---
    aliases: dict[str, str] = {}
    try:
        df_al = pd.read_excel(path, sheet_name="Aliases")
        for _, row in df_al.iterrows():
            variant = str(row.iloc[0]).strip()
            canon = str(row.iloc[1]).strip()
            if not variant or variant == "nan" or not canon or canon == "nan":
                continue
            key = _make_key(variant)
            if key:
                aliases[key] = _clean_label(canon)
    except Exception:
        pass

    # --- Force Split ---
    force_split: dict[str, list[str]] = {}
    try:
        df_fs = pd.read_excel(path, sheet_name="Force Split")
        for _, row in df_fs.iterrows():
            original = str(row.iloc[0]).strip()
            parts = []
            for i in range(1, len(row)):
                v = str(row.iloc[i]).strip()
                if v and v != "nan":
                    parts.append(v)
            key = _make_key(original)
            if key and parts:
                force_split[key] = parts
    except Exception:
        pass

    # --- Paqueteria ---
    paqueteria: set[str] = set()
    try:
        df_pq = pd.read_excel(path, sheet_name="Paquetería")
    except Exception:
        try:
            df_pq = pd.read_excel(path, sheet_name="Paqueteria")
        except Exception:
            df_pq = pd.DataFrame()
    for _, row in df_pq.iterrows():
        name = str(row.iloc[0]).strip()
        if name and name != "nan":
            paqueteria.add(_clean_label(name))

    return aliases, force_split, paqueteria


# =====================================================================
# CANONICALIZACION Y TOKENIZACION
# =====================================================================

def canonical(name: str, aliases: dict) -> str:
    """Resuelve un nombre de proveedor a su forma canonica."""
    key = _make_key(name)
    if "nordex" in key:
        return "Nordex"
    if key in aliases:
        return aliases[key]
    return _clean_label(name)


# Separadores: + / , ; & y 'y'/'Y'
_SEP_STAGE1 = re.compile(r"\s*(?:\+|/|,|;|&|\by\b|\bY\b)\s*")
# Guion pegado sin espacios (A-B)
_HYPHEN_NO_SPACE = re.compile(r"(?<!\s)-(?!\s)")


def tokenize_suppliers(
    raw: str,
    aliases: dict,
    force_split: dict,
) -> list[str]:
    """
    Divide un texto de proveedor en nombres canonicos individuales.

    1) Regla force-split si existe.
    2) Separa por conectores (+, /, ,, ;, &, y/Y).
    3) Dentro de cada parte, divide por guion pegado (A-B).
    4) Canonicaliza cada trozo y deduplica preservando orden.
    """
    if pd.isna(raw) or not str(raw).strip():
        return []
    s = str(raw).strip()

    # Regla explicita para 'Woka Electronic'
    if re.search(r"\bwoka\s+electronic\b", s, flags=re.I):
        return [canonical("Woka", aliases), canonical("Electronic", aliases)]

    # Force-split exacto (normalizado)
    full_key = _make_key(s)
    if full_key in force_split:
        # Filtrar codigos numericos (ej: "0000096656") que no son nombres
        return [
            canonical(x, aliases)
            for x in force_split[full_key]
            if not x.strip().isdigit()
        ]

    # 1a fase: conectores
    parts1 = [p for p in _SEP_STAGE1.split(s) if p]

    out: list[str] = []
    for p in parts1:
        # 2a fase: guion pegado
        if _HYPHEN_NO_SPACE.search(p):
            subs = [sp for sp in _HYPHEN_NO_SPACE.split(p) if sp]
        else:
            subs = [p]
        out.extend(canonical(sp, aliases) for sp in subs if str(sp).strip())

    # Deduplicar preservando orden
    seen: set[str] = set()
    result: list[str] = []
    for x in out:
        if x not in seen:
            seen.add(x)
            result.append(x)
    return result


def owner_from_raw(raw: str, aliases: dict) -> str | None:
    """
    Devuelve el proveedor 'owner' de la incidencia:
    primer proveedor del texto original (antes de +, /, etc.).
    Se usa para Pareto: la incidencia se asigna solo al owner.
    """
    if pd.isna(raw) or not str(raw).strip():
        return None
    s = str(raw).strip()
    first = _SEP_STAGE1.split(s)[0]
    if _HYPHEN_NO_SPACE.search(first):
        first = _HYPHEN_NO_SPACE.split(first)[0]
    return canonical(first, aliases)


# =====================================================================
# PIPELINE PRINCIPAL
# =====================================================================

def _find_col(columns, candidates: list[str]):
    """Busqueda flexible de columna por nombre."""
    cols_lower = {str(c).lower().strip(): c for c in columns}
    for cand in candidates:
        cl = cand.lower().strip()
        if cl in cols_lower:
            return cols_lower[cl]
        for k, v in cols_lower.items():
            if cl in k or k in cl:
                return v
    return None


def process_bip(
    df_raw: pd.DataFrame,
    aliases: dict,
    force_split: dict,
    paqueteria: set,
) -> dict:
    """
    Pipeline principal BIP.

    Parameters
    ----------
    df_raw : DataFrame del extracto SGA (hoja Recepciones)
    aliases : dict de aliases normalizados
    force_split : dict de reglas force-split
    paqueteria : set de nombres canonicos de paqueteria

    Returns
    -------
    dict con claves:
        df_resultado_final, df_pareto, df_mensual_split,
        df_ajuste_split, metadata
    """
    df = df_raw.copy()

    # ---- Renombrar columnas (flexible) ----
    rename_map = {
        "Fecha Documento": "FechaDoc",
        "Tipo Recepcion": "TipoRecepcion",
        "Recepción con problemas": "Incidencia",
        "Recepcion con problemas": "Incidencia",
        "Estado Incidencia": "EstadoIncidencia",
        "Proveedor(texto)": "ProveedorTexto",
    }
    df.rename(
        columns={k: v for k, v in rename_map.items() if k in df.columns},
        inplace=True,
        errors="ignore",
    )

    # Busqueda flexible de columnas clave
    if "ProveedorTexto" not in df.columns:
        c = _find_col(df.columns, ["Proveedor(texto)", "Proveedor", "Supplier"])
        if c:
            df.rename(columns={c: "ProveedorTexto"}, inplace=True)

    if "Incidencia" not in df.columns:
        c = _find_col(
            df.columns,
            ["Recepción con problemas", "Recepcion con problemas", "Incidencia"],
        )
        if c:
            df.rename(columns={c: "Incidencia"}, inplace=True)

    if "EstadoIncidencia" not in df.columns:
        c = _find_col(df.columns, ["Estado Incidencia", "Estado incidencia", "EstadoIncidencia"])
        if c:
            df.rename(columns={c: "EstadoIncidencia"}, inplace=True)

    if "FechaDoc" not in df.columns:
        c = _find_col(df.columns, ["Fecha Documento", "Fecha", "FechaDoc"])
        if c:
            df.rename(columns={c: "FechaDoc"}, inplace=True)

    # ---- Preparar columnas (con guards) ----
    if "ProveedorTexto" not in df.columns:
        # Ultimo recurso: buscar cualquier columna con "proveedor" o "supplier"
        for c in df.columns:
            if "proveedor" in str(c).lower() or "supplier" in str(c).lower():
                df.rename(columns={c: "ProveedorTexto"}, inplace=True)
                break
        if "ProveedorTexto" not in df.columns:
            raise ValueError(
                "No se encontró columna de proveedor en el extracto SGA. "
                f"Columnas disponibles: {list(df.columns)}"
            )

    df["ProveedorTexto"] = df["ProveedorTexto"].astype(str).str.strip()
    df["Supplier_raw"] = df["ProveedorTexto"]

    if "FechaDoc" not in df.columns:
        # Ultimo recurso: buscar cualquier columna datetime
        for c in df.columns:
            if "fecha" in str(c).lower() or "date" in str(c).lower():
                df.rename(columns={c: "FechaDoc"}, inplace=True)
                break
        if "FechaDoc" not in df.columns:
            raise ValueError(
                "No se encontró columna de fecha en el extracto SGA. "
                f"Columnas disponibles: {list(df.columns)}"
            )

    df["FechaDoc"] = pd.to_datetime(df["FechaDoc"], dayfirst=True, errors="coerce")
    df["Año y mes"] = df["FechaDoc"].dt.to_period("M")
    df["Month"] = df["FechaDoc"].dt.strftime("%m/%Y")

    # Clasificacion: Nordex vs Supplier
    df["ResponsableTipo"] = (
        df["ProveedorTexto"]
        .str.contains("nordex", case=False, na=False)
        .map({True: "Nordex", False: "Supplier"})
    )

    # Normalizar Incidencia a Si / No
    if "Incidencia" in df.columns:
        inc = df["Incidencia"].astype(str).str.strip()
        si_set = {"sí", "si", "yes", "1", "true", "Sí", "SI", "Si"}
        no_set = {"no", "0", "false", "No", "NO"}
        df["Incidencia"] = inc.apply(
            lambda x: "Si" if x in si_set else ("No" if x in no_set else x)
        )
    if "EstadoIncidencia" in df.columns:
        df["EstadoIncidencia"] = df["EstadoIncidencia"].astype(str).str.strip()

    # ==================================================================
    # PARTE A — ResultadoFinal (KPIs mensuales)
    # ==================================================================
    res_a: list[dict] = []
    for mes in sorted(df["Año y mes"].dropna().unique()):
        dm = df[df["Año y mes"] == mes]
        df_val = dm[dm["Incidencia"].isin(["Si", "No"])]
        B = len(df_val)

        df_inc = df_val[df_val["Incidencia"] == "Si"]
        C = len(df_inc)

        df_supp = df_inc[df_inc["ResponsableTipo"] == "Supplier"]
        D = len(df_supp)
        E = int((df_supp["EstadoIncidencia"] == "Solucionada").sum()) if "EstadoIncidencia" in df_supp.columns else 0
        F = int((df_supp["EstadoIncidencia"] == "Abierta").sum()) if "EstadoIncidencia" in df_supp.columns else 0

        df_nx = df_inc[df_inc["ResponsableTipo"] == "Nordex"]
        G = len(df_nx)
        H = int((df_nx["EstadoIncidencia"] == "Solucionada").sum()) if "EstadoIncidencia" in df_nx.columns else 0
        I_val = int((df_nx["EstadoIncidencia"] == "Abierta").sum()) if "EstadoIncidencia" in df_nx.columns else 0

        K = (C / B) if B else 0.0
        L = (F / D) if D else 0.0
        M = (I_val / G) if G else 0.0

        res_a.append({
            "Año y mes": str(mes),
            "Total Receipts": B,
            "Total Receipts with incidents": C,
            "Supplier (with incidents)": D,
            "Supplier Resolved": E,
            "Supplier Unresolved": F,
            "Nordex (with incidents)": G,
            "Nordex Resolved": H,
            "Nordex Unresolved": I_val,
            "With incidents about the total": round(K, 4),
            "With incidents UNRESOLVED about the total of incidents (Supplier)": round(L, 4),
            "With incidents UNRESOLVED about the total of incidents (Nordex)": round(M, 4),
            "Total Receipts Correct": B - C,
        })

    df_resultado_final = pd.DataFrame(res_a)

    # ==================================================================
    # PARTE B — Pareto
    # ==================================================================
    base_b = df[["Incidencia", "Supplier_raw"]].dropna(subset=["Supplier_raw"]).copy()
    base_b["parts"] = base_b["Supplier_raw"].apply(
        lambda x: tokenize_suppliers(x, aliases, force_split)
    )
    base_b["parts"] = base_b["parts"].apply(
        lambda L: list(dict.fromkeys([p for p in L if p]))
    )
    base_b["Owner"] = base_b["Supplier_raw"].apply(
        lambda x: owner_from_raw(x, aliases)
    )
    base_b["HasIncident"] = (base_b["Incidencia"] == "Si").astype(int)

    # Total recepciones por proveedor (tras split)
    p_rows = (
        base_b.explode("parts")
        .rename(columns={"parts": "Supplier"})
        .dropna(subset=["Supplier"])
    )
    df_totales = (
        p_rows.groupby("Supplier", as_index=False)
        .size()
        .rename(columns={"size": "Total Receipts"})
    )

    # Incidencias por owner
    df_incid_owner = (
        base_b.groupby("Owner", as_index=False)["HasIncident"]
        .sum()
        .rename(columns={"Owner": "Supplier", "HasIncident": "Incidents"})
    )

    # Merge
    df_pareto = (
        df_totales.merge(df_incid_owner, on="Supplier", how="left")
        .fillna({"Incidents": 0})
    )
    df_pareto["Incidents"] = df_pareto["Incidents"].astype(int)
    df_pareto["Total Receipts"] = df_pareto["Total Receipts"].astype(int)

    df_pareto["% Incidents own"] = (
        (df_pareto["Incidents"] / df_pareto["Total Receipts"])
        .replace([np.inf, -np.inf], 0)
        .round(4)
    )

    total_incid_global = (df["Incidencia"] == "Si").sum()
    df_pareto["Global Incidents %"] = (
        (df_pareto["Incidents"] / total_incid_global).round(4)
        if total_incid_global
        else 0.0
    )

    df_pareto = df_pareto.sort_values("Incidents", ascending=False).reset_index(drop=True)

    # Acumulado para linea Pareto (0-1 decimal, Excel lo formatea como %)
    if df_pareto["Incidents"].sum() > 0:
        df_pareto["Cumulative %"] = (
            df_pareto["Incidents"].cumsum() / df_pareto["Incidents"].sum()
        ).round(4)
    else:
        df_pareto["Cumulative %"] = 0.0

    # Marcar paqueteria
    df_pareto["Paquetería"] = df_pareto["Supplier"].isin(paqueteria)

    # Numero de proveedores en 80%
    n_80 = 0
    if len(df_pareto) > 0 and df_pareto["Incidents"].sum() > 0:
        n_80 = int((df_pareto["Cumulative %"] <= 0.80).sum())
        if n_80 < len(df_pareto):
            n_80 += 1
        n_80 = min(n_80, len(df_pareto))

    # ==================================================================
    # PARTE C — Mensual por proveedor (split) + Observaciones
    # ==================================================================
    base_c = df[["Month", "Supplier_raw"]].dropna(subset=["Supplier_raw"]).copy()
    base_c = base_c.reset_index(drop=False).rename(
        columns={"index": "RowId", "Supplier_raw": "Supplier_orig"}
    )
    base_c["__parts"] = base_c["Supplier_orig"].apply(
        lambda x: tokenize_suppliers(x, aliases, force_split)
    )

    exploded = (
        base_c.explode("__parts")
        .rename(columns={"__parts": "Supplier"})
        .dropna(subset=["Supplier"])
    )
    exploded["Total Receipts"] = 1

    # Conteos de compañeros (singles vs grupos)
    ind_counts: dict = defaultdict(Counter)
    grp_counts: dict = defaultdict(Counter)
    for _, r in base_c.iterrows():
        month = r["Month"]
        parts = [p for p in r["__parts"] if p]
        uniq = list(dict.fromkeys(parts))
        for s in uniq:
            others = [x for x in uniq if x != s]
            if len(others) == 1:
                ind_counts[(month, s)][others[0]] += 1
            elif len(others) >= 2:
                grp_counts[(month, s)][tuple(sorted(others))] += 1

    monthly_split = (
        exploded.groupby(["Month", "Supplier"])["Total Receipts"]
        .sum()
        .reset_index()
        .sort_values(["Month", "Total Receipts"], ascending=[True, False])
    )

    def _obs(month: str, supplier: str) -> str:
        singles = ind_counts.get((month, supplier), Counter())
        groups = grp_counts.get((month, supplier), Counter())
        blocks: list[str] = []
        if singles:
            blocks.append(
                "; ".join(f"{cnt} con {name}" for name, cnt in singles.most_common())
            )
        if groups:
            blocks.append(
                "; ".join(
                    f"{cnt} con " + " y ".join(g) for g, cnt in groups.most_common()
                )
            )
        return ". ".join(blocks)

    monthly_split["Observaciones"] = [
        _obs(m, s)
        for m, s in zip(monthly_split["Month"], monthly_split["Supplier"])
    ]
    monthly_split["Paquetería"] = monthly_split["Supplier"].isin(paqueteria)

    # ==================================================================
    # PARTE D — Ajuste split
    # ==================================================================
    orig_by_month = (
        df.dropna(subset=["Month"])
        .groupby("Month")
        .size()
        .rename("Original Receipts")
        .reset_index()
    )
    after_by_month = (
        exploded.groupby("Month")
        .size()
        .rename("Rows after split")
        .reset_index()
    )
    df_ajuste = orig_by_month.merge(after_by_month, on="Month", how="outer").fillna(0)
    df_ajuste["Original Receipts"] = df_ajuste["Original Receipts"].astype(int)
    df_ajuste["Rows after split"] = df_ajuste["Rows after split"].astype(int)
    df_ajuste["Extra to subtract"] = (
        df_ajuste["Rows after split"] - df_ajuste["Original Receipts"]
    ).astype(int)

    # ==================================================================
    # METADATA
    # ==================================================================
    total_rows = len(df)
    total_valid = len(df[df["Incidencia"].isin(["Si", "No"])])
    total_incidents = int((df["Incidencia"] == "Si").sum())
    total_supplier = int(df_resultado_final["Supplier (with incidents)"].sum()) if len(df_resultado_final) else 0
    total_nordex = int(df_resultado_final["Nordex (with incidents)"].sum()) if len(df_resultado_final) else 0

    metadata = {
        "total_rows": total_rows,
        "total_valid": total_valid,
        "total_incidents": total_incidents,
        "total_supplier_incidents": total_supplier,
        "total_nordex_incidents": total_nordex,
        "total_suppliers": len(df_pareto),
        "n_suppliers_80pct": n_80,
        "pct_suppliers_80": round(n_80 / len(df_pareto) * 100, 1) if len(df_pareto) else 0,
    }

    return {
        "df_resultado_final": df_resultado_final,
        "df_pareto": df_pareto,
        "df_mensual_split": monthly_split[
            ["Month", "Supplier", "Total Receipts", "Observaciones", "Paquetería"]
        ],
        "df_ajuste_split": df_ajuste,
        "metadata": metadata,
    }


# =====================================================================
# EXPORTACION A EXCEL
# =====================================================================

def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    """Exporta un DataFrame a bytes xlsx formateado."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        wb = writer.book
        ws = writer.sheets[sheet_name[:31]]

        hdr_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "#FFFFFF",
            "border": 1,
            "text_wrap": True,
            "valign": "vcenter",
        })
        pct_fmt = wb.add_format({"num_format": "0.00%", "border": 1})

        for ci, col in enumerate(df.columns):
            ws.write(0, ci, col, hdr_fmt)
            max_len = max(
                len(str(col)),
                int(df[col].astype(str).str.len().max()) if len(df) else 0,
            )
            ws.set_column(ci, ci, min(max_len + 2, 50))

            # Formato porcentaje para columnas que son decimales 0-1
            col_l = col.lower()
            if "%" in col or "incidents about" in col_l or "unresolved about" in col_l:
                for ri in range(len(df)):
                    val = df.iloc[ri, ci]
                    if pd.notna(val):
                        ws.write(ri + 1, ci, val, pct_fmt)

    return buf.getvalue()


def to_pareto_excel_bytes(
    df_pareto: pd.DataFrame,
    metadata: dict,
) -> bytes:
    """
    Exporta el Pareto a Excel con formato y resumen al final:
    - N proveedores hacen el 80% de las incidencias
    - El X% de los proveedores hace el 80%
    - Fecha de actualizacion
    """
    import datetime

    buf = BytesIO()
    sn = "Pareto"
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df_pareto.to_excel(writer, index=False, sheet_name=sn)
        wb = writer.book
        ws = writer.sheets[sn]

        # Formatos
        hdr_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#4472C4",
            "font_color": "#FFFFFF",
            "border": 1,
            "text_wrap": True,
            "valign": "vcenter",
        })
        pct_fmt = wb.add_format({"num_format": "0.00%", "border": 1})
        summary_fmt = wb.add_format({
            "bold": True,
            "bg_color": "#FFFF00",
            "border": 1,
            "text_wrap": True,
            "font_size": 12,
        })

        # Cabeceras
        for ci, col in enumerate(df_pareto.columns):
            ws.write(0, ci, col, hdr_fmt)
            max_len = max(
                len(str(col)),
                int(df_pareto[col].astype(str).str.len().max()) if len(df_pareto) else 0,
            )
            ws.set_column(ci, ci, min(max_len + 3, 55))

        # Formato % a columnas que son decimales 0-1
        for ci, col in enumerate(df_pareto.columns):
            col_l = col.lower()
            if "%" in col or "incidents own" in col_l:
                for ri in range(len(df_pareto)):
                    val = df_pareto.iloc[ri, ci]
                    if pd.notna(val):
                        ws.write(ri + 1, ci, val, pct_fmt)

        # --- Resumen al final (2 filas vacias + 3 filas resumen) ---
        sr = len(df_pareto) + 3  # +1 header +2 blank
        n_80 = metadata.get("n_suppliers_80pct", 0)
        total_s = metadata.get("total_suppliers", len(df_pareto))
        pct_80 = round(n_80 / total_s * 100, 1) if total_s else 0

        ws.write(sr, 0, n_80, summary_fmt)
        ws.merge_range(
            sr, 1, sr, 4,
            f"Suppliers make the 80% of the incidents",
            summary_fmt,
        )

        ws.write(sr + 1, 0, "So the", summary_fmt)
        ws.merge_range(
            sr + 1, 1, sr + 1, 4,
            f"{pct_80}%  of the suppliers makes the 80% of the incidents",
            summary_fmt,
        )

        # Fila con fecha
        sr2 = sr + 3
        today = datetime.date.today().strftime("%d/%m/%y")
        ws.merge_range(sr2, 0, sr2, 2, f"Updated to {today}", summary_fmt)

    return buf.getvalue()


def to_zip_bytes(files: list[tuple[str, bytes]]) -> bytes:
    """Empaqueta multiples (nombre, bytes) en un ZIP."""
    buf = BytesIO()
    with ZipFile(buf, "w", ZIP_DEFLATED) as zf:
        for name, data in files:
            zf.writestr(name, data)
    return buf.getvalue()
