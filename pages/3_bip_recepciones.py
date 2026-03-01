"""
3_bip_recepciones.py
====================
Pagina Streamlit — Analisis de Recepciones BIP.
Sube un extracto SGA y obtiene KPIs, graficos Plotly y Excel descargables.
"""
from __future__ import annotations

import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

from bip_utils import load_aliases, process_bip, to_excel_bytes, to_pareto_excel_bytes, to_zip_bytes

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
DEFAULT_ALIASES = _PROJECT_ROOT / "data" / "aliases_proveedores.xlsx"

MONTH_NAMES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

# ---------------------------------------------------------------------------
# Helpers cacheados
# ---------------------------------------------------------------------------
@st.cache_data
def _read_sga(data: bytes) -> pd.DataFrame:
    """Lee el extracto SGA: intenta hoja Recepciones, si no, la primera."""
    try:
        return pd.read_excel(BytesIO(data), sheet_name="Recepciones")
    except Exception:
        return pd.read_excel(BytesIO(data), sheet_name=0)


@st.cache_data
def _cached_load_aliases(path: str) -> tuple:
    return load_aliases(path)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Archivos — BIP Recepciones")
    f_sga = st.file_uploader(
        "Extracto SGA (.xlsx)", type=["xlsx", "xls"], key="bip_sga"
    )
    f_aliases = st.file_uploader(
        "Aliases proveedores (.xlsx)",
        type=["xlsx", "xls"],
        key="bip_aliases",
        help="Opcional. Si no se sube, se usa el fichero por defecto.",
    )

    # Botón para descargar el template de aliases actual
    if DEFAULT_ALIASES.exists():
        st.download_button(
            "Descargar aliases actual",
            data=DEFAULT_ALIASES.read_bytes(),
            file_name="aliases_proveedores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    btn = st.button("Procesar", type="primary", use_container_width=True)

# ---------------------------------------------------------------------------
# Titulo
# ---------------------------------------------------------------------------
st.title("Analisis de Recepciones BIP")

# ---------------------------------------------------------------------------
# Procesamiento
# ---------------------------------------------------------------------------
if btn and f_sga is not None:
    try:
        with st.spinner("Procesando extracto SGA..."):
            # Cargar aliases
            if f_aliases is not None:
                tmp_path = _PROJECT_ROOT / "data" / "_tmp_aliases.xlsx"
                tmp_path.parent.mkdir(parents=True, exist_ok=True)
                tmp_path.write_bytes(f_aliases.getvalue())
                aliases, force_split, paqueteria = load_aliases(str(tmp_path))
            else:
                aliases, force_split, paqueteria = _cached_load_aliases(
                    str(DEFAULT_ALIASES)
                )

            # Cargar SGA
            df_raw = _read_sga(f_sga.getvalue())

            # Ejecutar pipeline
            results = process_bip(df_raw, aliases, force_split, paqueteria)

            # Guardar en session_state
            st.session_state["bip_results"] = results
            st.session_state["bip_processed"] = True

        st.success(
            f"Procesadas **{results['metadata']['total_valid']:,}** recepciones "
            f"de **{results['metadata']['total_suppliers']}** proveedores."
        )
    except Exception as e:
        st.error(f"Error al procesar: {e}")
        import traceback
        st.code(traceback.format_exc())
        st.stop()

elif btn and f_sga is None:
    st.warning("Sube un extracto SGA antes de pulsar Procesar.")

# Guard
if not st.session_state.get("bip_processed"):
    st.info("Sube un extracto SGA en la barra lateral y pulsa **Procesar** para comenzar.")
    st.stop()

# ---------------------------------------------------------------------------
# Recuperar resultados
# ---------------------------------------------------------------------------
results = st.session_state["bip_results"]
meta = results["metadata"]
df_resultado_full = results["df_resultado_final"].copy()
df_pareto_full = results["df_pareto"].copy()
df_mensual_full = results["df_mensual_split"].copy()
df_ajuste_full = results["df_ajuste_split"].copy()

# ---------------------------------------------------------------------------
# Filtros de año / mes
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Filtros")

# Extraer años y meses disponibles del df_resultado_final
avail_years, avail_months = [], []
if not df_resultado_full.empty and "Año y mes" in df_resultado_full.columns:
    try:
        periods = pd.PeriodIndex(df_resultado_full["Año y mes"], freq="M")
        avail_years = sorted(periods.year.unique().tolist())
        avail_months = sorted(periods.month.unique().tolist())
    except Exception:
        # Fallback: parsear strings "YYYY-MM" manualmente
        raw = df_resultado_full["Año y mes"].dropna().astype(str).unique()
        years_set, months_set = set(), set()
        for p in raw:
            parts = p.split("-")
            if len(parts) == 2:
                years_set.add(int(parts[0]))
                months_set.add(int(parts[1]))
        avail_years = sorted(years_set)
        avail_months = sorted(months_set)

fc1, fc2 = st.columns(2)
with fc1:
    sel_years = st.multiselect(
        "Año",
        options=avail_years,
        default=avail_years,
        key="bip_years",
    )
with fc2:
    sel_months = st.multiselect(
        "Mes",
        options=avail_months,
        default=avail_months,
        format_func=lambda m: MONTH_NAMES_ES.get(m, str(m)),
        key="bip_months",
    )

if not sel_years or not sel_months:
    st.warning("Selecciona al menos un año y un mes.")
    st.stop()

# Construir periodos validos
valid_periods = set()
for y in sel_years:
    for m in sel_months:
        valid_periods.add(f"{y}-{m:02d}")


def _filt(df: pd.DataFrame, col: str = "Año y mes") -> pd.DataFrame:
    if col in df.columns:
        return df[df[col].isin(valid_periods)].copy()
    return df


# Filtrar ResultadoFinal y Mensual
df_resultado = _filt(df_resultado_full)

# Filtrar mensual por Month (formato MM/YYYY)
valid_months_slash = set()
for y in sel_years:
    for m in sel_months:
        valid_months_slash.add(f"{m:02d}/{y}")

df_mensual = df_mensual_full[df_mensual_full["Month"].isin(valid_months_slash)].copy() if "Month" in df_mensual_full.columns else df_mensual_full.copy()

# Recalcular Pareto filtrado desde mensual
if not df_mensual.empty:
    _par = (
        df_mensual.groupby("Supplier", as_index=False)["Total Receipts"]
        .sum()
    )
    # Incidencias del pareto completo (ya están asignadas por owner)
    # Usamos los datos del pareto completo pero filtrado por periodos
    # Para filtrado correcto: el pareto no tiene columna de mes, lo mantenemos tal cual
    # pero filtramos mostrando solo proveedores que aparecen en el periodo filtrado
    provs_in_period = set(df_mensual["Supplier"].unique())
    df_pareto = df_pareto_full[df_pareto_full["Supplier"].isin(provs_in_period)].copy()
    df_pareto = df_pareto.sort_values("Incidents", ascending=False).reset_index(drop=True)
    if df_pareto["Incidents"].sum() > 0:
        df_pareto["Cumulative %"] = (
            df_pareto["Incidents"].cumsum() / df_pareto["Incidents"].sum()
        ).round(4)
    else:
        df_pareto["Cumulative %"] = 0.0
else:
    df_pareto = df_pareto_full.copy()

# Ajuste: filtrar por Month
df_ajuste = df_ajuste_full[df_ajuste_full["Month"].isin(valid_months_slash)].copy() if "Month" in df_ajuste_full.columns else df_ajuste_full.copy()

# Recalcular metadata filtrada
f_total_valid = int(df_resultado["Total Receipts"].sum()) if not df_resultado.empty else 0
f_total_incidents = int(df_resultado["Total Receipts with incidents"].sum()) if not df_resultado.empty else 0
f_supplier = int(df_resultado["Supplier (with incidents)"].sum()) if not df_resultado.empty else 0
f_nordex = int(df_resultado["Nordex (with incidents)"].sum()) if not df_resultado.empty else 0
f_total_suppliers = len(df_pareto) if not df_pareto.empty else 0
f_n80 = 0
if not df_pareto.empty and df_pareto["Incidents"].sum() > 0:
    f_n80 = int((df_pareto["Cumulative %"] <= 0.80).sum())
    if f_n80 < len(df_pareto):
        f_n80 += 1

# ---------------------------------------------------------------------------
# KPI Cards
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Indicadores Clave")

kc1, kc2, kc3, kc4 = st.columns(4)
pct_inc = (f_total_incidents / f_total_valid * 100) if f_total_valid > 0 else 0.0

kc1.metric("Total recepciones", f"{f_total_valid:,}")
kc2.metric("Con incidencia", f"{pct_inc:.1f} %")
kc3.metric(
    "Supplier vs Nordex",
    f"{f_supplier:,} / {f_nordex:,}",
)
kc4.metric(
    "Proveedores 80% Pareto",
    f"{f_n80} / {f_total_suppliers}",
)

# ---------------------------------------------------------------------------
# Graficos — Fila 1
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Graficos")

ch1, ch2 = st.columns(2)

# Chart 1: Tarta correctas vs con incidencia
with ch1:
    correctas = f_total_valid - f_total_incidents
    fig1 = go.Figure(
        data=[go.Pie(
            labels=["Correctas", "Con incidencia"],
            values=[correctas, f_total_incidents],
            marker_colors=["#4CAF50", "#FF5722"],
            hole=0.35,
            textinfo="label+percent",
            textposition="inside",
        )]
    )
    fig1.update_layout(
        title_text="Recepciones correctas vs con incidencia",
        title_x=0.5,
        margin=dict(t=50, b=20, l=20, r=20),
        height=400,
        showlegend=True,
        legend=dict(orientation="h", y=-0.1, x=0.5, xanchor="center"),
    )
    st.plotly_chart(fig1, use_container_width=True)

# Chart 2: Barras apiladas tendencia mensual
with ch2:
    if not df_resultado.empty:
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            x=df_resultado["Año y mes"],
            y=df_resultado["Total Receipts Correct"],
            name="Correctas",
            marker_color="#4CAF50",
        ))
        fig2.add_trace(go.Bar(
            x=df_resultado["Año y mes"],
            y=df_resultado["Total Receipts with incidents"],
            name="Con incidencia",
            marker_color="#FF5722",
        ))
        fig2.update_layout(
            barmode="stack",
            title_text="Tendencia mensual de recepciones",
            title_x=0.5,
            xaxis_title="Mes",
            yaxis_title="Recepciones",
            margin=dict(t=50, b=20, l=20, r=20),
            height=400,
            legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
        )
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Sin datos para el periodo seleccionado.")

# ---------------------------------------------------------------------------
# Graficos — Fila 2
# ---------------------------------------------------------------------------
ch3, ch4 = st.columns(2)

# Chart 3: Tarta desglose incidencias
with ch3:
    if not df_resultado.empty:
        sr = int(df_resultado["Supplier Resolved"].sum())
        su = int(df_resultado["Supplier Unresolved"].sum())
        nr = int(df_resultado["Nordex Resolved"].sum())
        nu = int(df_resultado["Nordex Unresolved"].sum())

        fig3 = go.Figure(
            data=[go.Pie(
                labels=[
                    "Supplier Resueltas",
                    "Supplier Abiertas",
                    "Nordex Resueltas",
                    "Nordex Abiertas",
                ],
                values=[sr, su, nr, nu],
                marker_colors=["#2196F3", "#FF9800", "#00BCD4", "#F44336"],
                hole=0.35,
                textinfo="label+percent",
                textposition="inside",
            )]
        )
        fig3.update_layout(
            title_text="Desglose de incidencias",
            title_x=0.5,
            margin=dict(t=50, b=20, l=20, r=20),
            height=400,
            showlegend=True,
            legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
        )
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Sin datos.")

# Chart 4: Barras horizontales top 15
with ch4:
    if not df_pareto.empty:
        top15 = df_pareto.head(15).copy()
        top15 = top15.sort_values("Incidents", ascending=True)
        top15["Tipo"] = top15["Paquetería"].map(
            {True: "Paquetería", False: "Proveedor"}
        )
        fig4 = px.bar(
            top15,
            x="Incidents",
            y="Supplier",
            color="Tipo",
            color_discrete_map={"Paquetería": "#FF9800", "Proveedor": "#2196F3"},
            orientation="h",
            title="Top 15 proveedores por incidencias",
        )
        fig4.update_layout(
            margin=dict(t=50, b=20, l=20, r=20),
            height=400,
            yaxis_title="",
            xaxis_title="Incidencias",
            legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
        )
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info("Sin datos.")

# ---------------------------------------------------------------------------
# Chart 5: Pareto completo (barras + linea acumulada)
# ---------------------------------------------------------------------------
st.divider()

if not df_pareto.empty and df_pareto["Incidents"].sum() > 0:
    pareto_show = df_pareto[df_pareto["Incidents"] > 0].head(30).copy()

    fig5 = go.Figure()
    fig5.add_trace(go.Bar(
        x=pareto_show["Supplier"],
        y=pareto_show["Incidents"],
        name="Incidencias",
        marker_color="#2196F3",
        yaxis="y",
    ))
    fig5.add_trace(go.Scatter(
        x=pareto_show["Supplier"],
        y=pareto_show["Cumulative %"] * 100,  # 0-1 → 0-100 para el grafico
        name="% Acumulado",
        mode="lines+markers",
        marker=dict(color="#FF5722", size=6),
        line=dict(color="#FF5722", width=2),
        yaxis="y2",
    ))
    fig5.add_hline(
        y=80,
        line_dash="dash",
        line_color="#F44336",
        annotation_text="80 %",
        annotation_position="top right",
        yref="y2",
    )
    fig5.update_layout(
        title_text="Pareto de incidencias por proveedor",
        title_x=0.5,
        xaxis=dict(title="Proveedor", tickangle=-45),
        yaxis=dict(title="Incidencias", side="left"),
        yaxis2=dict(
            title="% Acumulado",
            side="right",
            overlaying="y",
            range=[0, 105],
            showgrid=False,
            ticksuffix="%",
        ),
        margin=dict(t=60, b=120, l=60, r=60),
        height=500,
        legend=dict(orientation="h", y=-0.3, x=0.5, xanchor="center"),
        barmode="group",
    )
    st.plotly_chart(fig5, use_container_width=True)
else:
    st.info("No hay incidencias para mostrar el Pareto.")

# ---------------------------------------------------------------------------
# Tablas de datos
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Tablas de datos")

tab_res, tab_par, tab_mens, tab_aj = st.tabs([
    "Resultado Final", "Pareto", "Mensual por Proveedor", "Ajuste Split"
])

with tab_res:
    if not df_resultado.empty:
        st.dataframe(df_resultado, use_container_width=True, hide_index=True)
    else:
        st.info("Sin datos para el periodo seleccionado.")

with tab_par:
    if not df_pareto.empty:
        # Mostrar con % formateado para la tabla visual
        df_pareto_display = df_pareto.copy()
        for c in ["% Incidents own", "Global Incidents %", "Cumulative %"]:
            if c in df_pareto_display.columns:
                df_pareto_display[c] = (df_pareto_display[c] * 100).round(2).astype(str) + " %"
        st.dataframe(df_pareto_display, use_container_width=True, hide_index=True)
        st.info(
            f"**{f_n80}** proveedores hacen el 80% de las incidencias "
            f"({round(f_n80 / f_total_suppliers * 100, 1) if f_total_suppliers else 0}% del total)"
        )
    else:
        st.info("Sin datos para el periodo seleccionado.")

with tab_mens:
    if not df_mensual.empty:
        st.dataframe(df_mensual, use_container_width=True, hide_index=True)
    else:
        st.info("Sin datos para el periodo seleccionado.")

with tab_aj:
    if not df_ajuste.empty:
        st.dataframe(df_ajuste, use_container_width=True, hide_index=True)
    else:
        st.info("Sin datos.")

# ---------------------------------------------------------------------------
# Descargas
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Descargas")

MIME_XL = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

b_res = to_excel_bytes(df_resultado, "ResultadoFinal")
b_par = to_pareto_excel_bytes(df_pareto, meta)
b_men = to_excel_bytes(df_mensual, "MensualProveedor")
b_aju = to_excel_bytes(df_ajuste, "AjusteSplit")

d1, d2, d3, d4, d5 = st.columns(5)

with d1:
    st.download_button(
        "ResultadoFinal.xlsx",
        data=b_res,
        file_name="ResultadoFinal.xlsx",
        mime=MIME_XL,
        use_container_width=True,
    )
with d2:
    st.download_button(
        "Pareto.xlsx",
        data=b_par,
        file_name="ParetoIncidenciasProveedores.xlsx",
        mime=MIME_XL,
        use_container_width=True,
    )
with d3:
    st.download_button(
        "Mensual Split.xlsx",
        data=b_men,
        file_name="Recepciones_por_mes_y_proveedor_SPLIT.xlsx",
        mime=MIME_XL,
        use_container_width=True,
    )
with d4:
    st.download_button(
        "Ajuste Split.xlsx",
        data=b_aju,
        file_name="Ajuste_split_resumen.xlsx",
        mime=MIME_XL,
        use_container_width=True,
    )
with d5:
    z = to_zip_bytes([
        ("ResultadoFinal.xlsx", b_res),
        ("ParetoIncidenciasProveedores.xlsx", b_par),
        ("Recepciones_por_mes_y_proveedor_SPLIT.xlsx", b_men),
        ("Ajuste_split_resumen.xlsx", b_aju),
    ])
    st.download_button(
        "Descargar TODO (.zip)",
        data=z,
        file_name="BIP_Recepciones_Completo.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )
