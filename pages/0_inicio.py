import streamlit as st

st.title("Herramientas de Gestión")
st.markdown("Selecciona una herramienta en el menú lateral para comenzar.")

st.markdown("""
### Auditorías Internas de Inventario

- **Generar Segregaciones** — Genera muestras de auditoría a partir del extracto de stock
- **Generar Reporte Consolidado** — Genera el reporte de KPIs a partir del Excel rellenado

### Análisis de Recepciones (BIP)

- **BIP Recepciones** — Analiza recepciones del SGA: KPIs mensuales, Pareto de incidencias por proveedor, desglose mensual y gráficos interactivos. Genera los Excel con nombres exactos para alimentar Power BI.

### Análisis de Absentismo

- **Absentismo por Centro** — Sube los cuadrantes mensuales de horas (un Excel por centro de trabajo) y genera análisis de absentismo: plantilla, días laborables, desglose de ausencias (bajas, vacaciones, permisos, asuntos propios, excedencias), % de absentismo con y sin vacaciones, y reporte consolidado multi-centro.
""")
