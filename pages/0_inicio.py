import streamlit as st

st.title("Herramientas de Gestión")
st.markdown("Selecciona una herramienta en el menú lateral para comenzar.")

st.markdown("""
### Auditorías Internas de Inventario

- **Generar Segregaciones** — Genera muestras de auditoría a partir del extracto de stock
- **Generar Reporte Consolidado** — Genera el reporte de KPIs a partir del Excel rellenado

### Análisis de Recepciones (BIP)

- **BIP Recepciones** — Analiza recepciones del SGA: KPIs mensuales, Pareto de incidencias por proveedor, desglose mensual y gráficos interactivos. Genera los Excel con nombres exactos para alimentar Power BI.
""")
