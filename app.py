import streamlit as st

st.set_page_config(page_title="Herramientas de Gestión", layout="wide")

pages = {
    "": [
        st.Page("pages/0_inicio.py", title="Inicio", default=True),
    ],
    "Auditorías Internas de Inventario": [
        st.Page("pages/1_segregaciones.py", title="Generar Segregaciones"),
        st.Page("pages/2_reporte_consolidado.py", title="Generar Reporte Consolidado"),
    ],
    "Análisis de Recepciones": [
        st.Page("pages/3_bip_recepciones.py", title="BIP Recepciones"),
    ],
}

pg = st.navigation(pages)
pg.run()
