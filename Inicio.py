"""
Inicio.py — Página de inicio de la app "Procesos Varios".

Esta es la página raíz de una app Streamlit multipágina. Para agregar nuevas
funciones en el futuro:
  1. Crea un archivo .py dentro de la carpeta pages/ (ej: pages/1_Avance_curricular.py).
     Aparecerá automáticamente en el menú de navegación de la barra lateral.
  2. Al principio de ese archivo, después de st.set_page_config(), llama a
     render_sidebar() (importado desde sidebar_uah) para que el logo, el
     fondo negro y el texto de la unidad también aparezcan en esa página.
     Mira este mismo archivo como ejemplo del patrón a seguir.

Ejecutar con:
    streamlit run Inicio.py
"""

import streamlit as st
from sidebar_uah import render_sidebar

st.set_page_config(
    page_title="Procesos Varios — UAH",
    page_icon="📊",
    layout="wide",
)

render_sidebar()

# ── Contenido principal ──────────────────────────────────────────────────

st.title("Procesos Varios")

st.markdown(
    "En esta aplicación se pueden realizar las siguientes funciones:\n\n"
    "- Generar tablas para actualizar UAH en cifras.\n"
    "- Otras en construcción" 

)

st.markdown("---")

st.markdown(
    "Usa el menú de la barra lateral para navegar entre las distintas funciones "
    "disponibles."
)

st.markdown("---")
st.caption("Dirección de Gestión y Desarrollo — Unidad de Inteligencia de Datos")
