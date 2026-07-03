"""
sidebar_uah.py — Componente de sidebar compartido para todas las páginas.

Como cada archivo dentro de pages/ se ejecuta de forma independiente,
Streamlit no comparte automáticamente el código de Inicio.py con el resto
de las páginas. Para que el logo, el fondo negro del sidebar y el texto
de la unidad aparezcan en TODAS las páginas (no solo en Inicio), cada
página debe llamar a render_sidebar() al principio de su script.

Uso en cualquier página nueva (ej: pages/1_Busqueda_Rut.py):

    import streamlit as st
    from sidebar_uah import render_sidebar

    st.set_page_config(page_title="Búsqueda por Rut", page_icon="🔎", layout="wide")
    render_sidebar()

    st.title("Búsqueda por Rut en Ucampus")
    ... resto de la página ...

Nota: st.set_page_config() debe llamarse ANTES de render_sidebar() (y antes
de cualquier otro comando de Streamlit) en cada página, ya que solo puede
invocarse una vez por ejecución de página y debe ser el primer comando.
"""

import os
import streamlit as st

# Ruta del logo relativa a la raíz del proyecto (streamlit_app/). Como
# Streamlit ejecuta cada página con esa misma raíz como working directory,
# esta misma ruta funciona tanto desde Inicio.py como desde pages/*.py.
LOGO_PATH = "UAH-logo.png"


def render_sidebar() -> None:
    """Dibuja el fondo negro del sidebar, el logo institucional agrandado,
    y el texto de la unidad. Llamar una vez al principio de cada página,
    después de st.set_page_config()."""

    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {
            background-color: #000000;
        }
        [data-testid="stSidebar"] * {
            color: #FFFFFF;
        }
        /* Espacio entre el borde superior del sidebar y el logo */
        [data-testid="stSidebarHeader"] {
            padding-top: 1.5rem !important;
            padding-bottom: 0.5rem !important;
        }
        /* Agrandar el logo puesto por st.logo() (selector puede variar según versión) */
        [data-testid="stLogo"],
        [data-testid="stSidebarHeader"] img,
        [data-testid="stHeader"] img {
            height: 4.5rem !important;
            width: auto !important;
            max-width: none !important;
            margin-top: 1.5rem !important;
        }
        /* Agrandar el logo cuando se usa el fallback st.sidebar.image() */
        [data-testid="stSidebar"] [data-testid="stImage"] img {
            width: 100%;
            max-width: 220px;
        }
        /* Espacio entre el logo y el menú de navegación (Inicio, Búsqueda...) */
        [data-testid="stSidebarNav"] {
            padding-top: 1.5rem !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    if os.path.exists(LOGO_PATH):
        try:
            # st.logo coloca el logo en la esquina superior izquierda y en
            # la barra lateral automáticamente (Streamlit >= 1.37).
            st.logo(LOGO_PATH, link="https://www.uahurtado.cl/")
        except AttributeError:
            # Compatibilidad con versiones de Streamlit anteriores a st.logo
            st.sidebar.image(LOGO_PATH, use_container_width=True)
    else:
        st.sidebar.warning(f"No se encontró el logo en: {LOGO_PATH}")

    st.sidebar.markdown("---")
    st.sidebar.caption("Dirección de Gestión y Desarrollo")
    st.sidebar.caption("Unidad de Inteligencia de Datos")
