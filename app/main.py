"""Investment Lab: simulador de estrategias de inversión bajo incertidumbre.

Ejecutar con:  streamlit run app/main.py
"""
import sys
from pathlib import Path

import streamlit as st

APP = Path(__file__).resolve().parent
sys.path.insert(0, str(APP.parent / "src"))
sys.path.insert(0, str(APP))

st.set_page_config(page_title="Investment Lab", page_icon="📈", layout="wide")

from components import estado, ui  # noqa: E402

ui.estilo()
estado.iniciar()

paginas = {
    "Simulador": [
        st.Page(str(APP / "pages" / "01_inicio.py"), title="Inicio", default=True),
        st.Page(str(APP / "pages" / "02_inversionista.py"), title="Inversionista"),
        st.Page(str(APP / "pages" / "03_modelos.py"), title="Modelos y economía"),
        st.Page(str(APP / "pages" / "04_estrategias.py"), title="Estrategias"),
        st.Page(str(APP / "pages" / "05_experimento.py"), title="Experimento"),
    ],
    "Análisis": [
        st.Page(str(APP / "pages" / "06_resultados.py"), title="Resultados"),
        st.Page(str(APP / "pages" / "07_sensibilidad.py"), title="Sensibilidad y capital-meta"),
    ],
}

with st.sidebar:
    p = estado.params()
    st.markdown("**Configuración actual**")
    st.caption(f"Capital {ui.cop(p.f('inv.capital'))} · {p.i('inv.horizonte') // 12} años · "
               f"meta {ui.cop(p.f('inv.meta_real'))} · liquidez mínima {ui.pct(p.f('inv.liquidez_min'), 0)}")
    n = len(estado.cambios())
    st.caption(f"{n} parámetro(s) cambiado(s) frente a la referencia" if n else "Todos los parámetros en su valor de referencia")
    if n and st.button("Volver a la referencia", width="stretch"):
        estado.restablecer()
        st.rerun()

st.navigation(paginas).run()
