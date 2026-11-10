import subprocess
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

from components import estado, ui
from investment_sim.engine.events import Tipo
from investment_sim.experiments.runner import una_replica
from investment_sim.experiments.validation import validar

p = estado.params()
RAIZ = Path(__file__).resolve().parents[2]

st.title("Simulación, escenarios y validación")
t1, t2, t3 = st.tabs(["Eventos discretos", "Escenarios", "Verificación y validación"])

with t1:
    st.write("El simulador es de eventos discretos: un reloj avanza de evento en evento (no mes a mes), una lista ordenada por "
             "tiempo y prioridad programa los eventos futuros, y cada evento actualiza las variables de estado. Si dos eventos "
             "caen en el mismo instante, se procesan por el orden de prioridad.")
    ev = pd.DataFrame([{"N.°": t.numero, "Evento": t.nombre, "Prioridad en empates": t.prioridad}
                       for t in sorted(Tipo, key=lambda t: t.prioridad)])
    st.markdown("**Eventos**")
    st.dataframe(ev, hide_index=True, width="stretch")
    st.caption("La cobertura de liquidez, la compra y la venta de activos y el rebalanceo ocurren dentro de estos eventos: la "
               "necesidad de liquidez dispara las ventas, la revisión anual de E4 dispara el rebalanceo y el vencimiento del CDT "
               "dispara su renovación.")
    st.markdown("**Variables de estado**")
    st.write("Tiempo actual · efectivo en la cuenta · saldo en CDT · valor del ETF · valor del inmueble · deuda · ingresos y gastos "
             "acumulados · inflación acumulada (índice de precios) · patrimonio nominal y real · estado económico vigente · "
             "estado de ocupación del inmueble (ocupado o vacante) · TPM e inflación vigentes.")
    st.markdown("**Traza de una réplica**")
    c = st.columns(3)
    with c[0]:
        est = st.selectbox("Estrategia", ["E1", "E2", "E3", "E4"], format_func=ui.NOMBRES.get, key="est_traza")
    with c[1]:
        esc = st.selectbox("Escenario", list(ui.ESCENARIOS), format_func=ui.ESCENARIOS.get, key="esc_traza")
    with c[2]:
        sem = st.number_input("Semilla", 0, 2**31 - 1, p.i("exp.semilla"), key="sem_traza")
    if st.button("Mostrar eventos de una réplica"):
        r = una_replica(estado.params(esc), est, int(sem), traza=True)
        st.session_state["traza"] = pd.DataFrame(r.traza)
    if "traza" in st.session_state:
        z = st.session_state["traza"].copy()
        z = z.rename(columns={"t": "Mes", "n": "N.°", "evento": "Evento", "detalle": "Qué pasó", "estado": "Estado", "cuenta": "Cuenta", "W": "Patrimonio"})
        st.dataframe(z[["Mes", "N.°", "Evento", "Qué pasó", "Estado", "Cuenta", "Patrimonio"]], hide_index=True, width="stretch", height=420)

with t2:
    st.write("Cada escenario cambia parámetros del modelo; la tabla se genera con los valores vigentes. El escenario base usa las "
             "condiciones estimadas con los datos. El adverso combina condiciones desfavorables. El de estrés fuerza una crisis "
             "de 12 a 24 meses que empieza en un momento aleatorio, con caída del ETF y del inmueble y recortes de la TPM.")
    filas = []
    for k, t in p.tabla.items():
        if t.adverso is None and t.estres is None:
            continue
        filas.append({"Parámetro": t.nombre, "Base": str(t.base), "Adverso": str(t.adverso) if t.adverso is not None else "igual",
                      "Estrés": str(t.estres) if t.estres is not None else "igual"})
    st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch", height=520)
    st.caption("«Igual» indica que el escenario usa el valor base. En el escenario de estrés, los valores propios rigen solo mientras dura la crisis.")

with t3:
    st.markdown("##### Verificación: ¿está bien implementado?")
    st.write("Pruebas automáticas sobre la ejecución de eventos y la actualización de estados, el reloj y la programación, los "
             "cálculos financieros, los generadores aleatorios, las reglas de decisión y casos controlados con resultados conocidos.")
    if st.button("Ejecutar pruebas"):
        with st.spinner("Corriendo pruebas..."):
            r = subprocess.run([sys.executable, "-m", "pytest", "-q", "--no-header", str(RAIZ / "tests")], capture_output=True,
                               text=True, cwd=str(RAIZ))
        st.session_state["pytest"] = (r.returncode, r.stdout[-1500:])
    if "pytest" in st.session_state:
        rc, out = st.session_state["pytest"]
        (st.success if rc == 0 else st.error)("Todas las pruebas pasaron." if rc == 0 else "Hay pruebas que fallan.")
        st.code(out)
    st.markdown("##### Validación: ¿representa bien el fenómeno?")
    st.write("Compara trayectorias simuladas con lo observado: inflación, TPM, rendimiento del ETF, valorización, vacancia y estados "
             "de la economía. «Reproduce insumos» verifica que los generadores devuelvan los datos con que se estimaron; "
             "«Contraste» compara con información que el modelo no usó.")
    n_v = st.number_input("Réplicas de validación", 50, 5000, 300, 50)
    if st.button("Validar"):
        with st.spinner("Simulando..."):
            st.session_state["val"] = validar(estado.params(), n=int(n_v), semilla=p.i("exp.semilla"))
    if "val" in st.session_state:
        v = st.session_state["val"].rename(columns={"proceso": "Proceso", "medida": "Medida", "tipo": "Tipo", "simulado": "Simulado",
                                                    "observado": "Observado", "referencia": "Referencia", "lectura": "Lectura"})
        st.dataframe(v.astype(str), hide_index=True, width="stretch")
    st.markdown("##### Sensibilidad")
    st.write("El análisis de sensibilidad (inflación, rendimiento del ETF, tasas, valorización inmobiliaria y vacancia, entre otros) "
             "está en la página Sensibilidad y capital-meta, pestaña Tornado.")
