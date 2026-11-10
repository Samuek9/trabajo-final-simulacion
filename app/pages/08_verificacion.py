import pandas as pd
import streamlit as st

from components import estado, ui
from investment_sim.engine.events import Tipo
from investment_sim.experiments.runner import una_replica
from investment_sim.experiments.validation import validar
from investment_sim.experiments.verification import verificar

p = estado.params()

st.title("Simulación, escenarios y validación")
t1, t2, t3 = st.tabs(["Eventos discretos", "Escenarios", "Verificación y validación"])

with t1:
    c = st.columns(3)
    with c[0]:
        ui.tarjeta("Reloj", "Evento a evento", "avanza al siguiente evento programado")
    with c[1]:
        ui.tarjeta("Lista de eventos", "Tiempo y prioridad", "los empates se resuelven por prioridad")
    with c[2]:
        ui.tarjeta("Estado", "En cada evento", "se actualizan las variables de estado")
    st.write("")
    ev = pd.DataFrame([{"N.°": t.numero, "Evento": t.nombre, "Prioridad en empates": t.prioridad}
                       for t in sorted(Tipo, key=lambda t: t.prioridad)])
    st.markdown("##### Eventos y prioridades")
    st.dataframe(ev, hide_index=True, width="stretch", height=(len(ev) + 1) * 35 + 3,
                 column_config={"N.°": st.column_config.NumberColumn(width="small", format="%d"),
                                "Prioridad en empates": st.column_config.NumberColumn(width="small", format="%d")})
    st.caption("La necesidad de liquidez dispara las ventas, la revisión de E4 dispara el rebalanceo y el vencimiento del CDT su renovación.")
    st.markdown("##### Variables de estado")
    VARS = {"Tiempo": ["Tiempo actual", "Estado económico vigente", "TPM e inflación vigentes"],
            "Activos y deuda": ["Efectivo en la cuenta", "Saldo en CDT", "Valor del ETF", "Valor del inmueble", "Deuda"],
            "Inmueble": ["Estado de ocupación (ocupado o vacante)"],
            "Resultado": ["Ingresos y gastos acumulados", "Índice de precios acumulado", "Patrimonio nominal y real"]}
    cv = st.columns(len(VARS))
    for col, (g, xs) in zip(cv, VARS.items()):
        with col:
            st.markdown(f"**{g}**")
            st.markdown(chr(10).join(f"- {x}" for x in xs))
    st.markdown("##### Traza de una réplica")
    c = st.columns(3)
    with c[0]:
        est = st.selectbox("Estrategia", ["E1", "E2", "E3", "E4"], format_func=ui.NOMBRES.get, key="est_traza")
    with c[1]:
        esc = st.selectbox("Escenario", list(ui.ESCENARIOS), format_func=ui.ESCENARIOS.get, key="esc_traza")
    with c[2]:
        sem = st.number_input("Semilla", 0, 2**31 - 1, p.i("exp.semilla"), key="sem_traza")
    if st.button("Mostrar traza"):
        r = una_replica(estado.params(esc), est, int(sem), traza=True)
        st.session_state["traza"] = pd.DataFrame(r.traza)
    if "traza" in st.session_state:
        z = st.session_state["traza"].copy()
        z = z.rename(columns={"t": "Mes", "n": "N.°", "evento": "Evento", "detalle": "Qué pasó", "estado": "Estado", "cuenta": "Cuenta", "W": "Patrimonio"})
        st.dataframe(z[["Mes", "N.°", "Evento", "Qué pasó", "Estado", "Cuenta", "Patrimonio"]], hide_index=True, width="stretch", height=360)
        st.caption(f"{len(z)} eventos; se muestran con desplazamiento.")

with t2:
    c = st.columns(3)
    for col, (t, n, d) in zip(c, (("Base", "Condiciones estimadas", "Parámetros estimados con los datos."),
                                  ("Adverso", "Condiciones desfavorables", "Combina varios parámetros en contra."),
                                  ("Estrés", "Crisis forzada", "12 a 24 meses desde un momento aleatorio."))):
        with col:
            ui.tarjeta(t, n, d)
    st.write("")
    filas = []
    for k, t in p.tabla.items():
        if t.adverso is None and t.estres is None:
            continue
        filas.append({"Parámetro": t.nombre, "Base": str(t.base), "Adverso": str(t.adverso) if t.adverso is not None else "igual",
                      "Estrés": str(t.estres) if t.estres is not None else "igual"})
    st.dataframe(pd.DataFrame(filas), hide_index=True, width="stretch", height=420)
    st.caption("«Igual»: usa el valor base. En estrés, los valores propios rigen solo mientras dura la crisis.")

with t3:
    st.markdown("##### Verificación")
    st.caption("Comprueba la implementación con la configuración actual: capital, horizonte, pesos y liquidez del inversionista.")
    if st.button("Verificar"):
        with st.spinner("Verificando..."):
            st.session_state["verif"] = verificar(estado.params())
    if "verif" in st.session_state:
        vf = st.session_state["verif"]
        cumplen = int((vf["Resultado"] == "Cumple").sum())
        c = st.columns(3)
        with c[0]:
            ui.tarjeta("Comprobaciones que cumplen", f"{cumplen} de {len(vf)}")
        with c[1]:
            ui.tarjeta("Que no cumplen", str(len(vf) - cumplen))
        with c[2]:
            ui.tarjeta("Componentes", str(vf["Componente"].nunique()))
        st.write("")
        vista = vf[["Componente", "Comprobación", "Resultado", "Esperado", "Obtenido"]].copy()
        vista["Componente"] = vista["Componente"].where(vista["Componente"] != vista["Componente"].shift(), "")
        st.dataframe(vista, hide_index=True, width="stretch", height=(len(vista) + 1) * 35 + 3,
                     column_config={"Componente": st.column_config.TextColumn(width=150), "Comprobación": st.column_config.TextColumn(width=280),
                                    "Resultado": st.column_config.TextColumn(width=90), "Esperado": st.column_config.TextColumn(width=240),
                                    "Obtenido": st.column_config.TextColumn(width=220)})
    st.markdown("##### Validación")
    st.caption("«Reproduce insumos»: los generadores devuelven los datos con que se estimaron. «Contraste»: se compara con información que el modelo no usó.")
    n_v = st.number_input("Réplicas de validación", 50, 5000, 300, 50)
    if st.button("Validar"):
        with st.spinner("Simulando..."):
            st.session_state["val"] = validar(estado.params(), n=int(n_v), semilla=p.i("exp.semilla"))
    if "val" in st.session_state:
        v = st.session_state["val"].rename(columns={"proceso": "Proceso", "medida": "Medida", "tipo": "Tipo", "simulado": "Simulado",
                                                    "observado": "Observado", "referencia": "Referencia", "lectura": "Lectura"})
        def legible(x):
            return f"{x:.3f}".replace(".", ",") if isinstance(x, float) else str(x)
        st.dataframe(v.map(legible), hide_index=True, width="stretch", height=420,
                     column_config={c: st.column_config.TextColumn(width=w) for c, w in
                                    {"Proceso": "small", "Medida": "medium", "Tipo": "small", "Simulado": "small",
                                     "Observado": "small", "Referencia": "small", "Lectura": "large"}.items() if c in v.columns})
