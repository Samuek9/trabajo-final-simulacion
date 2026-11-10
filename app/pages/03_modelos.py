import pandas as pd
import streamlit as st

from components import estado

estado.params()

st.title("Modelos y economía")
st.write("Cada proceso aleatorio se modela con un solo modelo, el que quedó elegido con las pruebas de hipótesis sobre los datos "
         "y con el diseño experimental. Sus parámetros están estimados con esos datos y quedan fijos: lo que se cambia en el "
         "simulador son las decisiones y la situación del inversionista (páginas Inversionista y Estrategias).")

MODELOS = [
    ("Estado de la economía", "PIB trimestral del DANE, 2006 a 2026",
     "Cadena de Markov de tres estados (expansión, normal y crisis) que cambia cada trimestre. El estado mueve la volatilidad del "
     "ETF, la vacancia, la frecuencia de las necesidades de liquidez y el reparto de la estrategia dinámica."),
    ("Inflación", "IPC anual del DANE, 2001 a 2025",
     "AR(1) en logaritmos: la inflación de cada diciembre depende de la del año anterior y de una sorpresa lognormal. Entre "
     "diciembres, línea recta."),
    ("Decisión de la Junta del Banco de la República", "202 decisiones de 2001 a 2026",
     "Logit ordenado (baja, mantiene o sube) con la tendencia de la inflación y la brecha de Taylor; el tamaño del cambio sale "
     "de las frecuencias observadas."),
    ("Tasa de la cuenta remunerada", "19 tasas de cuatro cuentas digitales, 2023 a 2026",
     "Una proporción estable de la tasa del Banco de la República."),
    ("Margen del CDT sobre la TPM", "CDT a 360 días y TPM semanal, 2008 a 2026",
     "Margen = constante + coeficiente por el cambio de la TPM en tres meses + residuo sorteado de los observados."),
    ("Rendimiento del ETF", "Valor neto del iShares COLCAP, 2011 a 2026",
     "t de Student con cinco grados de libertad, rendimiento medio constante y volatilidad que depende del estado de la economía."),
    ("Valorización del inmueble", "Índice de precios de la vivienda usada, 2001 a 2025",
     "AR(1) del logaritmo de la valorización anual: los años buenos y malos vienen seguidos."),
    ("Reajuste del canon", "IPC de arriendos del DANE, 2010 a 2025",
     "Regresión sobre el IPC del año anterior, sin pasar del tope legal."),
    ("Vacancia", "Tiempo en oferta de apartamentos en Medellín, 2025 a 2026",
     "Al vencer cada contrato el inquilino no renueva con una probabilidad que depende del estado; la vacancia dura una gamma."),
    ("Necesidades extraordinarias de liquidez", "Sin datos de monto; referencia de frecuencia en la encuesta IEFIC del DANE",
     "Proceso de Poisson con monto lognormal. La frecuencia y el monto mediano se ajustan en la página Inversionista."),
]
st.dataframe(pd.DataFrame(MODELOS, columns=["Proceso", "Datos", "Modelo elegido"]), hide_index=True, width="stretch")
st.caption("El detalle de las pruebas que respaldan cada elección está en el informe y en el libro de datos del proyecto.")
