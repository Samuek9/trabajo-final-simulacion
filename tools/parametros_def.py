"""Tabla de parámetros del simulador.

Este archivo solo sirve para escribir la hoja PARAMETROS de data/master_datos.xlsx
(tools/construir_master.py). El simulador nunca lo importa: lee la hoja del Excel.
Si un valor cambia, se cambia aquí, se reconstruye el Excel y queda registrado en
docs/decision_log.md.

Columnas: id, clave, grupo, nombre, base, adverso, estres, unidad, estado, fuente, soporte.
adverso / estres = None significa "igual que la base". En estrés el valor aplica solo
mientras dura la crisis forzada.
Estados: Configurable, Verificado, Estimado, Calibrado, Supuesto, Pendiente.
"""

X = None  # igual que la base

FILAS = [
    # ---------------- Inversionista y experimento (punto 4 y 14) ----------------
    ("P001", "inv.capital", "Inversionista", "Capital inicial", 700_000_000, X, X, "COP", "Configurable", "Equipo", "Punto 4"),
    ("P002", "inv.horizonte", "Inversionista", "Horizonte", 120, X, X, "meses", "Configurable", "Equipo", "Punto 4"),
    ("P003", "inv.meta_real", "Inversionista", "Meta: patrimonio real al final (pesos de hoy)", 1_000_000_000, X, X, "COP", "Configurable", "Equipo", "Punto 4"),
    ("P004", "inv.liquidez_min", "Inversionista", "Liquidez mínima en la cuenta", 0.05, X, X, "fracción del patrimonio", "Configurable", "Equipo", "Punto 12"),
    ("P005", "inv.modo_inmueble", "Inversionista", "Valor del inmueble: proporcional al capital o fijo", "proporcional", X, X, "texto", "Configurable", "Equipo", "Punto 11"),
    ("P006", "inv.valor_inmueble_fijo", "Inversionista", "Valor del inmueble si el modo es fijo", 300_000_000, X, X, "COP", "Configurable", "Hoja INM (precio de referencia)", "5.4"),
    ("P007", "exp.replicas", "Experimento", "Número de réplicas", 1068, X, X, "réplicas", "Calculado", "n = 1,96² × 0,25 ÷ 0,03²", "Punto 14"),
    ("P008", "exp.semilla", "Experimento", "Semilla base (réplica r usa semilla + r)", 20261107, X, X, "entero", "Configurable", "Equipo", "Punto 14"),
    ("P009", "eco.estado_inicial", "Economía", "Estado económico inicial", "Normal", X, X, "texto", "Verificado", "DANE: PIB 2026-II creció 3,5 %", "Punto 8"),

    # ---------------- Estados económicos (punto 8) ----------------
    ("P010", "eco.umbral_expansion", "Economía", "Expansión si el PIB real crece más de", 0.04, X, X, "crecimiento anual", "Supuesto", "Equipo (hoja MARKOV)", "Punto 8"),
    ("P011", "eco.umbral_crisis", "Economía", "Crisis si el PIB real crece menos de", 0.01, X, X, "crecimiento anual", "Supuesto", "Mismo umbral en todo el libro y en los datos. Con 1 % cuentan como crisis 2008-09, 2020 y 2023-24 (12 trimestres desde 2006); con 0 % solo quedarían 4 trimestres (pandemia y uno de 2023) y la crisis de 2009 no contaría, así que la matriz de Markov no tendría con qué estimarse. El estado se lee como crisis o estancamiento", "Punto 8"),
    ("P012", "eco.paso", "Economía", "Cada cuánto cambia el estado", 3, X, X, "meses", "Verificado", "PIB trimestral", "Punto 8"),
    ("P013", "markov.E_E", "Economía", "P(Expansión → Expansión)", 25 / 30, X, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (25 de 30)", "MARKOV"),
    ("P014", "markov.E_N", "Economía", "P(Expansión → Normal)", 5 / 30, X, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (5 de 30)", "MARKOV"),
    ("P015", "markov.E_C", "Economía", "P(Expansión → Crisis)", 0.0, X, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (0 de 30)", "MARKOV"),
    ("P016", "markov.N_E", "Economía", "P(Normal → Expansión)", 4 / 39, X, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (4 de 39)", "MARKOV"),
    ("P017", "markov.N_N", "Economía", "P(Normal → Normal)", 32 / 39, 1 - 4 / 39 - 0.16, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (32 de 39); adverso: complemento", "MARKOV"),
    ("P018", "markov.N_C", "Economía", "P(Normal → Crisis)", 3 / 39, 0.16, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (3 de 39); adverso: el doble", "Punto 13"),
    ("P019", "markov.C_E", "Economía", "P(Crisis → Expansión)", 0.0, X, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (0 de 12)", "MARKOV"),
    ("P020", "markov.C_N", "Economía", "P(Crisis → Normal)", 3 / 12, 0.15, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (3 de 12); adverso: salida más lenta", "Punto 13"),
    ("P021", "markov.C_C", "Economía", "P(Crisis → Crisis)", 9 / 12, 0.85, X, "probabilidad", "Estimado", "DANE PIB 2006-2026 (9 de 12)", "MARKOV"),

    # ---------------- Inflación (7.1) ----------------
    ("P030", "inf.c", "Inflación", "Constante a de ln(π) = a + φ·ln(π anterior) + σ·Z", 0.8866950789741039, (1 - 0.40811296853944157) * 1.8718021769015913, X, "ln de %", "Estimado", "DANE IPC 2001-2025; adverso: nivel normal 6,5 %", "INF, 7.1"),
    ("P031", "inf.phi", "Inflación", "Memoria φ", 0.40811296853944157, X, X, "coeficiente", "Estimado", "DANE IPC 2001-2025", "INF, 7.1"),
    ("P032", "inf.sigma", "Inflación", "σ de las sorpresas", 0.46887476395876315, X, X, "ln de %", "Estimado", "DANE IPC 2001-2025", "INF, 7.1"),
    ("P033", "inf.inicial", "Inflación", "Inflación anual al iniciar (último dato)", 6.28979, X, X, "% anual", "Verificado", "DANE, septiembre de 2026", "INF"),

    # ---------------- TPM y Junta (7.2) ----------------
    ("P040", "tpm.inicial", "TPM", "TPM al iniciar", 12.25, X, X, "% anual", "Verificado", "Banco de la República, desde el 1 de octubre de 2026", "5.1"),
    ("P039", "tpm.hace_3m", "TPM", "TPM tres meses antes de iniciar", 12.0, X, X, "% anual", "Verificado", "Banco de la República, julio de 2026", "7.4"),
    ("P041", "tpm.min", "TPM", "TPM mínima", 1.75, X, X, "% anual", "Verificado", "Mínimo histórico (2020)", "7.2"),
    ("P042", "tpm.max", "TPM", "TPM máxima", 15.0, X, X, "% anual", "Supuesto", "Equipo", "7.2"),
    ("P043", "tpm.meses_junta", "TPM", "Meses con reunión de la Junta", "1;3;4;6;7;9;10;12", X, X, "lista", "Verificado", "Calendario del Banco de la República", "7.2"),
    ("P044", "tpm.umbral_tendencia", "TPM", "Umbral de tendencia de la inflación en 3 meses", 0.5, X, X, "puntos", "Supuesto", "Hoja TPM_Cuenta", "7.2"),
    ("P045", "junta.bajando", "TPM", "P(baja; mantiene; sube) con inflación bajando", "0.615;0.385;0", X, X, "lista", "Estimado", "46 reuniones (8 de 13; 5 de 13)", "7.2"),
    ("P046", "junta.estable", "TPM", "P(baja; mantiene; sube) con inflación estable", "0.125;0.625;0.25", X, X, "lista", "Estimado", "46 reuniones (2; 10; 4 de 16)", "7.2"),
    ("P047", "junta.subiendo", "TPM", "P(baja; mantiene; sube) con inflación subiendo", "0;0.176;0.824", X, X, "lista", "Estimado", "46 reuniones (3; 14 de 17)", "7.2"),
    ("P048", "junta.subidas", "TPM", "Tamaños de subida", "0.25;0.5;0.75;1;1.5", X, X, "lista de puntos", "Estimado", "Subidas observadas", "7.2"),
    ("P049", "junta.p_subidas", "TPM", "Probabilidad de cada tamaño de subida", "0.2222;0.1111;0.1111;0.4444;0.1112", X, X, "lista", "Estimado", "Subidas observadas (2; 1; 1; 4; 1 de 9)", "7.2"),
    ("P050", "junta.bajadas", "TPM", "Tamaños de bajada", "0.25;0.5", X, X, "lista de puntos", "Estimado", "Bajadas observadas", "7.2"),
    ("P051", "junta.p_bajadas", "TPM", "Probabilidad de cada tamaño de bajada", "0.4;0.6", X, X, "lista", "Estimado", "Bajadas observadas (4 y 6 de 10)", "7.2"),
    ("P054", "junta.modelo", "TPM", "Modelo de la Junta: logit (tendencia y brecha de Taylor, 2001-2026) o tendencia (tabla de 7.2)", "logit", X, X, "texto", "Estimado", "Validación: con la tabla la TPM se pega al tope (ver VALIDACION)", "7.2, punto 16"),
    ("P055", "junta.logit", "TPM", "Logit ordenado: b tendencia; b brecha de Taylor; k1; k2", "1.267202;-0.140946;-1.621086;1.753157", X, X, "lista", "Estimado", "202 decisiones 2001-2026 con el calendario de 8 reuniones (tools/estimar_junta.py); p tendencia < 0,001, p Taylor = 0.062".replace(".", ",", 1) if False else "202 decisiones 2001-2026 con el calendario de 8 reuniones (tools/estimar_junta.py); p tendencia < 0,001, p Taylor = 0,062", "TPM_Logit"),
    ("P056", "junta.r_neutral", "TPM", "Tasa real neutral r* de la brecha de Taylor", 1.0, X, X, "puntos", "Supuesto", "Solo desplaza los umbrales; el ajuste no cambia con r*", "TPM_Logit"),
    ("P057", "junta.meta_inflacion", "TPM", "Meta de inflación del Banco de la República", 3.0, X, X, "% anual", "Verificado", "Banco de la República, meta de 3 % desde 2010", "TPM_Logit"),
    ("P058", "junta.subidas_logit", "TPM", "Logit: tamaños de subida", "0.25;0.5;0.75;1;1.5;2.5", X, X, "lista de puntos", "Estimado", "42 subidas 2001-2026 (calendario de 8 reuniones)", "TPM_Logit"),
    ("P059", "junta.p_subidas_logit", "TPM", "Logit: probabilidad de cada subida", "0.404762;0.142857;0.119048;0.285714;0.02381;0.02381", X, X, "lista", "Estimado", "42 subidas 2001-2026 (calendario de 8 reuniones)", "TPM_Logit"),
    ("P064", "junta.bajadas_logit", "TPM", "Logit: tamaños de bajada", "0.25;0.5;0.75;1;1.5;2.5", X, X, "lista de puntos", "Estimado", "43 bajadas 2001-2026 (calendario de 8 reuniones)", "TPM_Logit"),
    ("P065", "junta.p_bajadas_logit", "TPM", "Logit: probabilidad de cada bajada", "0.325581;0.348837;0.186047;0.069767;0.023256;0.046512", X, X, "lista", "Estimado", "43 bajadas 2001-2026 (calendario de 8 reuniones)", "TPM_Logit"),
    ("P052", "junta.bajadas_estres", "TPM", "Tamaños de bajada durante la crisis forzada", "0.5;1;1.5;2", X, X, "lista de puntos", "Supuesto", "Recortes de 1 a 2 puntos en 2009 (TPM oficial)", "Punto 13"),
    ("P053", "junta.p_bajadas_estres", "TPM", "Probabilidad de cada bajada durante la crisis forzada", "0.25;0.25;0.25;0.25", X, X, "lista", "Supuesto", "Equipo: sin datos para repartirlas", "Punto 13"),

    # ---------------- Cuenta remunerada (7.3) ----------------
    ("P060", "cta.alfa", "Cuenta", "Tasa de la cuenta = α × TPM", 0.8978, 0.80, X, "fracción", "Estimado", "19 tasas de 4 cuentas digitales (Nu, Pibank, Lulo, RappiPay) en 5 fechas, dic-2023 a oct-2026; IC 95 % por entidad 0,755 a 1,046. Ualá se excluye: es un depósito de bajo monto con tope cercano a $10 M. En las otras 4 no se encontró tope de saldo (comparadores, oct-2026); adverso 0,80", "TPM_Cuenta, 7.3"),
    ("P061", "imp.retencion", "Costos", "Retención en la fuente sobre intereses del CDT", 0.04, X, X, "fracción", "Verificado", "4 % para CDT y CDAT (Decreto 572 de 2025, DUR 1625 de 2016 art. 1.2.4.2.83)", "5.2"),
    ("P066", "imp.retencion_cuenta", "Costos", "Retención en la fuente sobre intereses de la cuenta de ahorros", 0.07, X, X, "fracción", "Verificado", "7 %, tarifa general de rendimientos financieros (DUR 1625 de 2016 art. 1.2.4.2.5); se cobra sobre todo el interés porque el del día supera la base mínima (art. 1.2.4.2.87)", "5.2"),
    ("P062", "imp.gmf", "Costos", "Gravamen a los movimientos financieros (4x1000)", 0.004, X, X, "fracción", "Verificado", "Estatuto Tributario, art. 871", "5.1"),
    ("P063", "imp.gmf_exencion", "Costos", "Retiros mensuales exentos del 4x1000 (350 UVT)", 350 * 52_374, X, X, "COP por mes", "Verificado", "E.T. art. 879; UVT 2026 = $52.374 (por confirmar)", "5.1"),

    # ---------------- CDT (7.4) ----------------
    ("P070", "cdt.plazo", "CDT", "Plazo del CDT", 12, X, X, "meses", "Configurable", "Equipo", "5.2"),
    ("P071", "cdt.modelo", "CDT", "Modelo del margen: m5 (TPM + a + b·ΔTPM + ε) o fase", "m5", X, X, "texto", "Estimado", "Selección de 7.4", "7.4"),
    ("P072", "cdt.a", "CDT", "M5: constante a", 1.0596008700709776, X, X, "puntos", "Estimado", "BanRep CDT 360 y TPM, 982 semanas 2008-2026", "CDT, 7.4"),
    ("P073", "cdt.b", "CDT", "M5: pendiente b sobre ΔTPM de 3 meses", 0.6718857912760832, X, X, "coeficiente", "Estimado", "BanRep, 982 semanas (p = 0,021 con HAC)", "CDT, 7.4"),
    ("P074", "cdt.residuos", "CDT", "M5: cómo se sortea ε", "empirico", X, X, "texto", "Estimado", "Residuos observados (data/raw/cdt_semanal_2008.csv)", "7.4"),
    ("P075", "cdt.t_nu", "CDT", "M5 alternativo: ε ~ t, grados de libertad", 3.530694440831547, X, X, "ν", "Estimado", "Ajuste de máxima verosimilitud", "7.4"),
    ("P076", "cdt.t_escala", "CDT", "M5 alternativo: escala de ε", 0.6582663941647997, X, X, "puntos", "Estimado", "Ajuste de máxima verosimilitud", "7.4"),
    ("P077", "cdt.fase_subida", "CDT", "Fase: margen medio; desviación (subida)", "1.5511;1.5495", X, X, "puntos", "Estimado", "317 semanas", "7.4"),
    ("P078", "cdt.fase_estable", "CDT", "Fase: margen medio; desviación (estable)", "0.9419;0.5000", X, X, "puntos", "Estimado", "343 semanas", "7.4"),
    ("P079", "cdt.fase_bajada", "CDT", "Fase: margen medio; desviación (bajada)", "0.7752;0.8475", X, X, "puntos", "Estimado", "322 semanas", "7.4"),
    ("P080", "cdt.margen_min", "CDT", "Margen mínimo observado", -1.96, X, X, "puntos", "Estimado", "Mínimo 2008-2026", "7.4"),
    ("P081", "cdt.margen_max", "CDT", "Margen máximo observado", 6.40, X, X, "puntos", "Estimado", "Máximo 2008-2026", "7.4"),
    ("P082", "cdt.castigo", "CDT", "Castigo c en la venta anticipada", 0.01, X, X, "fracción anual", "Supuesto", "Equipo; sensibilidad 0 a 2 puntos", "7.4"),

    # ---------------- ETF (7.5 y punto 8) ----------------
    ("P090", "etf.modelo", "ETF", "Modelo: t5_estado (t con volatilidad del estado) o jf", "t5_estado", X, X, "texto", "Estimado", "Punto 8: la volatilidad cambia con el estado", "7.5, punto 8"),
    ("P091", "etf.mu", "ETF", "Rendimiento logarítmico medio mensual", 0.0036327192628493733, 0.0, X, "ln mensual", "Estimado", "BlackRock ICOLCAP ago-2011 a jun-2026 (179 meses, la misma muestra de las volatilidades por estado); no significativo (p = 0,37); la media 2024-2026 es mucho mayor que la de 2011-2023, por eso el adverso es 0", "ETF, 7.5"),
    ("P092", "etf.nu", "ETF", "Grados de libertad de la t", 5.0, X, X, "ν", "Estimado", "Máxima verosimilitud (ν = 5,34; redondeado a 5)", "7.5"),
    ("P093", "etf.vol_E", "ETF", "Volatilidad mensual en expansión", 0.04706, X, X, "desviación mensual", "Estimado", "Hoja MARKOV, expansión + normal (155 meses)", "Punto 8"),
    ("P094", "etf.vol_N", "ETF", "Volatilidad mensual en normal", 0.04706, X, X, "desviación mensual", "Estimado", "Hoja MARKOV, expansión + normal (155 meses)", "Punto 8"),
    ("P095", "etf.vol_C", "ETF", "Volatilidad mensual en crisis", 0.08783, X, X, "desviación mensual", "Estimado", "Hoja MARKOV, crisis (24 meses). Levene p = 0,004; sin marzo de 2020 p = 0,080: descansa en pocos meses", "Punto 8"),
    ("P096", "etf.jf", "ETF", "Alternativa Jones-Faddy: a; b; ubicación; escala", "2.5447;3.4477;0.024712;0.040623", X, X, "lista", "Estimado", "Máxima verosimilitud, 182 meses", "7.5"),
    ("P097", "etf.caida_estres", "ETF", "Caída del ETF al iniciar la crisis forzada", 0.0, X, 0.25, "fracción", "Supuesto", "Marzo de 2020: −27 %", "Punto 13"),
    ("P098", "etf.comision", "ETF", "Comisión por compra o venta", 0.005, X, X, "fracción", "Supuesto", "Tarifas de comisionistas", "5.3"),

    # ---------------- Inmueble (5.4, 7.6) ----------------
    ("P110", "inm.costo_compra", "Inmueble", "Costos de compra", 0.025, X, X, "fracción del valor", "Verificado", "Tarifas notariales y de registro 2026", "5.4"),
    ("P111", "inm.canon_inicial", "Inmueble", "Canon de un contrato nuevo", 0.006, X, X, "fracción mensual del valor", "Calibrado", "Lonja de Propiedad Raíz de Medellín, edición 275: canon ÷ precio de venta por m² en 11 barrios y tamaños de Medellín (30 a 70 m²), mediana 0,59 %, rango 0,43 % a 0,66 %; sensibilidad 0,5 % a 0,7 %", "5.4"),
    ("P112", "inm.comision_inmobiliaria", "Inmueble", "Comisión de la inmobiliaria (10 % + IVA)", 0.119, X, X, "fracción del canon", "Verificado", "Tarifas publicadas", "5.4"),
    ("P113", "inm.ph", "Inmueble", "Administración del edificio (pesos de hoy)", 185_000, X, X, "COP por mes", "Verificado", "Lonja de Medellín; sensibilidad 165.000 a 245.000", "5.4"),
    ("P114", "inm.predial", "Inmueble", "Predial", 0.006, X, X, "fracción anual del valor", "Verificado", "Acuerdo municipal + avalúo", "5.4"),
    ("P115", "inm.mantenimiento", "Inmueble", "Mantenimiento preventivo", 0.003, X, X, "fracción anual del valor", "Supuesto", "Equipo; 0,2 % a 0,5 %", "5.4"),
    ("P116", "inm.pintura", "Inmueble", "Pintura al cambiar de inquilino (pesos de hoy)", 1_800_000, X, X, "COP", "Calibrado", "Tarifas por m²", "5.4"),
    ("P117", "inm.costo_venta", "Inmueble", "Costos de venta (3 % + IVA, notaría 0,27 %, retención 1 %)", 0.0484, X, X, "fracción del precio", "Verificado", "Tarifas y norma", "5.4"),
    ("P118", "inm.descuento_urgente", "Inmueble", "Descuento por venta urgente", 0.10, X, X, "fracción", "Supuesto", "Sin dato público del precio de cierre frente al pedido. Representa lo que se cede para vender rápido, incluida la comisión más alta que se le da a la inmobiliaria; se paga además de los costos de venta normales. Sensibilidad 5 % a 20 %", "5.4"),
    ("P119", "inm.venta_min", "Inmueble", "Tiempo de venta urgente: mínimo", 3, X, X, "meses", "Supuesto", "Más corto que una venta normal (4 a 7 meses, Valora Analitik con datos de Fincaraíz) porque el descuento compra velocidad", "Punto 12"),
    ("P120", "inm.venta_max", "Inmueble", "Tiempo de venta urgente: máximo", 6, X, X, "meses", "Supuesto", "Más corto que una venta normal (4 a 7 meses, Valora Analitik con datos de Fincaraíz) porque el descuento compra velocidad", "Punto 12"),
    ("P121", "inm.tope_canon", "Inmueble", "Tope del canon", 0.01, X, X, "fracción mensual del valor", "Verificado", "Ley 820 de 2003, art. 18", "5.4"),
    ("P122", "inm.reajuste_modelo", "Inmueble", "Reajuste del canon: dane (a + b·IPC previo + ε) o ipc (todo el IPC)", "dane", X, X, "texto", "Estimado", "IPC de arriendo del DANE 2010-2025", "5.4.3"),
    ("P123", "inm.reajuste_a", "Inmueble", "Reajuste DANE: constante a", 1.6784667368428081, X, X, "puntos", "Estimado", "DANE, 16 años", "INM_Arriendo"),
    ("P124", "inm.reajuste_b", "Inmueble", "Reajuste DANE: pendiente b sobre el IPC previo", 0.47782486531386226, X, X, "coeficiente", "Estimado", "DANE, 16 años (R² = 0,71)", "INM_Arriendo"),
    ("P125", "inm.reajuste_sigma", "Inmueble", "Reajuste DANE: σ de ε", 0.9409701223606414, X, X, "puntos", "Estimado", "DANE, 16 años", "INM_Arriendo"),

    # ---------------- Valorización (7.6 y punto 8) ----------------
    ("P130", "val.modelo", "Valorización", "Modelo: ar1 (IPVU, vivienda usada) o estado (normal por estado, IPVNBR)", "ar1", X, X, "texto", "Estimado", "IPVU: mide vivienda usada, como el inmueble del modelo. Con el IPVU el estado económico no cambia la valorización (Welch, p = 0,275); el IPVNBR mide vivienda nueva y queda solo como alternativa", "7.6, punto 8"),
    ("P131", "val.mu_E", "Valorización", "Valorización logarítmica media en expansión", 0.11935, 0.09935, X, "ln anual", "Estimado", "Solo con val.modelo = estado (alternativa). IPVNBR Medellín 2006-2025, 8 años; adverso −2 puntos", "MARKOV"),
    ("P132", "val.mu_N", "Valorización", "Valorización logarítmica media en normal", 0.08316, 0.06316, X, "ln anual", "Estimado", "Solo con val.modelo = estado (alternativa). IPVNBR Medellín, 12 años normal + crisis; adverso −2 puntos", "MARKOV"),
    ("P133", "val.mu_C", "Valorización", "Valorización logarítmica media en crisis", 0.08316, 0.06316, -0.05129, "ln anual", "Estimado", "Solo con val.modelo = estado (alternativa). Igual a normal (2 años de crisis); estrés: −5 % anual", "MARKOV, punto 13"),
    ("P134", "val.sd_E", "Valorización", "Desviación en expansión", 0.03466, X, X, "ln anual", "Estimado", "Solo con val.modelo = estado (alternativa). IPVNBR Medellín", "MARKOV"),
    ("P135", "val.sd_N", "Valorización", "Desviación en normal", 0.02152, X, X, "ln anual", "Estimado", "Solo con val.modelo = estado (alternativa). IPVNBR Medellín", "MARKOV"),
    ("P136", "val.sd_C", "Valorización", "Desviación en crisis", 0.02152, X, X, "ln anual", "Estimado", "Solo con val.modelo = estado (alternativa). IPVNBR Medellín", "MARKOV"),
    ("P137", "val.ar1", "Valorización", "AR(1) del IPVU: media; φ; σ", "0.0755;0.676;0.0236", "0.0555;0.676;0.0236", "-0.05129;0;0.0236", "lista", "Estimado", "IPVU nominal 2001-2025 (n = 25, máxima verosimilitud); adverso: media −2 puntos; estrés: −5 % anual mientras dura la crisis", "INM_Valoriz, 7.6"),

    # ---------------- Vacancia (7.7) ----------------
    ("P140", "vac.p_E", "Vacancia", "Probabilidad de no renovar en expansión", 0.05, 0.075, X, "probabilidad", "Supuesto", "Sin datos de no renovación; escenario a sensibilidad (Ciencuadras 2026 reporta 0,6 % de vacancia, no la tasa de no renovación); adverso × 1,5", "7.7"),
    ("P141", "vac.p_N", "Vacancia", "Probabilidad de no renovar en normal", 0.10, 0.15, X, "probabilidad", "Supuesto", "Sin datos de no renovación; escenario a sensibilidad; adverso × 1,5", "7.7"),
    ("P142", "vac.p_C", "Vacancia", "Probabilidad de no renovar en crisis", 0.30, 0.45, 0.30, "probabilidad", "Supuesto", "Sin datos; escenario de crisis a sensibilidad; adverso × 1,5", "7.7"),
    ("P143", "vac.forma", "Vacancia", "Forma de la gamma de la duración", 1.10, X, X, "forma", "Estimado", "Curva de tiempo en oferta de apartamentos, Medellín (Palomma, ago-2025 a ago-2026; publicada en un artículo sobre oficinas)", "7.7"),
    ("P144", "vac.media_E", "Vacancia", "Duración media en expansión", 1.0, 1.5, X, "meses", "Supuesto", "Sin datos por estado; escenario a sensibilidad; adverso × 1,5", "7.7"),
    ("P145", "vac.media_N", "Vacancia", "Duración media en normal", 1.5, 2.25, X, "meses", "Calibrado", "Palomma; adverso × 1,5", "7.7"),
    ("P146", "vac.media_C", "Vacancia", "Duración media en crisis", 6.0, 9.0, 9.0, "meses", "Supuesto", "Sin datos por estado; escenario de crisis a sensibilidad", "7.7"),

    # ---------------- Liquidez extraordinaria (punto 12) ----------------
    ("P150", "liq.tasa_E", "Liquidez", "Necesidades por mes en expansión", 1 / 24, 1.5 / 24, X, "eventos por mes", "Calibrado", "IEFIC 2017-2018 (DANE), Medellín, hogares del 10 % de ingreso más alto: un mes con gastos mayores que los ingresos cubierto con ahorros o venta de activos cada 16 a 23 meses. Es una cota alta (cuenta déficits de cualquier tamaño); se usa 1 cada 24. Adverso × 1,5 (tools/iefic_liquidez.py)", "Punto 12"),
    ("P151", "liq.tasa_N", "Liquidez", "Necesidades por mes en normal", 1 / 24, 1.5 / 24, X, "eventos por mes", "Calibrado", "IEFIC 2017-2018 (DANE), Medellín, hogares del 10 % de ingreso más alto: un mes con gastos mayores que los ingresos cubierto con ahorros o venta de activos cada 16 a 23 meses. Es una cota alta (cuenta déficits de cualquier tamaño); se usa 1 cada 24. Adverso × 1,5 (tools/iefic_liquidez.py)", "Punto 12"),
    ("P152", "liq.tasa_C", "Liquidez", "Necesidades por mes en crisis", 1 / 12, 1.5 / 12, 1 / 12, "eventos por mes", "Supuesto", "Sin datos por estado (la IEFIC es de 2017-2018, sin crisis): el doble en crisis es supuesto; adverso × 1,5", "Punto 12"),
    ("P153", "liq.mediana", "Liquidez", "Mediana del monto (pesos de hoy)", 20_000_000, X, X, "COP", "Supuesto", "Sin fuente: los resultados se reportan con y sin necesidades", "Punto 12"),
    ("P154", "liq.sigma", "Liquidez", "σ del logaritmo del monto", 0.8, X, X, "ln", "Supuesto", "Equipo (percentil 90 ≈ $56 M)", "Punto 12"),
    ("P155", "liq.mult_monto", "Liquidez", "Multiplicador del monto", 1.0, X, 2.0, "veces", "Supuesto", "Estrés: monto × 2", "Punto 13"),

    # ---------------- Reparaciones (evento 10) ----------------
    ("P160", "rep.tasa", "Reparación", "Reparaciones extraordinarias por mes", 0.0, X, X, "eventos por mes", "Pendiente", "Sin datos: el evento existe pero está apagado", "Diagrama, evento 10"),
    ("P161", "rep.mediana", "Reparación", "Mediana del costo (pesos de hoy)", 0.0, X, X, "COP", "Pendiente", "Sin datos", "Diagrama, evento 10"),
    ("P162", "rep.sigma", "Reparación", "σ del logaritmo del costo", 0.0, X, X, "ln", "Pendiente", "Sin datos", "Diagrama, evento 10"),

    # ---------------- Escenario de estrés (punto 13) ----------------
    ("P170", "estres.activo", "Escenario", "¿Se fuerza una crisis?", 0, X, 1, "0 o 1", "Configurable", "Punto 13", "Punto 13"),
    ("P171", "estres.inicio_min", "Escenario", "Inicio de la crisis forzada: mínimo", 1, X, X, "mes", "Supuesto", "Equipo", "Punto 13"),
    ("P172", "estres.margen_final", "Escenario", "Inicio máximo = H menos este valor", 12, X, X, "meses", "Supuesto", "Equipo", "Punto 13"),
    ("P173", "estres.dur_min", "Escenario", "Duración mínima de la crisis forzada", 12, X, X, "meses", "Supuesto", "Promedio histórico: 12 meses", "Punto 13"),
    ("P174", "estres.dur_max", "Escenario", "Duración máxima de la crisis forzada", 24, X, X, "meses", "Supuesto", "Equipo", "Punto 13"),

    # ---------------- Estrategias (punto 11) ----------------
    ("P200", "E1.pesos", "Estrategias", "E1 Seguridad: cuenta; CDT; ETF; inmueble", "0.20;0.60;0.20;0", X, X, "fracción del capital", "Configurable", "Equipo", "Punto 11"),
    ("P201", "E2.pesos", "Estrategias", "E2 Moderada: cuenta; CDT; ETF; inmueble", "0.10;0.26;0.21;0.43", X, X, "fracción del capital", "Configurable", "Equipo", "Punto 11"),
    ("P202", "E3.pesos", "Estrategias", "E3 Arriesgada: cuenta; CDT; ETF; inmueble", "0.07;0.08;0.30;0.55", X, X, "fracción del capital", "Configurable", "Equipo", "Punto 11"),
    ("P203", "E4.pesos", "Estrategias", "E4 Dinámica, inicio: cuenta; CDT; ETF; inmueble", "0.10;0.26;0.21;0.43", X, X, "fracción del capital", "Configurable", "Equipo (igual que E2)", "Punto 11"),
    ("P204", "E4.obj_E", "Estrategias", "E4 en expansión: cuenta; CDT; ETF del portafolio financiero", "0.10;0.35;0.55", X, X, "fracción", "Configurable", "Equipo", "Punto 11"),
    ("P205", "E4.obj_N", "Estrategias", "E4 en normal: cuenta; CDT; ETF", "0.175;0.45;0.375", X, X, "fracción", "Configurable", "Equipo", "Punto 11"),
    ("P206", "E4.obj_C", "Estrategias", "E4 en crisis: cuenta; CDT; ETF", "0.25;0.55;0.20", X, X, "fracción", "Configurable", "Equipo (volatilidad del ETF casi el doble)", "Punto 11"),
    ("P207", "E4.banda", "Estrategias", "E4: solo se ajusta si se aleja más de", 0.05, X, X, "puntos (fracción)", "Configurable", "Equipo", "Punto 11"),
    ("P208", "E4.cada", "Estrategias", "E4: cada cuánto revisa", 12, X, X, "meses", "Configurable", "Equipo (coincide con el CDT)", "Punto 11"),
]
