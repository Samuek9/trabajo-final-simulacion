# Junta en data/master_datos.xlsx:
#   - las hojas nuevas que arma tools/construir_master.py,
#   - las hojas del consolidado de los puntos 1 a 7 (sin la vieja Parametros_simulador),
#   - las hojas del análisis de Markov del equipo, renombradas MARKOV*.
# Usa Excel porque openpyxl borra los gráficos al guardar. Los archivos de origen no se modifican.
#
#   powershell -File tools/unir_master.ps1 -Nuevas hojas_nuevas.xlsx
param(
    [Parameter(Mandatory = $true)][string]$Nuevas,
    [string]$Consolidado,
    [string]$Markov,
    [string]$Destino
)
$ErrorActionPreference = "Stop"
$aqui = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Consolidado) { $Consolidado = "$aqui\..\..\insumos_equipo\07_consolidado_puntos_1_a_7\Trabajo_Final_Simulacion_puntos_1_a_7.xlsx" }
if (-not $Markov) { $Markov = "$aqui\..\..\insumos_equipo\01_estados_markov\Analisis_completo_estados_economicos_Markov.xlsx" }
if (-not $Destino) { $Destino = "$aqui\..\data\master_datos.xlsx" }
$Nuevas = (Resolve-Path $Nuevas).Path
$Consolidado = (Resolve-Path $Consolidado).Path
$Markov = (Resolve-Path $Markov).Path
$Destino = [IO.Path]::GetFullPath($Destino)
$falta = [Type]::Missing

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $dest = $xl.Workbooks.Open($Consolidado, 0, $true)
    $nue = $xl.Workbooks.Open($Nuevas, 0, $true)
    $mar = $xl.Workbooks.Open($Markov, 0, $true)

    # Renombrar en la copia abierta (no se guarda) antes de copiar, para que no choquen con PIB y ETF.
    $nombres = @{ "Resultado" = "MARKOV"; "PIB" = "MARKOV_PIB"; "ETF" = "MARKOV_ETF"; "Vivienda" = "MARKOV_Vivienda"; "Fuentes" = "MARKOV_Fuentes" }
    foreach ($ws in @($mar.Worksheets)) { if ($nombres.ContainsKey($ws.Name)) { $ws.Name = $nombres[$ws.Name] } }

    $nue.Worksheets.Copy($dest.Worksheets.Item(1))
    $mar.Worksheets.Copy($falta, $dest.Worksheets.Item($dest.Worksheets.Count))
    $dest.Worksheets.Item("TPM_Logit").Move($falta, $dest.Worksheets.Item("TPM_Cuenta"))
    $ultima = $dest.Worksheets.Item($dest.Worksheets.Count)
    $dest.Worksheets.Item("VALIDACION").Move($falta, $ultima)
    $ultima = $dest.Worksheets.Item($dest.Worksheets.Count)
    $dest.Worksheets.Item("RESULTADOS").Move($falta, $ultima)
    $dest.Worksheets.Item("Parametros_simulador").Delete()

    # Correcciones de la auditoría sobre las copias (los originales del equipo no se tocan).
    # B1: el conteo de años dentro del rango meta apuntaba a una columna vacía.
    $dest.Worksheets.Item("INF").Range("A77").Value2 = "Con otras pruebas la conclusión cambia: Shapiro-Wilk da p = 0,057 y Lilliefors p = 0,29, y ninguna rechaza la normal al 5 %. Con 25 años las pruebas tienen poca potencia. La asimetría (1,21) sí se ve, y en logaritmos ninguna prueba rechaza (Shapiro-Wilk p = 0,81), por eso el modelo trabaja con ln(π)."
    $dest.Worksheets.Item("INF").Range("C40").Formula = "=SUMPRODUCT((B17:B32>=G17:G32)*(B17:B32<=H17:H32))"
    # B6: los ejemplos de generación usaban ALEATORIO() y cambiaban en cada recálculo; quedan con u = 0,5.
    $etf = $dest.Worksheets.Item("ETF")
    $etf.Range("B52").Formula = "=BETA.INV(0.5,B51,C51)"
    $etf.Range("B57").Formula = "=B56+C56*T.INV(0.5,B55)"
    $etf.Range("A52").Value2 = [string]$etf.Range("A52").Value2 + " (ejemplo con ALEATORIO() = 0,5)"
    $etf.Range("A57").Value2 = [string]$etf.Range("A57").Value2 + " (ejemplo con ALEATORIO() = 0,5)"
    # A3: un solo umbral de crisis (PIB < 1 %) en todo el libro, como MARKOV y PARAMETROS.
    $dest.Worksheets.Item("PIB").Range("H4").Value2 = 1
    $filas = @(@(44, -0.003843, 0.052647), @(111, 0.006340, 0.044575), @(24, 0.004815, 0.087832))
    for ($i = 0; $i -lt 3; $i++) { for ($j = 0; $j -lt 3; $j++) { $col = @("B", "C", "D")[$j]; $etf.Range("$col$(25 + $i)").Value2 = [double]$filas[$i][$j] } }
    $etf.Range("A28").Value2 = "Con el umbral de 1 % (el mismo de MARKOV y PARAMETROS) la media del ETF no cambia con el estado (Kruskal-Wallis p = 0,39), pero la desviación en crisis (24 meses) es mayor que en los demás estados (Levene p = 0,004). Sin marzo de 2020 esa diferencia deja de ser significativa (p = 0,080), así que la volatilidad de crisis descansa en pocos meses y se reporta con esa advertencia."
    # M8: α sin Ualá (depósito de bajo monto con tope cercano a $10 M). COUNT, AVERAGE, STDEV y CORREL saltan el texto.
    $cta = $dest.Worksheets.Item("TPM_Cuenta")
    foreach ($r in 5..28) { if ([string]$cta.Range("AS$r").Value2 -eq "Ualá") { $cta.Range("AV$r").Value2 = "excluida" } }
    $cta.Range("AF6").Value2 = "19 tasas de 4 cuentas digitales (Nu, Pibank, Lulo, RappiPay) en 5 fechas (dic-2023 a oct-2026). Ualá se excluye: es un depósito de bajo monto con tope cercano a `$10 M."
    # M6: el canon sale de la Lonja (edición 275); se quitan los avisos de Aptología de la copia.
    $inm = $dest.Worksheets.Item("INM")
    $inm.Range("A1:K1206").ClearContents()
    $inm.Range("A1").Value2 = "Inmueble: canon frente al precio de venta en barrios de Medellín (Lonja, edición 275)"
    $inm.Range("A2").Value2 = "Fuente: Lonja de Propiedad Raíz de Medellín y Antioquia, Mercado inmobiliario residencial, edición 275 (agosto de 2025). En el boletín los encabezados de canon y precio de esta tabla están intercambiados; se ubicaron por su orden de magnitud."
    $enc = @("Barrio", "Tamaño (m²)", "Precio de venta por m²", "Canon mensual por m²", "Canon ÷ precio")
    for ($k = 0; $k -lt 5; $k++) { $inm.Cells.Item(4, $k + 1).Value2 = $enc[$k] }
    $lonja = @(@("Fátima", "30 a 50", 8500000, 36664), @("La Mota", "50 a 70", 6000000, 37941), @("La Palma", "50 a 70", 5500000, 32227), @("Loma de los Bernal", "30 a 50", 7576529, 46921), @("Loma de los Bernal", "50 a 70", 6950718, 40494), @("Malibú", "50 a 70", 6481481, 38095), @("La América", "30 a 50", 5943478, 36073), @("La América", "50 a 70", 5081266, 33757), @("Calasanz", "30 a 50", 7121212, 44353), @("Calasanz", "50 a 70", 6004653, 35461), @("La Floresta", "50 a 70", 6562893, 34665))
    for ($k = 0; $k -lt $lonja.Count; $k++) {
        $f = 5 + $k
        $inm.Range("A$f").Value2 = $lonja[$k][0]; $inm.Range("B$f").Value2 = $lonja[$k][1]
        $inm.Range("C$f").Value2 = [double]$lonja[$k][2]; $inm.Range("D$f").Value2 = [double]$lonja[$k][3]
        $inm.Range("E$f").Formula = "=D$f/C$f"
    }
    $u = 4 + $lonja.Count
    $sup = @(@("G4", "Supuesto del equipo"), @("G5", "Precio de compra"), @("G6", "Área m²"), @("G7", "Canon ÷ valor mensual"), @("G8", "Canon inicial"),
             @("G10", "Estadístico"), @("H10", "Canon ÷ precio"), @("G11", "n"), @("G12", "Mínimo"), @("G13", "Mediana"), @("G14", "Máximo"),
             @("G16", "Rentabilidad bruta anual con la mediana"), @("G17", "Rentabilidad bruta anual con el canon inicial"))
    foreach ($s in $sup) { $inm.Range($s[0]).Value2 = $s[1] }
    $inm.Range("H5").Value2 = 300000000; $inm.Range("H6").Value2 = 50; $inm.Range("H7").Value2 = 0.006; $inm.Range("H8").Formula = "=H5*H7"
    $inm.Range("H11").Formula = "=COUNT(E5:E$u)"; $inm.Range("H12").Formula = "=MIN(E5:E$u)"
    $inm.Range("H13").Formula = "=MEDIAN(E5:E$u)"; $inm.Range("H14").Formula = "=MAX(E5:E$u)"
    $inm.Range("H16").Formula = "=12*H13"; $inm.Range("H17").Formula = "=12*H8/H5"
    foreach ($r in @("E5:E$u", "H12:H14", "H16:H17")) { $inm.Range($r).NumberFormat = "0.00%" }
    # MARKOV: el estado no mueve la valorización en el modelo de referencia (AR(1) del IPVU, sección 7.6).
    $mk = $dest.Worksheets.Item("MARKOV")
    $mk.Range("A2").Value2 = "El enunciado no permite suponer que todas las variables son independientes. En una crisis la bolsa se vuelve más volátil, cuesta más arrendar y llegan más gastos imprevistos, todo al mismo tiempo. Para representarlo, la economía pasa por tres estados y cada estado cambia el comportamiento de varias variables a la vez."
    $mk.Range("A72").Value2 = "Paso 4. ¿Cambia la valorización de la vivienda según el estado? (vivienda nueva de Medellín, IPVNBR)"
    $mk.Range("A84").Value2 = "Lectura: en expansión la vivienda nueva en Medellín sube cerca de 12,7 % al año; en el resto de los años, cerca de 8,7 %. Pero el inmueble del modelo es usado, y con el índice de vivienda usada (IPVU, hoja INM) el estado no cambia la valorización (Welch, p = 0,275). Por eso el simulador usa el AR(1) del IPVU, igual en los tres estados, y estos valores por estado quedan como alternativa en la sensibilidad."
    $mk.Range("B90").Value2 = "Valorización anual media (ln), AR(1) del IPVU"
    $mk.Range("B91").Value2 = "Desviación del choque anual del AR(1)"
    foreach ($col in @("C", "D", "E")) { $mk.Range("$col" + "90").Value2 = 0.0755; $mk.Range("$col" + "91").Value2 = 0.0236 }
    $mk.Range("F90").Value2 = "Hoja INM: no cambia con el estado"
    $mk.Range("F91").Value2 = "Hoja INM: no cambia con el estado"
    $mk.Range("F92").Value2 = "Supuesto (Ciencuadras 2026 como referencia)"
    $mk.Range("A99").Value2 = "3. El estado vigente fija la volatilidad del ETF, la probabilidad y la duración de la vacancia, la frecuencia de las necesidades de liquidez y el reparto objetivo de E4."
    $mk.Range("A100").Value2 = "4. Así, en una crisis el ETF se vuelve más riesgoso, cuesta más arrendar y llegan más gastos imprevistos, todo al mismo tiempo. La valorización, la inflación y la TPM no dependen del estado."
    $mar.Close($false)
    $nue.Close($false)
    # "Pendiente" en el sentido de coeficiente de una regresión se confunde con "trabajo pendiente".
    $dest.Worksheets.Item("CDT").Range("A30").Value2 = "H0: el coeficiente del cambio de la TPM en 13 semanas es 0 (errores HAC)"
    $dest.Worksheets.Item("INM").Range("AM11").Value2 = "AR(1) por MCO: φ (coeficiente)"
    $dest.Worksheets.Item("INM").Range("AU12").Value2 = "H0: φ = 0 en el AR(1) (t del coeficiente)"
    $s23 = [string]$dest.Worksheets.Item("INM").Range("S23").Value2
    $dest.Worksheets.Item("INM").Range("S23").Value2 = $s23.Replace("Pendiente con errores robustos", "Coeficiente con errores robustos")
    # Menos hojas: las fuentes de Markov ya están en FUENTES.
    $dest.Worksheets.Item("MARKOV_Fuentes").Delete()
    # Sin referencias al enunciado ni al equipo en los textos de las hojas.
    $dest.Worksheets.Item("CDT").Range("A1").Value2 = "CDT: definición del margen sobre la TPM, pruebas, modelos candidatos y modelo elegido"
    $dest.Worksheets.Item("ETF").Range("A1").Value2 = "ICOLCAP: pruebas de hipótesis, modelos candidatos, modelo elegido y dependencia"
    $dest.Worksheets.Item("INM").Range("AU1").Value2 = "Valorización del inmueble: pruebas, modelos candidatos, modelo elegido y dependencia"
    $dest.Worksheets.Item("INM").Range("G4").Value2 = "Supuesto"
    foreach ($r in 23..50) { if ([string]$cta.Range("R$r").Value2 -like "Archivo del equipo*") { $cta.Range("R$r").Value2 = "Calendario de decisiones de la Junta" } }
    $mk.Range("A2").Value2 = "Las variables no se modelan como independientes: el estado de la economía mueve varias a la vez. En una crisis la bolsa se vuelve más volátil, cuesta más arrendar y llegan más gastos imprevistos."
    # Orden de lectura: resumen, parámetros, procesos en el orden del informe, estados, validación, resultados y referencia.
    $orden = @("LEEME", "MODELOS", "PARAMETROS", "MODELO", "INF", "TPM_Cuenta", "TPM_Logit", "CDT", "ETF", "INM", "PIB", "MARKOV", "MARKOV_PIB", "MARKOV_ETF", "MARKOV_Vivienda", "VALIDACION", "RESULTADOS", "FUENTES", "DICCIONARIO")
    $dest.Worksheets.Item($orden[0]).Move($dest.Worksheets.Item(1))
    for ($k = 1; $k -lt $orden.Count; $k++) { $dest.Worksheets.Item($orden[$k]).Move($falta, $dest.Worksheets.Item($orden[$k - 1])) }

    $dest.Worksheets.Item("MODELO").Visible = 0
    # Las hojas copiadas dejan un vínculo vacío al libro de Markov; ninguna fórmula lo usa.
    try { $dest.LinkSources(1) | ForEach-Object { $dest.BreakLink([string]$_, 1) } } catch { Write-Output "sin vinculos para romper" }
    $dest.Worksheets.Item("LEEME").Activate()
    if (Test-Path $Destino) { Remove-Item $Destino -Force }
    $dest.SaveAs($Destino, 51)
    $lista = @($dest.Worksheets | ForEach-Object { $_.Name }) -join ", "
    $dest.Close($false)
    Write-Output "ok $Destino"
    Write-Output $lista
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
