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
    # ---- Hojas de presentación de la inflación: leen de INF con fórmulas, no duplican datos ni mueven nada de INF.
    function Nueva-Hoja($libro, $nombre, $titulo, $desc) {
        $h = $libro.Worksheets.Add([Type]::Missing, $libro.Worksheets.Item($libro.Worksheets.Count))
        $h.Name = $nombre
        $h.Cells.Font.Name = "Arial"
        $h.Cells.Font.Size = 10
        $h.Range("A1").Value2 = $titulo
        $h.Range("A1").Font.Size = 16
        $h.Range("A1").Font.Bold = $true
        $h.Range("A1").Font.Color = 0x404040
        $h.Range("A2").Value2 = $desc
        [void]$h.Hyperlinks.Add($h.Range("A3"), "", "'LEEME'!A1", "", "Volver a LEEME")
        return $h
    }
    function Encabezado($h, $fila, $textos) {
        for ($k = 0; $k -lt $textos.Count; $k++) {
            $c = $h.Cells.Item($fila, $k + 1)
            $c.Value2 = $textos[$k]
            $c.Font.Bold = $true
            $c.Font.Color = 0xFFFFFF
            $c.Interior.Color = 0x404040
            $c.WrapText = $true
            $c.VerticalAlignment = -4108
        }
    }
    function Subtitulo($h, $fila, $texto) {
        $h.Cells.Item($fila, 1).Value2 = $texto
        $h.Cells.Item($fila, 1).Font.Bold = $true
        $h.Cells.Item($fila, 1).Font.Size = 12
        $h.Cells.Item($fila, 1).Font.Color = 0x404040
    }

    # 1. Datos mensuales, vertical
    $d = Nueva-Hoja $dest "INF_Datos" "Inflación anual mensual, enero de 2001 a septiembre de 2026" "Una fila por mes. Los valores se leen de la hoja INF (columnas Z a AE); no se copian. Fuente: DANE (IPC) y Banco de la República (meta y límites)."
    Encabezado $d 5 @("Mes", "Año", "Inflación anual (%)", "Meta (%)", "Límite inferior (%)", "Límite superior (%)", "Observación")
    $d.Range("A6:A314").Formula = "=INF!Z5"
    $d.Range("B6:B314").Formula = "=INF!AA5"
    $d.Range("C6:C314").Formula = "=INF!AB5"
    $d.Range("D6:D314").Formula = '=IF(INF!AC5="","",INF!AC5)'
    $d.Range("E6:E314").Formula = '=IF(INF!AD5="","",INF!AD5)'
    $d.Range("F6:F314").Formula = '=IF(INF!AE5="","",INF!AE5)'
    $d.Range("G6:G314").Formula = '=IF(AND(YEAR(A6)=YEAR($A$314),MONTH($A$314)<12,YEAR(A6)=YEAR($A$314)),"Año incompleto: último dato de mes "&MONTH($A$314),"")'
    $d.Range("A6:A314").NumberFormat = "yyyy-mm"
    $d.Range("C6:F314").NumberFormat = "0.00"
    $d.Range("A5:G314").AutoFilter() | Out-Null
    $d.Columns.Item(1).ColumnWidth = 12; $d.Columns.Item(2).ColumnWidth = 8
    foreach ($k in 3..6) { $d.Columns.Item($k).ColumnWidth = 16 }
    $d.Columns.Item(7).ColumnWidth = 44
    $d.Activate(); $xl.ActiveWindow.FreezePanes = $false; $d.Range("A6").Select(); $xl.ActiveWindow.FreezePanes = $true

    # 2. Inflación anual y estadística descriptiva
    $a = Nueva-Hoja $dest "INF_Anual" "Inflación anual y estadística descriptiva" "Una fila por año. 2026 es un año incompleto: su valor es el del último mes publicado, no el de diciembre. Los valores se leen de INF."
    Encabezado $a 5 @("Año", "Inflación al cierre del dato (%)", "Promedio del año (%)", "Mínimo del año (%)", "Máximo del año (%)", "Meta (%)", "Límite inferior (%)", "Límite superior (%)", "Cierre del dato")
    $a.Range("A6:A31").Formula = "=INF!A8"
    $a.Range("B6:B31").Formula = "=INF!B8"
    $a.Range("C6:C31").Formula = "=INF!C8"
    $a.Range("D6:D31").Formula = "=INF!D8"
    $a.Range("E6:E31").Formula = "=INF!E8"
    $a.Range("F6:F31").Formula = "=INF!F8"
    $a.Range("G6:G31").Formula = "=INF!G8"
    $a.Range("H6:H31").Formula = "=INF!H8"
    $a.Range("I6:I31").Formula = '=IF(AND(A6=YEAR(INF_Datos!$A$314),MONTH(INF_Datos!$A$314)<12),"Parcial: mes "&MONTH(INF_Datos!$A$314),"Diciembre")'
    $a.Range("B6:H31").NumberFormat = "0.00"
    $a.Range("A5:I31").AutoFilter() | Out-Null
    Subtitulo $a 34 "Estadística descriptiva de la inflación de diciembre, 2001 a 2025"
    Encabezado $a 35 @("Estadístico", "Valor", "Año")
    $est = @(@("Promedio (%)", "=INF!C35", ""), @("Mediana (%)", "=INF!C36", ""), @("Desviación estándar (puntos)", "=INF!C37", ""),
             @("Mínimo (%)", "=INF!C38", "=INF!D38"), @("Máximo (%)", "=INF!C39", "=INF!D39"))
    for ($k = 0; $k -lt $est.Count; $k++) {
        $a.Cells.Item(36 + $k, 1).Value2 = $est[$k][0]
        $a.Cells.Item(36 + $k, 2).Formula = $est[$k][1]
        if ($est[$k][2] -ne "") { $a.Cells.Item(36 + $k, 3).Formula = $est[$k][2] }
    }
    $a.Range("B36:B40").NumberFormat = "0.00"
    $a.Columns.Item(1).ColumnWidth = 30
    foreach ($k in 2..9) { $a.Columns.Item($k).ColumnWidth = 17 }

    # 3. Pruebas de hipótesis
    $p = Nueva-Hoja $dest "INF_Pruebas" "Pruebas estadísticas de la inflación" "Una fila por prueba. Estadísticos y valores p se leen de INF; la decisión compara el valor p con el nivel de significancia de INF!C4."
    Encabezado $p 5 @("Variable", "Prueba", "Hipótesis nula", "Estadístico", "Valor p", "Significancia", "Decisión", "Limitación")
    $filas = @(
        @("Inflación de diciembre 2010-2025", "t de una muestra", "El promedio es igual a la meta de 3 %", "=INF!C62", "=INF!C63", "Con 16 años y un solo ciclo de inflación alta, la prueba tiene poco respaldo."),
        @("Inflación de diciembre, en %", "Jarque-Bera (normal)", "La inflación en % sigue una normal", "=INF!C73", "=INF!C74", "Otras pruebas (Shapiro-Wilk, Lilliefors) no rechazan la normal al 5 %."),
        @("ln(inflación de diciembre)", "Jarque-Bera (lognormal)", "El logaritmo de la inflación sigue una normal", "=INF!D73", "=INF!D74", "No rechazar no prueba que sea lognormal; con 25 datos la prueba tiene poca potencia."),
        @("Inflación de diciembre, por rangos", "Chi-cuadrado (normal)", "Los años se reparten por rangos como una normal", "=INF!G85", "=INF!G86", "Rangos fijados a mano; pocos datos por rango."),
        @("Inflación de diciembre, por rangos", "Chi-cuadrado (lognormal)", "Los años se reparten por rangos como una lognormal", "=INF!H85", "=INF!H86", "Rangos fijados a mano; pocos datos por rango."),
        @("ln(inflación) frente a la del año anterior", "Correlación de Pearson (t)", "No hay correlación entre un año y el siguiente", "=INF!C93", "=INF!C94", "24 pares de años consecutivos; la correlación es de 0,42.")
    )
    for ($k = 0; $k -lt $filas.Count; $k++) {
        $r = 6 + $k
        $p.Cells.Item($r, 1).Value2 = $filas[$k][0]
        $p.Cells.Item($r, 2).Value2 = $filas[$k][1]
        $p.Cells.Item($r, 3).Value2 = $filas[$k][2]
        $p.Cells.Item($r, 4).Formula = $filas[$k][3]
        $p.Cells.Item($r, 5).Formula = $filas[$k][4]
        $p.Cells.Item($r, 6).Formula = "=INF!`$C`$4"
        $p.Cells.Item($r, 7).Formula = "=IF(E$r<F$r,""Se rechaza H0"",""No se rechaza H0"")"
        $p.Cells.Item($r, 8).Value2 = $filas[$k][5]
    }
    $p.Range("D6:E11").NumberFormat = "0.0000"
    $p.Range("F6:F11").NumberFormat = "0.00"
    $p.Range("A6:H11").WrapText = $true
    $p.Range("A6:H11").VerticalAlignment = -4160
    $p.Range("A5:H11").AutoFilter() | Out-Null
    $p.Cells.Item(13, 1).Value2 = "No rechazar H0 significa evidencia insuficiente en contra, no que H0 sea cierta."
    foreach ($k in 1..8) { $p.Columns.Item($k).ColumnWidth = @(30, 26, 40, 13, 13, 13, 18, 55)[$k - 1] }

    # 4. Modelo AR(1)
    $m = Nueva-Hoja $dest "INF_Modelo" "Modelo de la inflación: AR(1) en logaritmos" "Parámetros estimados en INF por mínimos cuadrados con 24 pares de años consecutivos (inflación de diciembre, 2002 a 2025). El simulador usa estos valores."
    Encabezado $m 5 @("Parámetro", "Símbolo", "Valor", "Significado")
    $par = @(@("Memoria del año anterior", "φ", "=INF!C102", "Cuánto se parece un año al anterior: 0 = nada, 1 = igual."),
             @("Constante", "a", "=INF!C103", "Junto con φ fija el nivel normal de la inflación."),
             @("Tamaño de las sorpresas", "σ", "=INF!C104", "Desviación de la sorpresa de un año a otro, en logaritmos."),
             @("Nivel normal de la inflación (%)", "", "=INF!C105", "Valor al que la inflación tiende a volver."))
    for ($k = 0; $k -lt $par.Count; $k++) {
        $m.Cells.Item(6 + $k, 1).Value2 = $par[$k][0]
        $m.Cells.Item(6 + $k, 2).Value2 = $par[$k][1]
        $m.Cells.Item(6 + $k, 3).Formula = $par[$k][2]
        $m.Cells.Item(6 + $k, 4).Value2 = $par[$k][3]
    }
    $m.Range("C6:C9").NumberFormat = "0.0000"
    $m.Cells.Item(11, 1).Value2 = "Ecuación"
    $m.Cells.Item(11, 2).Value2 = "ln π(t) = a + φ·ln π(t−1) + σ·Z, con Z normal estándar"
    $m.Cells.Item(12, 1).Value2 = "Por qué este modelo"
    $m.Cells.Item(12, 2).Value2 = "La inflación nunca es negativa, sus saltos son hacia arriba y un año depende del anterior (p = 0,043)."
    $m.Columns.Item(1).ColumnWidth = 34; $m.Columns.Item(2).ColumnWidth = 12; $m.Columns.Item(3).ColumnWidth = 12; $m.Columns.Item(4).ColumnWidth = 70

    # 5. Validación
    $v = Nueva-Hoja $dest "INF_Validacion" "Validación del modelo de inflación" "Dos pruebas distintas. La primera usa parámetros estimados con todos los años, así que es dentro de muestra. La segunda estima con 2001-2015 y evalúa en 2016-2025."
    Subtitulo $v 5 "Prueba 1. Pronóstico año a año con los parámetros de toda la muestra (dentro de muestra)"
    Encabezado $v 6 @("Año", "Observada (%)", "Pronóstico (%)", "Mínimo del 90 % (%)", "Máximo del 90 % (%)", "¿Dentro del rango?", "Error en logaritmos")
    $v.Range("A7:A30").Formula = "=INF!K6"
    $v.Range("B7:B30").Formula = "=INF!L6"
    $v.Range("C7:C30").Formula = "=INF!O6"
    $v.Range("D7:D30").Formula = "=INF!P6"
    $v.Range("E7:E30").Formula = "=INF!Q6"
    $v.Range("F7:F30").Formula = "=INF!R6"
    $v.Range("G7:G30").Formula = "=INF!S6"
    $v.Range("B7:E30").NumberFormat = "0.00"; $v.Range("G7:G30").NumberFormat = "0.000"
    $v.Cells.Item(31, 1).Value2 = "Años dentro del rango"
    $v.Cells.Item(31, 2).Formula = "=INF!C124"; $v.Cells.Item(31, 3).Formula = "=INF!D124"; $v.Cells.Item(31, 4).Formula = "=INF!E124"
    $v.Cells.Item(31, 4).NumberFormat = "0.0%"
    Subtitulo $v 34 "Prueba 2. Fuera de muestra: estimación con 2001-2015 y evaluación de 2016 a 2025"
    Encabezado $v 35 @("Medida", "Valor")
    $fm = @(@("φ estimado con 2001-2015", "=INF!C126"), @("a estimada con 2001-2015", "=INF!C127"), @("σ estimada con 2001-2015", "=INF!C128"),
            @("Años 2016-2025 dentro del rango", "=INF!C129"), @("Años evaluados", "=INF!D129"), @("Proporción dentro del rango", "=INF!E129"))
    for ($k = 0; $k -lt $fm.Count; $k++) {
        $v.Cells.Item(36 + $k, 1).Value2 = $fm[$k][0]
        $v.Cells.Item(36 + $k, 2).Formula = $fm[$k][1]
    }
    $v.Range("B36:B38").NumberFormat = "0.0000"; $v.Range("B41").NumberFormat = "0.0%"
    $v.Columns.Item(1).ColumnWidth = 34
    foreach ($k in 2..7) { $v.Columns.Item($k).ColumnWidth = 18 }

    # Orden de lectura: resumen, parámetros, procesos en el orden del informe, estados, validación, resultados y referencia.
    $orden = @("LEEME", "MODELOS", "PARAMETROS", "MODELO", "INF", "INF_Datos", "INF_Anual", "INF_Pruebas", "INF_Modelo", "INF_Validacion", "TPM_Cuenta", "TPM_Logit", "CDT", "ETF", "INM", "PIB", "MARKOV", "MARKOV_PIB", "MARKOV_ETF", "MARKOV_Vivienda", "VALIDACION", "RESULTADOS", "FUENTES", "DICCIONARIO", "AUDITORIA")
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
