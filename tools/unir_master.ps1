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
    # ---- Presentación de la inflación: dos hojas que leen de INF con fórmulas.
    $azul = 0x784E1F
    $gris = 0x808080
    function Nueva-Hoja($libro, $nombre) {
        $h = $libro.Worksheets.Add([Type]::Missing, $libro.Worksheets.Item($libro.Worksheets.Count))
        $h.Name = $nombre
        $h.Cells.Font.Name = "Arial"
        $h.Cells.Font.Size = 10
        $h.Activate()
        $xl.ActiveWindow.DisplayGridlines = $false
        return $h
    }
    function Titulo($h, $texto, $sub) {
        $h.Range("A1").Value2 = $texto
        $h.Range("A1").Font.Size = 18
        $h.Range("A1").Font.Bold = $true
        $h.Range("A1").Font.Color = $azul
        $h.Range("A2").Value2 = $sub
        $h.Range("A2").Font.Italic = $true
        $h.Range("A2").Font.Color = $gris
        $h.Rows.Item(1).RowHeight = 28
    }
    function Seccion($h, $fila, $texto) {
        $c = $h.Cells.Item($fila, 1)
        $c.Value2 = $texto
        $c.Font.Size = 12
        $c.Font.Bold = $true
        $c.Font.Color = $azul
    }
    function Encabezado($h, $fila, $primera, $ultima, $textos) {
        for ($k = 0; $k -lt $textos.Count; $k++) { $h.Cells.Item($fila, $primera + $k).Value2 = $textos[$k] }
        $r = $h.Range($h.Cells.Item($fila, $primera), $h.Cells.Item($fila, $ultima))
        $r.Font.Bold = $true
        $r.Font.Color = 0xFFFFFF
        $r.Interior.Color = $azul
        $r.VerticalAlignment = -4108
        $r.WrapText = $true
    }
    function Nota($h, $fila, $texto) {
        $h.Cells.Item($fila, 1).Value2 = $texto
        $h.Cells.Item($fila, 1).Font.Size = 9
        $h.Cells.Item($fila, 1).Font.Color = $gris
    }

    # ------------------------------------------------ Inflación
    $s = Nueva-Hoja $dest "Inflación"
    Titulo $s "Inflación en Colombia" "Evolución de los precios al consumidor, enero de 2001 – septiembre de 2026"

    $s.Range("A4").Value2 = "Promedio histórico"
    $s.Range("C4").Value2 = "Máximo"
    $s.Range("E4").Value2 = "Último dato"
    $s.Range("A5").Formula = "=INF!C35"
    $s.Range("C5").Formula = "=INF!C39"
    $s.Range("E5").Formula = "=INF!AB313"
    $s.Range("A6").Value2 = "Diciembres de 2001 a 2025"
    $s.Range("C6").Formula = '="Diciembre de "&INF!D39'
    $s.Range("E6").Value2 = "Septiembre de 2026"
    foreach ($c in @("A4", "C4", "E4")) { $s.Range($c).Font.Color = $gris }
    foreach ($c in @("A5", "C5", "E5")) {
        $x = $s.Range($c)
        $x.Font.Size = 20
        $x.Font.Bold = $true
        $x.NumberFormatLocal = '0,00" %"'
        $x.HorizontalAlignment = -4131
    }
    foreach ($c in @("A6", "C6", "E6")) { $s.Range($c).Font.Size = 9; $s.Range($c).Font.Color = $gris }

    Seccion $s 8 "Evolución de la inflación anual"
    $s.Range("A9").Value2 = "Variación del IPC en los últimos doce meses, frente a la meta del Banco de la República."
    $s.Range("A9").Font.Color = $gris

    Seccion $s 32 "Serie mensual"
    Encabezado $s 33 1 5 @("Fecha", "Inflación", "Meta", "Límite inferior", "Límite superior")
    $s.Range("A34:A342").Formula = "=INF!Z5"
    $s.Range("B34:B342").Formula = "=INF!AB5"
    $s.Range("C34:C342").Formula = '=IF(INF!AC5="","",INF!AC5)'
    $s.Range("D34:D342").Formula = '=IF(INF!AD5="","",INF!AD5)'
    $s.Range("E34:E342").Formula = '=IF(INF!AE5="","",INF!AE5)'
    $s.Range("A34:A342").NumberFormat = "mmm-yyyy"
    $s.Range("A34:A342").HorizontalAlignment = -4131
    $s.Range("B34:E342").NumberFormatLocal = '0,00" %"'
    $s.Range("B34:E342").HorizontalAlignment = -4152
    $s.Range("B33:E33").HorizontalAlignment = -4152
    $s.Range("A33:E342").AutoFilter() | Out-Null
    Nota $s 344 "Fuente: DANE. Metas y rangos: Banco de la República."
    $s.Columns.Item(1).ColumnWidth = 22
    foreach ($k in 2..5) { $s.Columns.Item($k).ColumnWidth = 22 }
    $s.Columns.Item(6).ColumnWidth = 12

    # Gráfico con la serie completa
    $izq = $s.Range("A10").Left; $arr = $s.Range("A10").Top
    $ancho = $s.Range("F10").Left + $s.Range("F10").Width - $izq
    $alto = $s.Range("A31").Top - $arr
    $co = $s.ChartObjects().Add($izq, $arr, $ancho, $alto)
    $g = $co.Chart
    $g.ChartType = 4
    while ($g.SeriesCollection().Count -gt 0) { $g.SeriesCollection(1).Delete() | Out-Null }
    $a1 = $g.SeriesCollection().NewSeries()
    $a1.Name = "Inflación anual"
    $a1.Values = $s.Range("B34:B342")
    $a1.XValues = $s.Range("A34:A342")
    $a1.Format.Line.ForeColor.RGB = $azul
    $a1.Format.Line.Weight = 1.75
    $a2 = $g.SeriesCollection().NewSeries()
    $a2.Name = "Meta"
    $a2.Values = $s.Range("C34:C342")
    $a2.XValues = $s.Range("A34:A342")
    $a2.Format.Line.ForeColor.RGB = $gris
    $a2.Format.Line.Weight = 1.25
    $a2.Format.Line.DashStyle = 4
    $g.HasTitle = $false
    $g.HasLegend = $true
    $g.Legend.Position = -4107
    $ejeX = $g.Axes(1)
    $ejeX.CategoryType = 3
    $ejeX.BaseUnit = 0
    $ejeX.MajorUnitScale = 2
    $ejeX.MajorUnit = 5
    $ejeX.TickLabels.NumberFormat = "yyyy"
    $ejeY = $g.Axes(2)
    $ejeY.TickLabels.NumberFormat = '0" %"'
    $ejeY.MajorGridlines.Format.Line.ForeColor.RGB = 0xE0E0E0
    $ejeY.HasTitle = $true
    $ejeY.AxisTitle.Text = "Variación anual (%)"
    $co.Border.LineStyle = -4142
    $s.Range("A1").Select()

    # ------------------------------------------------ Inflación, modelo (composición en dos paneles)
    function Bloque($h, $fila, $col1, $col2, $texto) {
        $c = $h.Cells.Item($fila, $col1)
        $c.Value2 = $texto
        $c.Font.Size = 12
        $c.Font.Bold = $true
        $c.Font.Color = $azul
        $r = $h.Range($h.Cells.Item($fila, $col1), $h.Cells.Item($fila, $col2))
        $r.Borders.Item(9).LineStyle = 1
        $r.Borders.Item(9).Color = $azul
        $h.Rows.Item($fila).RowHeight = 20
    }
    function Cabecera($h, $fila, $col1, $textos, $derecha) {
        for ($k = 0; $k -lt $textos.Count; $k++) {
            $c = $h.Cells.Item($fila, $col1 + $k)
            $c.Value2 = $textos[$k]
            $c.Font.Bold = $true
            $c.Font.Color = 0xFFFFFF
            $c.Interior.Color = $azul
            if ($derecha -contains $k) { $c.HorizontalAlignment = -4152 }
        }
    }
    $m = Nueva-Hoja $dest "Inflación, modelo"
    Titulo $m "Inflación: modelo probabilístico" "AR(1) en logaritmos · inflación de diciembre, 2001–2025"
    foreach ($k in 1..9) { $m.Columns.Item($k).ColumnWidth = @(28, 48, 12, 9, 16, 3, 24, 14, 28)[$k - 1] }

    # Cifras principales, juntas
    $m.Range("G4").Value2 = "Nivel de largo plazo"
    $m.Range("H4").Value2 = "Memoria (φ)"
    $m.Range("I4").Value2 = "Volatilidad (σ)"
    $m.Range("G5").Formula = "=INF!C105"
    $m.Range("H5").Formula = "=INF!C102"
    $m.Range("I5").Formula = "=INF!C104"
    foreach ($c in @("G4", "H4", "I4")) { $m.Range($c).Font.Color = $gris }
    foreach ($c in @("G5", "H5", "I5")) {
        $x = $m.Range($c); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    $m.Range("G5").NumberFormatLocal = '0,00" %"'
    $m.Range("H5").NumberFormatLocal = "0,000"
    $m.Range("I5").NumberFormatLocal = "0,000"
    $m.Rows.Item(5).RowHeight = 28

    # Panel izquierdo: evidencia
    Bloque $m 8 1 5 "Evidencia estadística"
    Cabecera $m 9 1 @("Prueba", "Qué se contrasta", "Estadístico", "p", "Decisión") @(2, 3)
    $pr = @(
        @("t de una muestra", "Promedio de diciembre 2010–2025 frente a la meta de 3 %", "=INF!C62", "=INF!C63"),
        @("Jarque-Bera", "Normalidad de la inflación en %", "=INF!C73", "=INF!C74"),
        @("Jarque-Bera", "Normalidad del logaritmo de la inflación", "=INF!D73", "=INF!D74"),
        @("Chi-cuadrado", "Ajuste de la normal, por rangos", "=INF!G85", "=INF!G86"),
        @("Chi-cuadrado", "Ajuste de la lognormal, por rangos", "=INF!H85", "=INF!H86")
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 10 + $k
        $m.Cells.Item($r, 1).Value2 = $pr[$k][0]
        $m.Cells.Item($r, 2).Value2 = $pr[$k][1]
        $m.Cells.Item($r, 3).Formula = $pr[$k][2]
        $m.Cells.Item($r, 4).Formula = $pr[$k][3]
        $m.Cells.Item($r, 5).Formula = "=IF(D$r<INF!`$C`$4,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $m.Range("C10:D14").NumberFormatLocal = "0,000"
    $m.Range("E9:E14").IndentLevel = 1
    $m.Range("A15").Value2 = "Significancia de 5 %. No rechazar H₀ no demuestra que sea cierta."
    $m.Range("A15").Font.Size = 9; $m.Range("A15").Font.Color = $gris

    Bloque $m 17 1 5 "Dependencia entre años consecutivos"
    Cabecera $m 18 1 @("Medida", "Qué mide", "Valor", "", "") @(2)
    $dep = @(
        @("Correlación de Pearson (r)", "Entre ln πₜ y ln πₜ₋₁", "=INF!C91", "0,000"),
        @("Estadístico t", "Contraste de H₀: r = 0", "=INF!C93", "0,000"),
        @("Valor p", "Del contraste anterior", "=INF!C94", "0,000"),
        @("Observaciones", "Pares de años consecutivos", "=INF!C92", "0")
    )
    for ($k = 0; $k -lt $dep.Count; $k++) {
        $r = 19 + $k
        $m.Cells.Item($r, 1).Value2 = $dep[$k][0]
        $m.Cells.Item($r, 2).Value2 = $dep[$k][1]
        $m.Cells.Item($r, 3).Formula = $dep[$k][2]
        $m.Cells.Item($r, 3).NumberFormatLocal = $dep[$k][3]
    }

    # Panel derecho: especificación y validación
    Bloque $m 8 7 9 "Especificación"
    $m.Range("G9").Value2 = "ln πₜ = a + φ · ln πₜ₋₁ + σ · Zₜ"
    $m.Range("G9").Font.Italic = $true
    Cabecera $m 10 7 @("Parámetro", "Estimación", "Significado") @(1)
    $m.Range("I10").IndentLevel = 1
    $par = @(
        @("Constante (a)", "=INF!C103", "0,000", "Intercepto en logaritmos"),
        @("Memoria (φ)", "=INF!C102", "0,000", "Peso del año anterior"),
        @("Volatilidad (σ)", "=INF!C104", "0,000", "Escala de la perturbación Z"),
        @("Largo plazo", "=INF!C105", "0,00"" %""", "Nivel al que tiende el proceso")
    )
    for ($k = 0; $k -lt $par.Count; $k++) {
        $m.Cells.Item(11 + $k, 7).Value2 = $par[$k][0]
        $m.Cells.Item(11 + $k, 8).Formula = $par[$k][1]
        $m.Cells.Item(11 + $k, 8).NumberFormatLocal = $par[$k][2]
        $m.Cells.Item(11 + $k, 9).Value2 = $par[$k][3]
    }
    $m.Range("I11:I14").IndentLevel = 1

    Bloque $m 17 7 9 "Validación"
    Cabecera $m 18 7 @("Evaluación", "Periodo", "Dentro del 90 %") @(2)
    $m.Range("G19").Value2 = "Año a año"
    $m.Range("H19").Value2 = "2002–2025"
    $m.Range("I19").Formula = '=INF!C124&" de "&INF!D124&" · "&TEXT(INF!E124*100,"0,0")&" %"'
    $m.Range("G20").Value2 = "Fuera de muestra"
    $m.Range("H20").Value2 = "2016–2025"
    $m.Range("I20").Formula = '=INF!C129&" de "&INF!D129&" · "&TEXT(INF!E129*100,"0,0")&" %"'
    $m.Range("I19:I20").HorizontalAlignment = -4152
    $m.Range("G21").Value2 = "Proporción de observaciones dentro del intervalo predictivo nominal de 90 %."
    $m.Range("G22").Value2 = "Fuera de muestra: parámetros estimados con 2001–2015."
    $m.Range("G23").Value2 = "La cobertura observada no garantiza la precisión futura."
    $m.Range("G21:G23").Font.Size = 9; $m.Range("G21:G23").Font.Color = $gris

    # Lectura
    Bloque $m 25 1 9 "Lectura"
    $lec = $m.Range("A26:I28")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $m.Range("A26").Value2 = "Los logaritmos de la inflación son compatibles con una distribución normal, y modelar sobre ellos mantiene positiva la inflación simulada. La dependencia significativa entre años consecutivos justifica incluir memoria en el proceso. Con una muestra anual tan corta, los parámetros se estiman con incertidumbre, y la cobertura observada no garantiza el desempeño en un régimen distinto al de estos años."
    $m.Range("A1").Select()

    # ------------------------------------------------ ETF, datos
    $e = Nueva-Hoja $dest "ETF, datos"
    Titulo $e "ETF de acciones colombianas" "Rendimiento total mensual del NAV del ICOLCAP, agosto de 2011 – septiembre de 2026"
    $e.Range("A4").Value2 = "Rendimiento medio mensual"
    $e.Range("B4").Value2 = "Desviación mensual"
    $e.Range("C4").Value2 = "Caída máxima acumulada"
    $e.Range("A5").Formula = "=ETF!B44"
    $e.Range("B5").Formula = "=ETF!C44"
    $e.Range("C5").Formula = "=ETF!B64"
    $e.Range("A6").Value2 = "Del logaritmo del rendimiento"
    $e.Range("B6").Value2 = "Del logaritmo del rendimiento"
    $e.Range("C6").Value2 = "Del NAV, 182 meses"
    foreach ($c in @("A4", "B4", "C4")) { $e.Range($c).Font.Color = $gris }
    foreach ($c in @("A5", "B5", "C5")) {
        $x = $e.Range($c); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
        $x.NumberFormatLocal = "0,00 %"
    }
    foreach ($c in @("A6", "B6", "C6")) { $e.Range($c).Font.Size = 9; $e.Range($c).Font.Color = $gris }
    $e.Rows.Item(5).RowHeight = 28
    foreach ($k in 1..5) { $e.Columns.Item($k).ColumnWidth = @(26, 26, 26, 22, 4)[$k - 1] }

    Seccion $e 8 "Valor del fondo"
    $e.Range("A9").Value2 = "NAV de fin de mes, en pesos por participación."
    $e.Range("A9").Font.Color = $gris
    Seccion $e 29 "Rendimiento mensual"
    $e.Range("A30").Value2 = "Rendimiento total publicado de cada mes."
    $e.Range("A30").Font.Color = $gris

    Seccion $e 51 "Serie mensual"
    Encabezado $e 52 1 4 @("Mes", "Rendimiento total", "NAV de fin de mes", "g = ln(1 + R)")
    $e.Range("A53:A234").Formula = "=ETF!N5"
    $e.Range("B53:B234").Formula = "=ETF!O5"
    $e.Range("C53:C234").Formula = "=ETF!P5"
    $e.Range("D53:D234").Formula = "=ETF!Q5"
    $e.Range("A53:A234").NumberFormat = "mmm-yyyy"
    $e.Range("A53:A234").HorizontalAlignment = -4131
    $e.Range("B53:B234").NumberFormatLocal = "0,00 %"
    $e.Range("C53:C234").NumberFormatLocal = "#.##0,00"
    $e.Range("D53:D234").NumberFormatLocal = "0,0000"
    $e.Range("B53:D234").HorizontalAlignment = -4152
    $e.Range("B52:D52").HorizontalAlignment = -4152
    $e.Range("A52:D234").AutoFilter() | Out-Null
    Nota $e 236 "Fuente: BlackRock, archivo oficial del NAV del fondo. g = ln(1 + R) se calcula con fórmula."

    $izq = $e.Range("A10").Left; $ancho = $e.Range("E10").Left - $izq
    foreach ($par in @(@(10, 28, "NAV", "C53:C234", 4, "#.##0"), @(31, 49, "Rendimiento mensual", "B53:B234", 51, "0%"))) {
        $arr = $e.Range("A" + $par[0]).Top; $alto = $e.Range("A" + $par[1]).Top - $arr
        $co = $e.ChartObjects().Add($izq, $arr, $ancho, $alto)
        $g = $co.Chart
        $g.ChartType = $par[4]
        while ($g.SeriesCollection().Count -gt 0) { $g.SeriesCollection(1).Delete() | Out-Null }
        $a1 = $g.SeriesCollection().NewSeries()
        $a1.Name = $par[2]
        $a1.Values = $e.Range($par[3])
        $a1.XValues = $e.Range("A53:A234")
        if ($par[4] -eq 4) { $a1.Format.Line.ForeColor.RGB = $azul; $a1.Format.Line.Weight = 1.75 }
        else { $a1.Format.Fill.ForeColor.RGB = $azul; $g.ChartGroups(1).GapWidth = 40 }
        $g.HasTitle = $false
        $g.HasLegend = $false
        $ejeX = $g.Axes(1)
        $ejeX.CategoryType = 3
        $ejeX.BaseUnit = 1
        $ejeX.MajorUnitScale = 2
        $ejeX.MajorUnit = 2
        $ejeX.TickLabels.NumberFormat = "yyyy"
        $ejeX.TickLabelPosition = -4134
        $ejeY = $g.Axes(2)
        $ejeY.TickLabels.NumberFormat = $par[5]
        $ejeY.MajorGridlines.Format.Line.ForeColor.RGB = 0xE0E0E0
        $g.HasTitle = $true
        $g.HasTitle = $false
        $co.Border.LineStyle = -4142
    }
    $e.Range("A1").Select()

    # ------------------------------------------------ ETF, modelo
    $f = Nueva-Hoja $dest "ETF, modelo"
    Titulo $f "ETF: modelo probabilístico" "Distribución de g = ln(1 + R) con 182 rendimientos mensuales"
    foreach ($k in 1..12) { $f.Columns.Item($k).ColumnWidth = @(48, 12, 10, 16, 14, 12, 3, 28, 11, 13, 13, 17)[$k - 1] }
    $f.Range("H4").Value2 = "Rendimiento medio mensual"
    $f.Range("J4").Value2 = "Desviación"
    $f.Range("K4").Value2 = "Asimetría"
    $f.Range("H5").Formula = "=ETF!B44"
    $f.Range("J5").Formula = "=ETF!C44"
    $f.Range("K5").Formula = "=ETF!D44"
    foreach ($c in @("H4", "J4", "K4")) { $f.Range($c).Font.Color = $gris }
    foreach ($c in @("H5", "J5", "K5")) {
        $x = $f.Range($c); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    $f.Range("H5").NumberFormatLocal = "0,00 %"
    $f.Range("J5").NumberFormatLocal = "0,00 %"
    $f.Range("K5").NumberFormatLocal = "0,00"
    $f.Rows.Item(5).RowHeight = 28
    $pfmt = '[<0,001]"<0,001";0,000'

    # Evidencia
    Bloque $f 8 1 4 "Evidencia estadística"
    Cabecera $f 9 1 @("Prueba y contraste", "Estadístico", "p", "Decisión") @(1, 2)
    $pr = @(
        "t de una muestra: rendimiento medio mensual igual a 0",
        "Ljung-Box: sin autocorrelación a 1 mes",
        "ARCH: sin autocorrelación de g² a 1 mes",
        "Asimetría: g simétrica",
        "Curtosis: g con curtosis normal",
        "Jarque-Bera: g normal",
        "Shapiro-Wilk: g normal",
        "ADF: raíz unitaria en el NAV (logaritmo)",
        "ADF: raíz unitaria en g",
        "Kruskal-Wallis: media 2011–2023 frente a 2024–2026",
        "Kruskal-Wallis: g igual entre estados del PIB",
        "Correlación de g con el crecimiento anual del PIB",
        "Razón de verosimilitud: t asimétrica frente a t simétrica"
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 10 + $k; $o = 6 + $k
        $f.Cells.Item($r, 1).Value2 = $pr[$k]
        $f.Cells.Item($r, 2).Formula = "=IF(ETF!B$o="""","""",ETF!B$o)"
        $f.Cells.Item($r, 3).Formula = "=ETF!C$o"
        $f.Cells.Item($r, 4).Formula = "=IF(C$r<0.05,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $f.Range("B10:B22").NumberFormatLocal = "0,000"
    $f.Range("C10:C22").NumberFormatLocal = $pfmt
    $f.Range("D9:D22").IndentLevel = 1
    $f.Range("A23").Value2 = "Significancia de 5 %. No rechazar H₀ no demuestra que sea cierta."
    $f.Range("A23").Font.Size = 9; $f.Range("A23").Font.Color = $gris

    # Candidatos
    Bloque $f 25 1 6 "Modelos candidatos"
    Cabecera $f 26 1 @("Modelo", "Parámetros", "log L", "AIC", "BIC", "KS p") @(1, 2, 3, 4, 5)
    for ($k = 0; $k -lt 8; $k++) {
        $r = 27 + $k; $o = 32 + $k
        foreach ($j in 1..6) { $f.Cells.Item($r, $j).Formula = "=ETF!" + [char](64 + $j) + $o }
    }
    $f.Range("B27:B34").NumberFormatLocal = "0"
    $f.Range("C27:E34").NumberFormatLocal = "0,0"
    $f.Range("F27:F34").NumberFormatLocal = "0,000"
    $f.Range("A35").Value2 = "Máxima verosimilitud sobre los 182 meses. AIC y BIC no separan con claridad los modelos de colas pesadas."
    $f.Range("A35").Font.Size = 9; $f.Range("A35").Font.Color = $gris

    # Reproducción
    Bloque $f 37 1 6 "Cuánto reproduce cada modelo"
    Cabecera $f 38 1 @("Modelo", "Media", "Desviación", "Asimetría", "P(caída ≥ 15 %)", "Meses de 182") @(1, 2, 3, 4, 5)
    for ($k = 0; $k -lt 5; $k++) {
        $r = 39 + $k; $o = 44 + $k
        foreach ($j in 1..6) { $f.Cells.Item($r, $j).Formula = "=ETF!" + [char](64 + $j) + $o }
    }
    $f.Range("B39:C43").NumberFormatLocal = "0,00 %"
    $f.Range("D39:D43").NumberFormatLocal = "0,00"
    $f.Range("E39:E43").NumberFormatLocal = "0,00 %"
    $f.Range("F39:F43").NumberFormatLocal = "0,00"

    # Dependencia
    Bloque $f 45 1 6 "Dependencia con otras variables"
    Cabecera $f 46 1 @("Variable (correlación con g)", "n", "Pearson", "p", "Spearman", "p") @(1, 2, 3, 4, 5)
    for ($k = 0; $k -lt 11; $k++) {
        $r = 47 + $k; $o = 69 + $k
        foreach ($j in 1..6) { $f.Cells.Item($r, $j).Formula = "=IF(ETF!" + [char](64 + $j) + "$o="""","""",ETF!" + [char](64 + $j) + "$o)" }
    }
    $f.Range("B47:B57").NumberFormatLocal = "0"
    $f.Range("C47:C57").NumberFormatLocal = "0,000"
    $f.Range("E47:E57").NumberFormatLocal = "0,000"
    $f.Range("D47:D57").NumberFormatLocal = $pfmt
    $f.Range("F47:F57").NumberFormatLocal = $pfmt
    $f.Range("A58").Value2 = "Δlog: variación logarítmica mensual."
    $f.Range("A58").Font.Size = 9; $f.Range("A58").Font.Color = $gris

    # Especificación
    Bloque $f 8 8 12 "Especificación"
    $f.Range("H9").Value2 = "g = posición + escala · T,   T ~ t asimétrica (a, b)"
    $f.Range("H9").Font.Italic = $true
    Cabecera $f 10 8 @("Parámetro", "Estimación", "Significado", "", "") @(1)
    $f.Range("J10").IndentLevel = 1
    $par = @(
        @("a", "=ETF!B51", "0,000", "Forma de la cola izquierda"),
        @("b", "=ETF!C51", "0,000", "Forma de la cola derecha"),
        @("Posición", "=ETF!D51", "0,000", "Centro de g"),
        @("Escala", "=ETF!E51", "0,000", "Dispersión de g"),
        @("ν (alternativa)", "=ETF!B55", "0", "Grados de libertad de la t simétrica")
    )
    for ($k = 0; $k -lt $par.Count; $k++) {
        $f.Cells.Item(11 + $k, 8).Value2 = $par[$k][0]
        $f.Cells.Item(11 + $k, 9).Formula = $par[$k][1]
        $f.Cells.Item(11 + $k, 9).NumberFormatLocal = $par[$k][2]
        $f.Cells.Item(11 + $k, 10).Value2 = $par[$k][3]
    }
    $f.Range("J11:J15").IndentLevel = 1
    $f.Range("H16").Value2 = "El valor se acumula como V(t+1) = V(t) · exp(g)."
    $f.Range("H16").Font.Size = 9; $f.Range("H16").Font.Color = $gris

    # Validación
    Bloque $f 18 8 12 "Validación"
    Cabecera $f 19 8 @("Medida", "Observado", "t (ν = 5)", "Jones-Faddy", "") @(1, 2, 3)
    $f.Range("H20").Value2 = "Caída máxima acumulada"
    $f.Range("I20").Formula = "=ETF!B64"
    $f.Range("J20").Formula = "=ETF!C64"
    $f.Range("K20").Formula = "=ETF!E64"
    $f.Range("H21").Value2 = "Rendimiento anual geométrico"
    $f.Range("I21").Formula = "=ETF!B65"
    $f.Range("J21").Formula = "=ETF!C65"
    $f.Range("K21").Formula = "=ETF!E65"
    $f.Range("I20:K21").NumberFormatLocal = "0,0 %"
    $f.Range("H22").Value2 = "Mediana de 3.000 trayectorias simuladas de 182 meses."
    $f.Range("H23").Value2 = "El ajuste a la muestra no garantiza el comportamiento futuro."
    $f.Range("H22:H23").Font.Size = 9; $f.Range("H22:H23").Font.Color = $gris

    # Estados
    Bloque $f 25 8 12 "Estados de la economía"
    Cabecera $f 26 8 @("Estado (PIB)", "Meses", "Media de g", "Desviación", "Meses negativos") @(1, 2, 3, 4)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 27 + $k; $o = 25 + $k
        foreach ($j in 1..5) { $f.Cells.Item($r, 7 + $j).Formula = "=ETF!" + [char](64 + $j) + $o }
    }
    $f.Range("H27:H29").HorizontalAlignment = -4131
    $f.Range("I27:I29").NumberFormatLocal = "0"
    $f.Range("J27:K29").NumberFormatLocal = "0,00 %"
    $f.Range("L27:L29").NumberFormatLocal = "0 %"
    $f.Range("H30").Value2 = "Tres meses quedan fuera de los límites de Tukey: marzo de 2020, junio de 2022 y enero de 2026."
    $f.Range("H30").Font.Size = 9; $f.Range("H30").Font.Color = $gris

    # Lectura
    Bloque $f 33 8 12 "Lectura"
    $lec = $f.Range("H34:L46")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $f.Range("H34").Value2 = "El rendimiento mensual no muestra autocorrelación ni agrupamiento de volatilidad, pero sí asimetría y colas pesadas, así que la normal no lo describe bien. La comparación por AIC y BIC no separa con claridad los modelos de colas pesadas; se eligió la t asimétrica porque reproduce la asimetría y la frecuencia de las caídas profundas, y la t simétrica de cinco grados queda como alternativa. La media mensual no es significativa y depende de los años 2024 a 2026. Además, el fondo se mueve con el dólar y el petróleo, que el simulador no modela, y la volatilidad de las crisis descansa en muy pocos meses."
    $f.Rows.Item(18).RowHeight = 17
    $f.Rows.Item(33).RowHeight = 17
    $f.Range("A1").Select()

    # ------------------------------------------------ CDT, datos
    $d = Nueva-Hoja $dest "CDT, datos"
    Titulo $d "CDT a 360 días y tasa de política monetaria" "Serie semanal, octubre de 2007 – octubre de 2026"
    $d.Range("A4").Value2 = "CDT a 360 días"
    $d.Range("B4").Value2 = "TPM"
    $d.Range("C4").Value2 = "Margen medio"
    $d.Range("A5").Formula = "=CDT!T999"
    $d.Range("B5").Formula = "=CDT!U999"
    $d.Range("C5").Formula = "=CDT!B6"
    $d.Range("A6").Formula = "=CDT!S999"
    $d.Range("A6").NumberFormat = '"Último dato: "dd/mm/yyyy'
    $d.Range("A6").HorizontalAlignment = -4131
    $d.Range("B6").Value2 = "Misma fecha"
    $d.Range("C6").Value2 = "CDT menos TPM, 2008-2026"
    foreach ($c in @("A4", "B4", "C4")) { $d.Range($c).Font.Color = $gris }
    foreach ($c in @("A5", "B5")) { $d.Range($c).NumberFormatLocal = '0,00" %"' }
    $d.Range("C5").NumberFormatLocal = '0,00" pp"'
    foreach ($c in @("A5", "B5", "C5")) {
        $x = $d.Range($c); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    foreach ($c in @("A6", "B6", "C6")) { $d.Range($c).Font.Size = 9; $d.Range($c).Font.Color = $gris }
    $d.Rows.Item(5).RowHeight = 28
    foreach ($k in 1..5) { $d.Columns.Item($k).ColumnWidth = @(26, 26, 26, 22, 4)[$k - 1] }

    Seccion $d 8 "Tasas"
    $d.Range("A9").Value2 = "CDT a 360 días y TPM, en % efectivo anual."
    $d.Range("A9").Font.Color = $gris
    Seccion $d 29 "Margen del CDT sobre la TPM"
    $d.Range("A30").Value2 = "Diferencia entre el CDT y la TPM vigente, en puntos porcentuales."
    $d.Range("A30").Font.Color = $gris

    Seccion $d 51 "Serie semanal"
    Encabezado $d 52 1 4 @("Fecha (viernes)", "CDT a 360 días", "TPM", "Margen (pp)")
    $d.Range("A53:A1047").Formula = "=CDT!S5"
    $d.Range("B53:B1047").Formula = "=CDT!T5"
    $d.Range("C53:C1047").Formula = "=CDT!U5"
    $d.Range("D53:D1047").Formula = "=CDT!V5"
    $d.Range("A53:A1047").NumberFormat = "dd-mmm-yyyy"
    $d.Range("A53:A1047").HorizontalAlignment = -4131
    $d.Range("B53:C1047").NumberFormatLocal = '0,00" %"'
    $d.Range("D53:D1047").NumberFormatLocal = "0,00"
    $d.Range("B53:D1047").HorizontalAlignment = -4152
    $d.Range("B52:D52").HorizontalAlignment = -4152
    $d.Range("A52:D1047").AutoFilter() | Out-Null
    Nota $d 1049 "Fuente: Banco de la República (SUAMECA), CDT a 360 días semanal y TPM diaria, valor del viernes. El margen se calcula con fórmula."

    $izq = $d.Range("A10").Left; $ancho = $d.Range("E10").Left - $izq
    foreach ($par in @(@(10, 28, 1), @(31, 49, 2))) {
        $arr = $d.Range("A" + $par[0]).Top; $alto = $d.Range("A" + $par[1]).Top - $arr
        $co = $d.ChartObjects().Add($izq, $arr, $ancho, $alto)
        $g = $co.Chart
        $g.ChartType = 4
        while ($g.SeriesCollection().Count -gt 0) { $g.SeriesCollection(1).Delete() | Out-Null }
        if ($par[2] -eq 1) { $series = @(@("CDT a 360 días", "B53:B1047", $azul), @("TPM", "C53:C1047", $gris)) }
        else { $series = @(,@("Margen (pp)", "D53:D1047", $azul)) }
        foreach ($s1 in $series) {
            $a1 = $g.SeriesCollection().NewSeries()
            $a1.Name = $s1[0]
            $a1.Values = $d.Range($s1[1])
            $a1.XValues = $d.Range("A53:A1047")
            $a1.Format.Line.ForeColor.RGB = $s1[2]
            $a1.Format.Line.Weight = 1.5
        }
        $g.HasLegend = ($par[2] -eq 1)
        if ($par[2] -eq 1) { $g.Legend.Position = -4107 }
        $ejeX = $g.Axes(1)
        $ejeX.CategoryType = 3
        $ejeX.BaseUnit = 0
        $ejeX.MajorUnitScale = 2
        $ejeX.MajorUnit = 2
        $ejeX.TickLabels.NumberFormat = "yyyy"
        $ejeX.TickLabelPosition = -4134
        $ejeY = $g.Axes(2)
        $ejeY.TickLabels.NumberFormat = '0'
        $ejeY.MajorGridlines.Format.Line.ForeColor.RGB = 0xE0E0E0
        $g.HasTitle = $true
        $g.HasTitle = $false
        $co.Border.LineStyle = -4142
    }
    $d.Range("A1").Select()

    # ------------------------------------------------ CDT, modelo
    $c = Nueva-Hoja $dest "CDT, modelo"
    Titulo $c "CDT: modelo probabilístico" "Margen sobre la TPM = CDT a 360 días menos TPM vigente, en puntos porcentuales, 2008-2026"
    foreach ($k in 1..12) { $c.Columns.Item($k).ColumnWidth = @(52, 13, 12, 16, 22, 16, 3, 28, 11, 13, 13, 17)[$k - 1] }
    $c.Range("H4").Value2 = "Margen medio"
    $c.Range("J4").Value2 = "Desviación"
    $c.Range("K4").Value2 = "R² del modelo"
    $c.Range("H5").Formula = "=CDT!B6"
    $c.Range("J5").Formula = "=CDT!B7"
    $c.Range("K5").Formula = "=CDT!B51"
    foreach ($z in @("H4", "J4", "K4")) { $c.Range($z).Font.Color = $gris }
    foreach ($z in @("H5", "J5", "K5")) {
        $x = $c.Range($z); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    $c.Range("H5").NumberFormatLocal = '0,00" pp"'
    $c.Range("J5").NumberFormatLocal = '0,00" pp"'
    $c.Range("K5").NumberFormatLocal = "0 %"
    $c.Rows.Item(5).RowHeight = 28
    $pfmt = '[<0,001]"<0,001";0,000'

    # Evidencia
    Bloque $c 8 1 4 "Evidencia estadística"
    Cabecera $c 9 1 @("Prueba y contraste", "Estadístico", "p", "Decisión") @(1, 2)
    $pr = @(
        "t de una muestra (n): margen medio igual a 0",
        "t de una muestra (n efectivo): margen medio igual a 0",
        "ANOVA: margen igual en las tres fases",
        "ANOVA con bloques de 13 semanas no solapados",
        "Kruskal-Wallis: margen igual entre fases",
        "Coeficiente de ΔTPM a 13 semanas igual a 0 (Newey-West)",
        "ADF: raíz unitaria en el margen",
        "Ljung-Box (13 rezagos): sin autocorrelación",
        "Jarque-Bera: normalidad de los residuos de M5",
        "Quiebre: margen 1998-2007 frente a 2008-2026 (Welch)"
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 10 + $k; $o = 25 + $k
        $c.Cells.Item($r, 1).Value2 = $pr[$k]
        $c.Cells.Item($r, 2).Formula = "=IF(CDT!B$o="""","""",CDT!B$o)"
        $c.Cells.Item($r, 3).Formula = "=CDT!C$o"
        $c.Cells.Item($r, 4).Formula = "=IF(C$r<0.05,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $c.Range("B10:B19").NumberFormatLocal = "0,000"
    $c.Range("C10:C19").NumberFormatLocal = $pfmt
    $c.Range("D9:D19").IndentLevel = 1
    $c.Range("A20").Value2 = "Significancia de 5 %. La prueba con n ignora la autocorrelación; la de n efectivo la corrige."
    $c.Range("A21").Value2 = "Hay 90 semanas fuera de los límites de Tukey, casi todas en 2008, 2022, 2023 y 2024. Se conservan."
    $c.Range("A20:A21").Font.Size = 9; $c.Range("A20:A21").Font.Color = $gris

    # Fases
    Bloque $c 23 1 5 "Margen por fase del ciclo"
    Cabecera $c 24 1 @("Fase (cambio de la TPM en 13 semanas)", "Semanas", "Media (pp)", "Desviación (pp)", "Proporción") @(1, 2, 3, 4)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 25 + $k; $o = 19 + $k
        foreach ($j in 1..5) { $c.Cells.Item($r, $j).Formula = "=CDT!" + [char](64 + $j) + $o }
    }
    $c.Range("A25:A27").HorizontalAlignment = -4131
    $c.Range("B25:B27").NumberFormatLocal = "0"
    $c.Range("C25:D27").NumberFormatLocal = "0,00"
    $c.Range("E25:E27").NumberFormatLocal = "0 %"

    # Candidatos
    Bloque $c 29 1 5 "Modelos candidatos para el margen"
    Cabecera $c 30 1 @("Modelo", "RMSE dentro", "RMSE fuera", "AIC", "Parámetros") @(1, 2, 3)
    $c.Range("E30").IndentLevel = 1
    for ($k = 0; $k -lt 5; $k++) {
        $r = 31 + $k; $o = 41 + $k
        $c.Cells.Item($r, 1).Formula = "=CDT!A$o"
        $c.Cells.Item($r, 2).Formula = "=CDT!E$o"
        $c.Cells.Item($r, 3).Formula = "=CDT!F$o"
        $c.Cells.Item($r, 4).Formula = "=CDT!G$o"
        $c.Cells.Item($r, 5).Formula = "=TEXTJOIN(""; "",TRUE,ROUND(CDT!B$o,3),IF(CDT!C$o="""","""",ROUND(CDT!C$o,3)),IF(CDT!D$o="""","""",ROUND(CDT!D$o,3)))"
    }
    $c.Range("B31:C35").NumberFormatLocal = "0,00"
    $c.Range("D31:D35").NumberFormatLocal = "0,0"
    $c.Range("E31:E35").IndentLevel = 1
    $c.Range("A36").Value2 = "RMSE en puntos porcentuales; fuera de la muestra se estima hasta 2019 y se prueba desde 2020."
    $c.Range("A36").Font.Size = 9; $c.Range("A36").Font.Color = $gris

    # Distribución del error
    Bloque $c 38 1 6 "Distribución del error de M5"
    Cabecera $c 39 1 @("Distribución", "Parámetros", "AIC", "BIC", "KS p", "Peso de Akaike") @(1, 2, 3, 4, 5)
    for ($k = 0; $k -lt 5; $k++) {
        $r = 40 + $k; $o = 57 + $k
        foreach ($j in 1..6) { $c.Cells.Item($r, $j).Formula = "=CDT!" + [char](64 + $j) + $o }
    }
    $c.Range("B40:B44").NumberFormatLocal = "0"
    $c.Range("C40:D44").NumberFormatLocal = "0,0"
    $c.Range("E40:E44").NumberFormatLocal = "0,000"
    $c.Range("F40:F44").NumberFormatLocal = $pfmt
    $c.Range("A45").Value2 = "Ninguna distribución pasa la prueba KS: hay dos regímenes, tasas estables y episodios de ajuste."
    $c.Range("A46").Value2 = "Por eso el simulador remuestrea los residuos observados y usa la t como aproximación."
    $c.Range("A45:A46").Font.Size = 9; $c.Range("A45:A46").Font.Color = $gris

    # Dependencia
    Bloque $c 48 1 6 "Dependencia con otras variables"
    Cabecera $c 49 1 @("Variable", "Resultado", "p", "Lectura", "", "") @(1, 2)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 50 + $k; $o = 71 + $k
        $c.Cells.Item($r, 1).Formula = "=CDT!A$o"
        $c.Cells.Item($r, 2).Formula = "=CDT!B$o"
        $c.Cells.Item($r, 3).Formula = "=IF(CDT!C$o="""","""",CDT!C$o)"
        $c.Cells.Item($r, 4).Formula = "=CDT!D$o"
    }
    $c.Range("B50:B52").NumberFormatLocal = "0,000"
    $c.Range("C50:C52").NumberFormatLocal = $pfmt
    $c.Range("D49:D52").IndentLevel = 1

    Bloque $c 54 1 5 "Margen por estado de la economía"
    Cabecera $c 55 1 @("Estado (PIB)", "Semanas", "Media (pp)", "Desviación (pp)", "") @(1, 2, 3)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 56 + $k; $o = 76 + $k
        foreach ($j in 1..4) { $c.Cells.Item($r, $j).Formula = "=CDT!" + [char](64 + $j) + $o }
    }
    $c.Range("A56:A58").HorizontalAlignment = -4131
    $c.Range("B56:B58").NumberFormatLocal = "0"
    $c.Range("C56:D58").NumberFormatLocal = "0,00"
    $c.Range("A59").Value2 = "Son pocos episodios independientes; no entra al modelo y queda como sensibilidad."
    $c.Range("A59").Font.Size = 9; $c.Range("A59").Font.Color = $gris

    # Especificación
    Bloque $c 8 8 12 "Especificación (M5)"
    $c.Range("H9").Value2 = "s = a + b · ΔTPM13 + ε"
    $c.Range("H9").Font.Italic = $true
    Cabecera $c 10 8 @("Parámetro", "Estimación", "Significado", "", "") @(1)
    $c.Range("J10").IndentLevel = 1
    $par = @(
        @("a (pp)", "=CDT!B49", "0,000", "Margen cuando la TPM no cambia"),
        @("b", "=CDT!C49", "0,000", "Puntos de margen por punto de ΔTPM"),
        @("Error estándar de b", "=CDT!B50", "0,000", "Newey-West, 26 rezagos"),
        @("R²", "=CDT!B51", "0,000", "Parte del margen que explica"),
        @("Desviación de ε (pp)", "=CDT!B52", "0,000", "Tamaño del error"),
        @("Autocorrelación de ε, 1 semana", "=CDT!B53", "0,000", "Muy persistente"),
        @("Autocorrelación de ε, 52 semanas", "=CDT!C53", "0,000", "Independiente de un año a otro")
    )
    for ($k = 0; $k -lt $par.Count; $k++) {
        $c.Cells.Item(11 + $k, 8).Value2 = $par[$k][0]
        $c.Cells.Item(11 + $k, 9).Formula = $par[$k][1]
        $c.Cells.Item(11 + $k, 9).NumberFormatLocal = $par[$k][2]
        $c.Cells.Item(11 + $k, 10).Value2 = $par[$k][3]
    }
    $c.Range("J11:J17").IndentLevel = 1
    $c.Range("H18").Value2 = "El simulador sortea ε una vez por apertura o renovación del CDT."
    $c.Range("H18").Font.Size = 9; $c.Range("H18").Font.Color = $gris

    # Validación
    Bloque $c 23 8 12 "Validación"
    Cabecera $c 24 8 @("Evaluación", "", "", "Cobertura", "") @(3)
    $c.Range("H25").Value2 = "Dentro de la muestra"
    $c.Range("K25").Formula = "=CDT!B65"
    $c.Range("H26").Value2 = "Estimado hasta 2019, probado desde 2020"
    $c.Range("K26").Formula = "=CDT!B66"
    $c.Range("K25:K26").NumberFormatLocal = "0,0 %"
    $c.Range("H27").Value2 = "Proporción de observaciones dentro de la banda nominal de 90 %."
    $c.Range("H28").Value2 = "Fuera de la muestra pesa el ajuste de 2022 a 2024, mayor que cualquier episodio anterior."
    $c.Range("H27:H28").Font.Size = 9; $c.Range("H27:H28").Font.Color = $gris

    # Inversionista
    Bloque $c 30 8 12 "Qué le pasa al inversionista"
    Cabecera $c 31 8 @("Medida", "", "", "Valor", "") @(3)
    $inv = @(
        @("Rentabilidad real a 12 meses (926 compras)", "=CDT!B83", "0,0 %"),
        @("Peor rentabilidad real", "=CDT!B84", "0,0 %"),
        @("Compras con rentabilidad real negativa", "=CDT!B85", "0,0 %"),
        @("Venta a 6 meses: intereses perdidos (ψ)", "=CDT!B86", "0,000"),
        @("Ventas a 6 meses con pérdida de intereses", "=CDT!B87", "0,0 %"),
        @("Ventas a 6 meses con pérdida de capital", "=CDT!B88", "0,0 %"),
        @("ψ con la TPM subiendo", "=CDT!B89", "0,000"),
        @("ψ con la TPM bajando", "=CDT!B90", "0,000")
    )
    for ($k = 0; $k -lt $inv.Count; $k++) {
        $c.Cells.Item(32 + $k, 8).Value2 = $inv[$k][0]
        $c.Cells.Item(32 + $k, 11).Formula = $inv[$k][1]
        $c.Cells.Item(32 + $k, 11).NumberFormatLocal = $inv[$k][2]
    }
    $c.Range("H40").Value2 = "Compras de 2008 a 2025 con la inflación realizada; 956 casos de venta a 6 meses."
    $c.Range("H40").Font.Size = 9; $c.Range("H40").Font.Color = $gris

    # Lectura
    Bloque $c 42 8 12 "Lectura"
    $lec = $c.Range("H43:L57")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $c.Range("H43").Value2 = "El margen del CDT sobre la TPM es persistente y asimétrico, y cambia con el ciclo: es mayor cuando la TPM sube. M5 lo explica con el cambio de la TPM en las últimas trece semanas y es el mejor candidato dentro de la muestra, pero fuera de ella ningún modelo supera al margen fijo. El error es muy persistente de una semana a otra e independiente de un año a otro, y por eso se sortea una vez por apertura. Con solo tres ciclos de tasas completos, el episodio de 2022 a 2024 domina los extremos y la banda de 90 % fuera de la muestra cubre apenas la mitad de las observaciones. El modelo se usa con esa advertencia y con una sensibilidad que duplica la desviación del error."
    $c.Range("A8:A59").EntireRow.RowHeight = 15
    $c.Range("A1").Select()

    # ------------------------------------------------ PIB, estados
    $p = Nueva-Hoja $dest "PIB, estados"
    Titulo $p "Estado de la economía: PIB real de Colombia" "Crecimiento anual trimestral, primer trimestre de 1995 – segundo de 2026"
    $p.Range("A4").Value2 = "Crecimiento medio"
    $p.Range("B4").Value2 = "Desviación"
    $p.Range("C4").Value2 = "Mínimo"
    $p.Range("D4").Value2 = "Trimestres de crisis"
    $p.Range("A5").Formula = "=AVERAGE(B43:B168)/100"
    $p.Range("B5").Formula = "=STDEV.S(B43:B168)/100"
    $p.Range("C5").Formula = "=MIN(B43:B168)/100"
    $p.Range("D5").Formula = '=COUNTIFS(C43:C168,"crisis",A43:A168,">="&DATE(2006,1,1))'
    $p.Range("A6").Value2 = "Anual, 126 trimestres"
    $p.Range("B6").Value2 = "Del crecimiento anual"
    $p.Range("C6").Formula = "=INDEX(A43:A168,MATCH(MIN(B43:B168),B43:B168,0))"
    $p.Range("C6").NumberFormat = '"Terminado en "mm/yyyy'
    $p.Range("C6").HorizontalAlignment = -4131
    $p.Range("D6").Value2 = "Desde 2006"
    foreach ($z in @("A4", "B4", "C4", "D4")) { $p.Range($z).Font.Color = $gris }
    foreach ($z in @("A5", "B5", "C5")) { $p.Range($z).NumberFormatLocal = "0,00 %" }
    $p.Range("D5").NumberFormatLocal = "0"
    foreach ($z in @("A5", "B5", "C5", "D5")) {
        $x = $p.Range($z); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    foreach ($z in @("A6", "B6", "C6", "D6")) { $p.Range($z).Font.Size = 9; $p.Range($z).Font.Color = $gris }
    $p.Rows.Item(5).RowHeight = 28
    foreach ($k in 1..6) { $p.Columns.Item($k).ColumnWidth = @(24, 24, 24, 24, 24, 4)[$k - 1] }

    Seccion $p 8 "Crecimiento anual del PIB real"
    $p.Range("A9").Value2 = "Variación frente al mismo trimestre del año anterior, con los umbrales que definen el estado."
    $p.Range("A9").Font.Color = $gris

    Seccion $p 32 "Estados de la economía"
    $p.Range("A33").Value2 = "Cada trimestre se clasifica por su crecimiento anual. Los umbrales son parámetros del modelo."
    $p.Range("A33").Font.Color = $gris
    Encabezado $p 34 1 4 @("Estado", "Regla", "Trimestres", "Proporción")
    $p.Range("A35").Value2 = "crisis"
    $p.Range("A36").Value2 = "normal"
    $p.Range("A37").Value2 = "expansión"
    $p.Range("B35").Formula = '="Menor que "&TEXT(PIB!H4,"0")&" %"'
    $p.Range("B36").Formula = '="Entre "&TEXT(PIB!H4,"0")&" % y "&TEXT(PIB!H5,"0")&" %"'
    $p.Range("B37").Formula = '=TEXT(PIB!H5,"0")&" % o más"'
    foreach ($k in 35..37) {
        $p.Cells.Item($k, 3).Formula = "=COUNTIF(`$C`$43:`$C`$168,A$k)"
        $p.Cells.Item($k, 4).Formula = "=C$k/SUM(`$C`$35:`$C`$37)"
    }
    $p.Range("C35:C37").NumberFormatLocal = "0"
    $p.Range("D35:D37").NumberFormatLocal = "0 %"
    $p.Range("C34:D37").HorizontalAlignment = -4152
    $p.Range("A38").Value2 = "Desde 2006, un umbral de crisis en 0 % dejaría solo cuatro trimestres y la matriz de Markov no podría estimarse; con 1 % cuentan 2008-09, 2020 y 2023-24."
    $p.Range("A38").Font.Size = 9; $p.Range("A38").Font.Color = $gris

    Seccion $p 41 "Serie trimestral"
    Encabezado $p 42 1 5 @("Trimestre", "Crecimiento anual (%)", "Estado", "Umbral de crisis (%)", "Umbral de expansión (%)")
    $p.Range("A43:A168").Formula = "=PIB!A5"
    $p.Range("B43:B168").Formula = "=PIB!B5"
    $p.Range("C43:C168").Formula = "=PIB!C5"
    $p.Range("D43:D168").Formula = "=PIB!`$H`$4"
    $p.Range("E43:E168").Formula = "=PIB!`$H`$5"
    $p.Range("A43:A168").NumberFormat = "mmm-yyyy"
    $p.Range("A43:A168").HorizontalAlignment = -4131
    $p.Range("B43:B168").NumberFormatLocal = "0,00"
    $p.Range("D43:E168").NumberFormatLocal = "0,0"
    $p.Range("B43:B168").HorizontalAlignment = -4152
    $p.Range("D43:E168").HorizontalAlignment = -4152
    $p.Range("B42").HorizontalAlignment = -4152
    $p.Range("D42:E42").HorizontalAlignment = -4152
    $p.Range("C42:C168").IndentLevel = 1
    $p.Range("A42:E168").AutoFilter() | Out-Null
    Nota $p 170 "Fuente: DANE y Banco de la República (SUAMECA). Bases 1994, 2005 y 2015 encadenadas por la variación anual."

    $izq = $p.Range("A10").Left; $ancho = $p.Range("F10").Left - $izq
    $arr = $p.Range("A10").Top; $alto = $p.Range("A31").Top - $arr
    $co = $p.ChartObjects().Add($izq, $arr, $ancho, $alto)
    $g = $co.Chart
    $g.ChartType = 4
    while ($g.SeriesCollection().Count -gt 0) { $g.SeriesCollection(1).Delete() | Out-Null }
    foreach ($s1 in @(@("Crecimiento anual", "B43:B168", $azul, 1.75, 1), @("Umbral de crisis", "D43:D168", $gris, 1.0, 4), @("Umbral de expansión", "E43:E168", $gris, 1.0, 4))) {
        $a1 = $g.SeriesCollection().NewSeries()
        $a1.Name = $s1[0]
        $a1.Values = $p.Range($s1[1])
        $a1.XValues = $p.Range("A43:A168")
        $a1.Format.Line.ForeColor.RGB = $s1[2]
        $a1.Format.Line.Weight = $s1[3]
        $a1.Format.Line.DashStyle = $s1[4]
    }
    $g.HasLegend = $true
    $g.Legend.Position = -4107
    $ejeX = $g.Axes(1)
    $ejeX.CategoryType = 3
    $ejeX.BaseUnit = 0
    $ejeX.MajorUnitScale = 2
    $ejeX.MajorUnit = 5
    $ejeX.TickLabels.NumberFormat = "yyyy"
    $ejeX.TickLabelPosition = -4134
    $ejeY = $g.Axes(2)
    $ejeY.TickLabels.NumberFormat = '0" %"'
    $ejeY.MajorGridlines.Format.Line.ForeColor.RGB = 0xE0E0E0
    $g.HasTitle = $true
    $g.HasTitle = $false
    $co.Border.LineStyle = -4142
    $p.Range("A1").Select()

    # Orden de lectura: resumen, parámetros, procesos en el orden del informe, estados, validación, resultados y referencia.
    $orden = @("LEEME", "MODELOS", "PARAMETROS", "MODELO", "INF", "Inflación", "Inflación, modelo", "TPM_Cuenta", "TPM_Logit", "CDT, datos", "CDT, modelo", "CDT", "ETF, datos", "ETF, modelo", "ETF", "INM", "PIB, estados", "PIB", "MARKOV", "MARKOV_PIB", "MARKOV_ETF", "MARKOV_Vivienda", "VALIDACION", "RESULTADOS", "FUENTES", "DICCIONARIO", "AUDITORIA")
    $dest.Worksheets.Item($orden[0]).Move($dest.Worksheets.Item(1))
    for ($k = 1; $k -lt $orden.Count; $k++) { $dest.Worksheets.Item($orden[$k]).Move($falta, $dest.Worksheets.Item($orden[$k - 1])) }

    $dest.Worksheets.Item("MODELO").Visible = 0
    $dest.Worksheets.Item("INF").Visible = 0
    $dest.Worksheets.Item("ETF").Visible = 0
    $dest.Worksheets.Item("CDT").Visible = 0
    $dest.Worksheets.Item("PIB").Visible = 0
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
