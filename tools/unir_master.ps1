# ATENCIÓN: data/master_datos.xlsx se editó directamente después de generarse (hojas MODELOS y PARAMETROS, limitaciones, comparación de
# vacancia). Este script NO reproduce esos cambios: volver a correrlo los sobrescribe. El libro es la fuente vigente.
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
    $mk.Range("A84").Value2 = "Lectura: en expansión la vivienda nueva en Medellín sube cerca de 12,7 % al año; en el resto de los años, cerca de 8,7 %. Pero el inmueble del modelo es usado, y con el índice de vivienda usada (IPVU, hoja INM) el estado no cambia la valorización (Welch de una cola, p = 0,275). Por eso el simulador usa el AR(1) del IPVU, igual en los tres estados, y estos valores por estado quedan como alternativa en la sensibilidad."
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

    # ------------------------------------------------ TPM, datos
    $th = Nueva-Hoja $dest "TPM, datos"
    Titulo $th "Tasa de política monetaria" "Decisiones de la Junta del Banco de la República, 46 reuniones de enero de 2021 a septiembre de 2026"
    $th.Range("A4").Value2 = "TPM tras la última reunión"
    $th.Range("B4").Value2 = "Reuniones observadas"
    $th.Range("C4").Value2 = "Reuniones sin cambio"
    $th.Range("A5").Formula = "=TPM_Cuenta!T50"
    $th.Range("B5").Formula = "=TPM_Cuenta!B20"
    $th.Range("C5").Formula = "=TPM_Cuenta!C14"
    $th.Range("A6").Formula = "=TPM_Cuenta!Q50"
    $th.Range("A6").NumberFormat = 'mmmm "de" yyyy'
    $th.Range("B6").Value2 = "Ocho por año"
    $th.Range("C6").Value2 = "Proporción del total"
    foreach ($z in @("A4", "B4", "C4")) { $th.Range($z).Font.Color = $gris }
    $th.Range("A5").NumberFormatLocal = "0,00 %"
    $th.Range("B5").NumberFormatLocal = "0"
    $th.Range("C5").NumberFormatLocal = "0 %"
    foreach ($z in @("A5", "B5", "C5")) {
        $x = $th.Range($z); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    foreach ($z in @("A6", "B6", "C6")) { $th.Range($z).Font.Size = 9; $th.Range($z).Font.Color = $gris; $th.Range($z).HorizontalAlignment = -4131 }
    $th.Rows.Item(5).RowHeight = 28
    foreach ($k in 1..7) { $th.Columns.Item($k).ColumnWidth = @(22, 20, 20, 20, 20, 22, 4)[$k - 1] }

    Seccion $th 8 "Nivel de la TPM"
    $th.Range("A9").Value2 = "TPM vigente después de cada reunión."
    $th.Range("A9").Font.Color = $gris
    Seccion $th 29 "Cambio en cada reunión"
    $th.Range("A30").Value2 = "Puntos porcentuales que sube o baja la Junta. Casi siempre en múltiplos de 0,25."
    $th.Range("A30").Font.Color = $gris

    Seccion $th 51 "Reuniones de la Junta"
    Encabezado $th 52 1 6 @("Mes", "TPM antes", "TPM después", "Cambio (pp)", "Decisión", "Tendencia de la inflación")
    $th.Range("A53:A98").Formula = "=TPM_Cuenta!Q5"
    $th.Range("B53:B98").Formula = "=TPM_Cuenta!S5"
    $th.Range("C53:C98").Formula = "=TPM_Cuenta!T5"
    $th.Range("D53:D98").Formula = "=TPM_Cuenta!U5"
    $th.Range("E53:E98").Formula = "=TPM_Cuenta!V5"
    $th.Range("F53:F98").Formula = "=TPM_Cuenta!Z5"
    $th.Range("A53:A98").NumberFormat = "mmm-yyyy"
    $th.Range("A53:A98").HorizontalAlignment = -4131
    $th.Range("B53:C98").NumberFormatLocal = "0,00 %"
    $th.Range("D53:D98").NumberFormatLocal = "0,00"
    $th.Range("B53:D98").HorizontalAlignment = -4152
    $th.Range("B52:D52").HorizontalAlignment = -4152
    $th.Range("E52:F98").IndentLevel = 1
    $th.Range("A52:F98").AutoFilter() | Out-Null
    Nota $th 100 "Fuente: Banco de la República. Las primeras reuniones se reconstruyen con el IBR overnight; el resto sigue el calendario de decisiones de la Junta. La tendencia usa la inflación del DANE."

    $izq = $th.Range("A10").Left; $ancho = $th.Range("G10").Left - $izq
    foreach ($par in @(@(10, 28, 4, "C53:C98", "0%", 1), @(31, 49, 51, "D53:D98", "0,0", 0))) {
        $arr = $th.Range("A" + $par[0]).Top; $alto = $th.Range("A" + $par[1]).Top - $arr
        $co = $th.ChartObjects().Add($izq, $arr, $ancho, $alto)
        $g = $co.Chart
        $g.ChartType = $par[2]
        while ($g.SeriesCollection().Count -gt 0) { $g.SeriesCollection(1).Delete() | Out-Null }
        $a1 = $g.SeriesCollection().NewSeries()
        $a1.Name = "Serie"
        $a1.Values = $th.Range($par[3])
        $a1.XValues = $th.Range("A53:A98")
        if ($par[2] -eq 4) { $a1.Format.Line.ForeColor.RGB = $azul; $a1.Format.Line.Weight = 2 }
        else { $a1.Format.Fill.ForeColor.RGB = $azul; $g.ChartGroups(1).GapWidth = 60 }
        $g.HasLegend = $false
        $ejeX = $g.Axes(1)
        $ejeX.CategoryType = 3
        $ejeX.BaseUnit = 1
        $ejeX.MajorUnitScale = 2
        $ejeX.MajorUnit = 1
        $ejeX.TickLabels.NumberFormat = "yyyy"
        $ejeX.TickLabelPosition = -4134
        $ejeY = $g.Axes(2)
        $ejeY.TickLabels.NumberFormat = $par[4]
        $ejeY.MajorGridlines.Format.Line.ForeColor.RGB = 0xE0E0E0
        $g.HasTitle = $true
        $g.HasTitle = $false
        $co.Border.LineStyle = -4142
    }
    $th.Range("A1").Select()

    # ------------------------------------------------ TPM, modelo
    $tm = Nueva-Hoja $dest "TPM, modelo"
    Titulo $tm "TPM: modelo de las decisiones de la Junta" "Logit ordenado con la tendencia de la inflación y la brecha de Taylor, decisiones de 2001 a 2026"
    foreach ($k in 1..12) { $tm.Columns.Item($k).ColumnWidth = @(52, 12, 10, 16, 14, 12, 3, 30, 11, 13, 13, 17)[$k - 1] }
    $tm.Range("H4").Value2 = "Decisiones desde 2001"
    $tm.Range("J4").Value2 = "b tendencia"
    $tm.Range("K4").Value2 = "b Taylor"
    $tm.Range("H5").Formula = "=TPM_Logit!B13"
    $tm.Range("J5").Formula = "=TPM_Logit!B6"
    $tm.Range("K5").Formula = "=TPM_Logit!B7"
    foreach ($z in @("H4", "J4", "K4")) { $tm.Range($z).Font.Color = $gris }
    foreach ($z in @("H5", "J5", "K5")) {
        $x = $tm.Range($z); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    $tm.Range("H5").NumberFormatLocal = "0"
    $tm.Range("J5").NumberFormatLocal = "0,00"
    $tm.Range("K5").NumberFormatLocal = "0,00"
    $tm.Rows.Item(5).RowHeight = 28
    $pfmt = '[<0,001]"<0,001";0,000'

    # Evidencia
    Bloque $tm 8 1 4 "Evidencia estadística"
    Cabecera $tm 9 1 @("Prueba y contraste", "Estadístico", "p", "Decisión") @(1, 2)
    $pr = @(
        @("Chi-cuadrado: los cambios de la TPM siguen una normal", "=TPM_Cuenta!C45", "=TPM_Cuenta!C47"),
        @("Chi-cuadrado: la decisión es independiente de la tendencia", "=TPM_Cuenta!C65", "=TPM_Cuenta!C67"),
        @("t de Welch: subidas y bajadas tienen el mismo tamaño", "=TPM_Cuenta!C93", "=TPM_Cuenta!C95"),
        @("Razón de verosimilitud: la brecha de Taylor no mejora el logit", "=TPM_Logit!B12", "=TPM_Logit!C12")
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 10 + $k
        $tm.Cells.Item($r, 1).Value2 = $pr[$k][0]
        $tm.Cells.Item($r, 2).Formula = $pr[$k][1]
        $tm.Cells.Item($r, 3).Formula = $pr[$k][2]
        $tm.Cells.Item($r, 4).Formula = "=IF(C$r<TPM_Cuenta!`$C`$5,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $tm.Range("B10:B13").NumberFormatLocal = "0,000"
    $tm.Range("C10:C13").NumberFormatLocal = $pfmt
    $tm.Range("D9:D13").IndentLevel = 1
    $tm.Range("A14").Value2 = "Significancia de 5 %. En la prueba de independencia varias casillas esperadas tienen menos de 5 reuniones: el p es aproximado."
    $tm.Range("A14").Font.Size = 9; $tm.Range("A14").Font.Color = $gris

    # Coeficientes
    Bloque $tm 17 1 6 "Coeficientes del logit ordenado"
    Cabecera $tm 18 1 @("Coeficiente", "Valor", "p", "Significado", "", "") @(1, 2)
    $tm.Range("D18").IndentLevel = 1
    $co1 = @(
        @("b tendencia", "=TPM_Logit!B6", "=TPM_Logit!C6", "Inflación al alza: más probable subir"),
        @("b Taylor", "=TPM_Logit!B7", "=TPM_Logit!C7", "TPM sobre la regla: menos probable subir"),
        @("k1", "=TPM_Logit!B8", "", "Umbral para bajar la tasa"),
        @("k2", "=TPM_Logit!B9", "", "Umbral para subirla")
    )
    for ($k = 0; $k -lt $co1.Count; $k++) {
        $r = 19 + $k
        $tm.Cells.Item($r, 1).Value2 = $co1[$k][0]
        $tm.Cells.Item($r, 2).Formula = $co1[$k][1]
        if ($co1[$k][2] -ne "") { $tm.Cells.Item($r, 3).Formula = $co1[$k][2] }
        $tm.Cells.Item($r, 4).Value2 = $co1[$k][3]
    }
    $tm.Range("B19:B22").NumberFormatLocal = "0,000"
    $tm.Range("C19:C22").NumberFormatLocal = $pfmt
    $tm.Range("D19:D22").IndentLevel = 1

    # Tamaño del paso
    Bloque $tm 25 1 3 "Tamaño del paso cuando la Junta mueve la tasa"
    Cabecera $tm 26 1 @("Tamaño (pp)", "Si sube", "Si baja") @(1, 2)
    $tam = @(0.25, 0.5, 0.75, 1, 1.5, 1.75, 2)
    for ($k = 0; $k -lt $tam.Count; $k++) {
        $r = 27 + $k
        $tm.Cells.Item($r, 1).Formula = "=" + $tam[$k].ToString([Globalization.CultureInfo]::InvariantCulture)
        $tm.Cells.Item($r, 2).Formula = "=SUMIFS(TPM_Logit!`$C`$18:`$C`$29,TPM_Logit!`$A`$18:`$A`$29,""Subida"",TPM_Logit!`$B`$18:`$B`$29,`$A$r)"
        $tm.Cells.Item($r, 3).Formula = "=SUMIFS(TPM_Logit!`$C`$18:`$C`$29,TPM_Logit!`$A`$18:`$A`$29,""Bajada"",TPM_Logit!`$B`$18:`$B`$29,`$A$r)"
    }
    $tm.Range("A27:A33").NumberFormatLocal = "0,00"
    $tm.Range("A27:A33").HorizontalAlignment = -4131
    $tm.Range("B27:C33").NumberFormatLocal = "0,0 %"
    $tm.Range("A34").Value2 = "Probabilidad de cada tamaño, dado que la Junta sube o baja. Estimada con 52 subidas y 48 bajadas."
    $tm.Range("A34").Font.Size = 9; $tm.Range("A34").Font.Color = $gris

    # Tendencia
    Bloque $tm 36 1 5 "Decisiones de 2021 a 2026 según la tendencia de la inflación"
    Cabecera $tm 37 1 @("Tendencia de la inflación", "Reuniones", "P(baja)", "P(mantiene)", "P(sube)") @(1, 2, 3, 4)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 38 + $k; $o = 74 + $k; $n = 55 + $k
        $tm.Cells.Item($r, 1).Formula = "=TPM_Cuenta!A$o"
        $tm.Cells.Item($r, 2).Formula = "=TPM_Cuenta!E$n"
        $tm.Cells.Item($r, 3).Formula = "=TPM_Cuenta!B$o"
        $tm.Cells.Item($r, 4).Formula = "=TPM_Cuenta!C$o"
        $tm.Cells.Item($r, 5).Formula = "=TPM_Cuenta!D$o"
    }
    $tm.Range("B38:B40").NumberFormatLocal = "0"
    $tm.Range("C38:E40").NumberFormatLocal = "0 %"
    $tm.Range("A41").Value2 = "Frecuencias de las 46 reuniones. Con esta tabla la TPM simulada tiende al tope en horizontes largos."
    $tm.Range("A41").Font.Size = 9; $tm.Range("A41").Font.Color = $gris

    # Normalidad
    Bloque $tm 43 1 3 "Cambios de la TPM frente a la normal"
    Cabecera $tm 44 1 @("Grupo de reuniones", "Observadas", "Esperadas") @(1, 2)
    for ($k = 0; $k -lt 4; $k++) {
        $r = 45 + $k; $o = 41 + $k
        $tm.Cells.Item($r, 1).Formula = "=TPM_Cuenta!A$o"
        $tm.Cells.Item($r, 2).Formula = "=TPM_Cuenta!D$o"
        $tm.Cells.Item($r, 3).Formula = "=TPM_Cuenta!E$o"
    }
    $tm.Range("B45:B48").NumberFormatLocal = "0"
    $tm.Range("C45:C48").NumberFormatLocal = "0,0"
    $tm.Range("A49").Value2 = "Hay muchas más reuniones sin cambio de las que prevé la normal."
    $tm.Range("A49").Font.Size = 9; $tm.Range("A49").Font.Color = $gris

    # Especificación
    Bloque $tm 8 8 12 "Especificación"
    $tm.Range("H9").Value2 = "P(baja) = F(k₁ − x),   P(sube) = 1 − F(k₂ − x),   x = b_tend · tendencia + b_Taylor · brecha"
    $tm.Range("H9").Font.Italic = $true
    Cabecera $tm 10 8 @("Elemento", "Valor", "Significado", "", "") @(1)
    $tm.Range("J10").IndentLevel = 1
    $esp = @(
        @("F", "", "Función logística"),
        @("Tendencia", "3 meses", "Cambio de la inflación anual"),
        @("Brecha de Taylor", "", "TPM − (r* + π + 0,5 (π − meta))"),
        @("Reuniones por año", "8", "Calendario de la Junta"),
        @("TPM mínima", "=TPM_Cuenta!C7", "Límite inferior en la simulación"),
        @("TPM máxima", "=TPM_Cuenta!C8", "Límite superior en la simulación"),
        @("AIC solo con tendencia", "=TPM_Logit!B10", "Modelo sin brecha de Taylor"),
        @("AIC con brecha de Taylor", "=TPM_Logit!B11", "Modelo elegido")
    )
    for ($k = 0; $k -lt $esp.Count; $k++) {
        $r = 11 + $k
        $tm.Cells.Item($r, 8).Value2 = $esp[$k][0]
        if ($esp[$k][1] -like "=*") { $tm.Cells.Item($r, 9).Formula = $esp[$k][1] } elseif ($esp[$k][1] -ne "") { $tm.Cells.Item($r, 9).Value2 = $esp[$k][1] }
        $tm.Cells.Item($r, 10).Value2 = $esp[$k][2]
    }
    $tm.Range("I11:I18").HorizontalAlignment = -4152
    $tm.Range("I15:I16").NumberFormatLocal = "0,00 %"
    $tm.Range("I17:I18").NumberFormatLocal = "0,0"
    $tm.Range("J11:J18").IndentLevel = 1

    # Validación
    Bloque $tm 21 8 12 "Validación del simulador"
    Cabecera $tm 22 8 @("Medida", "", "", "Simulado", "Observado") @(3, 4)
    $val = @(
        @("TPM mediana en el año 10 (%)", 8, "0,00"),
        @("Tasa real media (puntos)", 9, "0,00"),
        @("Tasa real media, años 6 a 10 (puntos)", 10, "0,00"),
        @("Réplicas con la TPM en el tope de 15 %", 11, "0,0 %")
    )
    for ($k = 0; $k -lt $val.Count; $k++) {
        $r = 23 + $k; $o = $val[$k][1]
        $tm.Cells.Item($r, 8).Value2 = $val[$k][0]
        $tm.Cells.Item($r, 11).Formula = "=VALIDACION!D$o"
        $tm.Cells.Item($r, 12).Formula = "=VALIDACION!E$o"
        $tm.Range($tm.Cells.Item($r, 11), $tm.Cells.Item($r, 12)).NumberFormatLocal = $val[$k][2]
    }
    $tm.Range("K23:L26").HorizontalAlignment = -4152
    $tm.Range("H27").Value2 = "500 réplicas con los parámetros de referencia, frente a la historia desde 2001."
    $tm.Range("H28").Formula = '="Con la tabla por tendencia, "&TEXT(TPM_Cuenta!C110*100,"0")&" % de las reuniones de 2021 a 2026 cayó en la banda de 90 %."'
    $tm.Range("H27:H28").Font.Size = 9; $tm.Range("H27:H28").Font.Color = $gris

    # Lectura
    Bloque $tm 31 8 12 "Lectura"
    $lec = $tm.Range("H32:L46")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $tm.Range("H32").Value2 = "La Junta sube la tasa cuando la inflación sube y, en la muestra de 2021 a 2026, nunca la subió cuando bajaba; la independencia entre decisión y tendencia se rechaza con claridad, aunque algunas casillas tienen pocas reuniones. El logit ordenado agrega la brecha de Taylor, que mejora el ajuste, y se estima con 25 años de decisiones en lugar de 46 reuniones. Las subidas son más grandes que las bajadas, y por eso sus tamaños se sortean por separado. Como limitación, la tasa real simulada queda por encima de la histórica: el modelo mantiene la TPM alta más tiempo del observado."
    $tm.Range("A8:A49").EntireRow.RowHeight = 15
    $tm.Range("A1").Select()

    # ------------------------------------------------ Cuenta, modelo
    $cm = Nueva-Hoja $dest "Cuenta, modelo"
    Titulo $cm "Cuenta digital: tasa en función de la TPM" "La cuenta paga una proporción fija de la TPM: r = α · TPM"
    foreach ($k in 1..11) { $cm.Columns.Item($k).ColumnWidth = @(46, 12, 12, 16, 22, 3, 28, 11, 20, 13, 17)[$k - 1] }
    $cm.Range("G4").Value2 = "α, proporción de la TPM"
    $cm.Range("I4").Value2 = "IC de 95 % por cuenta"
    $cm.Range("J4").Value2 = "Tasas observadas"
    $cm.Range("G5").Formula = "=TPM_Cuenta!AD7"
    $cm.Range("I5").Formula = "=B17"
    $cm.Range("J5").Formula = "=TPM_Cuenta!AD6"
    foreach ($z in @("G4", "I4", "J4")) { $cm.Range($z).Font.Color = $gris }
    foreach ($z in @("G5", "I5", "J5")) {
        $x = $cm.Range($z); $x.Font.Size = 18; $x.Font.Bold = $true; $x.HorizontalAlignment = -4131
    }
    $cm.Range("G5").NumberFormatLocal = "0,000"
    $cm.Range("I5").Font.Size = 14
    $cm.Range("J5").NumberFormatLocal = "0"
    $cm.Rows.Item(5).RowHeight = 28
    $pfmt = '[<0,001]"<0,001";0,000'

    Bloque $cm 8 1 5 "Estimación de α"
    Cabecera $cm 9 1 @("Concepto", "Valor", "Cómo se calcula", "", "") @(1)
    $cm.Range("C9").IndentLevel = 1
    $est = @(
        @("Número de tasas (n)", "=TPM_Cuenta!AD6", "0", "4 cuentas digitales en 5 fechas"),
        @("α = promedio de tasa ÷ TPM", "=TPM_Cuenta!AD7", "0,000", "Promedio de las razones individuales"),
        @("Desviación de las razones (s)", "=TPM_Cuenta!AD11", "0,000", "Cuánto varían las cuentas entre sí"),
        @("t crítico (95 %)", "=TPM_Cuenta!AD12", "0,000", "t de Student con n − 1 grados"),
        @("Margen de error (19 tasas)", "=TPM_Cuenta!AD13", "0,000", "t · s ÷ √n"),
        @("Límite inferior (19 tasas)", "=TPM_Cuenta!AD14", "0,000", "α − margen de error"),
        @("Límite superior (19 tasas)", "=TPM_Cuenta!AD15", "0,000", "α + margen de error")
    )
    for ($k = 0; $k -lt $est.Count; $k++) {
        $r = 10 + $k
        $cm.Cells.Item($r, 1).Value2 = $est[$k][0]
        $cm.Cells.Item($r, 2).Formula = $est[$k][1]
        $cm.Cells.Item($r, 2).NumberFormatLocal = $est[$k][2]
        $cm.Cells.Item($r, 3).Value2 = $est[$k][3]
    }
    $cm.Range("C10:C16").IndentLevel = 1
    $cm.Range("A17").Value2 = "IC de 95 % por cuenta"
    $cm.Range("A17").Font.Bold = $true
    $cm.Range("B17").Formula = '=ROUND(N14-N15,3)&" a "&ROUND(N14+N15,3)'
    $cm.Range("B17").HorizontalAlignment = -4152
    $cm.Range("C17").Value2 = "Promedio de cada cuenta, t con 3 grados de libertad"
    $cm.Range("C17").IndentLevel = 1
    $cm.Range("M9").Value2 = "Cuenta"
    $cm.Range("N9").Value2 = "α medio"
    $cn = @("Nu", "Pibank", "Lulo", "RappiPay")
    for ($k = 0; $k -lt 4; $k++) {
        $cm.Cells.Item(10 + $k, 13).Value2 = $cn[$k]
        $cm.Cells.Item(10 + $k, 14).Formula = '=AVERAGEIFS(TPM_Cuenta!$AV$5:$AV$28,TPM_Cuenta!$AS$5:$AS$28,M' + (10 + $k) + ')'
    }
    $cm.Range("M14").Value2 = "Promedio"
    $cm.Range("N14").Formula = "=AVERAGE(N10:N13)"
    $cm.Range("M15").Value2 = "Margen"
    $cm.Range("N15").Formula = "=STDEV.S(N10:N13)/SQRT(4)*T.INV.2T(0.05,3)"
    $cm.Range("M9:N15").Font.Color = $gris
    $cm.Range("M9:N15").Font.Size = 9
    $cm.Range("N10:N15").NumberFormatLocal = "0,000"

    Bloque $cm 19 1 5 "¿Depende α del nivel de la TPM?"
    Cabecera $cm 20 1 @("Concepto", "Valor", "Cómo se calcula", "", "") @(1)
    $cm.Range("C20").IndentLevel = 1
    $cm.Range("A21").Value2 = "Correlación entre α individual y TPM"
    $cm.Range("B21").Formula = "=TPM_Cuenta!AD21"
    $cm.Range("C21").Value2 = "H₀: la razón no depende de la TPM (correlación igual a 0)"
    $cm.Range("A22").Value2 = "Estadístico t"
    $cm.Range("B22").Formula = "=TPM_Cuenta!AD22"
    $cm.Range("C22").Value2 = "r · √(n − 2) ÷ √(1 − r²)"
    $cm.Range("A23").Value2 = "Valor p"
    $cm.Range("B23").Formula = "=TPM_Cuenta!AD23"
    $cm.Range("C23").Formula = '=IF(B23<TPM_Cuenta!$C$5,"Rechaza H₀","No rechaza H₀")'
    $cm.Range("B21:B22").NumberFormatLocal = "0,000"
    $cm.Range("B23").NumberFormatLocal = $pfmt
    $cm.Range("C21:C23").IndentLevel = 1
    $cm.Range("A24").Value2 = "Significancia de 5 %. No rechazar H₀ no demuestra que α sea constante."
    $cm.Range("A24").Font.Size = 9; $cm.Range("A24").Font.Color = $gris

    Bloque $cm 26 1 3 "Cómo lo usa el simulador"
    Cabecera $cm 27 1 @("TPM", "Cuenta (E.A.)", "Tasa mensual") @(1, 2)
    for ($k = 0; $k -lt 5; $k++) {
        $r = 28 + $k; $o = 28 + $k
        $cm.Cells.Item($r, 1).Formula = "=TPM_Cuenta!AC$o"
        $cm.Cells.Item($r, 2).Formula = "=TPM_Cuenta!AD$o"
        $cm.Cells.Item($r, 3).Formula = "=TPM_Cuenta!AE$o"
    }
    $cm.Range("A28:A32").NumberFormatLocal = "0,00 %"
    $cm.Range("A28:A32").HorizontalAlignment = -4131
    $cm.Range("B28:B32").NumberFormatLocal = "0,00 %"
    $cm.Range("C28:C32").NumberFormatLocal = "0,000 %"
    $cm.Range("A33").Value2 = "Cada vez que cambia la TPM, la tasa de la cuenta se recalcula con r = α · TPM; la tasa mensual es (1 + r)^(1/12) − 1."
    $cm.Range("A33").Font.Size = 9; $cm.Range("A33").Font.Color = $gris

    Bloque $cm 8 7 11 "Tasas observadas"
    Cabecera $cm 9 7 @("Fecha", "Cuenta", "TPM", "Tasa (E.A.)", "α individual") @(2, 3, 4)
    for ($k = 0; $k -lt 24; $k++) {
        $r = 10 + $k; $o = 5 + $k
        $cm.Cells.Item($r, 7).Formula = "=TPM_Cuenta!AR$o"
        $cm.Cells.Item($r, 8).Formula = "=TPM_Cuenta!AS$o"
        $cm.Cells.Item($r, 9).Formula = "=TPM_Cuenta!AT$o"
        $cm.Cells.Item($r, 10).Formula = "=IF(TPM_Cuenta!AU$o="""","""",TPM_Cuenta!AU$o)"
        $cm.Cells.Item($r, 11).Formula = "=IF(TPM_Cuenta!AV$o="""","""",TPM_Cuenta!AV$o)"
    }
    $cm.Range("I10:J33").NumberFormatLocal = "0,00 %"
    $cm.Range("K10:K33").NumberFormatLocal = "0,000"
    $cm.Range("I10:K33").HorizontalAlignment = -4152
    $cm.Range("H10:H33").IndentLevel = 1
    $cm.Range("H9").IndentLevel = 1
    $cm.Range("G34").Value2 = "Ualá se excluye: es un depósito de bajo monto con tope cercano a 10 millones."
    $cm.Range("G34").Font.Size = 9; $cm.Range("G34").Font.Color = $gris

    Bloque $cm 36 1 11 "Lectura"
    $lec = $cm.Range("A37:K39")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $cm.Range("A37").Value2 = "Las cuentas digitales pagan en promedio cerca de nueve décimas de la TPM, y la razón de cada cuenta no muestra una relación significativa con el nivel de la tasa, por lo que una proporción fija describe razonablemente los datos. Como las 19 tasas vienen de cuatro cuentas, el intervalo de confianza se calcula con el promedio de cada cuenta y es más ancho que el de las 19 tasas tratadas como independientes. La estimación descansa en pocas observaciones: cuatro cuentas, cinco fechas y una tasa pagada que cada entidad fija a su criterio."
    $cm.Range("A8:A39").EntireRow.RowHeight = 15
    $cm.Range("A1").Select()

    # ------------------------------------------------ Inmueble (cinco hojas)
    function Kpi($h, $col, $etq, $formula, $fmt, $sub) {
        $h.Range($col + "4").Value2 = $etq
        $h.Range($col + "4").Font.Color = $gris
        $h.Range($col + "5").Formula = $formula
        $h.Range($col + "5").NumberFormatLocal = $fmt
        $h.Range($col + "5").Font.Size = 18
        $h.Range($col + "5").Font.Bold = $true
        $h.Range($col + "5").HorizontalAlignment = -4131
        $h.Rows.Item(5).RowHeight = 28
        if ($sub) {
            $h.Range($col + "6").Value2 = $sub
            $h.Range($col + "6").Font.Size = 9
            $h.Range($col + "6").Font.Color = $gris
        }
    }
    function Nuevo-Grafico($h, $desde, $hasta, $col1, $col2, $tipo) {
        $izq = $h.Range($col1 + $desde).Left
        $arr = $h.Range($col1 + $desde).Top
        $ancho = $h.Range($col2 + $desde).Left + $h.Range($col2 + $desde).Width - $izq
        $alto = $h.Range($col1 + $hasta).Top - $arr
        $co = $h.ChartObjects().Add($izq, $arr, $ancho, $alto)
        $g = $co.Chart
        $g.ChartType = $tipo
        while ($g.SeriesCollection().Count -gt 0) { $g.SeriesCollection(1).Delete() | Out-Null }
        $co.Border.LineStyle = -4142
        return $g
    }
    function Nueva-Serie($g, $nombre, $vals, $xs, $color, $peso, $guion) {
        $s = $g.SeriesCollection().NewSeries()
        $s.Name = $nombre
        $s.Values = $vals
        $s.XValues = $xs
        if ($g.ChartType -eq 74) { $s.MarkerStyle = -4142 }
        if ($g.ChartType -eq 57 -or $g.ChartType -eq 51) { $s.Format.Fill.ForeColor.RGB = $color }
        else {
            $s.Format.Line.ForeColor.RGB = $color
            $s.Format.Line.Weight = [single]$peso
            $s.Format.Line.DashStyle = $guion
        }
    }
    function Cerrar-Grafico($g, $fmtY, $leyenda) {
        $g.HasTitle = $true
        $g.HasTitle = $false
        $g.HasLegend = $leyenda
        if ($leyenda) { $g.Legend.Position = -4107 }
        $g.Axes(2).TickLabels.NumberFormat = $fmtY
        $g.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = 0xE0E0E0
    }

    # ---- Inmueble, datos
    $ih = Nueva-Hoja $dest "Inmueble, datos"
    Titulo $ih "Inmueble: canon, precio y costos" "Apartamentos de Medellín, Lonja de Propiedad Raíz, edición 275 de agosto de 2025"
    foreach ($k in 1..5) { $ih.Columns.Item($k).ColumnWidth = @(36, 20, 20, 20, 4)[$k - 1] }
    Kpi $ih "A" "Rentabilidad bruta anual" "=INM!H16" "0,0 %" "Con la mediana de canon ÷ precio"
    Kpi $ih "B" "Canon ÷ precio mensual" "=INM!H13" "0,00 %" "Mediana de 11 barrios"
    Kpi $ih "C" "Costo total de venta" "=INM!W23" "0,00 %" "Del valor del inmueble"

    Seccion $ih 8 "Canon frente a precio de venta"
    $ih.Range("A9").Value2 = "Canon mensual dividido por el precio de venta, por barrio y tamaño."
    $ih.Range("A9").Font.Color = $gris
    Seccion $ih 30 "Canon y precio por barrio"
    Encabezado $ih 31 1 4 @("Barrio y tamaño (m²)", "Precio de venta por m²", "Canon mensual por m²", "Canon ÷ precio")
    for ($k = 0; $k -lt 11; $k++) {
        $r = 32 + $k; $o = 5 + $k
        $ih.Cells.Item($r, 1).Formula = "=INM!A$o&"", ""&INM!B$o"
        $ih.Cells.Item($r, 2).Formula = "=INM!C$o"
        $ih.Cells.Item($r, 3).Formula = "=INM!D$o"
        $ih.Cells.Item($r, 4).Formula = "=INM!E$o"
    }
    $ih.Range("B32:C42").NumberFormatLocal = "#.##0"
    $ih.Range("D32:D42").NumberFormatLocal = "0,00 %"
    $ih.Range("B31:D42").HorizontalAlignment = -4152
    Nota $ih 43 "Fuente: Lonja de Propiedad Raíz de Medellín y Antioquia. Los encabezados del original estaban intercambiados y se ubicaron por magnitud."

    Seccion $ih 45 "Costos de tenencia y venta"
    $ih.Range("A46").Value2 = "Como proporción del valor comercial del inmueble, para que no dependan de un precio fijo."
    $ih.Range("A46").Font.Color = $gris
    Encabezado $ih 47 1 4 @("Concepto", "Valor", "Base", "")
    $cos = @(
        @("Predial anual", "=INM!W8", "Tarifa de 9,5 por mil sobre el 65 % del valor"),
        @("Mantenimiento preventivo anual", "=INM!W14", "Del valor"),
        @("Gasto fijo anual del propietario", "=INM!W19", "Predial, administración y mantenimiento"),
        @("Comisión inmobiliaria", "=INM!W10", "Del canon, sin IVA de 19 %"),
        @("Comisión de venta, con IVA", "=INM!W20", "Del precio de venta"),
        @("Notaría, mitad del vendedor", "=INM!W21", "Del precio de venta"),
        @("Retención en la fuente", "=INM!W22", "Del precio de venta"),
        @("Costo total de venta", "=INM!W23", "Suma de los tres anteriores")
    )
    for ($k = 0; $k -lt $cos.Count; $k++) {
        $r = 48 + $k
        $ih.Cells.Item($r, 1).Value2 = $cos[$k][0]
        $ih.Cells.Item($r, 2).Formula = $cos[$k][1]
        $ih.Cells.Item($r, 3).Value2 = $cos[$k][2]
    }
    $ih.Range("B48:B55").NumberFormatLocal = "0,00 %"
    $ih.Range("B47:B55").HorizontalAlignment = -4152
    $ih.Range("C47:C55").IndentLevel = 1
    Nota $ih 56 "Fuente: Acuerdo 093 de 2023 del Concejo de Medellín, tarifas publicadas y OneEstimate 2026."

    $g = Nuevo-Grafico $ih 10 28 "A" "D" 57
    Nueva-Serie $g "Canon ÷ precio" $ih.Range("D32:D42") $ih.Range("A32:A42") $azul 0 0
    $g.Axes(1).ReversePlotOrder = $true
    $g.ChartGroups(1).GapWidth = 50
    Cerrar-Grafico $g "0,0%" $false
    $ih.Range("A1").Select()

    # ---- Reajuste, modelo
    $rm = Nueva-Hoja $dest "Reajuste, modelo"
    Titulo $rm "Canon: reajuste anual" "Arriendo efectivo y IPC total del DANE, diciembre de cada año, 2010 a 2025"
    foreach ($k in 1..10) { $rm.Columns.Item($k).ColumnWidth = @(12, 18, 16, 18, 20, 3, 44, 12, 14, 16)[$k - 1] }
    Kpi $rm "G" "Coeficiente del IPC previo (b)" "=INM!T17" "0,00" ""
    Kpi $rm "H" "R²" "=INM!T18" "0 %" ""
    Kpi $rm "I" "σ del error (pp)" "=INM!T19" "0,00" ""
    $pfmt = '[<0,001]"<0,001";0,000'

    Seccion $rm 8 "Arriendo frente al IPC"
    $rm.Range("A9").Value2 = "Variación de diciembre a diciembre, en %."
    $rm.Range("A9").Font.Color = $gris
    Seccion $rm 28 "Serie anual"
    Encabezado $rm 29 1 5 @("Año", "Arriendo efectivo (%)", "IPC total (%)", "IPC previo (%)", "Arriendo − IPC (pp)")
    for ($k = 0; $k -lt 16; $k++) {
        $r = 30 + $k; $o = 5 + $k
        $rm.Cells.Item($r, 1).Formula = "=INM!M$o"
        $rm.Cells.Item($r, 2).Formula = "=INM!N$o"
        $rm.Cells.Item($r, 3).Formula = "=INM!O$o"
        $rm.Cells.Item($r, 4).Formula = "=INM!P$o"
        $rm.Cells.Item($r, 5).Formula = "=INM!Q$o"
    }
    $rm.Range("A30:A45").NumberFormatLocal = "0"
    $rm.Range("A30:A45").HorizontalAlignment = -4131
    $rm.Range("B30:E45").NumberFormatLocal = "0,00"
    $rm.Range("B29:E45").HorizontalAlignment = -4152
    Nota $rm 46 "Fuente: DANE, IPC por clase y subclase; IPC total del Banco de la República con cifras del DANE. El dato de 2023 se deduce del índice de diciembre de 2024."
    $rm.Range("A47").Value2 = "Septiembre de 2026: arriendo efectivo de 5,12 % anual frente a 6,29 % del IPC total."
    $rm.Range("A47").Font.Size = 9; $rm.Range("A47").Font.Color = $gris
    $g = Nuevo-Grafico $rm 10 27 "A" "E" 4
    Nueva-Serie $g "Arriendo efectivo" $rm.Range("B30:B45") $rm.Range("A30:A45") $azul 2 1
    Nueva-Serie $g "IPC total" $rm.Range("C30:C45") $rm.Range("A30:A45") $gris 1.5 4
    Cerrar-Grafico $g '0" %"' $true

    Bloque $rm 8 7 10 "Evidencia estadística"
    Cabecera $rm 9 7 @("Prueba y contraste", "Valor", "p", "Decisión") @(1, 2)
    $pr = @(
        @("t de una muestra: arriendo menos IPC igual a 0", "=INM!T10", "=INM!T11"),
        @("Correlación del arriendo con el IPC del mismo año", "=INM!T12", "=INM!T13"),
        @("Correlación del arriendo con el IPC del año anterior", "=INM!T14", "=INM!T15")
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 10 + $k
        $rm.Cells.Item($r, 7).Value2 = $pr[$k][0]
        $rm.Cells.Item($r, 8).Formula = $pr[$k][1]
        $rm.Cells.Item($r, 9).Formula = $pr[$k][2]
        $rm.Cells.Item($r, 10).Formula = "=IF(I$r<0.05,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $rm.Range("H10:H12").NumberFormatLocal = "0,000"
    $rm.Range("I10:I12").NumberFormatLocal = $pfmt
    $rm.Range("J9:J12").IndentLevel = 1
    $rm.Range("G13").Value2 = "Significancia de 5 %. Con errores robustos (HC1), el coeficiente del IPC previo tiene p < 0,001."
    $rm.Range("G14").Value2 = "Wilcoxon de la diferencia: p = 0,23."
    $rm.Range("G13:G14").Font.Size = 9; $rm.Range("G13:G14").Font.Color = $gris

    Bloque $rm 16 7 10 "Especificación"
    $rm.Range("G17").Value2 = "arriendo(t) = a + b · IPC(t−1) + ε"
    $rm.Range("G17").Font.Italic = $true
    Cabecera $rm 18 7 @("Parámetro", "Estimación", "Significado", "") @(1)
    $rm.Range("I18").IndentLevel = 1
    $esp = @(
        @("a (%)", "=INM!T16", "0,00", "Reajuste con IPC previo en cero"),
        @("b", "=INM!T17", "0,000", "Puntos de arriendo por punto de IPC"),
        @("R²", "=INM!T18", "0,000", "Parte del arriendo que explica"),
        @("σ de ε (pp)", "=INM!T19", "0,00", "Tamaño del error"),
        @("f, sin constante", "=INM!T20", "0,000", "Fracción del IPC previo"),
        @("Error típico con f = 1 (pp)", "=INM!T21", "0,00", "Sensibilidad"),
        @("Error típico con el modelo (pp)", "=INM!T22", "0,00", "Modelo elegido")
    )
    for ($k = 0; $k -lt $esp.Count; $k++) {
        $r = 19 + $k
        $rm.Cells.Item($r, 7).Value2 = $esp[$k][0]
        $rm.Cells.Item($r, 8).Formula = $esp[$k][1]
        $rm.Cells.Item($r, 8).NumberFormatLocal = $esp[$k][2]
        $rm.Cells.Item($r, 9).Value2 = $esp[$k][3]
    }
    $rm.Range("I19:I25").IndentLevel = 1

    Bloque $rm 28 7 10 "Lectura"
    $lec = $rm.Range("G29:J43")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $rm.Range("G29").Value2 = "El arriendo sube menos que el IPC en promedio, pero la diferencia no es significativa. Lo que sí se observa es que sigue la inflación del año anterior, que es la que fija el tope legal del reajuste: casi la mitad de cada punto de IPC previo, más una constante. El simulador reajusta con este modelo sin pasar del IPC previo, y f = 1 queda como sensibilidad. Son 16 observaciones anuales y los años de 2021 a 2023, con la inflación más alta de la muestra, pesan mucho en el ajuste."
    $rm.Range("A8:A47").EntireRow.RowHeight = 15
    $rm.Range("A1").Select()

    # ---- Valorización, datos
    $vd = Nueva-Hoja $dest "Valorización, datos"
    Titulo $vd "Valorización de la vivienda" "Índice de precios de vivienda usada (IPVU), Banco de la República, 1988 a 2025"
    foreach ($k in 1..8) { $vd.Columns.Item($k).ColumnWidth = @(14, 16, 16, 16, 16, 16, 18, 4)[$k - 1] }
    Kpi $vd "A" "Variación nominal media" "=INM!AN5" "0,0 %" "Logarítmica anual, 2001-2025"
    Kpi $vd "C" "Desviación" "=INM!AN6" "0,0 %" "De la variación nominal"
    Kpi $vd "E" "Variación real media" "=INM!AN7" "0,0 %" "Descontada la inflación"
    Seccion $vd 8 "Variación anual del valor"
    $vd.Range("A9").Value2 = "Variación logarítmica anual del IPVU, nominal y real."
    $vd.Range("A9").Font.Color = $gris
    Seccion $vd 30 "Serie anual"
    Encabezado $vd 31 1 7 @("Año", "IPVU nominal", "IPVU real", "r nominal", "r real", "IPC de diciembre", "PIB real (%)")
    for ($k = 0; $k -lt 38; $k++) {
        $r = 32 + $k; $o = 5 + $k
        $vd.Cells.Item($r, 1).Formula = "=INM!AE$o"
        $vd.Cells.Item($r, 2).Formula = "=INM!AF$o"
        $vd.Cells.Item($r, 3).Formula = "=INM!AG$o"
        $vd.Cells.Item($r, 4).Formula = "=IF(INM!AH$o="""","""",INM!AH$o)"
        $vd.Cells.Item($r, 5).Formula = "=IF(INM!AI$o="""","""",INM!AI$o)"
        $vd.Cells.Item($r, 6).Formula = "=IF(INM!AJ$o="""","""",INM!AJ$o)"
        $vd.Cells.Item($r, 7).Formula = "=IF(INM!AL$o="""","""",INM!AL$o)"
    }
    $vd.Range("A32:A69").NumberFormatLocal = "0"
    $vd.Range("A32:A69").HorizontalAlignment = -4131
    $vd.Range("B32:C69").NumberFormatLocal = "#.##0,00"
    $vd.Range("D32:F69").NumberFormatLocal = "0,0 %"
    $vd.Range("G32:G69").NumberFormatLocal = "0,00"
    $vd.Range("B31:G69").HorizontalAlignment = -4152
    Nota $vd 71 "Fuente: Banco de la República (SUAMECA, series 17540 y 17541, descargadas el 30 de octubre de 2026). El IPVU agrega nueve ciudades y sus municipios vecinos; no hay serie solo de Medellín."
    $g = Nuevo-Grafico $vd 10 29 "A" "G" 4
    Nueva-Serie $g "Variación nominal" $vd.Range("D33:D69") $vd.Range("A33:A69") $azul 2 1
    Nueva-Serie $g "Variación real" $vd.Range("E33:E69") $vd.Range("A33:A69") $gris 1.5 4
    $g.Axes(1).TickLabelSpacing = 4
    $g.Axes(1).TickLabelPosition = -4134
    Cerrar-Grafico $g "0%" $true
    $vd.Range("A1").Select()

    # ---- Valorización, modelo
    $vm = Nueva-Hoja $dest "Valorización, modelo"
    Titulo $vm "Valorización: modelo probabilístico" "AR(1) sobre la variación logarítmica anual del IPVU, muestra 2001-2025"
    foreach ($k in 1..12) { $vm.Columns.Item($k).ColumnWidth = @(52, 12, 10, 16, 14, 12, 3, 34, 12, 13, 13, 17)[$k - 1] }
    Kpi $vm "H" "Media de largo plazo (μ)" "=INM!AV30" "0,00 %" ""
    Kpi $vm "I" "φ" "=INM!AW30" "0,00" ""
    Kpi $vm "J" "σ" "=INM!AX30" "0,000" ""

    Bloque $vm 8 1 4 "Evidencia estadística"
    Cabecera $vm 9 1 @("Prueba y contraste", "Estadístico", "p", "Decisión") @(1, 2)
    $pr = @(
        @("t de una muestra: variación nominal media igual a 0", 6),
        @("t de una muestra: variación real media igual a 0", 7),
        @("Jarque-Bera: r normal", 8),
        @("Ljung-Box: sin autocorrelación a 1 año", 9),
        @("ADF: raíz unitaria en r", 10),
        @("Kruskal-Wallis: variación trimestral igual entre trimestres del año", 11),
        @("t del coeficiente: φ igual a 0 en el AR(1)", 12),
        @("Correlación de r con el PIB del mismo año", 13),
        @("Correlación de r con la inflación del año", 14)
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 10 + $k; $o = $pr[$k][1]
        $vm.Cells.Item($r, 1).Value2 = $pr[$k][0]
        $vm.Cells.Item($r, 2).Formula = "=IF(INM!AV$o="""","""",INM!AV$o)"
        $vm.Cells.Item($r, 3).Formula = "=INM!AW$o"
        $vm.Cells.Item($r, 4).Formula = "=IF(C$r<0.05,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $vm.Range("B10:B18").NumberFormatLocal = "0,000"
    $vm.Range("C10:C18").NumberFormatLocal = $pfmt
    $vm.Range("D9:D18").IndentLevel = 1
    $vm.Range("A19").Value2 = "Significancia de 5 %. El ADF tiene poca potencia con 25 datos; el KPSS no rechaza estacionariedad."
    $vm.Range("A20").Value2 = "Ningún año pasa de 2,5 desviaciones: no hay atípicos en 2001-2025."
    $vm.Range("A19:A20").Font.Size = 9; $vm.Range("A19:A20").Font.Color = $gris

    Bloque $vm 22 1 5 "Modelos candidatos"
    Cabecera $vm 23 1 @("Modelo", "log L", "AIC", "BIC", "p") @(1, 2, 3, 4)
    for ($k = 0; $k -lt 5; $k++) {
        $r = 24 + $k; $o = 21 + $k
        $vm.Cells.Item($r, 1).Formula = "=INM!AU$o"
        $vm.Cells.Item($r, 2).Formula = "=IF(INM!AW$o="""","""",INM!AW$o)"
        $vm.Cells.Item($r, 3).Formula = "=IF(INM!AX$o="""","""",INM!AX$o)"
        $vm.Cells.Item($r, 4).Formula = "=IF(INM!AY$o="""","""",INM!AY$o)"
        $vm.Cells.Item($r, 5).Formula = "=IF(INM!AZ$o="""","""",INM!AZ$o)"
    }
    $vm.Range("B24:D28").NumberFormatLocal = "0,0"
    $vm.Range("E24:E28").NumberFormatLocal = "0,000"
    $vm.Range("A29").Value2 = "p de KS o de los residuos. El AR(1) nominal gana por AIC y BIC; sus residuos son normales y sin autocorrelación."
    $vm.Range("A29").Font.Size = 9; $vm.Range("A29").Font.Color = $gris

    Bloque $vm 31 1 4 "Variación según el estado de la economía"
    Cabecera $vm 32 1 @("Crecimiento medio anual del PIB", "Años", "r medio", "Desviación") @(1, 2, 3)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 33 + $k; $o = 37 + $k
        $vm.Cells.Item($r, 1).Formula = "=INM!AU$o"
        $vm.Cells.Item($r, 2).Formula = "=INM!AV$o"
        $vm.Cells.Item($r, 3).Formula = "=INM!AW$o"
        $vm.Cells.Item($r, 4).Formula = "=INM!AX$o"
    }
    $vm.Range("B33:B35").NumberFormatLocal = "0"
    $vm.Range("C33:D35").NumberFormatLocal = "0,0 %"
    $vm.Range("A36").Value2 = "La diferencia entre estados no es significativa (Kruskal-Wallis p = 0,50), aunque el efecto lineal del PIB sí lo es."
    $vm.Range("A36").Font.Size = 9; $vm.Range("A36").Font.Color = $gris

    Bloque $vm 38 1 4 "Sensibilidad: vivienda nueva en Medellín (IPVNBR)"
    Cabecera $vm 39 1 @("Medida", "Valor", "", "") @(1)
    $sen = @(
        @("Años usados, 2005 a 2025", "=INM!AV48", "0"),
        @("Variación nominal media", "=INM!AV49", "0,0 %"),
        @("Desviación estándar", "=INM!AV50", "0,0 %"),
        @("Autocorrelación a 1 año", "=INM!AV51", "0,00"),
        @("Jarque-Bera, p de normalidad", "=INM!AV52", "0,000"),
        @("Correlación con el IPVU nacional", "=INM!AV53", "0,00")
    )
    for ($k = 0; $k -lt $sen.Count; $k++) {
        $r = 40 + $k
        $vm.Cells.Item($r, 1).Value2 = $sen[$k][0]
        $vm.Cells.Item($r, 2).Formula = $sen[$k][1]
        $vm.Cells.Item($r, 2).NumberFormatLocal = $sen[$k][2]
    }
    $vm.Range("B39:B45").HorizontalAlignment = -4152
    $vm.Range("A46").Value2 = "Es vivienda nueva en construcción y se mueve poco con el IPVU: entra como escenario optimista."
    $vm.Range("A46").Font.Size = 9; $vm.Range("A46").Font.Color = $gris

    Bloque $vm 8 8 12 "Especificación"
    $vm.Range("H9").Value2 = "r(t) = μ + φ · (r(t−1) − μ) + e(t),   e ~ Normal(0, σ)"
    $vm.Range("H9").Font.Italic = $true
    Cabecera $vm 10 8 @("Parámetro", "Estimación", "Significado", "", "") @(1)
    $vm.Range("J10").IndentLevel = 1
    $esp = @(
        @("μ", "=INM!AV30", "0,00 %", "Variación media de largo plazo"),
        @("φ", "=INM!AW30", "0,000", "Memoria del año anterior"),
        @("σ", "=INM!AX30", "0,000", "Desviación de la perturbación"),
        @("r(0)", "=INM!AV31", "0,00 %", "Variación de 2025, punto de partida"),
        @("γ", "=INM!AV32", "0,000", "Ajuste por punto de PIB sobre su media"),
        @("Crisis (sensibilidad)", "=INM!AV33", "0,0 %", "Se resta a μ en el escenario adverso")
    )
    for ($k = 0; $k -lt $esp.Count; $k++) {
        $r = 11 + $k
        $vm.Cells.Item($r, 8).Value2 = $esp[$k][0]
        $vm.Cells.Item($r, 9).Formula = $esp[$k][1]
        $vm.Cells.Item($r, 9).NumberFormatLocal = $esp[$k][2]
        $vm.Cells.Item($r, 10).Value2 = $esp[$k][3]
    }
    $vm.Range("J11:J16").IndentLevel = 1
    $vm.Range("H17").Value2 = "El valor del inmueble se acumula como V(t) = V(t−1) · exp(r(t))."
    $vm.Range("H17").Font.Size = 9; $vm.Range("H17").Font.Color = $gris

    Bloque $vm 22 8 12 "Validación"
    Cabecera $vm 23 8 @("Evaluación", "", "", "Dentro del 90 %", "") @(3)
    $vm.Range("H24").Value2 = "Un paso adelante, 2002–2025"
    $vm.Range("K24").Formula = "=INM!AV43"
    $vm.Range("H25").Value2 = "Fuera de muestra, 2016–2025"
    $vm.Range("K25").Formula = "=INM!AV44"
    $vm.Range("K24:K25").HorizontalAlignment = -4152
    $vm.Range("H26").Value2 = "Proporción de años dentro del intervalo predictivo nominal de 90 %."
    $vm.Range("H27").Value2 = "Fuera de muestra: parámetros estimados con 2001–2015."
    $vm.Range("H26:H27").Font.Size = 9; $vm.Range("H26:H27").Font.Color = $gris

    Bloque $vm 29 8 12 "Lectura"
    $lec = $vm.Range("H30:L46")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $vm.Range("H30").Value2 = "La variación nominal es compatible con una normal y depende del año anterior, así que un AR(1) la describe mejor que los modelos de una sola distribución, incluidos los de colas pesadas. El PIB del año aporta un efecto pequeño pero significativo, que entra como ajuste de la media; la inflación del año no se relaciona con la variación nominal, y por eso se simula en términos nominales. Como limitaciones, la serie es corta y de avalúos hipotecarios, que suavizan los precios y exageran la persistencia, y la crisis de 1998-1999 queda fuera de la muestra, de modo que el escenario de crisis es un supuesto. El índice de vivienda nueva de Medellín sube más y se parece poco al nacional."
    $vm.Range("A8:A46").EntireRow.RowHeight = 15
    $vm.Range("A1").Select()

    # ---- Vacancia, modelo
    $va = Nueva-Hoja $dest "Vacancia, modelo"
    Titulo $va "Vacancia: tiempo hasta arrendar" "Curva de supervivencia de Palomma (Kaplan-Meier) y distribución elegida"
    foreach ($k in 1..10) { $va.Columns.Item($k).ColumnWidth = @(30, 14, 14, 14, 14, 14, 3, 40, 22, 22)[$k - 1] }
    Kpi $va "H" "Forma de la gamma" "=INM!BK6" "0,00" ""
    Kpi $va "I" "Media normal (meses)" "=INM!BE17" "0,0" ""
    Kpi $va "J" "Vacancia normal" "=INM!BJ17" "0,0 %" ""
    $va.Range("I4").WrapText = $false

    Bloque $va 8 1 6 "Probabilidad de seguir vacío, S(t)"
    Cabecera $va 9 1 @("Días en oferta", "Publicado", "Exponencial", "Gamma", "Weibull", "Lognormal") @(1, 2, 3, 4, 5)
    for ($k = 0; $k -lt 7; $k++) {
        $r = 10 + $k; $o = 5 + $k
        foreach ($j in 1..6) { $va.Cells.Item($r, $j).Formula = "=INM!" + @("BC", "BD", "BE", "BF", "BG", "BH")[$j - 1] + "$o" }
    }
    $va.Range("A10:A16").NumberFormatLocal = "0"
    $va.Range("A10:A16").HorizontalAlignment = -4131
    $va.Range("B10:F16").NumberFormatLocal = "0,0 %"
    $va.Range("A17").Value2 = "S(t) es la proporción de apartamentos que sigue sin arrendar después de t días."
    $va.Range("A17").Font.Size = 9; $va.Range("A17").Font.Color = $gris
    $g = Nuevo-Grafico $va 19 36 "A" "F" 74
    Nueva-Serie $g "Publicado" $va.Range("B10:B16") $va.Range("A10:A16") $azul 2.5 1
    Nueva-Serie $g "Exponencial" $va.Range("C10:C16") $va.Range("A10:A16") 0x808080 1.25 4
    Nueva-Serie $g "Gamma" $va.Range("D10:D16") $va.Range("A10:A16") 0xB07F4F 1.25 4
    Nueva-Serie $g "Weibull" $va.Range("E10:E16") $va.Range("A10:A16") 0x55A55A 1.25 4
    Nueva-Serie $g "Lognormal" $va.Range("F10:F16") $va.Range("A10:A16") 0x3C6FCF 1.25 4
    $g.Axes(2).MaximumScale = 1
    $g.Axes(1).MaximumScale = 270
    $g.Axes(1).MajorUnit = 30
    Cerrar-Grafico $g "0%" $true

    Bloque $va 38 1 6 "Duración de la vacancia por estado (gamma)"
    Cabecera $va 39 1 @("Estado de la economía", "p no renovar", "Media (meses)", "Escala", "Mediana", "P90") @(1, 2, 3, 4, 5)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 40 + $k; $o = 16 + $k
        $va.Cells.Item($r, 1).Formula = "=INM!BC$o"
        $va.Cells.Item($r, 2).Formula = "=INM!BD$o"
        $va.Cells.Item($r, 3).Formula = "=INM!BE$o"
        $va.Cells.Item($r, 4).Formula = "=INM!BG$o"
        $va.Cells.Item($r, 5).Formula = "=INM!BH$o"
        $va.Cells.Item($r, 6).Formula = "=INM!BI$o"
    }
    $va.Range("B40:B42").NumberFormatLocal = "0 %"
    $va.Range("C40:F42").NumberFormatLocal = "0,0"
    $va.Range("A43").Value2 = "La probabilidad de no renovar y las medias de expansión y crisis son supuestos."
    $va.Range("A44").Value2 = "La media normal se calibra con el 86 % de apartamentos de 45 a 60 m² arrendados en 90 días."
    $va.Range("A43:A44").Font.Size = 9; $va.Range("A43:A44").Font.Color = $gris

    Bloque $va 8 8 10 "Parámetros ajustados"
    Cabecera $va 9 8 @("Distribución", "Parámetro", "Valor") @(2)
    $par = @(
        @("Exponencial", "media (días)", 5, "0,0"), @("Gamma", "forma", 6, "0,00"), @("", "escala (días)", 7, "0,0"),
        @("Weibull", "forma", 8, "0,00"), @("", "escala (días)", 9, "0,0"), @("Lognormal", "μ", 10, "0,00"), @("", "σ", 11, "0,00")
    )
    for ($k = 0; $k -lt $par.Count; $k++) {
        $r = 10 + $k
        $va.Cells.Item($r, 8).Value2 = $par[$k][0]
        $va.Cells.Item($r, 9).Value2 = $par[$k][1]
        $va.Cells.Item($r, 10).Formula = "=INM!BK" + $par[$k][2]
        $va.Cells.Item($r, 10).NumberFormatLocal = $par[$k][3]
    }
    $va.Range("I9:I16").HorizontalAlignment = -4131

    Bloque $va 19 8 10 "Vacancia implícita por estado"
    Cabecera $va 20 8 @("Estado", "Tiempo vacío", "") @(1)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 21 + $k; $o = 16 + $k
        $va.Cells.Item($r, 8).Formula = "=INM!BC$o"
        $va.Cells.Item($r, 9).Formula = "=INM!BJ$o"
    }
    $va.Range("I21:I23").NumberFormatLocal = "0,0 %"
    $va.Range("H24").Value2 = "Fracción del tiempo que el apartamento queda sin arrendar."
    $va.Range("H24").Font.Size = 9; $va.Range("H24").Font.Color = $gris

    Bloque $va 26 8 10 "Arrendados en 90 días, por segmento"
    Cabecera $va 27 8 @("Segmento", "En 90 días", "Media (meses)") @(1, 2)
    for ($k = 0; $k -lt 12; $k++) {
        $r = 28 + $k; $o = 25 + $k
        $va.Cells.Item($r, 8).Formula = "=IF(INM!BC$o=""tamano"",""Tamaño"",IF(INM!BC$o=""canon"",""Canon"",""Canon, menos de 60 m²""))&"": ""&SUBSTITUTE(SUBSTITUTE(INM!BD$o,""Mas de"",""Más de""),""m2"",""m²"")"
        $va.Cells.Item($r, 9).Formula = "=INM!BE$o"
        $va.Cells.Item($r, 10).Formula = "=INM!BF$o"
    }
    $va.Range("I28:I39").NumberFormatLocal = "0 %"
    $va.Range("J28:J39").NumberFormatLocal = "0,0"
    $va.Range("I27:J39").HorizontalAlignment = -4152

    Bloque $va 41 8 10 "Lectura"
    $lec = $va.Range("H42:J52")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $va.Range("H42").Value2 = "El tiempo hasta arrendar se ajusta bien con una gamma de forma cercana a uno, casi una exponencial: la probabilidad de seguir vacío cae de manera regular. Las cuatro distribuciones reproducen la curva publicada con diferencias pequeñas. Lo que más pesa en la simulación son los supuestos por estado de la economía, no el ajuste, y la vacancia implícita pasa de menos de 1 % en expansión a más de 10 % en crisis."
    $va.Range("A8:A52").EntireRow.RowHeight = 15
    $va.Range("A1").Select()

    # ------------------------------------------------ Markov, modelo
    $mk2 = Nueva-Hoja $dest "Markov, modelo"
    Titulo $mk2 "Estados de la economía: cadena de Markov" "PIB real trimestral, 2006-2026, con tres estados: expansión, normal y crisis"
    foreach ($k in 1..12) { $mk2.Columns.Item($k).ColumnWidth = @(52, 14, 18, 18, 16, 3, 3, 40, 14, 22, 22, 4, 8, 12, 12, 12)[$k - 1] }
    Kpi $mk2 "H" "Permanencia en crisis" "=MARKOV!D41" "0 %" "Probabilidad de seguir en crisis"
    Kpi $mk2 "J" "Duración de la crisis" "=MARKOV!E41" "0,0" "Trimestres, en promedio"
    Kpi $mk2 "K" "Crisis a largo plazo" "=MARKOV!D146" "0 %" "Del tiempo, según la cadena"
    $pfmt = '[<0,001]"<0,001";0,000'

    Bloque $mk2 8 1 5 "Los tres estados"
    Cabecera $mk2 9 1 @("Estado y definición", "Trimestres", "Proporción", "", "") @(1, 2)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 10 + $k; $o = 11 + $k
        $mk2.Cells.Item($r, 1).Formula = "=MARKOV!A$o&"": ""&LOWER(MARKOV!B$o)"
        $mk2.Cells.Item($r, 2).Formula = "=MARKOV!C$o"
        $mk2.Cells.Item($r, 3).Formula = "=MARKOV!D$o"
    }
    $mk2.Range("B10:B12").NumberFormatLocal = "0"
    $mk2.Range("C10:C12").NumberFormatLocal = "0 %"
    $mk2.Range("A13").Value2 = "Con tres estados cada uno tiene datos suficientes; con cinco, algunos tendrían dos o tres trimestres."
    $mk2.Range("A14").Value2 = "La crisis incluye el estancamiento: con un umbral de 0 % solo quedarían cuatro trimestres."
    $mk2.Range("A13:A14").Font.Size = 9; $mk2.Range("A13:A14").Font.Color = $gris

    Bloque $mk2 16 1 5 "Transiciones observadas"
    Cabecera $mk2 17 1 @("Desde \ hacia", "Expansión", "Normal", "Crisis", "Total") @(1, 2, 3, 4)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 18 + $k; $o = 33 + $k
        foreach ($j in 1..5) { $mk2.Cells.Item($r, $j).Formula = "=MARKOV!" + [char](64 + $j) + $o }
    }
    $mk2.Range("B18:E20").NumberFormatLocal = "0"

    Bloque $mk2 22 1 5 "Matriz de transición"
    Cabecera $mk2 23 1 @("Desde \ hacia", "Expansión", "Normal", "Crisis", "Duración (trim.)") @(1, 2, 3, 4)
    for ($k = 0; $k -lt 3; $k++) {
        $r = 24 + $k; $o = 39 + $k
        foreach ($j in 1..5) { $mk2.Cells.Item($r, $j).Formula = "=MARKOV!" + [char](64 + $j) + $o }
    }
    $mk2.Range("B24:D26").NumberFormatLocal = "0,0 %"
    $mk2.Range("E24:E26").NumberFormatLocal = "0,0"
    $mk2.Range("A27").Value2 = "Probabilidad de pasar de un estado a otro en un trimestre. De expansión nunca se pasó directo a crisis."
    $mk2.Range("A27").Font.Size = 9; $mk2.Range("A27").Font.Color = $gris

    Bloque $mk2 29 1 4 "Evidencia estadística"
    Cabecera $mk2 30 1 @("Prueba y contraste", "Estadístico", "p", "Decisión") @(1, 2)
    $pr = @(
        @("Chi-cuadrado: el estado no depende del trimestre anterior", "=MARKOV!C50", "=MARKOV!C52"),
        @("ANOVA: el rendimiento del ETF es igual en los tres estados", "=MARKOV!C63", "=MARKOV!C64"),
        @("F de varianzas: la volatilidad del ETF en crisis no es mayor", "=MARKOV!C67", "=MARKOV!C68"),
        @("t de Welch: la vivienda nueva no sube más en expansión", "=MARKOV!C80", "=MARKOV!C82")
    )
    for ($k = 0; $k -lt $pr.Count; $k++) {
        $r = 31 + $k
        $mk2.Cells.Item($r, 1).Value2 = $pr[$k][0]
        $mk2.Cells.Item($r, 2).Formula = $pr[$k][1]
        $mk2.Cells.Item($r, 3).Formula = $pr[$k][2]
        $mk2.Cells.Item($r, 4).Formula = "=IF(C$r<MARKOV!`$C`$7,""Rechaza H₀"",""No rechaza H₀"")"
    }
    $mk2.Range("B31:B34").NumberFormatLocal = "0,000"
    $mk2.Range("C31:C34").NumberFormatLocal = $pfmt
    $mk2.Range("D30:D34").IndentLevel = 1
    $mk2.Range("A35").Value2 = "Significancia de 5 %. Las pruebas de varianza y de vivienda son de una cola."
    $mk2.Range("A35").Font.Size = 9; $mk2.Range("A35").Font.Color = $gris

    Bloque $mk2 37 1 4 "Efecto del estado en el ETF"
    Cabecera $mk2 38 1 @("Estado", "Meses", "Rendimiento medio", "Desviación") @(1, 2, 3)
    for ($k = 0; $k -lt 4; $k++) {
        $r = 39 + $k; $o = 57 + $k
        $mk2.Cells.Item($r, 1).Formula = "=MARKOV!A$o"
        $mk2.Cells.Item($r, 2).Formula = "=MARKOV!B$o"
        $mk2.Cells.Item($r, 3).Formula = "=MARKOV!C$o"
        $mk2.Cells.Item($r, 4).Formula = "=MARKOV!D$o"
    }
    $mk2.Range("B39:B42").NumberFormatLocal = "0"
    $mk2.Range("C39:D42").NumberFormatLocal = "0,00 %"
    $mk2.Range("A43").Value2 = "El estado cambia el riesgo del ETF, no su rendimiento esperado."
    $mk2.Range("A43").Font.Size = 9; $mk2.Range("A43").Font.Color = $gris

    Bloque $mk2 45 1 4 "Efecto del estado en la vivienda nueva de Medellín"
    Cabecera $mk2 46 1 @("Estado", "Años", "Valorización media", "Desviación") @(1, 2, 3)
    for ($k = 0; $k -lt 4; $k++) {
        $r = 47 + $k; $o = 74 + $k
        $mk2.Cells.Item($r, 1).Formula = "=MARKOV!A$o"
        $mk2.Cells.Item($r, 2).Formula = "=MARKOV!B$o"
        $mk2.Cells.Item($r, 3).Formula = "=MARKOV!C$o"
        $mk2.Cells.Item($r, 4).Formula = "=MARKOV!D$o"
    }
    $mk2.Range("B47:B50").NumberFormatLocal = "0"
    $mk2.Range("C47:D50").NumberFormatLocal = "0,0 %"
    $mk2.Range("A51").Value2 = "Es vivienda nueva. Con el IPVU, que describe el inmueble del modelo, el estado no cambia la valorización."
    $mk2.Range("A51").Font.Size = 9; $mk2.Range("A51").Font.Color = $gris

    Bloque $mk2 8 8 11 "Qué cambia en el simulador"
    Cabecera $mk2 9 8 @("Variable", "Expansión", "Normal", "Crisis") @(1, 2, 3)
    $cam = @(
        @("Volatilidad mensual del ETF", 88, "0,00 %"),
        @("Rendimiento medio mensual del ETF", 89, "0,00 %"),
        @("Vacancia: probabilidad de no renovar", 92, "0 %"),
        @("Vacancia: duración media (meses)", 93, "0,0")
    )
    for ($k = 0; $k -lt $cam.Count; $k++) {
        $r = 10 + $k; $o = $cam[$k][1]
        $mk2.Cells.Item($r, 8).Value2 = $cam[$k][0]
        foreach ($j in 1..3) {
            $mk2.Cells.Item($r, 8 + $j).Formula = "=MARKOV!" + [char](66 + $j) + $o
            $mk2.Cells.Item($r, 8 + $j).NumberFormatLocal = $cam[$k][2]
        }
    }
    $mk2.Range("I10:K13").HorizontalAlignment = -4152
    $mk2.Range("H14").Value2 = "También cambian la frecuencia de las necesidades de liquidez y el reparto objetivo de E4."
    $mk2.Range("H15").Value2 = "La valorización, la inflación y la TPM no dependen del estado."
    $mk2.Range("H14:H15").Font.Size = 9; $mk2.Range("H14:H15").Font.Color = $gris

    Bloque $mk2 17 8 11 "Convergencia de la cadena"
    $mk2.Range("M17").Value2 = "Trim."
    $mk2.Range("N17").Value2 = "Expansión"
    $mk2.Range("O17").Value2 = "Normal"
    $mk2.Range("P17").Value2 = "Crisis"
    for ($k = 0; $k -lt 41; $k++) {
        $r = 18 + $k; $o = 105 + $k
        $mk2.Cells.Item($r, 13).Formula = "=MARKOV!A$o"
        $mk2.Cells.Item($r, 14).Formula = "=MARKOV!B$o"
        $mk2.Cells.Item($r, 15).Formula = "=MARKOV!C$o"
        $mk2.Cells.Item($r, 16).Formula = "=MARKOV!D$o"
    }
    $mk2.Range("M17:P58").Font.Color = $gris
    $mk2.Range("N18:P58").NumberFormatLocal = "0,0 %"
    $g = Nuevo-Grafico $mk2 18 33 "H" "K" 4
    Nueva-Serie $g "Expansión" $mk2.Range("N18:N58") $mk2.Range("M18:M58") $azul 2 1
    Nueva-Serie $g "Normal" $mk2.Range("O18:O58") $mk2.Range("M18:M58") 0x808080 2 1
    Nueva-Serie $g "Crisis" $mk2.Range("P18:P58") $mk2.Range("M18:M58") 0x3C6FCF 2 1
    $g.Axes(1).TickLabelSpacing = 5
    Cerrar-Grafico $g "0%" $true
    Cabecera $mk2 34 8 @("Largo plazo", "Expansión", "Normal", "Crisis") @(1, 2, 3)
    $mk2.Range("H35").Value2 = "Según la cadena"
    $mk2.Range("H36").Value2 = "Observado 2006–2026"
    foreach ($j in 1..3) {
        $mk2.Cells.Item(35, 8 + $j).Formula = "=MARKOV!" + [char](65 + $j) + "146"
        $mk2.Cells.Item(36, 8 + $j).Formula = "=MARKOV!" + [char](65 + $j) + "147"
    }
    $mk2.Range("I35:K36").NumberFormatLocal = "0,0 %"
    $mk2.Range("I34:K36").HorizontalAlignment = -4152

    Bloque $mk2 38 8 11 "Lectura"
    $lec = $mk2.Range("H39:K45")
    $lec.Merge()
    $lec.WrapText = $true
    $lec.VerticalAlignment = -4160
    $mk2.Range("H39").Value2 = "Los estados persisten: una economía en crisis tiene tres probabilidades de cuatro de seguir en ella el trimestre siguiente, y por eso se usa una cadena de Markov y no un sorteo independiente cada trimestre. El estado cambia el riesgo del ETF, no su rendimiento esperado, y encarece el arriendo y los gastos imprevistos al mismo tiempo. A largo plazo la cadena reproduce la proporción de tiempo y las duraciones observadas en cada estado."

    Bloque $mk2 47 8 11 "Limitaciones"
    $lim = $mk2.Range("H48:K54")
    $lim.Merge()
    $lim.WrapText = $true
    $lim.VerticalAlignment = -4160
    $mk2.Range("H48").Value2 = "Solo hay tres episodios de crisis, 12 trimestres, y las probabilidades de entrar y salir de ella tienen incertidumbre; la sensibilidad duplica la probabilidad de entrar. La vacancia por estado se apoya en supuestos y no en datos históricos. El rebote de 2021 y 2022 cuenta como expansión aunque fue una recuperación de la pandemia."
    $mk2.Range("A8:A54").EntireRow.RowHeight = 15
    $mk2.Range("A1").Select()

    # ------------------------------------------------ Validación y resultados
    $naranja = 0x3C6FCF
    function Esc-Txt($ref) { return 'IF(' + $ref + '="base","Base",IF(' + $ref + '="adverso","Adverso","Estrés"))' }

    # ---- Validación
    $vl = Nueva-Hoja $dest "Validación del simulador"
    Titulo $vl "Validación del simulador" "500 réplicas con los parámetros de referencia, frente a los datos de cada proceso"
    foreach ($k in 1..7) { $vl.Columns.Item($k).ColumnWidth = @(14, 46, 20, 14, 14, 34, 70)[$k - 1] }
    Kpi $vl "A" "Medidas comparadas" '=COUNTA(VALIDACION!B5:B24)' "0" "Seis procesos"
    Kpi $vl "C" "Reproducen los datos de estimación" '=COUNTIF(VALIDACION!C5:C24,"Reproduce insumos")' "0" "Verifican el generador"
    Kpi $vl "F" "Contrastes" '=COUNTIF(VALIDACION!C5:C24,"Contraste")' "0" "Con algo que el modelo no usó"
    Seccion $vl 8 "Simulado frente a observado"
    Encabezado $vl 9 1 7 @("Proceso", "Medida", "Tipo", "Simulado", "Observado", "Referencia", "Lectura")
    $fm = @("0,00", "@", "0,000", "0,00", "0,00", "0,00", "0,0 %", "0,00 %", "0,00 %", "0,00 %", "0,00", "0,00", "0,0 %", "0,0 %", "0,0 %", "0,0 %", "0,0", "0,00 %", "0,00 %", "0,0 %")
    for ($k = 0; $k -lt 20; $k++) {
        $r = 10 + $k; $o = 5 + $k; $p = $o - 1
        $vl.Cells.Item($r, 1).Formula = '=IF(VALIDACION!A' + $o + '=VALIDACION!A' + $p + ',"",VALIDACION!A' + $o + ')'
        $vl.Cells.Item($r, 2).Formula = '=VALIDACION!B' + $o
        $vl.Cells.Item($r, 3).Formula = '=IF(VALIDACION!C' + $o + '="Contraste","Contraste","Reproduce los datos")'
        if ($k -eq 1) {
            $vl.Cells.Item($r, 4).Formula = '=SUBSTITUTE(VALIDACION!D' + $o + ',".",",")'
            $vl.Cells.Item($r, 5).Formula = '=SUBSTITUTE(VALIDACION!E' + $o + ',".",",")'
        } else {
            $vl.Cells.Item($r, 4).Formula = '=VALIDACION!D' + $o
            $vl.Cells.Item($r, 5).Formula = '=VALIDACION!E' + $o
            $vl.Range($vl.Cells.Item($r, 4), $vl.Cells.Item($r, 5)).NumberFormatLocal = $fm[$k]
        }
        $vl.Cells.Item($r, 6).Formula = '=VALIDACION!F' + $o
        $vl.Cells.Item($r, 7).Formula = '=IF(VALIDACION!G' + $o + '="","",VALIDACION!G' + $o + ')'
    }
    $vl.Range("D9:E29").HorizontalAlignment = -4152
    $vl.Range("A10:A29").Font.Bold = $true
    $vl.Range("C9:C29").IndentLevel = 1
    $vl.Range("F9:G29").IndentLevel = 2
    $vl.Range("F10:G29").WrapText = $true
    $vl.Range("A10:G29").VerticalAlignment = -4160
    $vl.Range("A10:G29").Borders.Item(12).LineStyle = 1
    $vl.Range("A10:G29").Borders.Item(12).Color = 0xE0E0E0
    $vl.Range("A10:G29").Borders.Item(9).LineStyle = 1
    $vl.Range("A10:G29").Borders.Item(9).Color = 0xE0E0E0
    $vl.Range("A10:G29").Rows.AutoFit() | Out-Null
    Nota $vl 31 "Reproduce los datos: compara con los mismos datos con que se estimó el proceso. Contraste: compara con algo que el modelo no usó al estimarse."
    Nota $vl 32 "Inflación y TPM en porcentaje; rendimiento del ETF y valorización, en logaritmo; vacancia, fracción del tiempo desocupado."
    $vl.Range("A1").Select()

    # ---- Resultados
    $rs = Nueva-Hoja $dest "Resultados, estrategias"
    Titulo $rs "Resultados de las estrategias" "1.068 réplicas por estrategia y escenario, con los parámetros de referencia. Montos en millones de pesos de hoy."
    foreach ($k in 1..12) { $rs.Columns.Item($k).ColumnWidth = @(14, 12, 12, 16, 15, 14, 14, 14, 14, 14, 14, 14)[$k - 1] }
    Kpi $rs "A" "Mayor P(meta), base" '=MAX(RESULTADOS!C17:C20)' "0,0 %" ""
    $rs.Range("A6").Formula = '="Estrategia "&INDEX(RESULTADOS!B17:B20,MATCH(A5,RESULTADOS!C17:C20,0))'
    Kpi $rs "D" "Mayor P(meta), adverso" '=MAX(RESULTADOS!C21:C24)' "0,0 %" ""
    $rs.Range("D6").Formula = '="Estrategia "&INDEX(RESULTADOS!B21:B24,MATCH(D5,RESULTADOS!C21:C24,0))'
    Kpi $rs "G" "Mayor P(meta), estrés" '=MAX(RESULTADOS!C25:C28)' "0,0 %" ""
    $rs.Range("G6").Formula = '="Estrategia "&INDEX(RESULTADOS!B25:B28,MATCH(G5,RESULTADOS!C25:C28,0))'
    foreach ($z in @("A6", "D6", "G6")) { $rs.Range($z).Font.Size = 9; $rs.Range($z).Font.Color = $gris }
    Kpi $rs "J" "Réplicas por configuración" "=1068" "#.##0" "4 estrategias en 3 escenarios"
    $rs.Range("J6").Font.Size = 9; $rs.Range("J6").Font.Color = $gris

    Seccion $rs 8 "Probabilidad de alcanzar la meta"
    $rs.Range("A9").Value2 = "Con la configuración de referencia, ninguna estrategia pasa de 20 % en el escenario base, y en el adverso y el de estrés queda por debajo de 3 %."
    $rs.Range("A9").Font.Color = $gris
    Encabezado $rs 11 1 4 @("Estrategia", "Base", "Adverso", "Estrés")
    for ($k = 0; $k -lt 4; $k++) {
        $r = 12 + $k; $o = 17 + $k
        $rs.Cells.Item($r, 1).Formula = "=RESULTADOS!B$o"
        $rs.Cells.Item($r, 2).Formula = "=RESULTADOS!C$o"
        $rs.Cells.Item($r, 3).Formula = "=RESULTADOS!C" + ($o + 4)
        $rs.Cells.Item($r, 4).Formula = "=RESULTADOS!C" + ($o + 8)
    }
    $rs.Range("B12:D15").NumberFormatLocal = "0,0 %"
    $rs.Range("B11:D15").HorizontalAlignment = -4152
    $g = Nuevo-Grafico $rs 17 33 "A" "H" 51
    Nueva-Serie $g "Base" $rs.Range("B12:B15") $rs.Range("A12:A15") $azul 0 0
    Nueva-Serie $g "Adverso" $rs.Range("C12:C15") $rs.Range("A12:A15") 0x808080 0 0
    Nueva-Serie $g "Estrés" $rs.Range("D12:D15") $rs.Range("A12:A15") $naranja 0 0
    $g.ChartGroups(1).GapWidth = 80
    Cerrar-Grafico $g "0%" $true

    Seccion $rs 35 "Indicadores por estrategia y escenario"
    $rs.Range("A36").Value2 = "Patrimonio real al final del horizonte, en millones de pesos de hoy. CVaR 5 %: promedio del 5 % peor de los casos."
    $rs.Range("A36").Font.Color = $gris
    Encabezado $rs 37 1 12 @("Escenario", "Estrategia", "P(meta)", "IC de 95 %", "Patrimonio P5", "Patrimonio P50", "Patrimonio P95", "CVaR 5 %", "P(pérdida real)", "Caída máx. P50", "P(iliquidez)", "Años a la meta")
    for ($k = 0; $k -lt 12; $k++) {
        $r = 38 + $k; $o = 17 + $k
        $rs.Cells.Item($r, 1).Formula = '=' + (Esc-Txt "RESULTADOS!A$o")
        $rs.Cells.Item($r, 2).Formula = "=RESULTADOS!B$o"
        $rs.Cells.Item($r, 3).Formula = "=RESULTADOS!C$o"
        $rs.Cells.Item($r, 4).Formula = '=ROUND(RESULTADOS!D' + $o + '*100,1)&" a "&ROUND(RESULTADOS!E' + $o + '*100,1)&" %"'
        $rs.Cells.Item($r, 5).Formula = "=RESULTADOS!G$o"
        $rs.Cells.Item($r, 6).Formula = "=RESULTADOS!I$o"
        $rs.Cells.Item($r, 7).Formula = "=RESULTADOS!K$o"
        $rs.Cells.Item($r, 8).Formula = "=RESULTADOS!L$o"
        $rs.Cells.Item($r, 9).Formula = "=RESULTADOS!N$o"
        $rs.Cells.Item($r, 10).Formula = "=RESULTADOS!O$o"
        $rs.Cells.Item($r, 11).Formula = "=RESULTADOS!P$o"
        $rs.Cells.Item($r, 12).Formula = "=RESULTADOS!T$o"
    }
    $rs.Range("C38:C49").NumberFormatLocal = "0,0 %"
    $rs.Range("E38:H49").NumberFormatLocal = "#.##0"
    $rs.Range("I38:K49").NumberFormatLocal = "0,0 %"
    $rs.Range("L38:L49").NumberFormatLocal = "0,0"
    $rs.Range("C37:L49").HorizontalAlignment = -4152
    foreach ($k in @(41, 45)) { $rs.Range($rs.Cells.Item($k, 1), $rs.Cells.Item($k, 12)).Borders.Item(9).LineStyle = 1; $rs.Range($rs.Cells.Item($k, 1), $rs.Cells.Item($k, 12)).Borders.Item(9).Color = 0xC0C0C0 }
    $rs.Range("A38:A49").Font.Bold = $true

    Seccion $rs 52 "Capital necesario para alcanzar la meta"
    $rs.Range("A53").Value2 = "Capital inicial mínimo, en millones, para llegar a cada probabilidad. «No alcanza» indica que queda fuera del rango de 400 a 1.000 millones."
    $rs.Range("A53").Font.Color = $gris
    $rs.Range("C55").Value2 = "Para 80 % de probabilidad"
    $rs.Range("H55").Value2 = "Para 50 % de probabilidad"
    $rs.Range("C55").Font.Bold = $true; $rs.Range("H55").Font.Bold = $true
    Encabezado $rs 56 1 12 @("Escenario", "Estrategia", "", "", "", "", "", "", "", "", "", "")
    for ($j = 0; $j -lt 5; $j++) {
        $rs.Cells.Item(56, 3 + $j).Formula = '="Meta "&TEXT(RESULTADOS!C' + (32 + $j) + ',"#.##0")'
        $rs.Cells.Item(56, 8 + $j).Formula = '="Meta "&TEXT(RESULTADOS!C' + (32 + $j) + ',"#.##0")'
    }
    for ($gI = 0; $gI -lt 8; $gI++) {
        $r = 57 + $gI; $o = 32 + 5 * $gI
        $rs.Cells.Item($r, 1).Formula = '=' + (Esc-Txt "RESULTADOS!A$o")
        $rs.Cells.Item($r, 2).Formula = "=RESULTADOS!B$o"
        for ($j = 0; $j -lt 5; $j++) {
            $rs.Cells.Item($r, 3 + $j).Formula = '=IF(RESULTADOS!D' + ($o + $j) + '="","no alcanza",RESULTADOS!D' + ($o + $j) + ')'
            $rs.Cells.Item($r, 8 + $j).Formula = '=IF(RESULTADOS!E' + ($o + $j) + '="","no alcanza",RESULTADOS!E' + ($o + $j) + ')'
        }
    }
    $rs.Range("C57:L64").NumberFormatLocal = "#.##0"
    $rs.Range("C56:L64").HorizontalAlignment = -4152
    $rs.Range("A57:A64").Font.Bold = $true
    $rs.Range($rs.Cells.Item(60, 1), $rs.Cells.Item(60, 12)).Borders.Item(9).LineStyle = 1
    $rs.Range($rs.Cells.Item(60, 1), $rs.Cells.Item(60, 12)).Borders.Item(9).Color = 0xC0C0C0
    Nota $rs 65 "No se corrió el escenario adverso en esta malla: 500 réplicas por celda, 7 capitales de 400 a 1.000 millones y 5 metas de 700 a 1.500 millones."

    Seccion $rs 67 "Diseño del experimento"
    for ($k = 0; $k -lt 8; $k++) {
        $r = 68 + $k; $o = 6 + $k
        $rs.Cells.Item($r, 1).Formula = "=RESULTADOS!A$o"
        $rs.Cells.Item($r, 1).Font.Bold = $true
        $rs.Range($rs.Cells.Item($r, 3), $rs.Cells.Item($r, 12)).Merge()
        $rs.Cells.Item($r, 3).Formula = "=RESULTADOS!B$o"
        $rs.Cells.Item($r, 3).WrapText = $true
        $rs.Cells.Item($r, 3).VerticalAlignment = -4160
        $rs.Cells.Item($r, 1).VerticalAlignment = -4160
        $rs.Rows.Item($r).RowHeight = 30
    }
    $rs.Range("A1").Select()

    # ---- Comparación
    $cp = Nueva-Hoja $dest "Comparación de estrategias"
    Titulo $cp "Comparación entre estrategias" "Diferencia a − b con los mismos números aleatorios en cada réplica; significancia con corrección de Holm"
    foreach ($k in 1..8) { $cp.Columns.Item($k).ColumnWidth = @(14, 16, 22, 16, 22, 16, 22, 16)[$k - 1] }
    Kpi $cp "A" "Comparaciones significativas" '=COUNTIF(RESULTADOS!J82:J153,TRUE)+COUNTIF(RESULTADOS!J82:J153,"True")' "0" "De 72"
    Seccion $cp 8 "Diferencias por escenario"
    $cp.Range("A9").Value2 = "Patrimonio real en millones de pesos de hoy; probabilidad de la meta y caída máxima en puntos porcentuales."
    $cp.Range("A9").Font.Color = $gris
    Encabezado $cp 10 1 8 @("Escenario", "Comparación", "Dif. patrimonio real", "Significativa", "Dif. P(meta)", "Significativa", "Dif. caída máxima", "Significativa")
    for ($gI = 0; $gI -lt 18; $gI++) {
        $r = 11 + $gI; $o = 82 + 4 * $gI
        $cp.Cells.Item($r, 1).Formula = '=' + (Esc-Txt "RESULTADOS!A$o")
        $cp.Cells.Item($r, 2).Formula = '=RESULTADOS!B' + $o + '&" − "&RESULTADOS!C' + $o
        $cp.Cells.Item($r, 3).Formula = '=RESULTADOS!F' + $o + '/1000000'
        $cp.Cells.Item($r, 4).Formula = '=IF(OR(RESULTADOS!J' + $o + '=TRUE,RESULTADOS!J' + $o + '="True"),"Sí","No")'
        $cp.Cells.Item($r, 5).Formula = '=RESULTADOS!F' + ($o + 1) + '*100'
        $cp.Cells.Item($r, 6).Formula = '=IF(OR(RESULTADOS!J' + ($o + 1) + '=TRUE,RESULTADOS!J' + ($o + 1) + '="True"),"Sí","No")'
        $cp.Cells.Item($r, 7).Formula = '=RESULTADOS!F' + ($o + 3) + '*100'
        $cp.Cells.Item($r, 8).Formula = '=IF(OR(RESULTADOS!J' + ($o + 3) + '=TRUE,RESULTADOS!J' + ($o + 3) + '="True"),"Sí","No")'
    }
    $cp.Range("C11:C28").NumberFormatLocal = '+#.##0,0;-#.##0,0;0,0'
    $cp.Range("E11:E28").NumberFormatLocal = '+0,0;-0,0;0,0'
    $cp.Range("G11:G28").NumberFormatLocal = '+0,0;-0,0;0,0'
    $cp.Range("C10:H28").HorizontalAlignment = -4152
    $cp.Range("A11:A28").Font.Bold = $true
    foreach ($k in @(16, 22)) { $cp.Range($cp.Cells.Item($k, 1), $cp.Cells.Item($k, 8)).Borders.Item(9).LineStyle = 1; $cp.Range($cp.Cells.Item($k, 1), $cp.Cells.Item($k, 8)).Borders.Item(9).Color = 0xC0C0C0 }
    Nota $cp 29 "La primera estrategia de cada comparación es a y la segunda es b. Un signo positivo indica que a supera a b."

    Seccion $cp 31 "Precisión de P(meta) según el número de réplicas"
    $cp.Range("A32").Value2 = "Estrategia E2, escenario base."
    $cp.Range("A32").Font.Color = $gris
    Encabezado $cp 33 1 5 @("Réplicas", "P(meta)", "IC de 95 %", "Semiancho (pp)", "Semiancho máximo (pp)")
    for ($k = 0; $k -lt 4; $k++) {
        $r = 34 + $k; $o = 75 + $k
        $cp.Cells.Item($r, 1).Formula = "=RESULTADOS!A$o"
        $cp.Cells.Item($r, 2).Formula = "=RESULTADOS!B$o"
        $cp.Cells.Item($r, 3).Formula = '=ROUND(RESULTADOS!C' + $o + '*100,1)&" a "&ROUND(RESULTADOS!D' + $o + '*100,1)&" %"'
        $cp.Cells.Item($r, 4).Formula = "=RESULTADOS!E$o*100"
        $cp.Cells.Item($r, 5).Formula = "=RESULTADOS!F$o*100"
    }
    $cp.Range("A34:A37").NumberFormatLocal = "#.##0"
    $cp.Range("A34:A37").HorizontalAlignment = -4131
    $cp.Range("B34:B37").NumberFormatLocal = "0,0 %"
    $cp.Range("D34:E37").NumberFormatLocal = "0,0"
    $cp.Range("B33:E37").HorizontalAlignment = -4152
    $cp.Range("A1").Select()

    # ---- Sensibilidad
    $sn = Nueva-Hoja $dest "Sensibilidad"
    Titulo $sn "Sensibilidad" "Estrategia E2, escenario base: P(meta) al mover un factor a la vez, 400 réplicas por punto"
    foreach ($k in 1..6) { $sn.Columns.Item($k).ColumnWidth = @(84, 22, 22, 22, 22, 22)[$k - 1] }
    Kpi $sn "A" "P(meta) de referencia" '=RESULTADOS!C157' "0,0 %" "Con todos los factores en su valor base"
    Kpi $sn "B" "Mayor rango" '=RESULTADOS!F157' '0,0" puntos"' ""
    $sn.Range("B6").Formula = '=RESULTADOS!B157'
    $sn.Range("B6").Font.Size = 9; $sn.Range("B6").Font.Color = $gris
    Seccion $sn 8 "Factores que más mueven la probabilidad"
    $sn.Range("A9").Value2 = "Rango de P(meta), en puntos porcentuales, entre el valor bajo y el alto de cada factor. Los doce mayores."
    $sn.Range("A9").Font.Color = $gris
    Seccion $sn 32 "Todos los factores"
    Encabezado $sn 33 1 6 @("Factor", "Fuente de incertidumbre", "P(meta) referencia", "Valor bajo", "Valor alto", "Rango (puntos)")
    for ($k = 0; $k -lt 20; $k++) {
        $r = 34 + $k; $o = 157 + $k
        $sn.Cells.Item($r, 1).Formula = "=RESULTADOS!A$o"
        $sn.Cells.Item($r, 2).Formula = "=RESULTADOS!B$o"
        $sn.Cells.Item($r, 3).Formula = "=RESULTADOS!C$o"
        $sn.Cells.Item($r, 4).Formula = "=RESULTADOS!D$o"
        $sn.Cells.Item($r, 5).Formula = "=RESULTADOS!E$o"
        $sn.Cells.Item($r, 6).Formula = "=RESULTADOS!F$o"
    }
    $sn.Range("C34:E53").NumberFormatLocal = "0,0 %"
    $sn.Range("F34:F53").NumberFormatLocal = "0,0"
    $sn.Range("C33:F53").HorizontalAlignment = -4152
    $sn.Range("B33:B53").IndentLevel = 1
    $g = Nuevo-Grafico $sn 10 30 "A" "F" 57
    Nueva-Serie $g "Rango" $sn.Range("F34:F45") $sn.Range("A34:A45") $azul 0 0
    $g.Axes(1).ReversePlotOrder = $true
    $g.ChartGroups(1).GapWidth = 40
    Cerrar-Grafico $g "0" $false

    Seccion $sn 56 "Sensibilidad a la elección de modelos"
    $sn.Range("A57").Value2 = "P(meta) con el modelo de referencia y con la alternativa, por estrategia. 400 réplicas."
    $sn.Range("A57").Font.Color = $gris
    Encabezado $sn 58 1 5 @("Modelo alternativo", "E1", "E2", "E3", "E4")
    for ($mI = 0; $mI -lt 8; $mI++) {
        $r = 59 + $mI; $o = 180 + 4 * $mI
        $sn.Cells.Item($r, 1).Formula = "=RESULTADOS!A$o"
        foreach ($j in 0..3) {
            $q = $o + $j
            $sn.Cells.Item($r, 2 + $j).Formula = '=ROUND(RESULTADOS!C' + $q + '*100,1)&" a "&ROUND(RESULTADOS!D' + $q + '*100,1)&" %"'
        }
    }
    $sn.Range("B58:E66").HorizontalAlignment = -4152
    Nota $sn 67 "Cada celda muestra la probabilidad con el modelo de referencia y con la alternativa."

    Seccion $sn 69 "Interacción entre inflación y necesidades de liquidez"
    $sn.Range("A70").Value2 = "P(meta) con inflación de 3,5 % o 6,5 % y necesidades de liquidez de 0,5 o 1,5 veces la referencia."
    $sn.Range("A70").Font.Color = $gris
    Encabezado $sn 71 1 6 @("Estrategia", "", "", "", "", "Interacción (puntos)")
    foreach ($j in 0..3) { $sn.Cells.Item(71, 2 + $j).Formula = '=RESULTADOS!' + [char](66 + $j) + '214' }
    for ($k = 0; $k -lt 4; $k++) {
        $r = 72 + $k; $o = 215 + $k
        foreach ($j in 1..6) { $sn.Cells.Item($r, $j).Formula = '=RESULTADOS!' + [char](64 + $j) + $o }
    }
    $sn.Range("B72:E75").NumberFormatLocal = "0,0 %"
    $sn.Range("F72:F75").NumberFormatLocal = "0,0"
    $sn.Range("B71:F75").HorizontalAlignment = -4152
    $sn.Range("A1").Select()

    # Orden de lectura: resumen, parámetros, procesos en el orden del informe, estados, validación, resultados y referencia.
    $orden = @("LEEME", "MODELOS", "PARAMETROS", "MODELO", "INF", "Inflación", "Inflación, modelo", "TPM, datos", "TPM, modelo", "Cuenta, modelo", "TPM_Cuenta", "TPM_Logit", "CDT, datos", "CDT, modelo", "CDT", "ETF, datos", "ETF, modelo", "ETF", "Inmueble, datos", "Reajuste, modelo", "Valorización, datos", "Valorización, modelo", "Vacancia, modelo", "Markov, modelo", "INM", "PIB, estados", "PIB", "Validación del simulador", "Resultados, estrategias", "Comparación de estrategias", "Sensibilidad", "AUDITORIA", "MARKOV", "MARKOV_PIB", "MARKOV_ETF", "MARKOV_Vivienda", "VALIDACION", "RESULTADOS", "FUENTES", "DICCIONARIO")
    $dest.Worksheets.Item($orden[0]).Move($dest.Worksheets.Item(1))
    for ($k = 1; $k -lt $orden.Count; $k++) { $dest.Worksheets.Item($orden[$k]).Move($falta, $dest.Worksheets.Item($orden[$k - 1])) }

    $dest.Worksheets.Item("MODELO").Visible = 0
    $dest.Worksheets.Item("INF").Visible = 0
    $dest.Worksheets.Item("ETF").Visible = 0
    $dest.Worksheets.Item("CDT").Visible = 0
    $dest.Worksheets.Item("PIB").Visible = 0
    $dest.Worksheets.Item("TPM_Cuenta").Visible = 0
    $dest.Worksheets.Item("TPM_Logit").Visible = 0
    $dest.Worksheets.Item("INM").Visible = 0
    foreach ($n in @("MARKOV", "MARKOV_PIB", "MARKOV_ETF", "MARKOV_Vivienda", "VALIDACION", "RESULTADOS")) { $dest.Worksheets.Item($n).Visible = 0 }
    # Las hojas copiadas dejan un vínculo vacío al libro de Markov; ninguna fórmula lo usa.
    try { $dest.LinkSources(1) | ForEach-Object { $dest.BreakLink([string]$_, 1) } } catch { Write-Output "sin vinculos para romper" }
    foreach ($n in @("LEEME", "FUENTES", "DICCIONARIO", "Validación del simulador", "Resultados, estrategias", "Comparación de estrategias", "Sensibilidad", "RESULTADOS", "AUDITORIA")) { $dest.Worksheets.Item($n).Delete() }
    $dest.Worksheets.Item("MODELOS").Activate()
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
