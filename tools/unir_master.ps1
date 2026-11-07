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
if (-not $Consolidado) { $Consolidado = "$aqui\..\..\Consolidado\Trabajo_Final_Simulacion_puntos_1_a_7.xlsx" }
if (-not $Markov) { $Markov = "$aqui\..\..\Analisis_completo_estados_economicos_Markov.xlsx" }
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
    $mar.Worksheets.Copy($falta, $dest.Worksheets.Item("DECISIONES"))
    $dest.Worksheets.Item("TPM_Logit").Move($falta, $dest.Worksheets.Item("TPM_Cuenta"))
    $ultima = $dest.Worksheets.Item($dest.Worksheets.Count)
    $dest.Worksheets.Item("VALIDACION").Move($falta, $ultima)
    $ultima = $dest.Worksheets.Item($dest.Worksheets.Count)
    $dest.Worksheets.Item("RESULTADOS").Move($falta, $ultima)
    $dest.Worksheets.Item("Parametros_simulador").Delete()
    $mar.Close($false)
    $nue.Close($false)

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
