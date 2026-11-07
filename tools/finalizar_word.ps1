# Actualiza la tabla de contenido y los campos de docs/informe_final.docx y exporta el PDF.
#   powershell -File tools/finalizar_word.ps1
param([string]$Docx)
$ErrorActionPreference = "Stop"
$aqui = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Docx) { $Docx = "$aqui\..\docs\informe_final.docx" }
$Docx = (Resolve-Path $Docx).Path
$pdf = [IO.Path]::ChangeExtension($Docx, ".pdf")

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Docx)
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17)
    $paginas = $doc.ComputeStatistics(2)
    $doc.Close($false)
    Write-Output "ok $paginas paginas"
    Write-Output $pdf
}
finally {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
