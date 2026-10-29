# Recalculate the workbook with desktop Excel and save the cached values (Windows only).
$p = Join-Path $PSScriptRoot "DataMoments_Analyst_Calculations.xlsx"
$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
$wb = $x.Workbooks.Open($p)
$x.CalculateFullRebuild()
$wb.Save(); $wb.Close($false); $x.Quit()
Write-Output "recalculated $p"
