$outputReport = '.\result.xlsx'
Remove-Item $outputReport -ErrorAction Ignore

$data = Import-Csv .\sales.csv

$data |
Where-Object { $_.Region -ne 'West' } |
Export-Excel $outputReport -WorksheetName Sales -AutoSize `
    -PivotRows Region -PivotData @{ Units = 'Sum' } -Activate -Show 