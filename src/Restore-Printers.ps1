function Restore-Printers {
    param([string]$Path)
    
    $printerPath = Join-Path $Path "Printers.txt"
    if (Test-Path $printerPath) {
        $printers = Get-Content -Path $printerPath
        $wsNet = New-Object -ComObject WScript.Network
        
        foreach ($printer in $printers) {
            try {
                Write-Host "Adding printer: $printer"
                $wsNet.AddWindowsPrinterConnection($printer)
                Write-Host "Successfully added printer" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to add printer $printer`: $($_.Exception.Message)"
            }
        }
    }
}
