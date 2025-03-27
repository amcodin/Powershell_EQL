function Restore-NetworkDrives {
    param([string]$Path)
    
    $drivePath = Join-Path $Path "Drives.csv"
    if (Test-Path $drivePath) {
        $drives = Import-Csv -Path $drivePath
        foreach ($drive in $drives) {
            try {
                $letter = $drive.Name.Substring(0, 1)
                Write-Host "Mapping drive $letter to $($drive.ProviderName)"
                
                # Remove existing mapping if present
                $existing = Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='$($letter):'"
                if ($existing) {
                    $net = New-Object -ComObject WScript.Network
                    $net.RemoveNetworkDrive($letter)
                }
                
                # Add new mapping
                New-PSDrive -Name $letter -PSProvider FileSystem -Root $drive.ProviderName -Persist -Scope Global -ErrorAction Stop
                Write-Host "Successfully mapped drive $letter" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to map drive $letter`: $($_.Exception.Message)"
            }
        }
    }
}
