function Restore-NetworkDrives {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupLocation
    )

    $csvPath = Join-Path -Path $BackupLocation -ChildPath "Drives.csv"

    # Verify backup file exists
    if (-not (Test-Path -Path $csvPath)) {
        Write-Warning "Network drives backup file not found at: $csvPath"
        return
    }

    try {
        # Import the CSV file
        $driveMappings = Import-Csv -Path $csvPath

        foreach ($mapping in $driveMappings) {
            $driveLetter = $mapping.Name
            $networkPath = $mapping.ProviderName

            # Validate drive letter format
            if ($driveLetter -notmatch '^[A-Z]:$') {
                Write-Warning "Invalid drive letter format: $driveLetter - Skipping"
                continue
            }

            # Validate network path format
            if ($networkPath -notmatch '^\\\\\w+\\\w+') {
                Write-Warning "Invalid network path format: $networkPath - Skipping"
                continue
            }

            # Check if drive letter is already in use
            $existingDrive = Get-PSDrive -Name $driveLetter.TrimEnd(':') -ErrorAction SilentlyContinue
            if ($existingDrive) {
                Write-Warning "Drive letter $driveLetter already in use - Skipping"
                continue
            }

            # Test network path accessibility with retry logic
            $maxRetries = 3
            $retryCount = 0
            $success = $false

            while (-not $success -and $retryCount -lt $maxRetries) {
                try {
                    if (Test-Path -Path $networkPath -ErrorAction Stop) {
                        # Map the network drive with persistence and global scope
                        $driveParams = @{
                            Name = $driveLetter.TrimEnd(':')
                            PSProvider = 'FileSystem'
                            Root = $networkPath
                            Persist = $true
                            Scope = 'Global'
                            ErrorAction = 'Stop'
                        }

                        New-PSDrive @driveParams
                        Write-Verbose "Successfully mapped drive $driveLetter to $networkPath"
                        $success = $true
                    } else {
                        Write-Warning "Network path not accessible: $networkPath"
                        $retryCount++
                        if ($retryCount -lt $maxRetries) {
                            Start-Sleep -Seconds ($retryCount * 2)
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to map drive $driveLetter - Attempt $($retryCount + 1): $($_.Exception.Message)"
                    $retryCount++
                    if ($retryCount -lt $maxRetries) {
                        Start-Sleep -Seconds ($retryCount * 2)
                    }
                }
            }

            if (-not $success) {
                Write-Warning "Failed to map drive $driveLetter after $maxRetries attempts"
            }
        }
    }
    catch {
        Write-Error "Error processing network drive mappings: $($_.Exception.Message)"
    }
}
