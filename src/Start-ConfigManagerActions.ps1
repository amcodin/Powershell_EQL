function Start-ConfigManagerActions {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseModule
    )

    try {
        Write-Verbose "Initiating Configuration Manager actions"
        
        if ($UseModule) {
            # Try using Configuration Manager module
            try {
                $modulePath = "$($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5))\ConfigurationManager.psd1"
                Import-Module $modulePath -ErrorAction Stop
                
                $scheduleActions = @(
                    'MachinePolicy',
                    'UserPolicy',
                    'HardwareInventory',
                    'SoftwareInventory',
                    'SoftwareUpdatesScan'
                )

                foreach ($action in $scheduleActions) {
                    Write-Verbose "Triggering $action"
                    Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID=$action}
                }
            }
            catch {
                Write-Warning "Failed to use CM module, falling back to ccmexec.exe: $($_.Exception.Message)"
                $UseModule = $false
            }
        }
        
        if (-not $UseModule) {
            # Use ccmexec.exe direct execution
            $actions = @(
                @{Name = "Machine Policy Retrieval"; Flag = "MachinePolicy"},
                @{Name = "User Policy Retrieval"; Flag = "UserPolicy"},
                @{Name = "Software Inventory"; Flag = "SoftwareInventory"},
                @{Name = "Hardware Inventory"; Flag = "HardwareInventory"},
                @{Name = "Software Updates Scan"; Flag = "SoftwareUpdatesScan"}
            )

            foreach ($action in $actions) {
                Write-Verbose "Triggering $($action.Name)"
                $process = Start-Process -FilePath "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction $($action.Flag)" -NoNewWindow -PassThru
                $process.WaitForExit()

                if ($process.ExitCode -ne 0) {
                    Write-Warning "$($action.Name) action failed with exit code $($process.ExitCode)"
                }
            }
        }

        Write-Verbose "Configuration Manager actions completed"
        return $true
    }
    catch {
        Write-Error "Failed to execute Configuration Manager actions: $($_.Exception.Message)"
        return $false
    }
}
