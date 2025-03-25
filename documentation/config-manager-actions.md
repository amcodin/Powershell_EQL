# Using PowerShell to Trigger Configuration Manager Actions (via CCMExec)
If you're using PowerShell to trigger actions like software inventory, hardware inventory, policy retrieval, etc., you can call ccmexec.exe and use specific flags to start these actions.

Here’s a general example of how you can trigger common Configuration Manager actions:
# Trigger Software Inventory
Invoke-Command -ScriptBlock {
    Start-Process "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction SoftwareInventory"
}

# Trigger Hardware Inventory
Invoke-Command -ScriptBlock {
    Start-Process "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction HardwareInventory"
}

# Trigger Policy Retrieval (Machine Policy Retrieval & Evaluation)
Invoke-Command -ScriptBlock {
    Start-Process "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction MachinePolicy"
}

# Trigger User Policy Retrieval
Invoke-Command -ScriptBlock {
    Start-Process "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction UserPolicy"
}

# Trigger Software Update Scan
Invoke-Command -ScriptBlock {
    Start-Process "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction SoftwareUpdatesScan"
}

# Trigger Client Actions (such as triggering compliance evaluation or running other tasks)
Invoke-Command -ScriptBlock {
    Start-Process "C:\Windows\CCM\ccmexec.exe" -ArgumentList "-TriggerAction ClientActions"
}


If you have the Configuration Manager PowerShell module installed, you can use the cmdlets directly to trigger actions on the client.

Here’s how to trigger different actions using the Configuration Manager cmdlets:

Import Configuration Manager Module (if it's not already loaded):

powershell
Copy
# Import the Configuration Manager module (only if not already imported)
Import-Module "$($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5))\ConfigurationManager.psd1"
Trigger Machine Policy Retrieval:

powershell
Copy
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="MachinePolicy"}
Trigger User Policy Retrieval:

powershell
Copy
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="UserPolicy"}
Trigger Software Inventory:

powershell
Copy
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="SoftwareInventory"}
Trigger Hardware Inventory:

powershell
Copy
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="HardwareInventory"}
Trigger Software Update Scan:

powershell
Copy
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="SoftwareUpdatesScan"}
3. List of Common Action Schedules and Their IDs
Here’s a list of some common Configuration Manager action schedule IDs:

Machine Policy Retrieval & Evaluation: MachinePolicy

User Policy Retrieval & Evaluation: UserPolicy

Hardware Inventory: HardwareInventory

Software Inventory: SoftwareInventory

Software Update Scan: SoftwareUpdatesScan

Compliance Evaluation: ComplianceEvaluation

These schedule IDs are used in the TriggerSchedule method to trigger each action.

Example: Triggering All Actions at Once
You can combine the above PowerShell cmdlets into one script to trigger all major actions at once:

powershell
Copy
# Import Configuration Manager PowerShell module
Import-Module "$($env:SMS_ADMIN_UI_PATH.Substring(0,$env:SMS_ADMIN_UI_PATH.Length-5))\ConfigurationManager.psd1"

# Trigger some actions
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="MachinePolicy"}
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="UserPolicy"}
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="HardwareInventory"}
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="SoftwareInventory"}
Invoke-CimMethod -Namespace "root\ccm\clientSDK" -ClassName CCM_Client -MethodName TriggerSchedule -Arguments @{scheduleID="SoftwareUpdatesScan"}


$computername = "RemoteComputerName"

Invoke-Command -ComputerName $computername -ScriptBlock {
    $ClientActionPath = Join-Path $env:windir "CCM\ClientAction.exe"

    $Actions = @(
        "Machine Policy Retrieval & Evaluation Cycle",
        "User Policy Retrieval & Evaluation Cycle",
        "Application Deployment Evaluation Cycle",
        "Software Updates Deployment Evaluation Cycle",
        "Software Updates Scan Cycle",
        "Hardware Inventory Cycle",
        "Software Inventory Cycle",
        "File Collection Cycle",
        "Discovery Data Collection Cycle"
    )

    foreach ($Action in $Actions) {
        Write-Host "Triggering: $Action"
        try {
            Start-Process -FilePath $ClientActionPath -ArgumentList "$Action" -Wait -ErrorAction Stop
            Write-Host "$Action triggered successfully."
        } catch {
            Write-Warning "Failed to trigger $Action: $($_.Exception.Message)"
        }
    }

    Write-Host "All Configuration Manager actions triggered (or attempted)."
}