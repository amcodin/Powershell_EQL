﻿#Requires -Version 3.0

# Load required assemblies 
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Configure script behavior
$VerbosePreference = 'Continue'
$script:DefaultSavePath = $env:HOMEDRIVE
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Define backup paths
$script:BackupPaths = @{
    Required = @(
        @{
            Path = "$([Environment]::GetFolderPath('Personal'))"
            Description = 'Documents'
        }
        @{
            Path = "$([Environment]::GetFolderPath('Favorites'))"
            Description = 'Favorites'
        }
    )
    Optional = @(
        @{
            Path = "$env:APPDATA\Microsoft\Signatures"
            Description = 'Outlook signatures'
        }
        @{
            Path = "$env:SystemDrive\User"
            Description = 'User folder'
        }
        @{
            Path = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"
            Description = 'Quick Access'
        }
        @{
            Path = "$env:SystemDrive\Temp"
            Description = 'Temp folder'
        }
        @{
            Path = "$env:APPDATA\google\googleearth\myplaces.kml"
            Description = 'Google Earth KML'
        }
        @{
            Path = "$ENV:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
            Description = 'Chrome Bookmarks'
        }
        @{
            Path = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
            Description = 'Sticky Notes'
        }
    )
}

# Helper Functions
function Show-UserPrompt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [string]$Title = 'User Prompt',
        [string[]]$ButtonText = @('OK')
    )
    
    try {
        if ($ButtonText.Count -eq 1 -and $ButtonText[0] -eq 'OK') {
            $null = [System.Windows.MessageBox]::Show($Message, $Title, 'OK')
            return 'OK'
        }

        $buttonElements = foreach ($btn in $ButtonText) {
            "<Button Name='Btn$($btn -replace '\W')' Content='$btn' Margin='5' Padding='10,5' MinWidth='60'/>"
        }

        [xml]$xaml = @"
<Window 
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    Title='$Title'
    SizeToContent='WidthAndHeight'
    WindowStartupLocation='CenterScreen'
    MaxWidth='500'>
    <StackPanel Margin='10'>
        <TextBlock Text='$Message' TextWrapping='Wrap' Margin='0,0,0,20'/>
        <StackPanel Orientation='Horizontal' HorizontalAlignment='Right'>
            $buttonElements
        </StackPanel>
    </StackPanel>
</Window>
"@

        $reader = New-Object System.Xml.XmlNodeReader $xaml
        $window = [Windows.Markup.XamlReader]::Load($reader)
        $script:result = $null

        foreach ($btn in $ButtonText) {
            $button = $window.FindName("Btn$($btn -replace '\W')")
            if ($button) {
                $button.Add_Click({
                    $script:result = $this.Content
                    $window.Close()
                })
            }
        }

        $null = $window.ShowDialog()
        return $script:result
    }
    catch {
        Write-Warning ("Failed to show prompt: {0}" -f $_.Exception.Message)
        return $null
    }
}

function Set-GPupdate {
    [CmdletBinding()]
    param()
    
    try {
        Write-Host 'Running Group Policy Update...' -ForegroundColor Cyan
        $null = Start-Process -FilePath 'cmd.exe' -ArgumentList '/k gpupdate /force' -PassThru
        Write-Host 'Group Policy update initiated' -ForegroundColor Green
    }
    catch {
        Write-Warning ("Failed to run GPUpdate: {0}" -f $_.Exception.Message)
    }
}

function Backup-Files {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    try {
        foreach ($category in $script:BackupPaths.Keys) {
            foreach ($item in $script:BackupPaths[$category]) {
                if (Test-Path $item.Path) {
                    $destPath = Join-Path -Path $Path -ChildPath (Split-Path -Leaf $item.Path)
                    Write-Verbose ("Backing up {0}: {1}" -f $item.Description, $item.Path)
                    Copy-Item -Path $item.Path -Destination $destPath -Force -Recurse
                }
            }
        }
        return $true
    }
    catch {
        Write-Warning ("Failed to backup files: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Restore-Files {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    try {
        foreach ($category in $script:BackupPaths.Keys) {
            foreach ($item in $script:BackupPaths[$category]) {
                $sourcePath = Join-Path -Path $Path -ChildPath (Split-Path -Leaf $item.Path)
                if (Test-Path $sourcePath) {
                    Write-Verbose ("Restoring {0}" -f $item.Description)
                    Copy-Item -Path $sourcePath -Destination $item.Path -Force -Recurse
                }
            }
        }
        return $true
    }
    catch {
        Write-Warning ("Failed to restore files: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Invoke-NetworkBackupOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        
        [Parameter(Mandatory)]
        [ValidateSet('Backup', 'Restore')]
        [string]$Operation
    )

    try {
        if ($Operation -eq 'Backup') {
            Write-Verbose 'Starting network drives backup...'
            $networkDrives = Get-WmiObject -Class 'Win32_MappedLogicalDisk' | 
                Select-Object -Property Name, ProviderName
            
            if ($networkDrives) {
                $drivePath = Join-Path -Path $Path -ChildPath 'Drives.csv'
                $networkDrives | Export-Csv -Path $drivePath -NoTypeInformation
                Write-Verbose 'Network drives backed up successfully'
                return $true
            }
            Write-Verbose 'No network drives found to backup'
            return $true
        }
        else {
            Write-Verbose 'Starting network drives restore...'
            $drivePath = Join-Path -Path $Path -ChildPath 'Drives.csv'
            
            if (Test-Path $drivePath) {
                $drives = Import-Csv -Path $drivePath
                foreach ($drive in $drives) {
                    try {
                        $letter = $drive.Name.Substring(0,1)
                        Write-Verbose ("Mapping drive {0}" -f $letter)
                        New-PSDrive -Name $letter -PSProvider FileSystem -Root $drive.ProviderName -Persist -Scope Global -ErrorAction Stop
                    }
                    catch {
                        Write-Warning ("Failed to map drive {0}: {1}" -f $letter, $_.Exception.Message)
                    }
                }
                return $true
            }
            Write-Verbose 'No network drives backup found'
            return $true
        }
    }
    catch {
        Write-Warning ("{0} of network drives failed: {1}" -f $Operation, $_.Exception.Message)
        return $false
    }
}

function Invoke-PrinterBackupOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        
        [Parameter(Mandatory)]
        [ValidateSet('Backup', 'Restore')]
        [string]$Operation
    )

    try {
        if ($Operation -eq 'Backup') {
            Write-Verbose 'Starting printer mappings backup...'
            $printers = Get-CimInstance -ClassName Win32_Printer -Filter "Network=True" |
                Select-Object -ExpandProperty Name
                
            if ($printers) {
                $printerPath = Join-Path -Path $Path -ChildPath 'Printers.txt'
                $printers | Set-Content -Path $printerPath
                Write-Verbose 'Printer mappings backed up successfully'
                return $true
            }
            Write-Verbose 'No network printers found to backup'
            return $true
        }
        else {
            Write-Verbose 'Starting printer mappings restore...'
            $printerPath = Join-Path -Path $Path -ChildPath 'Printers.txt'
            
            if (Test-Path $printerPath) {
                $printers = Get-Content -Path $printerPath
                $wsNet = New-Object -ComObject WScript.Network
                
                foreach ($printer in $printers) {
                    try {
                        Write-Verbose ("Adding printer: {0}" -f $printer)
                        $wsNet.AddWindowsPrinterConnection($printer)
                    }
                    catch {
                        Write-Warning ("Failed to add printer {0}: {1}" -f $printer, $_.Exception.Message)
                    }
                }
                return $true
            }
            Write-Verbose 'No printer mappings backup found'
            return $true
        }
    }
    catch {
        Write-Warning ("{0} of printer mappings failed: {1}" -f $Operation, $_.Exception.Message)
        return $false
    }
}

function Start-BackupRestore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        
        [Parameter(Mandatory)]
        [ValidateSet('Backup', 'Restore')]
        [string]$Operation
    )

    try {
        Write-Verbose 'Processing files...'
        if ($Operation -eq 'Backup') {
            if (-not (Backup-Files -Path $Path)) {
                Write-Warning 'File backup completed with errors'
            }
        }
        else {
            if (-not (Restore-Files -Path $Path)) {
                Write-Warning 'File restore completed with errors'
            }
        }

        Write-Verbose 'Processing network drives...'
        if (-not (Invoke-NetworkBackupOperation -Path $Path -Operation $Operation)) {
            Write-Warning 'Network drives operation completed with errors'
        }

        Write-Verbose 'Processing printer mappings...'
        if (-not (Invoke-PrinterBackupOperation -Path $Path -Operation $Operation)) {
            Write-Warning 'Printer mappings operation completed with errors'
        }

        Write-Verbose ("{0} operation completed" -f $Operation)
        Show-UserPrompt -Message ("{0} completed successfully" -f $Operation) -Title 'Operation Complete'
        return $true
    }
    catch {
        Write-Warning ("{0} operation failed: {1}" -f $Operation, $_.Exception.Message)
        Show-UserPrompt -Message "$Operation failed: $($_.Exception.Message)" -Title 'Operation Failed'
        return $false
    }
}

# Main script execution
try {
    $operation = Show-UserPrompt -Message 'Select operation' -Title 'Backup/Restore' -ButtonText @('Backup', 'Restore')
    
    if ($operation) {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = ("Select folder for {0}" -f $operation)
        $dialog.SelectedPath = $env:USERPROFILE
        
        if ($dialog.ShowDialog() -eq 'OK') {
            $result = Start-BackupRestore -Path $dialog.SelectedPath -Operation $operation
            
            if ($result -and $operation -eq 'Restore') {
                $runGpUpdate = Show-UserPrompt -Message 'Run GPUpdate now?' -Title 'GPUpdate' -ButtonText @('Yes', 'No')
                if ($runGpUpdate -eq 'Yes') {
                    Set-GPupdate
                }
            }
        }
    }
}
catch {
    Write-Warning ("Script execution failed: {0}" -f $_.Exception.Message)
    Show-UserPrompt -Message "Script execution failed: $($_.Exception.Message)" -Title 'Error'
}
