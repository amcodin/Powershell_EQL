﻿#requires -Version 3.0
##########################################
#        Client Data Backup Tool         #
#       Written By Stephen Onions        #
#    Edited By Kevin King 04.12.2023     #  
#    Edited by Jared Vosters 6/12/23  #  Prompt user to run GP Update & rearranged persist & global scope for New-PSDrive to fix energyq drive issues
#  Edited by Jared Vosters 6/3/25  #  All changes commented with #JV2025
#    Changelog- Updated path to localdata, removing the need to set it for backup, and 2 less clicks during restore
# Moved the GPUPDATE to only trigger on a restore
# Updated the QUickaccesslinks to overwrite the existing file on restore

# JV2025 Changes:
# 1) Configuration-driven backup paths            #configchanges
# 2) Improved error handling and logging         #errorhandling
# 3) Enhanced GP Update integration              #gpupdate
# 4) Simplified LocalData path handling          #localdata
# 5) Optimized UI/UX with WPF                   #uiupdate
# 6) Better progress reporting                   #progress
# 7) Added Sticky Notes backup support           #stickynotes
##########################################

# Required assemblies #JV2025
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Script configuration #JV2025
$VerbosePreference = 'Continue'
$script:DefaultSavePath = $env:HOMEDRIVE

# Configuration structure for backup paths #JV2025 #configchanges
$script:BackupPaths = @{
    Required = @(
        @{Path = [Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop)}
        @{Path = [Environment]::GetFolderPath([Environment+SpecialFolder]::Documents)}
        @{Path = [Environment]::GetFolderPath([Environment+SpecialFolder]::Favorites)}
    )
    Optional = @(
        @{Path = "$env:APPDATA\Microsoft\Signatures"; Description = "Outlook signatures"}
        @{Path = "$env:SystemDrive\User"; Description = "User folder"}
        @{Path = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"; Description = "Quick Access"}
        @{Path = "$env:SystemDrive\Temp"; Description = "Temp folder"}
        @{Path = "$env:APPDATA\google\googleearth\myplaces.kml"; Description = "Google Earth KML"}
        @{Path = "$ENV:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"; Description = "Chrome Bookmarks"}
        @{Path = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"; Description = "Windows 10/11 Sticky Notes"} #stickynotes
        @{Path = "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt"; Description = "Legacy Sticky Notes"} #stickynotes
    )
}

# Network and printer backup functions
Function Backup-NetworkDrives {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SaveLocation
    )
    
    try {
        Write-Verbose "Backing up network drives..."
        $drives = Get-WmiObject -Class 'Win32_MappedLogicalDisk' | 
                  Select-Object -Property Name, ProviderName
                  
        if ($drives) {
            $drives | Export-Csv -Path "$SaveLocation\Drives.csv" -NoTypeInformation
            Write-Verbose "Network drives backed up successfully"
            return $true
        } else {
            Write-Verbose "No network drives found to backup"
            return $false
        }
    }
    catch {
        Write-Warning "Failed to backup network drives: $_"
        return $false
    }
}

Function Backup-PrinterMappings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SaveLocation
    )
    
    try {
        Write-Verbose "Backing up printer mappings..."
        $printers = Get-WmiObject -Query "Select * FROM Win32_Printer WHERE Local=$false" | 
                    Select-Object -ExpandProperty Name
                    
        if ($printers) {
            $printers | Set-Content -Path "$SaveLocation\Printers.txt"
            Write-Verbose "Printer mappings backed up successfully"
            return $true
        } else {
            Write-Verbose "No printer mappings found to backup"
            return $false
        }
    }
    catch {
        Write-Warning "Failed to backup printer mappings: $_"
        return $false
    }
}

Function Restore-NetworkDrives {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BackupLocation
    )
    
    try {
        $drivesFile = "$BackupLocation\Drives.csv"
        if (Test-Path $drivesFile) {
            Write-Verbose "Restoring network drives..."
            $drives = Import-Csv -Path $drivesFile
            
            foreach ($drive in $drives) {
                $letter = $drive.Name.Substring(0, 1)
                $path = $drive.ProviderName
                Write-Verbose "Mapping drive $letter to $path"
                
                # Use New-PSDrive with Persist parameter and Global scope
                New-PSDrive -Persist -Name $letter -PSProvider FileSystem -Root $path -Scope Global -ErrorAction Continue
            }
            Write-Verbose "Network drives restored successfully"
            return $true
        } else {
            Write-Verbose "No network drives backup file found"
            return $false
        }
    }
    catch {
        Write-Warning "Failed to restore network drives: $_"
        return $false
    }
}

Function Restore-PrinterMappings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$BackupLocation
    )
    
    try {
        $printersFile = "$BackupLocation\Printers.txt"
        if (Test-Path $printersFile) {
            Write-Verbose "Restoring printer mappings..."
            $printers = Get-Content -Path $printersFile
            
            if ($printers) {
                $wsNet = New-Object -ComObject WScript.Network
                foreach ($printer in $printers) {
                    Write-Verbose "Mapping printer: $printer"
                    $wsNet.AddWindowsPrinterConnection($printer)
                }
                Write-Verbose "Printer mappings restored successfully"
            } else {
                Write-Verbose "No printer mappings found in backup"
            }
            return $true
        } else {
            Write-Verbose "No printer mappings backup file found"
            return $false
        }
    }
    catch {
        Write-Warning "Failed to restore printer mappings: $_"
        return $false
    }
}

# Remote info class definition
Add-Type -Language CSharp -TypeDefinition @'
public class RemoteInfo {
    public string Computername;
    public string ProfileFolder;
    public string ProfileName;
    public bool IsInUse;
    public string SSID;
    public Microsoft.Win32.RegistryKey UserHive;
}
'@

#region Helper Functions

Function Set-GPupdate { #gpupdate
    [CmdletBinding()]
    param()
    
    Write-Host "Running Group Policy Update..." -ForegroundColor Cyan
    #JV2025 spit out new cmd window for gpupdate
    Start-Process -FilePath "cmd.exe" -ArgumentList "/k gpupdate /force" -PassThru | Out-Null
    
    Write-Host "Group Policy update initiated" -ForegroundColor Green
}

Function Show-UserPrompt {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory)][String]$Message,
        [string]$Button1 = $null,
        [string]$Button2 = $null,
        [string]$Button3 = 'OK'
    )
    
    # Try simple MessageBox first
    if ([string]::IsNullOrEmpty($Button1) -and [string]::IsNullOrEmpty($Button2) -and $Button3 -eq 'OK') {
        try {
            $null = [System.Windows.MessageBox]::Show($Message, 'Prompt', 'OK')
            return 'OK'
        }
        catch {
            Write-Verbose "Falling back to custom dialog"
        }
    }

    # Custom dialog XAML
    [xml]$xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        Title='User Prompt' SizeToContent='WidthAndHeight' MaxWidth='500'
        WindowStartupLocation='CenterScreen'>
    <StackPanel Margin='10'>
        <TextBlock Text='$Message' TextWrapping='Wrap' Margin='0,0,0,20'/>
        <StackPanel Orientation='Horizontal' HorizontalAlignment='Right'>
            $(if($Button1){"<Button Name='Btn1' Content='$Button1' Margin='5' Padding='10,5' MinWidth='60'/>"})
            $(if($Button2){"<Button Name='Btn2' Content='$Button2' Margin='5' Padding='10,5' MinWidth='60'/>"})
            $(if($Button3){"<Button Name='Btn3' Content='$Button3' Margin='5' Padding='10,5' MinWidth='60'/>"})
        </StackPanel>
    </StackPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $result = $null
    
    @('Btn1','Btn2','Btn3') | ForEach-Object {
        $btn = $window.FindName($_)
        if ($btn) {
            $btn.Add_Click({
                $script:result = $this.Content
                $window.Close()
            })
        }
    }

    $null = $window.ShowDialog()
    return $result
}

Function New-FilePathObject { #errorhandling
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    
    try {
        $fso = New-Object -ComObject Scripting.FileSystemObject
        $item = Get-Item -Path $Path -ErrorAction Stop
        
        if ($item.PSIsContainer) {
            $folder = $fso.GetFolder($item.FullName)
            $size = [math]::Round($folder.Size/1MB, 3)
            $type = 'Folder'
        }
        else {
            $file = $fso.GetFile($item.FullName)
            $size = [math]::Round($file.Size/1MB, 3)
            $type = 'File'
        }

        [PSCustomObject]@{
            Name = $item.Name
            FullPath = $item.FullName
            Type = $type
            Selected = ($item.Name -notmatch 'Redirected')
            Size = $size
        }
    }
    catch {
        Write-Warning "Failed to process path '$Path': $_"
        return $null
    }
}

# Helper function for adding backup locations #JV2025
Function Add-BackupLocation {
    [CmdletBinding()]
    param()
    
    try {
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select Folder to Add"
        $folderBrowser.SelectedPath = $env:USERPROFILE
        
        if ($folderBrowser.ShowDialog() -eq 'OK') {
            $newItem = New-FilePathObject -Path $folderBrowser.SelectedPath
            if ($newItem) {
                $script:FilePaths += $newItem
                Update-ListView
                Update-BackupSize
                Write-Verbose "Added new backup location: $($newItem.FullPath)"
            }
        }
    }
    catch {
        Write-Warning "Failed to add backup location: $_"
        Show-UserPrompt -Message "Failed to add location: $_"
    }
}

Function Set-SaveLocation {
    [CmdletBinding()]
    param()
    
    try {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = "Select Backup Location"
        $dialog.SelectedPath = $script:DefaultSavePath
        
        if ($dialog.ShowDialog() -eq 'OK') {
            $script:DefaultSavePath = $dialog.SelectedPath
            $TxtSaveLoc.Text = $script:DefaultSavePath
            
            $drive = [System.IO.DriveInfo]::new($script:DefaultSavePath)
            $freeSpace = [math]::Round($drive.AvailableFreeSpace/1GB, 2)
            $lblFreeSpace.Content = "Free Space: $freeSpace GB"
        }
    }
    catch {
        Write-Warning "Failed to set save location: $_"
    }
}

Function Set-InitialStateBackup {
    [CmdletBinding()]
    param()
    
    Write-Host "Initializing backup state..." -ForegroundColor Cyan
    $script:FilePaths = @()

    # Process required paths
    $script:BackupPaths.Required | ForEach-Object {
        $script:FilePaths += New-FilePathObject -Path $_.Path
    }

    # Process optional paths
    $script:BackupPaths.Optional | Where-Object { Test-Path $_.Path } | ForEach-Object {
        Write-Verbose "Adding optional path: $($_.Description)"
        $script:FilePaths += New-FilePathObject -Path $_.Path
    }

    Update-BackupSize
    Update-ListView
}

Function Set-InitialStateAdminBackup {
    [CmdletBinding()]
    param()
    
    Write-Host "Initializing admin backup state..." -ForegroundColor Cyan
    Set-InitialStateBackup
    
    # Add admin-specific paths
    $script:BackupPaths.Optional | Where-Object { $_.Path -like "*\User" } | ForEach-Object {
        Write-Verbose "Adding admin path: $($_.Description)"
        $script:FilePaths += New-FilePathObject -Path $_.Path
    }
    
    Update-BackupSize
    Update-ListView
}

Function Set-InitialStateRestore {
    [CmdletBinding()]
    param()
    
    Write-Host "Initializing restore state..." -ForegroundColor Cyan
    
    if (-not $TxtSaveLoc.Text -or -not (Test-Path $TxtSaveLoc.Text)) {
        Show-UserPrompt -Message "Please select a valid backup location first."
        return
    }

    try {
        $script:FilePaths = Get-ChildItem -Path $TxtSaveLoc.Text -ErrorAction Stop | 
            ForEach-Object { New-FilePathObject -Path $_.FullName }
        
        Update-BackupSize
        Update-ListView
    }
    catch {
        Write-Warning "Failed to initialize restore state: $_"
        Show-UserPrompt -Message "Failed to process restore location: $_"
    }
}

Function Update-BackupSize {
    [float]$script:backupSize = ($script:FilePaths | 
        Where-Object { $_.Selected } | 
        Measure-Object -Property Size -Sum).Sum
    
    $lblrequiredSpace.Content = ("Required Space: {0:N2} GB" -f 
        ($script:backupSize/1KB))
}

Function Update-ListView {
    $Script:lvview = [Windows.Data.ListCollectionView]$script:FilePaths
    $lvwFileList.ItemsSource = $Script:lvview
}

# Form initialization function #JV2025
Function Initialize-Form {
    [CmdletBinding()]
    param()

    Write-Host "Initializing form..." -ForegroundColor Cyan

    # Set default save path #localdata
    if (Test-Path "$env:SystemDrive\LocalData") {
        $script:DefaultSavePath = "$env:SystemDrive\LocalData"
    }

    # Check admin status
    $script:Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    # Prompt for operation type
    $Script:Action = Show-UserPrompt -Message 'Are we running a Backup or Restore?' -Button2 'Restore' -Button3 'Backup'
    
    # Initialize based on admin status and operation type
    if ($script:Admin) {
        if ($Script:Action -eq 'Backup') {
            Write-Host "Starting admin backup..." -ForegroundColor Cyan
            Set-InitialStateAdminBackup
        }
        elseif ($Script:Action -eq 'Restore') {
            Write-Host "Starting admin restore..." -ForegroundColor Cyan
            Set-InitialStateRestore
            
            # Prompt for GP Update after restore
            $Script:Action2 = Show-UserPrompt -Message 'Run A GPUPDATE?' -Button2 'Yes' -Button3 'No'
            if ($Script:Action2 -eq 'Yes') {
                Write-Host "Running GP Update..." -ForegroundColor Cyan
                $null = Set-GPupdate
            }
        }
    }
    else {
        if ($Script:Action -eq 'Backup') {
            Write-Host "Starting user backup..." -ForegroundColor Cyan
            Set-InitialStateBackup
        }
        elseif ($Script:Action -eq 'Restore') {
            Write-Host "Starting user restore..." -ForegroundColor Cyan
            Set-InitialStateRestore
            
            # Prompt for GP Update after restore
            $Script:Action2 = Show-UserPrompt -Message 'Run A GPUPDATE?' -Button2 'Yes' -Button3 'No'
            if ($Script:Action2 -eq 'Yes') {
                Write-Host "Running GP Update..." -ForegroundColor Cyan
                $null = Set-GPupdate
            }
        }
    }
}

Function Start-BRProcess {
    [CmdletBinding()]
    param()
    
    if (-not $TxtSaveLoc.Text) {
        Show-UserPrompt -Message "Please select a location first."
        return
    }

    try {
        $selectedPaths = $script:FilePaths | Where-Object { $_.Selected }
        $total = $selectedPaths.Count
        
        # Add steps for network drives and printers if checked
        if ($chk_networkDrives.IsChecked) { $total++ }
        if ($chk_Printers.IsChecked) { $total++ }
        
        $current = 0

        if ($Script:Action -eq 'Backup') {
            # Process file/folder backups
            foreach ($path in $selectedPaths) {
                $current++
                $percent = ($current / $total) * 100
                
                Write-Progress -Activity "Backing up files" -Status "Processing $($path.Name)" -PercentComplete $percent
                
                $destination = $TxtSaveLoc.Text
                Copy-Item -Path $path.FullPath -Destination $destination -Recurse -Force
            }
            
            # Process network drives backup
            if ($chk_networkDrives.IsChecked) {
                $current++
                $percent = ($current / $total) * 100
                Write-Progress -Activity "Backing up files" -Status "Processing network drives" -PercentComplete $percent
                Backup-NetworkDrives -SaveLocation $TxtSaveLoc.Text
            }
            
            # Process printer mappings backup
            if ($chk_Printers.IsChecked) {
                $current++
                $percent = ($current / $total) * 100
                Write-Progress -Activity "Backing up files" -Status "Processing printer mappings" -PercentComplete $percent
                Backup-PrinterMappings -SaveLocation $TxtSaveLoc.Text
            }
        }
        else { # Restore
            # Process file/folder restores
            foreach ($path in $selectedPaths) {
                $current++
                $percent = ($current / $total) * 100
                
                Write-Progress -Activity "Restoring files" -Status "Processing $($path.Name)" -PercentComplete $percent
                
                $destination = $env:USERPROFILE
                Copy-Item -Path $path.FullPath -Destination $destination -Recurse -Force
            }
            
            # Process network drives restore
            if ($chk_networkDrives.IsChecked) {
                $current++
                $percent = ($current / $total) * 100
                Write-Progress -Activity "Restoring files" -Status "Processing network drives" -PercentComplete $percent
                Restore-NetworkDrives -BackupLocation $TxtSaveLoc.Text
            }
            
            # Process printer mappings restore
            if ($chk_Printers.IsChecked) {
                $current++
                $percent = ($current / $total) * 100
                Write-Progress -Activity "Restoring files" -Status "Processing printer mappings" -PercentComplete $percent
                Restore-PrinterMappings -BackupLocation $TxtSaveLoc.Text
            }
        }

        Write-Progress -Activity "Processing complete" -Completed
        Show-UserPrompt -Message "Operation completed successfully."
    }
    catch {
        Write-Warning "Operation failed: $_"
        Show-UserPrompt -Message "Operation failed: $_"
    }
}

#endregion Helper Functions

#region Main Form Logic

# Main form XAML definition #JV2025 #uiupdate
[xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="CanResize"
    Width="800"
    Height="600"
    Title="Data Backup Tool"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,0,0">
        <Label Width="Auto" Height="30" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Content="Backup Location (New or Existing):"/>
        <TextBox Name="TxtSaveLoc" Width="400" Height="30" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnSavBrowse" Width="60" Height="30" HorizontalAlignment="Left" Margin="420,40,0,0" VerticalAlignment="Top" Content="Browse"/>
        <Label Name="lblFreeSpace" Width="Auto" Height="30" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,10,0,0"/>
        <Label Name="lblRequiredSpace" Width="Auto" Height="30" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="490,40,0,0"/>
        <Label Width="200" Height="30" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Content="Files/Locations Selected for Backup:"/>
        <ListView Name="lvwFileList" Height="Auto" Width="Auto" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Margin="10,110,200,50">
            <ListView.Resources>
                <Style TargetType="{x:Type ListViewItem}">
                    <Setter Property="IsSelected" Value="{Binding Selected, Mode=TwoWay}"/>
                </Style>
            </ListView.Resources>
            <ListView.View>
                <GridView>
                    <GridView.Columns>
                        <GridViewColumn>
                            <GridViewColumn.CellTemplate>
                                <DataTemplate>
                                    <CheckBox Tag="{Binding Name}" IsChecked="{Binding RelativeSource={RelativeSource AncestorType={x:Type ListViewItem}}, Path=IsSelected}"/>
                                </DataTemplate>
                            </GridViewColumn.CellTemplate>
                        </GridViewColumn>
                        <GridViewColumn DisplayMemberBinding="{Binding Name}" Header="Name"/>
                        <GridViewColumn DisplayMemberBinding="{Binding FullPath}" Header="FullPath"/>
                        <GridViewColumn DisplayMemberBinding="{Binding Type}" Header="Type"/>
                        <GridViewColumn DisplayMemberBinding="{Binding Size}" Header="Size mb"/>
                    </GridView.Columns>
                </GridView>
            </ListView.View>
        </ListView>
        <Button Name="btnAddLoc" Width="80" Height="30" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,110,100,0" Content="Add File"/>
        <StackPanel Margin="0,180,10,0" HorizontalAlignment="Right">
            <Label Content="Non File Based Options"/>
            <CheckBox Name="chk_networkDrives" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,2,0,0" IsChecked="True">Network Drive Mappings</CheckBox>
            <CheckBox Name="chk_Printers" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="0,2,0,0" IsChecked="True">Printer Mappings</CheckBox>
        </StackPanel>
        <Button Name="btnStart" Width="70" Height="30" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="10,420,0,10" Content="Start"/>
        <Button Name="btnReset" Width="70" Height="30" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="85,420,0,10" Content="Reset Form"/>
    </Grid>
</Window>
'@

# Initialize form
$reader = New-Object System.Xml.XmlNodeReader $XAML
try {
    $Form = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Write-Error "Failed to load form: $_"
    exit 1
}

# Initialize form controls
$controls = @(
    'TxtSaveLoc', 'lblFreeSpace', 'lblRequiredSpace', 'lvwFileList',
    'btnSavBrowse', 'btnAddLoc', 'btnStart', 'btnReset',
    'chk_networkDrives', 'chk_Printers'
)

foreach ($control in $controls) {
    Set-Variable -Name $control -Value ($Form.FindName($control))
}

# Wire up event handlers
$btnSavBrowse.Add_click({ Set-SaveLocation })
$btnAddLoc.Add_Click({ Add-BackupLocation })
$btnStart.Add_Click({ Start-BRProcess })
$btnReset.Add_Click({
    $lvwFileList.ItemsSource = $null
    $TxtSaveLoc.Text = $null
    Initialize-Form
})
$lvwFileList.Add_SelectionChanged({ Update-BackupSize })
$Form.Add_ContentRendered({ Initialize-Form })

# Show the form
$Form.ShowDialog()

