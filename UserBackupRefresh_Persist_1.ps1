﻿#Requires -Version 3.0

# Load required assemblies 
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Configure script behavior
$VerbosePreference = 'Continue'
$script:DefaultSavePath = "$env:HOMEDRIVE\LocalData"
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Add required assembly and custom class
Add-Type -AssemblyName presentationframework
Add-Type -Language CSharp -TypeDefinition @' 
public class RemoteInfo {
    public string Computername;
    public string ProfileFolder;
    public string ProfileName;
    public bool IsInUse;
    public string SSID;
    public Microsoft.Win32.RegistryKey UserHive;
};
'@

#XAML For main screen.
[xml]$XAML = 
@'
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

# Show-UserPrompt: Displays a dialog with customizable buttons
# Usage:
#   Single OK button: Show-UserPrompt -Message "message"
#   Two buttons: Show-UserPrompt -Message "message" -Button2 "Yes" -Button3 "No"
#   Three buttons: Show-UserPrompt -Message "message" -Button1 "Retry" -Button2 "Ignore" -Button3 "Cancel"
# Note: Button ordering is right-to-left, with Button3 rightmost
Function Show-UserPrompt {
    Param(
        [parameter(Mandatory)][String]$Message,
        [string]$Button1 = $null,
        [string]$Button2 = $null,
        [string]$Button3 = 'OK'
    )
    
    Add-Type -AssemblyName PresentationFramework
    [xml]$ErrorForm = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="CanResize"
    SizeToContent="WidthAndHeight" 
    MaxWidth="500"
    Title="User Prompt"
    WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,0,0">
        <Label Margin="10,10,10,100">
            <TextBlock TextWrapping="WrapWithOverflow" Text="{Binding Path=Message}"/>
        </Label>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Bottom" HorizontalAlignment="Right" Margin="0,0,5,5" Height="100" Width="Auto" MaxWidth="250">
            <Button Name="BtnOne" Width="Auto" Height="Auto" HorizontalAlignment="Left" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="{Binding Path=ButtonOne}"/>
            <Button Name="BtnTwo" Width="Auto" Height="Auto" HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="{Binding Path=ButtonTwo}"/>
            <Button Name="BtnThree" Width="Auto" Height="Auto" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="5,5,5,5" Padding="10" Content="{Binding Path=ButtonThree}"/>
        </StackPanel>
    </Grid>
</Window>
'@

    try {
        $reader = New-Object System.Xml.XmlNodeReader $ErrorForm
        $ErrorWindow = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        # Fallback to standard MessageBox if XAML window fails
        $null = [System.Windows.MessageBox]::Show($Message, 'User Prompt', 'OK')
        Write-Warning "Failed to create custom dialog: $($_.Exception.Message)"
        return 'OK'
    }

    $ErrorDetails = [PsCustomObject]@{ 
        Message     = $Message
        ButtonOne   = $Button1
        ButtonTwo   = $Button2
        ButtonThree = $Button3
    }

    $btnOne = $ErrorWindow.Findname('BtnOne')
    $btnTwo = $ErrorWindow.Findname('BtnTwo')
    $btnThree = $ErrorWindow.Findname('BtnThree')

    if([string]::IsNullOrEmpty($Button1)) {
        $btnOne.Visibility = 'Hidden'
    } else {
        $btnOne.Visibility = 'Visible'
    }
    if([string]::IsNullOrEmpty($Button2)) {
        $btnTwo.Visibility = 'Hidden'
    } else {
        $btnTwo.Visibility = 'Visible'
    }
    if([string]::IsNullOrEmpty($Button3)) {
        $btnThree.Visibility = 'Hidden'
    } else {
        $btnThree.Visibility = 'Visible'
    }

    $ErrorWindow.DataContext = $ErrorDetails

    [string]$Return = ''
    $btnOne.Add_Click({
        $ErrorWindow.Close()
        Set-Variable -Name 'Return' -Value $btnOne.Content -Scope 1
    })
    $btnTwo.Add_Click({
        $ErrorWindow.Close()
        Set-Variable -Name 'Return' -Value $btnTwo.Content -Scope 1
    })
    $btnThree.Add_Click({
        $ErrorWindow.Close()
        Set-Variable -Name 'Return' -Value $btnThree.Content -Scope 1
    })

    if([string]::IsNullOrEmpty($Button1) -and [string]::IsNullOrEmpty($Button2) -and $Button3 -eq 'OK') {
        $null = [System.Windows.MessageBox]::Show($Message, 'User Prompt', 'OK')
        return 'OK'
    } else {
        $null = $ErrorWindow.ShowDialog()
        return $Return
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
        Show-UserPrompt -Message ("{0} completed successfully" -f $Operation)
        return $true
    }
    catch {
        Write-Warning ("{0} operation failed: {1}" -f $Operation, $_.Exception.Message)
        Show-UserPrompt -Message "$Operation failed: $($_.Exception.Message)"
        return $false
    }
}

# Helper function for space calculations
function Update-SpaceLabels {
    if ($script:controls.TxtSaveLoc.Text) {
        try {
            $driveLetter = $script:controls.TxtSaveLoc.Text[0]
            $filter = "DeviceID='$driveLetter`:'"
            $drive = Get-WmiObject -Class Win32_LogicalDisk -Filter $filter -ErrorAction Stop
            if ($drive) {
                $freeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
                $script:controls.lblFreeSpace.Content = "Free Space: $freeSpace GB"
            }
            else {
                $script:controls.lblFreeSpace.Content = "Free Space: Unable to determine"
            }
        }
        catch {
            Write-Warning ("Failed to get drive space: {0}" -f $_.Exception.Message)
            $script:controls.lblFreeSpace.Content = "Free Space: Error"
        }
    }
    
    $totalSize = 0
    foreach ($item in $script:controls.lvwFileList.Items) {
        if ($item.Selected) {
            $totalSize += $item.Size
        }
    }
    $script:controls.lblRequiredSpace.Content = "Required Space: $([math]::Round($totalSize / 1024, 2)) GB"
}

# Initialize Form
function Initialize-Form {
    # Check admin status first
    [bool]$script:Admin = ([Security.Principal.WindowsPrincipal]([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    # Get initial user choices
    [string]$Script:Action2 = Show-UserPrompt -Message 'Run A GPUPDATE?' -Button1 '' -Button2 'Yes' -Button3 'No'
    if ($Script:Action2 -eq 'Yes') {
        Write-Host "Running GPUpdate..." -ForegroundColor Cyan
        Set-GPupdate
    }

    [string]$Script:Action = Show-UserPrompt -Message 'Are we running a Backup or Restore?' -Button1 '' -Button2 'Restore' -Button3 'Backup'

    # Initialize GUI
    $reader = New-Object System.Xml.XmlNodeReader $XAML
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Get controls
    $script:controls = @{}
    $XAML.SelectNodes("//*[@Name]") | ForEach-Object {
        $script:controls[$_.Name] = $window.FindName($_.Name)
    }

    # Setup initial state based on operation
    if ($Script:Action -eq 'Backup') {
        # Scan for default paths
        foreach ($category in $script:BackupPaths.Keys) {
            foreach ($item in $script:BackupPaths[$category]) {
                if (Test-Path $item.Path) {
                    $size = 0
                    Get-ChildItem $item.Path -Recurse -File -ErrorAction SilentlyContinue | 
                        ForEach-Object { $size += $_.Length }
                    
                    $listItem = [PSCustomObject]@{
                        Name = Split-Path -Leaf $item.Path
                        FullPath = $item.Path
                        Type = if (Test-Path -Path $item.Path -PathType Container) {'Folder'} else {'File'}
                        Size = [math]::Round($size / 1MB, 2)
                        Selected = $true
                    }
                    $script:controls.lvwFileList.Items.Add($listItem)
                }
            }
        }

        # Set default save location
        if (-not (Test-Path $script:DefaultSavePath)) {
            New-Item -Path $script:DefaultSavePath -ItemType Directory -Force
        }
        $script:controls.TxtSaveLoc.Text = $script:DefaultSavePath
    }
    else {
        # Handle restore operation
        $script:controls.btnAddLoc.IsEnabled = $false
        $script:controls.btnSavBrowse.Content = "Select Backup"
    }

    # Setup event handlers
    $script:controls.btnSavBrowse.Add_Click({
        $dialog = $null
        try {
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($Script:Action -eq 'Backup') {"Select backup location"} else {"Select backup to restore"}
            $dialog.SelectedPath = $env:USERPROFILE
            
            if ($dialog.ShowDialog() -eq 'OK') {
                $script:controls.TxtSaveLoc.Text = $dialog.SelectedPath
                Update-SpaceLabels

                if ($Script:Action -eq 'Restore') {
                    # Load backup contents for restore
                    $script:controls.lvwFileList.Items.Clear()
                    Get-ChildItem -Path $dialog.SelectedPath -Exclude 'FileList*.csv','Drives.csv','Printers.txt' | ForEach-Object {
                        $size = if ($_.PSIsContainer) {
                            (Get-ChildItem $_.FullName -Recurse -File | Measure-Object -Property Length -Sum).Sum
                        } else {
                            $_.Length
                        }
                        
                        $script:controls.lvwFileList.Items.Add([PSCustomObject]@{
                            Name = $_.Name
                            FullPath = $_.FullName
                            Type = if ($_.PSIsContainer) {'Folder'} else {'File'}
                            Size = [math]::Round($size / 1MB, 2)
                            Selected = $true
                        })
                    }
                }
            }
        }
        catch {
            Write-Warning ("Failed to show folder browser: {0}" -f $_.Exception.Message)
        }
        finally {
            if ($dialog) {
                $dialog.Dispose()
            }
        }
    })

    $script:controls.btnAddLoc.Add_Click({
        $addType = Show-UserPrompt -Message "Select what to add" -Button1 '' -Button2 'File' -Button3 'Folder'
        
        if ($addType -eq 'Folder') {
            $dialog = $null
            try {
                $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
                $dialog.Description = "Select folder to backup"
                $dialog.SelectedPath = $env:USERPROFILE
                
                if ($dialog.ShowDialog() -eq 'OK') {
                    $folderInfo = Get-Item $dialog.SelectedPath
                    $size = 0
                    Get-ChildItem $dialog.SelectedPath -Recurse -File -ErrorAction SilentlyContinue | 
                        ForEach-Object { $size += $_.Length }
                    
                    $item = [PSCustomObject]@{
                        Name = $folderInfo.Name
                        FullPath = $folderInfo.FullName
                        Type = 'Folder'
                        Size = [math]::Round($size / 1MB, 2)
                        Selected = $true
                    }
                    $script:controls.lvwFileList.Items.Add($item)
                    Update-SpaceLabels
                }
            }
            catch {
                Write-Warning ("Failed to add folder: {0}" -f $_.Exception.Message)
            }
            finally {
                if ($dialog) {
                    $dialog.Dispose()
                }
            }
        }
        else {
            $dialog = $null
            try {
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog.Title = "Select file to backup"
                $dialog.InitialDirectory = $env:USERPROFILE
                $dialog.Multiselect = $true
                
                if ($dialog.ShowDialog() -eq 'OK') {
                    foreach ($file in $dialog.FileNames) {
                        $fileInfo = Get-Item $file
                        $item = [PSCustomObject]@{
                            Name = $fileInfo.Name
                            FullPath = $fileInfo.FullName
                            Type = 'File'
                            Size = [math]::Round($fileInfo.Length / 1MB, 2)
                            Selected = $true
                        }
                        $script:controls.lvwFileList.Items.Add($item)
                    }
                    Update-SpaceLabels
                }
            }
            catch {
                Write-Warning ("Failed to add file(s): {0}" -f $_.Exception.Message)
            }
            finally {
                if ($dialog) {
                    $dialog.Dispose()
                }
            }
        }
    })

    $script:controls.btnReset.Add_Click({
        $script:controls.TxtSaveLoc.Text = ""
        $script:controls.lvwFileList.Items.Clear()
        $script:controls.chk_networkDrives.IsChecked = $true
        $script:controls.chk_Printers.IsChecked = $true
        $script:controls.lblFreeSpace.Content = ""
        $script:controls.lblRequiredSpace.Content = ""
        
        if ($Script:Action -eq 'Backup') {
            Initialize-Form
        }
    })

    $script:controls.btnStart.Add_Click({
        if (-not $script:controls.TxtSaveLoc.Text) {
            Show-UserPrompt -Message "Please select a backup location"
            return
        }

        # Start backup/restore operation
        $result = Start-BackupRestore -Path $script:controls.TxtSaveLoc.Text -Operation $Script:Action
        
        if ($result -and $Script:Action -eq 'Restore') {
            $runGpUpdate = Show-UserPrompt -Message 'Run GPUpdate now?' -Button2 'Yes' -Button3 'No'
            if ($runGpUpdate -eq 'Yes') {
                Set-GPupdate
            }
        }
    })

    # Add event handler for ListView selection changes
    $script:controls.lvwFileList.Add_SelectionChanged({
        Update-SpaceLabels
    })

    # Initialize space labels
    Update-SpaceLabels

    return $window
}

# Main script execution
try {
    $mainWindow = Initialize-Form
    $mainWindow.ShowDialog()
}
catch {
    Write-Warning ("Script execution failed: {0}" -f $_.Exception.Message)
    Show-UserPrompt -Message "Script execution failed: $($_.Exception.Message)"
}
