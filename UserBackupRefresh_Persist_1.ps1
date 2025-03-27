﻿﻿#requires -Version 3.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# XAML UI Definition
[xml]$script:XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Data Backup Tool"
    Width="800"
    Height="600"
    WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Location:" Margin="10,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="txtSaveLoc" Width="400" Height="30" Margin="10,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnBrowse" Content="Browse" Width="60" Height="30" Margin="420,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Name="lblMode" Content="" Margin="500,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold"/>
        
        <Label Content="Files/Folders:" Margin="10,80,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,110,200,140" SelectionMode="Multiple">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="200"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="100"/>
                    <GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="300"/>
                </GridView>
            </ListView.View>
        </ListView>
        
        <StackPanel Margin="0,110,10,0" HorizontalAlignment="Right" Width="180">
            <Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Remove Selected" Width="120" Height="30" Margin="0,0,0,20"/>
            <CheckBox Name="chkNetwork" Content="Network Drives" IsChecked="True" Margin="0,0,0,5"/>
            <CheckBox Name="chkPrinters" Content="Printers" IsChecked="True" Margin="0,0,0,5"/>
        </StackPanel>
        
        <ProgressBar Name="prgProgress" Height="20" Margin="10,0,10,60" VerticalAlignment="Bottom"/>
        <TextBlock Name="txtProgress" Text="" Margin="10,0,10,85" VerticalAlignment="Bottom"/>
        <Button Name="btnStart" Content="Start" Width="100" Height="30" Margin="10,0,0,20" VerticalAlignment="Bottom"/>
        <Label Name="lblStatus" Content="" Margin="90,0,0,20" HorizontalAlignment="Left" VerticalAlignment="Bottom"/>
    </Grid>
</Window>
'@

# Dot source function files
. "$PSScriptRoot\src\Set-GPupdate.ps1"
. "$PSScriptRoot\src\Add-DefaultPaths.ps1"
. "$PSScriptRoot\src\Get-FileCount.ps1"
. "$PSScriptRoot\src\Restore-NetworkDrives.ps1"
. "$PSScriptRoot\src\Restore-Printers.ps1"
. "$PSScriptRoot\src\Initialize-MainWindow.ps1"

function New-FilePathObject {
    param (
        [string]$Path,
        [string]$Name,
        [string]$Type
    )
    
    if (-not $Name) { 
        $Name = Split-Path $Path -Leaf 
    }
    
    if (-not $Type) { 
        if (Test-Path -Path $Path -PathType Container) {
            $Type = "Folder"
        } else {
            $Type = "File"
        }
    }
    
    Write-Output ([PSCustomObject]@{
        Path = $Path
        Name = $Name
        Type = $Type
    })
}

# Initialize default paths array
$script:DefaultPaths = @()

# Add Outlook signatures if they exist
if(Test-Path -Path "$env:APPDATA\Microsoft\Signatures") {
    $script:DefaultPaths += New-FilePathObject -Path "$env:APPDATA\Microsoft\Signatures" -Name "Outlook Signatures" -Type "Folder"
}

# Add User folder if it exists
if(Test-Path -Path "$env:SystemDrive\User") {
    $script:DefaultPaths += New-FilePathObject -Path "$env:SystemDrive\User" -Name "User Directory" -Type "Folder"
}

# Add Quick Access if it exists
if(Test-Path -Path "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms") {
    $script:DefaultPaths += New-FilePathObject -Path "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms" -Name "Quick Access" -Type "File"
}

# Add Temp folder if it exists
if(Test-Path -Path "$env:SystemDrive\Temp") {
    $script:DefaultPaths += New-FilePathObject -Path "$env:SystemDrive\Temp" -Name "Temp Directory" -Type "Folder"
}

# Add Sticky Notes data file (older version) if it exists
if(Test-Path -Path "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt") {
    $script:DefaultPaths += New-FilePathObject -Path "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt" -Name "Sticky Notes (Legacy)" -Type "File"
}

# Add Sticky Notes data file (Windows 10 v1607+) if it exists
$stickyNotesPath = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
if(Test-Path -Path $stickyNotesPath) {
    $script:DefaultPaths += New-FilePathObject -Path $stickyNotesPath -Name "Sticky Notes" -Type "File"
}

# Add Google Earth KML if it exists
if(Test-Path -Path "$env:APPDATA\google\googleearth\myplaces.kml") {
    $script:DefaultPaths += New-FilePathObject -Path "$env:APPDATA\google\googleearth\myplaces.kml" -Name "Google Earth Places" -Type "File"
}

# Add Chrome bookmarks if they exist and are accessible
$chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
if (Test-Path $chromePath) {
    try {
        Get-Content $chromePath -ErrorAction Stop | Out-Null
        $script:DefaultPaths += New-FilePathObject -Path $chromePath -Name "Chrome Bookmarks" -Type "File"
    }
    catch {
        Write-Warning "Chrome bookmarks file not accessible"
    }
}

# Main Execution
try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation"
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(50,40)
    $btnBackup.Size = New-Object System.Drawing.Size(80,30)
    $btnBackup.Text = "Backup"
    $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(150,40)
    $btnRestore.Size = New-Object System.Drawing.Size(80,30)
    $btnRestore.Text = "Restore"
    $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnRestore)

    $form.AcceptButton = $btnBackup
    $form.CancelButton = $btnRestore

    $script:isBackup = $form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK
    $form.Dispose()
    
    $window = Initialize-MainWindow
    $window.FindName('lblMode').Content = if ($script:isBackup) { "Mode: Backup" } else { "Mode: Restore" }
    
    if ($script:isBackup) {
        Add-DefaultPaths
    }
    
    $window.ShowDialog() | Out-Null
}
catch {
    Write-Error $_.Exception.Message
    [System.Windows.MessageBox]::Show(
        "An error occurred: $($_.Exception.Message)",
        "Error",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    )
}
