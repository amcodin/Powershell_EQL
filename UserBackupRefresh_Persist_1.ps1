﻿﻿﻿﻿#requires -Version 3.0
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
        <Button Name="btnStart" Content="Start" Width="70" Height="30" Margin="10,0,0,20" VerticalAlignment="Bottom"/>
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

# Default paths to backup/restore with safer user folder handling
$script:DefaultPaths = @(
    @{
        Path = "$env:APPDATA\WinRAR"
        Name = "Outlook Signatures"
        Type = "Folder"
    }
)

# Add Chrome bookmarks only if accessible
$chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
if (Test-Path $chromePath) {
    try {
        $null = Get-Content $chromePath -ErrorAction Stop
        $script:DefaultPaths += @{Path = $chromePath; Name = "Chrome Bookmarks"; Type = "File"}
    }
    catch {
        Write-Warning "Chrome bookmarks file not accessible"
    }
}

# Main Execution
try {
    $result = [System.Windows.MessageBox]::Show("Would you like to perform a backup?`n`nClick 'Yes' for Backup`nClick 'No' for Restore", "Select Operation", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    $script:isBackup = $result -eq [System.Windows.MessageBoxResult]::Yes
    
    $window = Initialize-MainWindow
    $window.FindName('lblMode').Content = if ($script:isBackup) {"Mode: Backup"} else {"Mode: Restore"}
    
    if ($script:isBackup) {
        Add-DefaultPaths
    }
    
    $window.ShowDialog() | Out-Null
}
catch {
    Write-Error $_.Exception.Message
    [System.Windows.MessageBox]::Show("An error occurred: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}
