﻿#requires -Version 3.0
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
. "$PSScriptRoot\src\Get-BackupPaths.ps1"

# Show mode selection dialog
function Show-ModeDialog {
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

    $result = $form.ShowDialog()
    $form.Dispose()
    
    return ($result -eq [System.Windows.Forms.DialogResult]::OK)
}

# Show main window
function Show-MainWindow {
    param(
        [bool]$IsBackup
    )

    try {
        # Parse XAML
        $reader = New-Object System.Xml.XmlNodeReader $script:XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        
        # Find controls
        $controls = @{}
        "txtSaveLoc", "btnBrowse", "btnStart", "lblMode", "lblStatus", "lvwFiles",
        "btnAddFile", "btnAddFolder", "btnRemove", "chkNetwork", "chkPrinters",
        "prgProgress", "txtProgress" | ForEach-Object {
            $controls[$_] = $window.FindName($_)
        }

        # Set mode and button text
        $controls.lblMode.Content = if ($IsBackup) { "Mode: Backup" } else { "Mode: Restore" }
        $controls.btnStart.Content = if ($IsBackup) { "Backup" } else { "Restore" }

        # Set default path
        $defaultPath = "C:\LocalData"
        if (-not (Test-Path $defaultPath)) {
            New-Item -Path $defaultPath -ItemType Directory -Force | Out-Null
        }
        $controls.txtSaveLoc.Text = $defaultPath

        # Load initial items based on mode
        if ($IsBackup) {
            # Load default backup paths
            $paths = Get-BackupPaths
            foreach ($path in $paths) {
                $controls.lvwFiles.Items.Add($path)
            }
        }
        elseif (Test-Path $defaultPath) {
            # Look for most recent backup
            $latestBackup = Get-ChildItem -Path $defaultPath -Directory |
                Where-Object { $_.Name -like "Backup_*" } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                $controls.lvwFiles.Items.Clear()
                Get-ChildItem -Path $latestBackup.FullName | 
                    Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } | 
                    ForEach-Object {
                        $controls.lvwFiles.Items.Add([PSCustomObject]@{
                            Name = $_.Name
                            Type = if ($_.PSIsContainer) { "Folder" } else { "File" }
                            Path = $_.FullName
                        })
                    }
            }
        }

        # Add event handlers
        $controls.btnBrowse.Add_Click({
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($IsBackup) { "Select location to save backup" } else { "Select backup to restore from" }
            $dialog.SelectedPath = $controls.txtSaveLoc.Text
            
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $controls.txtSaveLoc.Text = $dialog.SelectedPath
                
                # If restoring, load backup contents
                if (-not $IsBackup) {
                    $controls.lvwFiles.Items.Clear()
                    Get-ChildItem -Path $dialog.SelectedPath | 
                        Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } | 
                        ForEach-Object {
                            $controls.lvwFiles.Items.Add([PSCustomObject]@{
                                Name = $_.Name
                                Type = if ($_.PSIsContainer) { "Folder" } else { "File" }
                                Path = $_.FullName
                            })
                        }
                }
            }
        })

        $controls.btnAddFile.Add_Click({
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Multiselect = $true
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                foreach ($file in $dialog.FileNames) {
                    $controls.lvwFiles.Items.Add([PSCustomObject]@{
                        Name = [System.IO.Path]::GetFileName($file)
                        Type = "File"
                        Path = $file
                    })
                }
            }
        })

        $controls.btnAddFolder.Add_Click({
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $controls.lvwFiles.Items.Add([PSCustomObject]@{
                    Name = [System.IO.Path]::GetFileName($dialog.SelectedPath)
                    Type = "Folder"
                    Path = $dialog.SelectedPath
                })
            }
        })

        $controls.btnRemove.Add_Click({
            while ($controls.lvwFiles.SelectedItems.Count -gt 0) {
                $controls.lvwFiles.Items.Remove($controls.lvwFiles.SelectedItems[0])
            }
        })

        $controls.btnStart.Add_Click({
            if ([string]::IsNullOrEmpty($controls.txtSaveLoc.Text)) {
                [System.Windows.MessageBox]::Show(
                    "Please select a location first.",
                    "Required Field",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning
                )
                return
            }

            try {
                $controls.lblStatus.Content = if ($IsBackup) { "Backing up..." } else { "Restoring..." }
                $backupPath = $controls.txtSaveLoc.Text

                if ($IsBackup) {
                    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
                    $backupPath = Join-Path $backupPath "Backup_$timestamp"
                    New-Item -Path $backupPath -ItemType Directory -Force | Out-Null

                    # Create CSV file for logging backup locations
                    try {
                        $csvPath = Join-Path $backupPath "FileList_Backup.csv"
                        "SourceLocation,FileName" | Set-Content -Path $csvPath
                    }
                    catch {
                        Write-Warning "Failed to create backup log file: $($_.Exception.Message)"
                    }
                }

                # Count total files
                $totalFiles = 0
                foreach ($item in $controls.lvwFiles.Items) {
                    if ($item.Type -eq "Folder") {
                        try {
                            $count = (Get-ChildItem -Path $item.Path -Recurse -File).Count
                            $totalFiles += $count
                        } catch {
                            Write-Warning "Could not access folder $($item.Path): $($_.Exception.Message)"
                        }
                    } else {
                        $totalFiles++
                    }
                }
                
                if ($controls.chkNetwork.IsChecked) { $totalFiles++ }
                if ($controls.chkPrinters.IsChecked) { $totalFiles++ }
                
                $controls.prgProgress.Maximum = $totalFiles
                $controls.prgProgress.Value = 0

                # Process files
                foreach ($item in $controls.lvwFiles.Items) {
                    try {
                        $controls.txtProgress.Text = "Processing: $($item.Name)"
                        
                        if ($IsBackup) {
                            $dest = Join-Path $backupPath $item.Name
                            if ($item.Type -eq "Folder") {
                                # Add folder to CSV log
                                try {
                                    $csvPath = Join-Path $backupPath "FileList_Backup.csv"
                                    "`"$($item.Path)`",`"$($item.Name)`"" | Add-Content -Path $csvPath
                                } catch {
                                    Write-Warning "Failed to log folder location: $($_.Exception.Message)"
                                }
                                Get-ChildItem -Path $item.Path -Recurse -File | ForEach-Object {
                                    $relativePath = $_.FullName.Substring($item.Path.Length)
                                    $targetPath = Join-Path $dest $relativePath
                                    $targetDir = [System.IO.Path]::GetDirectoryName($targetPath)
                                    
                                    if (-not (Test-Path $targetDir)) {
                                        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                                    }
                                    
                                    Copy-Item -Path $_.FullName -Destination $targetPath -Force
                                    $controls.prgProgress.Value++
                                    $controls.txtProgress.Text = "Processing: $($_.Name)"
                                }
                            } else {
                                # Add file to CSV log
                                try {
                                    $csvPath = Join-Path $backupPath "FileList_Backup.csv"
                                    "`"$([System.IO.Path]::GetDirectoryName($item.Path))`",`"$($item.Name)`"" | Add-Content -Path $csvPath
                                } catch {
                                    Write-Warning "Failed to log file location: $($_.Exception.Message)"
                                }
                                Copy-Item -Path $item.Path -Destination $dest -Force
                                $controls.prgProgress.Value++
                            }
                        } else {
                            $targetPath = Join-Path $env:USERPROFILE $item.Name
                            Copy-Item -Path $item.Path -Destination $targetPath -Recurse -Force
                            $controls.prgProgress.Value++
                        }
                    } catch {
                        Write-Warning "Failed to process $($item.Name): $($_.Exception.Message)"
                    }
                }

                # Process network drives
                if ($controls.chkNetwork.IsChecked) {
                    $controls.txtProgress.Text = "Processing network drives..."
                    if ($IsBackup) {
                        Get-WmiObject -Class Win32_MappedLogicalDisk | 
                            Select-Object Name, ProviderName |
                            Export-Csv -Path (Join-Path $backupPath "Drives.csv") -NoTypeInformation
                    } else {
                        $drivesPath = Join-Path $backupPath "Drives.csv"
                        if (Test-Path $drivesPath) {
                            Import-Csv $drivesPath | ForEach-Object {
                                $driveExists = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DeviceID -eq $_.Name }
                                if (-not $driveExists) {
                                    New-PSDrive -Name $_.Name.Substring(0,1) -PSProvider FileSystem -Root $_.ProviderName -Persist
                                }
                            }
                        }
                    }
                    $controls.prgProgress.Value++
                }

                # Process printers
                if ($controls.chkPrinters.IsChecked) {
                    $controls.txtProgress.Text = "Processing printers..."
                    if ($IsBackup) {
                        Get-WmiObject -Class Win32_Printer | 
                            Where-Object { -not $_.Local } |
                            Select-Object -ExpandProperty Name |
                            Set-Content -Path (Join-Path $backupPath "Printers.txt")
                    } else {
                        $printersPath = Join-Path $backupPath "Printers.txt"
                        if (Test-Path $printersPath) {
                            Get-Content $printersPath | ForEach-Object {
                                Add-Printer -ConnectionName $_
                            }
                        }
                    }
                    $controls.prgProgress.Value++
                }
                
                $controls.txtProgress.Text = "Operation completed successfully"
                $controls.prgProgress.Value = $controls.prgProgress.Maximum
                
                [System.Windows.MessageBox]::Show(
                    $(if ($IsBackup) { "Backup completed successfully!" } else { "Restore completed successfully!" }),
                    "Success",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Information
                )
                
                $controls.lblStatus.Content = "Operation completed successfully"
            }
            catch {
                $controls.txtProgress.Text = "Operation failed"
                $controls.lblStatus.Content = "Operation failed"
                [System.Windows.MessageBox]::Show(
                    "Operation failed: $($_.Exception.Message)",
                    "Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                )
            }
        })

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
}

# Main execution
try {
    $script:isBackup = Show-ModeDialog
    Show-MainWindow -IsBackup $script:isBackup
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
