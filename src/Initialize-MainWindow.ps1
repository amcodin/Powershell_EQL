function Initialize-MainWindow {
    # Create Window
    $reader = New-Object System.Xml.XmlNodeReader $XAML
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Set default path if C:\LocalData exists
    $defaultPath = "C:\LocalData"
    
    # Get Controls
    $script:txtSaveLoc = $window.FindName('txtSaveLoc')
    $script:btnBrowse = $window.FindName('btnBrowse')
    $script:btnStart = $window.FindName('btnStart')
    $script:lblMode = $window.FindName('lblMode')
    $script:lblStatus = $window.FindName('lblStatus')
    $script:lvwFiles = $window.FindName('lvwFiles')
    $script:btnAddFile = $window.FindName('btnAddFile')
    $script:btnAddFolder = $window.FindName('btnAddFolder')
    $script:btnRemove = $window.FindName('btnRemove')
    $script:chkNetwork = $window.FindName('chkNetwork')
    $script:chkPrinters = $window.FindName('chkPrinters')
    $script:prgProgress = $window.FindName('prgProgress')
    $script:txtProgress = $window.FindName('txtProgress')
    
    # Update button text to Backup/Restore
    $btnStart.Content = if ($script:isBackup) { "Backup" } else { "Restore" }

    # Set default path if it exists
    if (Test-Path $defaultPath) {
        $script:txtSaveLoc.Text = $defaultPath
    }

    # Setup Event Handlers
    $btnBrowse.Add_Click({
        if ($script:isBackup -and (Test-Path $defaultPath)) {
            $script:txtSaveLoc.Text = $defaultPath
        } else {
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($script:isBackup) {
                "Select location to save backup"
            } else {
                "Select backup to restore from"
            }
            
            $form = New-Object System.Windows.Forms.Form
            $result = $dialog.ShowDialog($form)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $script:txtSaveLoc.Text = $dialog.SelectedPath
                
                # If restoring, load the backup contents
                if (-not $script:isBackup) {
                    $script:lvwFiles.Items.Clear()
                    Get-ChildItem -Path $dialog.SelectedPath | 
                        Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } | 
                        ForEach-Object {
                            $script:lvwFiles.Items.Add([PSCustomObject]@{
                                Name = $_.Name
                                Type = if ($_.PSIsContainer) { "Folder" } else { "File" }
                                Path = $_.FullName
                            })
                        }
                }
            }
            
            $form.Dispose()
            $dialog.Dispose()
        }
    })
    
    $btnAddFile.Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Multiselect = $true
        
        $form = New-Object System.Windows.Forms.Form
        $result = $dialog.ShowDialog($form)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            foreach ($file in $dialog.FileNames) {
                $script:lvwFiles.Items.Add([PSCustomObject]@{
                    Name = [System.IO.Path]::GetFileName($file)
                    Type = "File"
                    Path = $file
                })
            }
        }
        
        $form.Dispose()
        $dialog.Dispose()
    })
    
    $btnAddFolder.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        
        $form = New-Object System.Windows.Forms.Form
        $result = $dialog.ShowDialog($form)
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:lvwFiles.Items.Add([PSCustomObject]@{
                Name = [System.IO.Path]::GetFileName($dialog.SelectedPath)
                Type = "Folder"
                Path = $dialog.SelectedPath
            })
        }
        
        $form.Dispose()
        $dialog.Dispose()
    })
    
    $btnRemove.Add_Click({
        while ($script:lvwFiles.SelectedItems.Count -gt 0) {
            $script:lvwFiles.Items.Remove($script:lvwFiles.SelectedItems[0])
        }
    })
    
    $btnStart.Add_Click({
        if ([string]::IsNullOrEmpty($script:txtSaveLoc.Text)) {
            [System.Windows.MessageBox]::Show(
                "Please select a location first.", 
                "Required Field",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        try {
            $script:lblStatus.Content = if ($script:isBackup) {"Backing up..."} else {"Restoring..."}
            $backupPath = $script:txtSaveLoc.Text
            
            if ($script:isBackup) {
                # Create backup folder with timestamp
                $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
                $backupPath = Join-Path $backupPath "Backup_$timestamp"
                New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
            }

            # Count total files for progress bar
            $totalFiles = 0
            foreach ($item in $script:lvwFiles.Items) {
                if ($item.Type -eq "Folder") {
                    try {
                        $totalFiles += Get-FileCount -Path $item.Path
                    }
                    catch {
                        Write-Warning "Could not access folder $($item.Path): $($_.Exception.Message)"
                    }
                } else {
                    $totalFiles++
                }
            }
            
            if ($script:chkNetwork.IsChecked) { $totalFiles++ }
            if ($script:chkPrinters.IsChecked) { $totalFiles++ }
            
            $script:prgProgress.Maximum = $totalFiles
            $script:prgProgress.Value = 0

            # Files/Folders
            foreach ($item in $script:lvwFiles.Items) {
                try {
                    $script:txtProgress.Text = "Processing: $($item.Name)"
                    
                    if ($script:isBackup) {
                        $dest = Join-Path $backupPath $item.Name
                        if ($item.Type -eq "Folder") {
                            Get-ChildItem -Path $item.Path -Recurse -File | ForEach-Object {
                                $relativePath = $_.FullName.Substring($item.Path.Length)
                                $targetPath = Join-Path $dest $relativePath
                                $targetDir = [System.IO.Path]::GetDirectoryName($targetPath)
                                
                                if (-not (Test-Path $targetDir)) {
                                    New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                                }
                                
                                Copy-Item -Path $_.FullName -Destination $targetPath -Force
                                $script:prgProgress.Value++
                                $script:txtProgress.Text = "Processing: $($_.Name)"
                            }
                        } else {
                            Copy-Item -Path $item.Path -Destination $dest -Force
                            $script:prgProgress.Value++
                        }
                    } else {
                        $targetPath = Join-Path $env:USERPROFILE $item.Name
                        Copy-Item -Path $item.Path -Destination $targetPath -Recurse -Force
                        $script:prgProgress.Value++
                    }
                }
                catch {
                    Write-Warning "Failed to process $($item.Name): $($_.Exception.Message)"
                }
            }

            # Network Drives
            if ($script:chkNetwork.IsChecked) {
                $script:txtProgress.Text = "Processing network drives..."
                if ($script:isBackup) {
                    Get-WmiObject -Class Win32_MappedLogicalDisk | 
                        Select-Object Name, ProviderName |
                        Export-Csv -Path (Join-Path $backupPath "Drives.csv") -NoTypeInformation
                } else {
                    Restore-NetworkDrives -Path $backupPath
                }
                $script:prgProgress.Value++
            }

            # Printers
            if ($script:chkPrinters.IsChecked) {
                $script:txtProgress.Text = "Processing printers..."
                if ($script:isBackup) {
                    Get-WmiObject -Class Win32_Printer | 
                        Where-Object { -not $_.Local } |
                        Select-Object -ExpandProperty Name |
                        Set-Content -Path (Join-Path $backupPath "Printers.txt")
                } else {
                    Restore-Printers -Path $backupPath
                }
                $script:prgProgress.Value++
            }
            
            $script:txtProgress.Text = "Operation completed successfully"
            $script:prgProgress.Value = $script:prgProgress.Maximum
            
            [System.Windows.MessageBox]::Show(
                $(if ($script:isBackup) {"Backup completed successfully!"} else {"Restore completed successfully!"}),
                "Success",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information)
            
            if ($script:isBackup) {
                Set-GPupdate
            }
            
            $script:lblStatus.Content = "Operation completed successfully"
        }
        catch {
            $script:txtProgress.Text = "Operation failed"
            $script:lblStatus.Content = "Operation failed"
            [System.Windows.MessageBox]::Show(
                "Operation failed: $($_.Exception.Message)",
                "Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error)
        }
    })
    
    return $window
}
