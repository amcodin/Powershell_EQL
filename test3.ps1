# working, need to update checkboxes in backup mode
#requires -Version 3.0
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

#region Functions

# --- Embedded Functions (Removed Dot-Sourcing) ---

function Get-BackupPaths {
    [CmdletBinding()]
    param ()

    # Define specific paths/files to check
    $specificPaths = @(
        "$env:APPDATA\Microsoft\Signatures",
        # "$env:SystemDrive\User", # Commented out - often too generic or permission issues. Add back if specifically needed.
        "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms", # Quick Access Pinned
        # "$env:SystemDrive\Temp", # Commented out - Temp folders usually not backed up. Add back if specifically needed.
        "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt", # Legacy Sticky Notes
        "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite", # Modern Sticky Notes
        "$env:APPDATA\google\googleearth\myplaces.kml" # Google Earth Places
        # Add other common paths as needed, e.g., specific application data
    )

    # *** FIX: Initialize list and loop through enums ***
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $standardFolderEnums = @(
        [Environment+SpecialFolder]::Desktop,
        [Environment+SpecialFolder]::Documents,
        [Environment+SpecialFolder]::Downloads,
        [Environment+SpecialFolder]::Pictures,
        [Environment+SpecialFolder]::Music,
        [Environment+SpecialFolder]::Videos,
        [Environment+SpecialFolder]::Favorites
    )

    foreach ($folderEnum in $standardFolderEnums) {
        try {
            $folderPath = [Environment]::GetFolderPath($folderEnum)
            if (Test-Path -Path $folderPath -PathType Container) {
                $result.Add([PSCustomObject]@{
                    Name = Split-Path $folderPath -Leaf
                    Path = $folderPath
                    Type = "Folder"
                    # IsSelected = $true # Keep commented unless needed for Backup mode checkboxes
                })
            } else {
                Write-Verbose "Standard folder path does not exist or is not a container: $folderPath (Enum: $folderEnum)"
            }
        } catch {
            Write-Warning "Could not get path for SpecialFolder '$folderEnum': $($_.Exception.Message)"
        }
    }
    # *** End FIX ***


    # Add specific paths if they exist
    foreach ($path in $specificPaths) {
        if (Test-Path -Path $path) {
            # Avoid adding duplicates
            if (-not ($result.Path -contains $path)) {
                $result.Add([PSCustomObject]@{
                    Name = Split-Path $path -Leaf
                    Path = $path
                    Type = if (Test-Path -Path $path -PathType Container) { "Folder" } else { "File" }
                    # IsSelected = $true
                })
            }
        } else {
             Write-Verbose "Specific path not found: $path"
        }
    }

    # Add Chrome bookmarks if accessible
    $chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
    try {
        if (Test-Path $chromePath -PathType Leaf) {
            Get-Content $chromePath -TotalCount 1 -ErrorAction Stop | Out-Null # Test read access silently
             if (-not ($result.Path -contains $chromePath)) {
                $result.Add([PSCustomObject]@{
                    Name = "Chrome Bookmarks"
                    Path = $chromePath
                    Type = "File"
                    # IsSelected = $true
                })
            }
        }
    } catch {
        Write-Warning "Chrome bookmarks file not found or not accessible at $chromePath"
    }

    # Add Edge favorites (Chromium Edge)
    $edgePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks"
     try {
        if (Test-Path $edgePath -PathType Leaf) {
            Get-Content $edgePath -TotalCount 1 -ErrorAction Stop | Out-Null # Test read access silently
             if (-not ($result.Path -contains $edgePath)) {
                $result.Add([PSCustomObject]@{
                    Name = "Edge Bookmarks"
                    Path = $edgePath
                    Type = "File"
                    # IsSelected = $true
                })
            }
        }
    } catch {
        Write-Warning "Edge bookmarks file not found or not accessible at $edgePath"
    }

    # Add Outlook Signatures explicitly if not already covered by APPDATA
    $outlookSignaturesPath = "$env:APPDATA\Microsoft\Signatures"
    if (Test-Path $outlookSignaturesPath -PathType Container) {
        if (-not ($result.Path -contains $outlookSignaturesPath)) {
             $result.Add([PSCustomObject]@{
                Name = "Outlook Signatures"
                Path = $outlookSignaturesPath
                Type = "Folder"
                # IsSelected = $true
            })
        }
    }

    return $result
}

function Set-GPupdate {
    Write-Host "Initiating Group Policy update..." -ForegroundColor Cyan
    try {
        # Using cmd /c ensures the window closes automatically after gpupdate finishes
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
        if ($process.ExitCode -eq 0) {
            Write-Host "Group Policy update completed successfully." -ForegroundColor Green
        } else {
            Write-Warning "Group Policy update process finished with exit code: $($process.ExitCode)."
        }
    } catch {
         Write-Error "Failed to start GPUpdate process: $($_.Exception.Message)"
    }
}

function Start-ConfigManagerActions {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseModule # Parameter to optionally try using the CM module (might require admin rights/module install)
    )

    Write-Verbose "Attempting to trigger Configuration Manager client actions..."
    $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"
    $clientSDKNamespace = "root\ccm\clientsdk"
    $clientClassName = "CCM_ClientUtilities" # More common class for triggering schedules
    $scheduleMethodName = "TriggerSchedule"

    # Check if CM client exists
    if (-not (Test-Path $ccmExecPath)) {
        Write-Warning "Configuration Manager client not found at $ccmExecPath. Skipping CM actions."
        return $true # Return success as CM is not applicable
    }

    $scheduleActions = @(
        @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}
        # Add more schedule IDs as needed
    )

    $success = $true
    $attemptedMethod = "WMI/CIM"

    try {
        # Attempt using WMI/CIM first (preferred, less prone to process issues)
        if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
             Write-Verbose "Using CIM method to trigger schedules."
             foreach ($action in $scheduleActions) {
                Write-Verbose "Triggering $($action.Name) (ID: $($action.ID))"
                try {
                    Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop
                    Write-Verbose "$($action.Name) triggered successfully via CIM."
                } catch {
                    Write-Warning "Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"
                    $success = $false # Mark as partially failed if any action fails
                }
            }
        } else {
             Write-Warning "CIM Class $clientClassName not found in $clientSDKNamespace. Falling back to ccmexec.exe."
             $attemptedMethod = "ccmexec.exe"
             # Fallback to ccmexec.exe direct execution
             foreach ($action in $scheduleActions) {
                Write-Verbose "Triggering $($action.Name) via ccmexec.exe"
                # Note: Triggering via ccmexec.exe is less reliable and provides less feedback
                try {
                    # Using -TriggerSchedule requires the GUID format
                    $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop
                    if ($process.ExitCode -ne 0) {
                        Write-Warning "$($action.Name) action via ccmexec.exe finished with exit code $($process.ExitCode)"
                        # Don't mark as failure here, as ccmexec often returns non-zero even if schedule triggers
                    } else {
                         Write-Verbose "$($action.Name) triggered via ccmexec.exe (Exit Code 0)."
                    }
                } catch {
                     Write-Warning "Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"
                     $success = $false
                }
            }
        }

    } catch {
        Write-Error "An unexpected error occurred during Configuration Manager actions ($attemptedMethod): $($_.Exception.Message)"
        return $false # Return failure on major error
    }

    Write-Verbose "Configuration Manager actions attempt finished."
    return $success # Return true if all actions triggered successfully (or CM not present), false otherwise
}

# --- GUI Functions ---

# Show mode selection dialog
function Show-ModeDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation"
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.HelpButton = $false # Prevent closing via Esc without selection

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(50, 40)
    $btnBackup.Size = New-Object System.Drawing.Size(80, 30)
    $btnBackup.Text = "Backup"
    $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(150, 40)
    $btnRestore.Size = New-Object System.Drawing.Size(80, 30)
    $btnRestore.Text = "Restore"
    $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnRestore)

    # Set Backup as the default button (activated by Enter key)
    $form.AcceptButton = $btnBackup
    # Allow closing via Esc key by setting CancelButton (maps to Restore)
    $form.CancelButton = $btnRestore

    $result = $form.ShowDialog()
    $form.Dispose()

    # Determine mode based on which button was effectively clicked
    $isBackupMode = ($result -eq [System.Windows.Forms.DialogResult]::OK)

    # If restore mode, run updates immediately in a non-blocking way
    if (-not $isBackupMode) {
        Write-Host "Restore mode selected. Initiating background system updates..."
        # Use Start-Job for non-blocking execution
        $updateJob = Start-Job -ScriptBlock {
            # Functions need to be defined within the job's scope or passed via -ArgumentList

            function Set-GPupdate {
                Write-Host "JOB: Initiating Group Policy update..." -ForegroundColor Cyan
                try {
                    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
                    if ($process.ExitCode -eq 0) {
                        Write-Host "JOB: Group Policy update completed successfully." -ForegroundColor Green
                    } else {
                        Write-Warning "JOB: Group Policy update process finished with exit code: $($process.ExitCode)."
                    }
                } catch {
                    Write-Error "JOB: Failed to start GPUpdate process: $($_.Exception.Message)"
                }
            }

            function Start-ConfigManagerActions {
                 param() # Keep param block even if empty for consistency
                 Write-Host "JOB: Attempting to trigger Configuration Manager client actions..."
                 $clientSDKNamespace = "root\ccm\clientsdk"
                 $clientClassName = "CCM_ClientUtilities"
                 $scheduleMethodName = "TriggerSchedule"
                 $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"

                 if (-not (Test-Path $ccmExecPath)) { Write-Warning "JOB: CCM client not found."; return }

                 $scheduleActions = @(
                     @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy'},
                     @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy'},
                     @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory'},
                     @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory'},
                     @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan'}
                 )
                 try {
                     if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                         Write-Host "JOB: Using CIM method."
                         foreach ($action in $scheduleActions) {
                             Write-Host "JOB: Triggering $($action.Name)"
                             try { Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop }
                             catch { Write-Warning "JOB: Failed CIM trigger for $($action.Name): $($_.Exception.Message)" }
                         }
                     } else {
                         Write-Warning "JOB: CIM Class not found. Using ccmexec.exe fallback."
                         foreach ($action in $scheduleActions) {
                              Write-Host "JOB: Triggering $($action.Name) via ccmexec"
                              try { Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop | Out-Null }
                              catch { Write-Warning "JOB: Failed ccmexec trigger for $($action.Name): $($_.Exception.Message)"}
                         }
                     }
                 } catch { Write-Error "JOB: Error during CM actions: $($_.Exception.Message)" }
                 Write-Host "JOB: CM Actions attempt finished."
            }

            # Execute the functions
            Set-GPupdate
            Start-ConfigManagerActions
        }
        Write-Host "Background update job started (ID: $($updateJob.Id)). Main window will load."
        # Optional: Display a brief message in the main window later indicating updates are running
    }

    return $isBackupMode
}

# Show main window
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$IsBackup
    )

    # XAML UI Definition with CheckBox column
    [xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Data Backup/Restore Tool"
    Width="800"
    Height="600"
    WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Location:" Margin="10,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="txtSaveLoc" Width="400" Height="30" Margin="10,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnBrowse" Content="Browse" Width="60" Height="30" Margin="420,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Name="lblMode" Content="" Margin="500,10,10,0" HorizontalAlignment="Right" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Name="lblStatus" Content="Ready" Margin="10,0,10,10" HorizontalAlignment="Left" VerticalAlignment="Bottom" FontStyle="Italic"/>

        <Label Content="Files/Folders to Process:" Margin="10,80,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,110,200,140" SelectionMode="Extended">
             <ListView.View>
                <GridView>
                    <!-- CheckBox Column -->
                    <GridViewColumn Width="30">
                        <GridViewColumn.CellTemplate>
                            <DataTemplate>
                                <!-- Bind CheckBox IsChecked to the IsSelected property of the data item -->
                                <!-- Set IsEnabled based on whether it's Restore mode -->
                                <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" IsEnabled="{Binding DataContext.IsRestoreMode, RelativeSource={RelativeSource AncestorType=Window}}"/>
                            </DataTemplate>
                        </GridViewColumn.CellTemplate>
                    </GridViewColumn>
                    <!-- Other Columns -->
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="180"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="70"/>
                    <GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="280"/>
                </GridView>
            </ListView.View>
        </ListView>

        <StackPanel Margin="0,110,10,0" HorizontalAlignment="Right" Width="180">
            <Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnRemove" Content="Remove Selected" Width="120" Height="30" Margin="0,0,0,20"/>
            <CheckBox Name="chkNetwork" Content="Network Drives" IsChecked="True" Margin="5,0,0,5"/>
            <CheckBox Name="chkPrinters" Content="Printers" IsChecked="True" Margin="5,0,0,5"/>
        </StackPanel>

        <ProgressBar Name="prgProgress" Height="20" Margin="10,0,10,60" VerticalAlignment="Bottom"/>
        <TextBlock Name="txtProgress" Text="" Margin="10,0,10,85" VerticalAlignment="Bottom" TextWrapping="Wrap"/>
        <!-- Set IsDefault=True for the Start button -->
        <Button Name="btnStart" Content="Start" Width="100" Height="30" Margin="10,0,0,20" VerticalAlignment="Bottom" HorizontalAlignment="Left" IsDefault="True"/>

    </Grid>
</Window>
'@

    try {
        # Parse XAML
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)

        # Set DataContext property for IsRestoreMode binding
        # Use a PSObject wrapper for the DataContext to make it mutable if needed later
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }

        # Use a hashtable to store controls for easy access
        $controls = @{}
        $window.FindName('txtSaveLoc') | ForEach-Object { $controls['txtSaveLoc'] = $_ }
        $window.FindName('btnBrowse') | ForEach-Object { $controls['btnBrowse'] = $_ }
        $window.FindName('btnStart') | ForEach-Object { $controls['btnStart'] = $_ }
        $window.FindName('lblMode') | ForEach-Object { $controls['lblMode'] = $_ }
        $window.FindName('lblStatus') | ForEach-Object { $controls['lblStatus'] = $_ }
        $window.FindName('lvwFiles') | ForEach-Object { $controls['lvwFiles'] = $_ }
        $window.FindName('btnAddFile') | ForEach-Object { $controls['btnAddFile'] = $_ }
        $window.FindName('btnAddFolder') | ForEach-Object { $controls['btnAddFolder'] = $_ }
        $window.FindName('btnRemove') | ForEach-Object { $controls['btnRemove'] = $_ }
        $window.FindName('chkNetwork') | ForEach-Object { $controls['chkNetwork'] = $_ }
        $window.FindName('chkPrinters') | ForEach-Object { $controls['chkPrinters'] = $_ }
        $window.FindName('prgProgress') | ForEach-Object { $controls['prgProgress'] = $_ }
        $window.FindName('txtProgress') | ForEach-Object { $controls['txtProgress'] = $_ }

        # --- Window Initialization ---
        $controls.lblMode.Content = if ($IsBackup) { "Mode: Backup" } else { "Mode: Restore" }
        $controls.btnStart.Content = if ($IsBackup) { "Backup" } else { "Restore" }

        # Enable/Disable Add/Remove buttons based on mode
        $controls.btnAddFile.IsEnabled = $IsBackup
        $controls.btnAddFolder.IsEnabled = $IsBackup
        $controls.btnRemove.IsEnabled = $IsBackup


        # Set default path
        $defaultPath = "C:\LocalData"
        if (-not (Test-Path $defaultPath)) {
            try {
                New-Item -Path $defaultPath -ItemType Directory -Force | Out-Null
            } catch {
                 Write-Warning "Could not create default path: $defaultPath. Please select a location manually."
                 $defaultPath = $env:USERPROFILE # Fallback
            }
        }
        $controls.txtSaveLoc.Text = $defaultPath

        # Load initial items based on mode
        if ($IsBackup) {
            # Load default backup paths
            $paths = Get-BackupPaths
            # Add IsSelected property for potential future use or consistency
            $pathsWithSelection = $paths | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru }
            # *** FIX: Initialize list correctly ***
            $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $pathsWithSelection | ForEach-Object { $itemsList.Add($_) }
            $controls.lvwFiles.ItemsSource = $itemsList
        }
        elseif (Test-Path $defaultPath) {
            # Restore Mode: Look for most recent backup in default path
            $latestBackup = Get-ChildItem -Path $defaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                # Load backup contents into ListView, adding IsSelected property
                $backupItems = Get-ChildItem -Path $latestBackup.FullName |
                    Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } |
                    ForEach-Object {
                        [PSCustomObject]@{
                            Name = $_.Name
                            Type = if ($_.PSIsContainer) { "Folder" } else { "File" }
                            Path = $_.FullName # Path within the backup folder
                            IsSelected = $true # Default to selected for restore
                        }
                    }
                # *** FIX: Initialize list correctly ***
                $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                $backupItems | ForEach-Object { $itemsList.Add($_) }
                $controls.lvwFiles.ItemsSource = $itemsList
            } else {
                 $controls.lblStatus.Content = "Restore mode: No backups found in $defaultPath. Please browse."
                 $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new() # Ensure it's an empty list
            }
        } else {
             # Ensure ItemsSource is an empty list if default path doesn't exist in restore mode
             $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new()
        }

        # --- Event Handlers ---
        $controls.btnBrowse.Add_Click({
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($IsBackup) { "Select location to save backup" } else { "Select backup folder to restore from" }
            # Ensure SelectedPath exists before assigning
            if(Test-Path $controls.txtSaveLoc.Text){
                 $dialog.SelectedPath = $controls.txtSaveLoc.Text
            } else {
                 $dialog.SelectedPath = $defaultPath # Fallback if current text isn't valid
            }
            $dialog.ShowNewFolderButton = $IsBackup # Only allow creating new folders in backup mode

            # Use a temporary hidden form as the owner for the dialog
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose() # Dispose the temporary owner form

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath
                $controls.txtSaveLoc.Text = $selectedPath

                # If restoring, load the contents of the selected backup folder
                if (-not $IsBackup) {
                    if (Test-Path -Path (Join-Path $selectedPath "FileList_Backup.csv")) {
                         $backupItems = Get-ChildItem -Path $selectedPath |
                            Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt)$' } |
                            ForEach-Object {
                                [PSCustomObject]@{
                                    Name = $_.Name
                                    Type = if ($_.PSIsContainer) { "Folder" } else { "File" }
                                    Path = $_.FullName
                                    IsSelected = $true # Default to selected
                                }
                            }
                        # *** FIX: Initialize list correctly ***
                        $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $backupItems | ForEach-Object { $itemsList.Add($_) }
                        $controls.lvwFiles.ItemsSource = $itemsList
                        $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                    } else {
                         $controls.lvwFiles.ItemsSource = $null # Clear list view
                         $controls.lblStatus.Content = "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         [System.Windows.MessageBox]::Show("The selected folder does not appear to be a valid backup. It's missing the 'FileList_Backup.csv' log file.", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                } else {
                     $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                }
            }
        })

        # Corrected Add File/Folder handlers
        $controls.btnAddFile.Add_Click({
            if (-not $IsBackup) { return }
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = "Select File(s) to Add to Backup"
            $dialog.Multiselect = $true

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                # Create a new list, copy existing items, add new ones
                $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    # Ensure existing items also have IsSelected if they somehow don't
                    $controls.lvwFiles.ItemsSource | ForEach-Object {
                        if (-not ($_.PSObject.Properties.Name -contains 'IsSelected')) {
                            $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true
                        }
                        $newItemsList.Add($_)
                    }
                }

                foreach ($file in $dialog.FileNames) {
                    # Avoid adding duplicates by path
                    if (-not ($newItemsList.Path -contains $file)) {
                        $newItemsList.Add([PSCustomObject]@{
                            Name = [System.IO.Path]::GetFileName($file)
                            Type = "File"
                            Path = $file
                            IsSelected = $true # Add IsSelected here too
                        })
                    }
                }
                $controls.lvwFiles.ItemsSource = $newItemsList # Reassign ItemsSource
            }
        })

        $controls.btnAddFolder.Add_Click({
             if (-not $IsBackup) { return }
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = "Select Folder to Add to Backup"
            $dialog.ShowNewFolderButton = $false

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath
                 # Create a new list, copy existing items, add new ones
                 $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                 if ($controls.lvwFiles.ItemsSource -ne $null) {
                     # Ensure existing items also have IsSelected if they somehow don't
                     $controls.lvwFiles.ItemsSource | ForEach-Object {
                         if (-not ($_.PSObject.Properties.Name -contains 'IsSelected')) {
                             $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true
                         }
                         $newItemsList.Add($_)
                     }
                 }

                 # Avoid adding duplicates by path
                 if (-not ($newItemsList.Path -contains $selectedPath)) {
                    $newItemsList.Add([PSCustomObject]@{
                        Name = [System.IO.Path]::GetFileName($selectedPath)
                        Type = "Folder"
                        Path = $selectedPath
                        IsSelected = $true # Add IsSelected here too
                    })
                 }
                 $controls.lvwFiles.ItemsSource = $newItemsList # Reassign ItemsSource
            }
        })

        $controls.btnRemove.Add_Click({
            # This button should only be enabled in Backup mode now
            if (-not $IsBackup) { return }
            if ($controls.lvwFiles.SelectedItems.Count -gt 0) {
                # Create a new list excluding the selected items
                $itemsToKeep = [System.Collections.Generic.List[PSCustomObject]]::new()
                # Get the actual selected objects, not just paths
                $selectedObjects = @($controls.lvwFiles.SelectedItems) # Ensure it's an array
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    # Filter by comparing the objects themselves
                    $controls.lvwFiles.ItemsSource | Where-Object { $selectedObjects -notcontains $_ } | ForEach-Object {
                        $itemsToKeep.Add($_)
                    }
                }
                $controls.lvwFiles.ItemsSource = $itemsToKeep # Reassign ItemsSource
            }
        })


        # --- Start Button Logic ---
        $controls.btnStart.Add_Click({
            $location = $controls.txtSaveLoc.Text
            if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) { # Check it's a directory
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Disable UI elements during operation
            $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }
            $window.Cursor = [System.Windows.Input.Cursors]::Wait

            try {
                $controls.lblStatus.Content = if ($IsBackup) { "Starting backup..." } else { "Starting restore..." }
                $controls.txtProgress.Text = "Initializing..."
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $true # Use indeterminate initially

                $operationPath = $location # Base path for operation

                # ==================
                # --- BACKUP Logic ---
                # ==================
                if ($IsBackup) {
                    Write-Host "DEBUG: Running BACKUP logic" # Debug line
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_' # Sanitize username for path
                    $backupRootPath = Join-Path $operationPath "Backup_${username}_$timestamp"

                    try {
                        if (-not (Test-Path $backupRootPath)) {
                            New-Item -Path $backupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }
                    } catch {
                        throw "Failed to create backup directory: $backupRootPath. Error: $($_.Exception.Message)"
                    }

                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    "OriginalFullPath,BackupRelativePath" | Set-Content -Path $csvLogPath -Encoding UTF8

                    # Use ItemsSource directly
                    $itemsToBackup = @($controls.lvwFiles.ItemsSource) # Ensure it's an array
                    if (-not $itemsToBackup -or $itemsToBackup.Count -eq 0) {
                         throw "No items selected or found for backup."
                    }

                    # Estimate total files for progress (can be inaccurate for large folders)
                    $totalFilesEstimate = 0
                    $itemsToBackup | ForEach-Object {
                        if ($_.Type -eq 'Folder' -and (Test-Path $_.Path)) {
                            try { $totalFilesEstimate += (Get-ChildItem $_.Path -Recurse -File -Force -ErrorAction SilentlyContinue).Count } catch {}
                        } elseif (Test-Path $_.Path -PathType Leaf) { # Only count existing files
                            $totalFilesEstimate++
                        }
                    }
                    if ($controls.chkNetwork.IsChecked) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked) { $totalFilesEstimate++ }

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 } # Avoid max=0
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Process Files/Folders for Backup
                    foreach ($item in $itemsToBackup) {
                        $controls.txtProgress.Text = "Processing: $($item.Name)"
                        $sourcePath = $item.Path

                        if (-not (Test-Path $sourcePath)) {
                            Write-Warning "Source path not found, skipping: $sourcePath"
                            continue
                        }

                        if ($item.Type -eq "Folder") {
                            try {
                                # Log the folder itself first (useful for restoring empty folders if needed)
                                # "`"$sourcePath`",`"$($item.Name)`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                Get-ChildItem -Path $sourcePath -Recurse -File -Force -ErrorAction Stop | ForEach-Object {
                                    $originalFileFullPath = $_.FullName
                                    # Calculate path relative to the *root* of the item being backed up
                                    $relativeFilePath = $originalFileFullPath.Substring($sourcePath.TrimEnd('\').Length).TrimStart('\')
                                    $backupRelativePath = Join-Path $item.Name $relativeFilePath # Path relative to backup root folder

                                    $targetBackupPath = Join-Path $backupRootPath $backupRelativePath
                                    $targetBackupDir = [System.IO.Path]::GetDirectoryName($targetBackupPath)

                                    if (-not (Test-Path $targetBackupDir)) {
                                        New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                    }

                                    Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                    "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                    $filesProcessed++
                                    if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                    $controls.txtProgress.Text = "Backed up: $($_.Name)"
                                }
                            } catch {
                                 Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                            }
                        } else { # Single File
                             try {
                                $originalFileFullPath = $sourcePath
                                $backupRelativePath = $item.Name # Store file directly under the backup root
                                $targetBackupPath = Join-Path $backupRootPath $backupRelativePath

                                Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                $controls.txtProgress.Text = "Backed up: $($item.Name)"
                             } catch {
                                 Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                             }
                        }
                    } # End foreach item

                    # Backup Network Drives
                    if ($controls.chkNetwork.IsChecked) {
                        $controls.txtProgress.Text = "Backing up network drives..."
                        try {
                            Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop |
                                Select-Object Name, ProviderName |
                                Export-Csv -Path (Join-Path $backupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                            $filesProcessed++
                            if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                        } catch {
                             Write-Warning "Failed to backup network drives: $($_.Exception.Message)"
                        }
                    }

                    # Backup Printers
                    if ($controls.chkPrinters.IsChecked) {
                        $controls.txtProgress.Text = "Backing up printers..."
                        try {
                            Get-WmiObject -Class Win32_Printer -Filter "Local = False" -ErrorAction Stop |
                                Select-Object -ExpandProperty Name |
                                Set-Content -Path (Join-Path $backupRootPath "Printers.txt") -Encoding UTF8 -ErrorAction Stop
                             $filesProcessed++
                             if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                        } catch {
                             Write-Warning "Failed to backup printers: $($_.Exception.Message)"
                        }
                    }

                    $controls.txtProgress.Text = "Backup completed successfully to: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }


                # ===================
                # --- RESTORE Logic ---
                # ===================
                } else {
                    Write-Host "DEBUG: Running RESTORE logic" # Debug line
                    $backupRootPath = $operationPath # In restore mode, the selected location IS the backup root
                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"

                    if (-not (Test-Path $csvLogPath -PathType Leaf)) {
                        throw "Backup log file 'FileList_Backup.csv' not found in the selected location: $backupRootPath"
                    }

                    # Restore mode updates (GPUpdate, CM Actions) were initiated earlier via Start-Job

                    $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8
                    if (-not $backupLog) {
                         throw "Backup log file is empty or could not be read: $csvLogPath"
                    }

                    # --- Get Selected Items from ListView ---
                    # Ensure ItemsSource is treated as a collection
                    $listViewItems = @($controls.lvwFiles.ItemsSource)
                    $selectedItemsFromListView = $listViewItems | Where-Object { $_.IsSelected }

                    if (-not $selectedItemsFromListView) {
                        throw "No items selected in the list for restore."
                    }
                    $selectedTopLevelNames = $selectedItemsFromListView | Select-Object -ExpandProperty Name

                    # --- Filter Log Entries Based on Selection ---
                    $logEntriesToRestore = $backupLog | Where-Object {
                        # Handle both files directly in root and files within folders
                        $topLevelName = ($_.BackupRelativePath -split '[\\/]', 2)[0]
                        $selectedTopLevelNames -contains $topLevelName
                    }

                    if (-not $logEntriesToRestore) {
                        throw "None of the selected items correspond to entries in the backup log."
                    }

                    # Estimate progress based on selected log entries
                    $totalFilesEstimate = $logEntriesToRestore.Count
                    if ($controls.chkNetwork.IsChecked) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked) { $totalFilesEstimate++ }

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 } # Avoid max=0
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Restore Files/Folders from Filtered Log
                    foreach ($entry in $logEntriesToRestore) {
                        $originalFileFullPath = $entry.OriginalFullPath
                        $backupRelativePath = $entry.BackupRelativePath
                        $sourceBackupPath = Join-Path $backupRootPath $backupRelativePath

                        $controls.txtProgress.Text = "Restoring: $(Split-Path $originalFileFullPath -Leaf)"

                        if (Test-Path $sourceBackupPath -PathType Leaf) { # Ensure source exists in backup
                            try {
                                $targetRestoreDir = [System.IO.Path]::GetDirectoryName($originalFileFullPath)
                                if (-not (Test-Path $targetRestoreDir)) {
                                    New-Item -Path $targetRestoreDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                }

                                # Add -Confirm:$false if you want to overwrite without prompts by default
                                Copy-Item -Path $sourceBackupPath -Destination $originalFileFullPath -Force -ErrorAction Stop

                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }

                            } catch {
                                Write-Warning "Failed to restore '$originalFileFullPath' from '$sourceBackupPath': $($_.Exception.Message)"
                                # Optionally, add to a list of failed files
                            }
                        } else {
                            Write-Warning "Source file not found in backup, skipping restore: $sourceBackupPath (Expected for: $originalFileFullPath)"
                        }
                    } # End foreach entry

                    # Restore Network Drives (Not dependent on ListView selection)
                    if ($controls.chkNetwork.IsChecked) {
                        $controls.txtProgress.Text = "Restoring network drives..."
                        $drivesCsvPath = Join-Path $backupRootPath "Drives.csv"
                        if (Test-Path $drivesCsvPath) {
                            try {
                                Import-Csv $drivesCsvPath | ForEach-Object {
                                    $driveLetter = $_.Name.TrimEnd(':')
                                    $networkPath = $_.ProviderName
                                    if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\' ) {
                                        if (-not (Test-Path -LiteralPath "$($driveLetter):")) {
                                            try {
                                                Write-Verbose "Mapping $driveLetter to $networkPath"
                                                New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -ErrorAction Stop
                                            } catch {
                                                 Write-Warning "Failed to map drive $driveLetter`: $($_.Exception.Message)"
                                            }
                                        } else {
                                             Write-Verbose "Drive $driveLetter already exists, skipping."
                                        }
                                    } else {
                                         Write-Warning "Skipping invalid drive mapping: Name='$($_.Name)', Provider='$networkPath'"
                                    }
                                }
                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                            } catch {
                                 Write-Warning "Error processing network drive restorations: $($_.Exception.Message)"
                            }
                        } else { Write-Warning "Network drives backup file (Drives.csv) not found." }
                    }

                    # Restore Printers (Not dependent on ListView selection)
                    if ($controls.chkPrinters.IsChecked) {
                        $controls.txtProgress.Text = "Restoring printers..."
                        $printersTxtPath = Join-Path $backupRootPath "Printers.txt"
                        if (Test-Path $printersTxtPath) {
                            try {
                                $wsNet = New-Object -ComObject WScript.Network # Use COM object for broader compatibility
                                Get-Content $printersTxtPath | ForEach-Object {
                                    $printerPath = $_
                                    if (-not ([string]::IsNullOrWhiteSpace($printerPath))) {
                                        try {
                                            Write-Verbose "Adding printer: $printerPath"
                                            # Check if printer already exists (optional, AddWindowsPrinterConnection might handle it)
                                            # if (-not (Get-Printer -Name $printerPath -ErrorAction SilentlyContinue)) {
                                                 $wsNet.AddWindowsPrinterConnection($printerPath)
                                            # } else { Write-Verbose "Printer '$printerPath' already exists." }
                                        } catch {
                                             Write-Warning "Failed to add printer '$printerPath': $($_.Exception.Message)"
                                        }
                                    }
                                }
                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                            } catch {
                                 Write-Warning "Error processing printer restorations: $($_.Exception.Message)"
                            }
                        } else { Write-Warning "Printers backup file (Printers.txt) not found." }
                    }

                    $controls.txtProgress.Text = "Restore completed from: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }

                } # End if/else ($IsBackup)

                # --- Operation Completion ---
                $controls.lblStatus.Content = "Operation completed successfully."
                [System.Windows.MessageBox]::Show("The $($controls.btnStart.Content) operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

            } catch {
                # --- Operation Failure ---
                $errorMessage = "Operation Failed: $($_.Exception.Message)"
                Write-Error $errorMessage
                $controls.lblStatus.Content = "Operation Failed!"
                $controls.txtProgress.Text = $errorMessage
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $false
                [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                # Re-enable UI elements
                 $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $true } }
                 # Restore Add/Remove button state based on mode
                 $controls.btnAddFile.IsEnabled = $IsBackup
                 $controls.btnAddFolder.IsEnabled = $IsBackup
                 $controls.btnRemove.IsEnabled = $IsBackup

                 $window.Cursor = [System.Windows.Input.Cursors]::Arrow
            }
        }) # End btnStart.Add_Click

        # --- Show Window ---
        $window.ShowDialog() | Out-Null

    } catch {
        # --- Window Load Failure ---
        $errorMessage = "Failed to load main window: $($_.Exception.Message)"
        Write-Error $errorMessage
        [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

#endregion Functions

# --- Main Execution ---
try {
    # Determine mode (Backup = $true, Restore = $false)
    # Show-ModeDialog now returns boolean directly
    [bool]$script:isBackupMode = Show-ModeDialog

    # Show the main window, passing the determined mode
    Show-MainWindow -IsBackup $script:isBackupMode

} catch {
    # Catch errors during initial mode selection or window loading
    $errorMessage = "An unexpected error occurred during startup: $($_.Exception.Message)"
    Write-Error $errorMessage
    [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
}

Write-Host "Script execution finished."

# --- Keep console open when double-clicked ---
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE) { # Check if running in console and not ISE
    Read-Host -Prompt "Press Enter to exit"
}