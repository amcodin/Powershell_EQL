

#requires -Version 3.0

# --- Load Assemblies FIRST ---
Write-Verbose "Attempting to load .NET Assemblies..."
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-Verbose ".NET Assemblies loaded successfully."
} catch {
    Write-Error "FATAL: Failed to load required .NET Assemblies. UI cannot be created."
    Write-Error "Error Message: $($_.Exception.Message)"
    # Optional: Pause if running directly in console
    if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
    Exit 1 # Exit script if assemblies fail to load
}


#region Functions

# --- Embedded Functions (Removed Dot-Sourcing) ---

function Get-BackupPaths {
    [CmdletBinding()]
    param ()
    

    # Define specific paths/files to check
    $specificPaths = @(
        "$env:APPDATA\Microsoft\Signatures",
        "$env:SystemDrive\User", # Check if this generic path is truly needed/valid
        "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms", # Quick Access Pinned
        "$env:SystemDrive\Temp",
        "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt", # Legacy Sticky Notes
        "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite", # Modern Sticky Notes
        "$env:APPDATA\google\googleearth\myplaces.kml", # Google Earth Places
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks", # Chrome Bookmarks
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks", # Edge Bookmarks
        # Add other common paths as needed
    )

    $result = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Add specific paths
    foreach ($path in $specificPaths) {
        if (Test-Path -Path $path) {
            $result.Add([PSCustomObject]@{
                Name = Split-Path $path -Leaf
                Path = $path
                Type = if (Test-Path -Path $path -PathType Container) { "Folder" } else { "File" }
            })
        } else {
             Write-Verbose "Specific path not found: $path"
        }
    }

    return $result
}

function Set-GPupdate {
    Write-Verbose "Entering Set-GPupdate function."
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
    Write-Verbose "Exiting Set-GPupdate function."
}

function Start-ConfigManagerActions {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$UseModule # Parameter to optionally try using the CM module (might require admin rights/module install)
    )
    Write-Verbose "Entering Start-ConfigManagerActions function."

    Write-Verbose "Attempting to trigger Configuration Manager client actions..."
    $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"
    $clientSDKNamespace = "root\ccm\clientsdk"
    $clientClassName = "CCM_ClientUtilities" # More common class for triggering schedules
    $scheduleMethodName = "TriggerSchedule"

    # Check if CM client exists
    if (-not (Test-Path $ccmExecPath)) {
        Write-Warning "Configuration Manager client not found at $ccmExecPath. Skipping CM actions."
        Write-Verbose "Exiting Start-ConfigManagerActions function (CM client not found)."
        return $true # Return success as CM is not applicable
    }
    Write-Verbose "CM Client found at $ccmExecPath."

    $scheduleActions = @(
        @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'},
        @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}
        # Add more schedule IDs as needed
    )
    Write-Verbose "Defined $($scheduleActions.Count) CM actions to trigger."

    $success = $true
    $attemptedMethod = "WMI/CIM"

    try {
        # Attempt using WMI/CIM first (preferred, less prone to process issues)
        Write-Verbose "Checking for CIM class '$clientClassName' in namespace '$clientSDKNamespace'."
        if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
             Write-Verbose "Using CIM method to trigger schedules."
             $attemptedMethod = "WMI/CIM"
             foreach ($action in $scheduleActions) {
                Write-Verbose "Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                try {
                    Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop
                    Write-Verbose "  $($action.Name) triggered successfully via CIM."
                } catch {
                    Write-Warning "  Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"
                    $success = $false # Mark as partially failed if any action fails
                }
            }
        } else {
             Write-Warning "CIM Class $clientClassName not found in $clientSDKNamespace. Falling back to ccmexec.exe."
             $attemptedMethod = "ccmexec.exe"
             # Fallback to ccmexec.exe direct execution
             foreach ($action in $scheduleActions) {
                Write-Verbose "Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                # Note: Triggering via ccmexec.exe is less reliable and provides less feedback
                try {
                    # Using -TriggerSchedule requires the GUID format
                    $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop
                    if ($process.ExitCode -ne 0) {
                        Write-Warning "  $($action.Name) action via ccmexec.exe finished with exit code $($process.ExitCode)"
                        # Don't mark as failure here, as ccmexec often returns non-zero even if schedule triggers
                    } else {
                         Write-Verbose "  $($action.Name) triggered via ccmexec.exe (Exit Code 0)."
                    }
                } catch {
                     Write-Warning "  Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"
                     $success = $false
                }
            }
        }

    } catch {
        Write-Error "An unexpected error occurred during Configuration Manager actions ($attemptedMethod): $($_.Exception.Message)"
        Write-Verbose "Exiting Start-ConfigManagerActions function (Error)."
        return $false # Return failure on major error
    }

    Write-Verbose "Configuration Manager actions attempt finished. Overall success: $success."
    Write-Verbose "Exiting Start-ConfigManagerActions function."
    return $success # Return true if all actions triggered successfully (or CM not present), false otherwise
}

# --- GUI Functions ---

# Show mode selection dialog
function Show-ModeDialog {
    Write-Verbose "Entering Show-ModeDialog function."
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

    Write-Verbose "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Verbose "Mode selection dialog closed with result: $result"

    # Determine mode based on which button was effectively clicked
    $isBackupMode = ($result -eq [System.Windows.Forms.DialogResult]::OK)
    # *** FIX: Use if/else for Write-Verbose ***
    $modeString = if ($isBackupMode) { 'Backup' } else { 'Restore' }
    Write-Verbose "Determined mode: $modeString"

    # If restore mode, run updates immediately in a non-blocking way
    if (-not $isBackupMode) {
        Write-Verbose "Restore mode selected. Initiating background system updates job..."
        # Use Start-Job for non-blocking execution
        $updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {
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

    Write-Verbose "Exiting Show-ModeDialog function."
    return $isBackupMode
}

# Show main window
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$IsBackup
    )
    # *** FIX: Use if/else for Write-Verbose ***
    $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
    Write-Verbose "Entering Show-MainWindow function. Mode: $modeString"

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
                                <!-- REMOVED IsEnabled binding to allow checking in Restore mode -->
                                <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" />
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
        Write-Verbose "Parsing XAML for main window."
        # Parse XAML
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Verbose "XAML loaded successfully."

        # Set DataContext property for IsRestoreMode binding (used by CheckBox IsEnabled)
        # Use a PSObject wrapper for the DataContext to make it mutable if needed later
        # $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }
        # Since IsEnabled was removed from CheckBox, this DataContext binding is no longer strictly needed for it,
        # but keeping it might be useful if other bindings depend on the mode.
        Write-Verbose "Setting Window DataContext."
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }


        # Use a hashtable to store controls for easy access
        Write-Verbose "Finding controls in main window."
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
        Write-Verbose "Controls found and stored in hashtable."

        # --- Window Initialization ---
        Write-Verbose "Initializing window controls based on mode."
        $controls.lblMode.Content = if ($IsBackup) { "Mode: Backup" } else { "Mode: Restore" }
        $controls.btnStart.Content = if ($IsBackup) { "Backup" } else { "Restore" }

        # Enable/Disable Add/Remove buttons based on mode
        $controls.btnAddFile.IsEnabled = $IsBackup
        $controls.btnAddFolder.IsEnabled = $IsBackup
        $controls.btnRemove.IsEnabled = $IsBackup
        Write-Verbose "Add/Remove buttons enabled state set based on mode."


        # Set default path
        $defaultPath = "C:\LocalData"
        Write-Verbose "Checking default path: $defaultPath"
        if (-not (Test-Path $defaultPath)) {
            Write-Verbose "Default path not found. Attempting to create."
            try {
                New-Item -Path $defaultPath -ItemType Directory -Force | Out-Null
                Write-Verbose "Default path created."
            } catch {
                 Write-Warning "Could not create default path: $defaultPath. Please select a location manually."
                 $defaultPath = $env:USERPROFILE # Fallback
                 Write-Verbose "Using fallback default path: $defaultPath"
            }
        } else { Write-Verbose "Default path exists."}
        $controls.txtSaveLoc.Text = $defaultPath

        # Load initial items based on mode
        Write-Verbose "Loading initial items for ListView based on mode."
        if ($IsBackup) {
            Write-Verbose "Backup Mode: Getting default paths using Get-BackupPaths."
            # Load default backup paths
            $paths = $null # Initialize to null
            try {
                $paths = Get-BackupPaths # Calls the function
                Write-Verbose "Get-BackupPaths call completed."
            } catch {
                 Write-Error "Error calling Get-BackupPaths: $($_.Exception.Message)"
                 # Decide how to handle this - maybe throw or just continue with empty list?
                 # For now, let's allow it to proceed to the null check below.
            }


            # *** FIX: Add more checks and logging ***
            if ($paths -ne $null -and $paths.Count -gt 0) {
                 Write-Verbose "Get-BackupPaths returned $($paths.Count) items. First item type: $($paths[0].GetType().FullName)"
                 Write-Verbose "Attempting to add 'IsSelected' property..."
                 # Add IsSelected property for potential future use or consistency
                 # Ensure Add-Member doesn't fail if input is single object vs collection
                 $pathsWithSelection = $null # Initialize
                 try {
                    # Ensure $paths is treated as a collection for the pipe
                    $pathsCollection = @($paths)
                    Write-Verbose "Processing $($pathsCollection.Count) items with Add-Member."
                    $pathsWithSelection = $pathsCollection | ForEach-Object {
                         # Write-Verbose "  Adding IsSelected to item: $($_.Path)" # Can be very noisy
                         $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                    }
                    Write-Verbose "'IsSelected' property added. Type of result: $($pathsWithSelection.GetType().FullName)"
                    # Check result count (ForEach-Object output might be single item or array)
                    $pathsWithSelectionCount = if ($pathsWithSelection -is [array]) { $pathsWithSelection.Count } elseif ($pathsWithSelection -ne $null) { 1 } else { 0 }
                    Write-Verbose "Count after Add-Member: $pathsWithSelectionCount"
                    if ($pathsWithSelection -eq $null -and $pathsCollection.Count -gt 0) {Write-Warning "pathsWithSelection is null after Add-Member despite input collection!"}

                 } catch {
                     Write-Error "Error during Add-Member: $($_.Exception.Message)"
                     # Handle error - maybe skip this step or use original $paths?
                     # For now, let it proceed to the next null check.
                 }


                 # Initialize list correctly
                 $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                 Write-Verbose "Initialized empty itemsList."

                 # Check if pathsWithSelection is valid before iterating
                 if ($pathsWithSelection -ne $null) {
                     Write-Verbose "Attempting to populate itemsList from pathsWithSelection..."
                     try {
                        # Ensure $pathsWithSelection is treated as a collection
                        $selectionCollection = @($pathsWithSelection)
                        Write-Verbose "Iterating through $($selectionCollection.Count) items to add to itemsList."
                        $selectionCollection | ForEach-Object {
                            if ($_ -ne $null) { # Add check for null items within the collection
                                # Write-Verbose "  Adding item to list: $($_.Path)" # Can be very noisy
                                $itemsList.Add($_)
                            } else {
                                Write-Warning "Skipping null item found in pathsWithSelection collection."
                            }
                        }
                        Write-Verbose "Finished populating itemsList. Count: $($itemsList.Count)"
                     } catch {
                         Write-Error "Error populating itemsList: $($_.Exception.Message)"
                         # Handle error - maybe clear itemsList?
                         $itemsList.Clear()
                     }
                 } else {
                     Write-Warning "Cannot populate itemsList because pathsWithSelection is null."
                 }


                 # Check if lvwFiles control exists before assigning
                 Write-Verbose "Checking if lvwFiles control exists..."
                 if ($controls['lvwFiles'] -ne $null) {
                    Write-Verbose "Assigning itemsList (Count: $($itemsList.Count)) to ListView ItemsSource."
                    $controls.lvwFiles.ItemsSource = $itemsList
                    Write-Verbose "Assigned itemsList to ListView ItemsSource."
                 } else {
                     Write-Error "ListView control ('lvwFiles') not found!"
                     # This would be a critical error, maybe throw?
                     throw "ListView control ('lvwFiles') could not be found in the XAML."
                 }

            } else {
                 Write-Warning "Get-BackupPaths returned no items or was null. ListView will be empty."
                 # Assign an empty list if no paths were found
                 Write-Verbose "Checking if lvwFiles control exists before assigning empty list..."
                 if ($controls['lvwFiles'] -ne $null) {
                    $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new()
                    Write-Verbose "Assigned empty list to ListView ItemsSource."
                 } else {
                     Write-Error "ListView control ('lvwFiles') not found when trying to assign empty list!"
                     throw "ListView control ('lvwFiles') could not be found in the XAML."
                 }
            }
            # *** End FIX ***
        }
        elseif (Test-Path $defaultPath) {
            Write-Verbose "Restore Mode: Checking for latest backup in $defaultPath."
            # Restore Mode: Look for most recent backup in default path
            $latestBackup = Get-ChildItem -Path $defaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                Write-Verbose "Found latest backup: $($latestBackup.FullName)"
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
                Write-Verbose "Populated ListView with $($itemsList.Count) items from latest backup."
            } else {
                 Write-Verbose "No backups found in $defaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $defaultPath. Please browse."
                 $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new() # Ensure it's an empty list
            }
        } else {
             Write-Verbose "Restore Mode: Default path $defaultPath does not exist."
             # Ensure ItemsSource is an empty list if default path doesn't exist in restore mode
             $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        Write-Verbose "Finished loading initial items."

        # --- Event Handlers ---
        Write-Verbose "Assigning event handlers."
        $controls.btnBrowse.Add_Click({
            Write-Verbose "Browse button clicked."
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
            Write-Verbose "Showing FolderBrowserDialog."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose() # Dispose the temporary owner form

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath
                Write-Verbose "Folder selected: $selectedPath"
                $controls.txtSaveLoc.Text = $selectedPath

                # If restoring, load the contents of the selected backup folder
                if (-not $IsBackup) {
                    Write-Verbose "Restore Mode: Loading items from selected backup folder."
                    $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                    Write-Verbose "Checking for log file: $logFilePath"
                    if (Test-Path -Path $logFilePath) {
                         Write-Verbose "Log file found. Populating ListView."
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
                        Write-Verbose "ListView updated with $($itemsList.Count) items from selected backup."
                    } else {
                         Write-Warning "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         $controls.lvwFiles.ItemsSource = $null # Clear list view
                         $controls.lblStatus.Content = "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         [System.Windows.MessageBox]::Show("The selected folder does not appear to be a valid backup. It's missing the 'FileList_Backup.csv' log file.", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                } else {
                     $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                }
            } else { Write-Verbose "Folder selection cancelled."}
        })

        # Corrected Add File/Folder handlers
        $controls.btnAddFile.Add_Click({
            Write-Verbose "Add File button clicked."
            if (-not $IsBackup) { Write-Verbose "Ignoring Add File click (not in Backup mode)."; return }
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = "Select File(s) to Add to Backup"
            $dialog.Multiselect = $true

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            Write-Verbose "Showing OpenFileDialog."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Verbose "$($dialog.FileNames.Count) file(s) selected."
                # Create a new list, copy existing items, add new ones
                $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    Write-Verbose "Copying existing items to new list."
                    # Ensure existing items also have IsSelected if they somehow don't
                    $controls.lvwFiles.ItemsSource | ForEach-Object {
                        if (-not ($_.PSObject.Properties.Name -contains 'IsSelected')) {
                            $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true
                        }
                        $newItemsList.Add($_)
                    }
                }

                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    # Avoid adding duplicates by path
                    if (-not ($newItemsList.Path -contains $file)) {
                        Write-Verbose "Adding file: $file"
                        $newItemsList.Add([PSCustomObject]@{
                            Name = [System.IO.Path]::GetFileName($file)
                            Type = "File"
                            Path = $file
                            IsSelected = $true # Add IsSelected here too
                        })
                        $addedCount++
                    } else { Write-Verbose "Skipping duplicate file: $file"}
                }
                $controls.lvwFiles.ItemsSource = $newItemsList # Reassign ItemsSource
                Write-Verbose "Updated ListView ItemsSource. Added $addedCount new file(s)."
            } else { Write-Verbose "File selection cancelled."}
        })

        $controls.btnAddFolder.Add_Click({
             Write-Verbose "Add Folder button clicked."
             if (-not $IsBackup) { Write-Verbose "Ignoring Add Folder click (not in Backup mode)."; return }
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = "Select Folder to Add to Backup"
            $dialog.ShowNewFolderButton = $false

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            Write-Verbose "Showing FolderBrowserDialog for adding folder."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath
                 Write-Verbose "Folder selected to add: $selectedPath"
                 # Create a new list, copy existing items, add new ones
                 $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                 if ($controls.lvwFiles.ItemsSource -ne $null) {
                     Write-Verbose "Copying existing items to new list."
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
                    Write-Verbose "Adding folder: $selectedPath"
                    $newItemsList.Add([PSCustomObject]@{
                        Name = [System.IO.Path]::GetFileName($selectedPath)
                        Type = "Folder"
                        Path = $selectedPath
                        IsSelected = $true # Add IsSelected here too
                    })
                    $controls.lvwFiles.ItemsSource = $newItemsList # Reassign ItemsSource
                    Write-Verbose "Updated ListView ItemsSource with new folder."
                 } else {
                    Write-Verbose "Skipping duplicate folder: $selectedPath"
                 }

            } else { Write-Verbose "Folder selection cancelled."}
        })

        $controls.btnRemove.Add_Click({
            Write-Verbose "Remove Selected button clicked."
            # This button should only be enabled in Backup mode now
            if (-not $IsBackup) { Write-Verbose "Ignoring Remove click (not in Backup mode)."; return }

            $selectedObjects = @($controls.lvwFiles.SelectedItems) # Ensure it's an array
            if ($selectedObjects.Count -gt 0) {
                Write-Verbose "Removing $($selectedObjects.Count) selected item(s)."
                # Create a new list excluding the selected items
                $itemsToKeep = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    # Filter by comparing the objects themselves
                    $controls.lvwFiles.ItemsSource | Where-Object { $selectedObjects -notcontains $_ } | ForEach-Object {
                        $itemsToKeep.Add($_)
                    }
                }
                $controls.lvwFiles.ItemsSource = $itemsToKeep # Reassign ItemsSource
                Write-Verbose "ListView ItemsSource updated after removal."
            } else { Write-Verbose "No items selected to remove."}
        })


        # --- Start Button Logic ---
        $controls.btnStart.Add_Click({
            # *** FIX: Use if/else for Write-Verbose ***
            $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
            Write-Verbose "Start button clicked. Mode: $modeString"

            $location = $controls.txtSaveLoc.Text
            Write-Verbose "Selected location: $location"
            if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) { # Check it's a directory
                Write-Warning "Invalid location selected."
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            Write-Verbose "Disabling UI controls and setting wait cursor."
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
                    Write-Verbose "--- Starting Backup Operation ---"
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_' # Sanitize username for path
                    $backupRootPath = Join-Path $operationPath "Backup_${username}_$timestamp"
                    Write-Verbose "Backup root path: $backupRootPath"

                    try {
                        Write-Verbose "Creating backup directory..."
                        if (-not (Test-Path $backupRootPath)) {
                            New-Item -Path $backupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                            Write-Verbose "Backup directory created."
                        } else { Write-Verbose "Backup directory already exists (should not happen with timestamp)."}
                    } catch {
                        throw "Failed to create backup directory: $backupRootPath. Error: $($_.Exception.Message)"
                    }

                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    Write-Verbose "Creating log file: $csvLogPath"
                    "OriginalFullPath,BackupRelativePath" | Set-Content -Path $csvLogPath -Encoding UTF8

                    # Use ItemsSource directly
                    $itemsToBackup = @($controls.lvwFiles.ItemsSource) # Ensure it's an array
                    if (-not $itemsToBackup -or $itemsToBackup.Count -eq 0) {
                         throw "No items selected or found for backup."
                    }
                    Write-Verbose "Found $($itemsToBackup.Count) items in ListView to process for backup."

                    # Estimate total files for progress (can be inaccurate for large folders)
                    Write-Verbose "Estimating total files for progress bar..."
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
                    Write-Verbose "Estimated total files/items: $totalFilesEstimate"

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 } # Avoid max=0
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Process Files/Folders for Backup
                    Write-Verbose "Starting processing of files/folders for backup..."
                    foreach ($item in $itemsToBackup) {
                        Write-Verbose "Processing item: $($item.Name) ($($item.Type)) - Path: $($item.Path)"
                        $controls.txtProgress.Text = "Processing: $($item.Name)"
                        $sourcePath = $item.Path

                        if (-not (Test-Path $sourcePath)) {
                            Write-Warning "Source path not found, skipping: $sourcePath"
                            continue
                        }

                        if ($item.Type -eq "Folder") {
                            Write-Verbose "  Item is a folder. Processing recursively..."
                            try {
                                Get-ChildItem -Path $sourcePath -Recurse -File -Force -ErrorAction Stop | ForEach-Object {
                                    $originalFileFullPath = $_.FullName
                                    # Calculate path relative to the *root* of the item being backed up
                                    $relativeFilePath = $originalFileFullPath.Substring($sourcePath.TrimEnd('\').Length).TrimStart('\')
                                    $backupRelativePath = Join-Path $item.Name $relativeFilePath # Path relative to backup root folder

                                    $targetBackupPath = Join-Path $backupRootPath $backupRelativePath
                                    $targetBackupDir = [System.IO.Path]::GetDirectoryName($targetBackupPath)

                                    if (-not (Test-Path $targetBackupDir)) {
                                        Write-Verbose "    Creating directory: $targetBackupDir"
                                        New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                    }

                                    Write-Verbose "    Copying '$($_.Name)' to '$targetBackupPath'"
                                    Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                    Write-Verbose "    Logging: `"$originalFileFullPath`",`"$backupRelativePath`""
                                    "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                    $filesProcessed++
                                    if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                    $controls.txtProgress.Text = "Backed up: $($_.Name)"
                                }
                                Write-Verbose "  Finished processing folder: $($item.Name)"
                            } catch {
                                 Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                            }
                        } else { # Single File
                             Write-Verbose "  Item is a file. Processing..."
                             try {
                                $originalFileFullPath = $sourcePath
                                $backupRelativePath = $item.Name # Store file directly under the backup root
                                $targetBackupPath = Join-Path $backupRootPath $backupRelativePath

                                Write-Verbose "    Copying '$($item.Name)' to '$targetBackupPath'"
                                Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                Write-Verbose "    Logging: `"$originalFileFullPath`",`"$backupRelativePath`""
                                "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                $controls.txtProgress.Text = "Backed up: $($item.Name)"
                             } catch {
                                 Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                             }
                        }
                    } # End foreach item
                    Write-Verbose "Finished processing files/folders for backup."

                    # Backup Network Drives
                    if ($controls.chkNetwork.IsChecked) {
                        Write-Verbose "Processing Network Drives backup..."
                        $controls.txtProgress.Text = "Backing up network drives..."
                        try {
                            Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop |
                                Select-Object Name, ProviderName |
                                Export-Csv -Path (Join-Path $backupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                            $filesProcessed++
                            if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                            Write-Verbose "Network drives backed up successfully."
                        } catch {
                             Write-Warning "Failed to backup network drives: $($_.Exception.Message)"
                        }
                    } else { Write-Verbose "Skipping Network Drives backup (unchecked)."}

                    # Backup Printers
                    if ($controls.chkPrinters.IsChecked) {
                        Write-Verbose "Processing Printers backup..."
                        $controls.txtProgress.Text = "Backing up printers..."
                        try {
                            Get-WmiObject -Class Win32_Printer -Filter "Local = False" -ErrorAction Stop |
                                Select-Object -ExpandProperty Name |
                                Set-Content -Path (Join-Path $backupRootPath "Printers.txt") -Encoding UTF8 -ErrorAction Stop
                             $filesProcessed++
                             if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                             Write-Verbose "Printers backed up successfully."
                        } catch {
                             Write-Warning "Failed to backup printers: $($_.Exception.Message)"
                        }
                    } else { Write-Verbose "Skipping Printers backup (unchecked)."}

                    $controls.txtProgress.Text = "Backup completed successfully to: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }
                    Write-Verbose "--- Backup Operation Finished ---"


                # ===================
                # --- RESTORE Logic ---
                # ===================
                } else {
                    Write-Verbose "--- Starting Restore Operation ---"
                    $backupRootPath = $operationPath # In restore mode, the selected location IS the backup root
                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    Write-Verbose "Restore source (backup root): $backupRootPath"
                    Write-Verbose "Log file path: $csvLogPath"

                    if (-not (Test-Path $csvLogPath -PathType Leaf)) {
                        throw "Backup log file 'FileList_Backup.csv' not found in the selected location: $backupRootPath"
                    }

                    # Restore mode updates (GPUpdate, CM Actions) were initiated earlier via Start-Job
                    Write-Verbose "Background updates (GPUpdate/CM) should have been initiated earlier."

                    Write-Verbose "Importing backup log file..."
                    $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8
                    if (-not $backupLog) {
                         throw "Backup log file is empty or could not be read: $csvLogPath"
                    }
                    Write-Verbose "Imported $($backupLog.Count) entries from log file."

                    # --- Get Selected Items from ListView ---
                    # Ensure ItemsSource is treated as a collection
                    $listViewItems = @($controls.lvwFiles.ItemsSource)
                    $selectedItemsFromListView = $listViewItems | Where-Object { $_.IsSelected }

                    if (-not $selectedItemsFromListView) {
                        throw "No items selected in the list for restore."
                    }
                    $selectedTopLevelNames = $selectedItemsFromListView | Select-Object -ExpandProperty Name
                    Write-Verbose "Found $($selectedItemsFromListView.Count) items selected in ListView for restore: $($selectedTopLevelNames -join ', ')"

                    # --- Filter Log Entries Based on Selection ---
                    Write-Verbose "Filtering log entries based on ListView selection..."
                    $logEntriesToRestore = $backupLog | Where-Object {
                        # Handle both files directly in root and files within folders
                        $topLevelName = ($_.BackupRelativePath -split '[\\/]', 2)[0]
                        $selectedTopLevelNames -contains $topLevelName
                    }

                    if (-not $logEntriesToRestore) {
                        throw "None of the selected items correspond to entries in the backup log."
                    }
                    Write-Verbose "Filtered log. $($logEntriesToRestore.Count) log entries will be processed for restore."

                    # Estimate progress based on selected log entries
                    $totalFilesEstimate = $logEntriesToRestore.Count
                    if ($controls.chkNetwork.IsChecked) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked) { $totalFilesEstimate++ }
                    Write-Verbose "Estimated total items for restore progress: $totalFilesEstimate"

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 } # Avoid max=0
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Restore Files/Folders from Filtered Log
                    Write-Verbose "Starting restore of files/folders from filtered log..."
                    foreach ($entry in $logEntriesToRestore) {
                        $originalFileFullPath = $entry.OriginalFullPath
                        $backupRelativePath = $entry.BackupRelativePath
                        $sourceBackupPath = Join-Path $backupRootPath $backupRelativePath

                        Write-Verbose "Processing restore entry: Source='$sourceBackupPath', Target='$originalFileFullPath'"
                        $controls.txtProgress.Text = "Restoring: $(Split-Path $originalFileFullPath -Leaf)"

                        if (Test-Path $sourceBackupPath -PathType Leaf) { # Ensure source exists in backup
                            try {
                                $targetRestoreDir = [System.IO.Path]::GetDirectoryName($originalFileFullPath)
                                if (-not (Test-Path $targetRestoreDir)) {
                                    Write-Verbose "  Creating target directory: $targetRestoreDir"
                                    New-Item -Path $targetRestoreDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                }

                                Write-Verbose "  Copying '$sourceBackupPath' to '$originalFileFullPath'"
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
                    Write-Verbose "Finished restoring files/folders from log."

                    # Restore Network Drives (Not dependent on ListView selection)
                    if ($controls.chkNetwork.IsChecked) {
                        Write-Verbose "Processing Network Drives restore..."
                        $controls.txtProgress.Text = "Restoring network drives..."
                        $drivesCsvPath = Join-Path $backupRootPath "Drives.csv"
                        if (Test-Path $drivesCsvPath) {
                            Write-Verbose "Found Drives.csv. Processing mappings..."
                            try {
                                Import-Csv $drivesCsvPath | ForEach-Object {
                                    $driveLetter = $_.Name.TrimEnd(':')
                                    $networkPath = $_.ProviderName
                                    Write-Verbose "  Checking mapping: $driveLetter -> $networkPath"
                                    if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\' ) {
                                        if (-not (Test-Path -LiteralPath "$($driveLetter):")) {
                                            try {
                                                Write-Verbose "    Mapping $driveLetter to $networkPath"
                                                New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -ErrorAction Stop
                                            } catch {
                                                 Write-Warning "    Failed to map drive $driveLetter`: $($_.Exception.Message)"
                                            }
                                        } else {
                                             Write-Verbose "    Drive $driveLetter already exists, skipping."
                                        }
                                    } else {
                                         Write-Warning "    Skipping invalid drive mapping: Name='$($_.Name)', Provider='$networkPath'"
                                    }
                                }
                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                Write-Verbose "Finished processing network drive mappings."
                            } catch {
                                 Write-Warning "Error processing network drive restorations: $($_.Exception.Message)"
                            }
                        } else { Write-Warning "Network drives backup file (Drives.csv) not found." }
                    } else { Write-Verbose "Skipping Network Drives restore (unchecked)."}

                    # Restore Printers (Not dependent on ListView selection)
                    if ($controls.chkPrinters.IsChecked) {
                        Write-Verbose "Processing Printers restore..."
                        $controls.txtProgress.Text = "Restoring printers..."
                        $printersTxtPath = Join-Path $backupRootPath "Printers.txt"
                        if (Test-Path $printersTxtPath) {
                             Write-Verbose "Found Printers.txt. Processing printers..."
                            try {
                                $wsNet = New-Object -ComObject WScript.Network # Use COM object for broader compatibility
                                Get-Content $printersTxtPath | ForEach-Object {
                                    $printerPath = $_
                                    if (-not ([string]::IsNullOrWhiteSpace($printerPath))) {
                                        Write-Verbose "  Attempting to add printer: $printerPath"
                                        try {
                                            # Check if printer already exists (optional, AddWindowsPrinterConnection might handle it)
                                            # if (-not (Get-Printer -Name $printerPath -ErrorAction SilentlyContinue)) {
                                                 $wsNet.AddWindowsPrinterConnection($printerPath)
                                                 Write-Verbose "    Added printer connection (or it already existed)."
                                            # } else { Write-Verbose "Printer '$printerPath' already exists." }
                                        } catch {
                                             Write-Warning "    Failed to add printer '$printerPath': $($_.Exception.Message)"
                                        }
                                    } else { Write-Verbose "  Skipping empty line in Printers.txt"}
                                }
                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                Write-Verbose "Finished processing printers."
                            } catch {
                                 Write-Warning "Error processing printer restorations: $($_.Exception.Message)"
                            }
                        } else { Write-Warning "Printers backup file (Printers.txt) not found." }
                    } else { Write-Verbose "Skipping Printers restore (unchecked)."}

                    $controls.txtProgress.Text = "Restore completed from: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }
                    Write-Verbose "--- Restore Operation Finished ---"

                } # End if/else ($IsBackup)

                # --- Operation Completion ---
                Write-Verbose "Operation completed. Displaying success message."
                $controls.lblStatus.Content = "Operation completed successfully."
                [System.Windows.MessageBox]::Show("The $($controls.btnStart.Content) operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

            } catch {
                # --- Operation Failure ---
                $errorMessage = "Operation Failed: $($_.Exception.Message)"
                Write-Error $errorMessage
                Write-Verbose "Operation failed: $errorMessage"
                $controls.lblStatus.Content = "Operation Failed!"
                $controls.txtProgress.Text = $errorMessage
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $false
                [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                Write-Verbose "Operation finished (finally block). Re-enabling UI controls."
                # Re-enable UI elements
                 $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $true } }
                 # Restore Add/Remove button state based on mode
                 $controls.btnAddFile.IsEnabled = $IsBackup
                 $controls.btnAddFolder.IsEnabled = $IsBackup
                 $controls.btnRemove.IsEnabled = $IsBackup
                 Write-Verbose "Add/Remove button states reset."

                 $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                 Write-Verbose "Cursor reset."
            }
        }) # End btnStart.Add_Click

        # --- Show Window ---
        Write-Verbose "Showing main window."
        $window.ShowDialog() | Out-Null
        Write-Verbose "Main window closed."

    } catch {
        # --- Window Load Failure ---
        $errorMessage = "Failed to load main window: $($_.Exception.Message)"
        Write-Error $errorMessage
        # Use Write-Host as fallback if MessageBox fails
        Write-Host "FATAL ERROR: Failed to load main window: $errorMessage" -ForegroundColor Red
        try {
             [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } catch {
             Write-Warning "Could not display error message box."
        }
    } finally {
         Write-Verbose "Exiting Show-MainWindow function."
    }
}

#endregion Functions

# --- Main Execution --- <--------------- THIS BLOCK MUST COME *AFTER* #endregion
Write-Verbose "--- Script Starting ---"
try {
    # Determine mode (Backup = $true, Restore = $false)
    Write-Verbose "Calling Show-ModeDialog to determine operation mode."
    # Show-ModeDialog now returns boolean directly
    [bool]$script:isBackupMode = Show-ModeDialog # CALLING the function

    # Show the main window, passing the determined mode
    Write-Verbose "Calling Show-MainWindow with IsBackup = $script:isBackupMode"
    Show-MainWindow -IsBackup $script:isBackupMode # CALLING the function

} catch {
    # Catch errors during initial mode selection or window loading
    $errorMessage = "An unexpected error occurred during startup: $($_.Exception.Message)"
    Write-Error $errorMessage
    # Use Write-Host as fallback if MessageBox fails
    Write-Host "FATAL ERROR during startup: $errorMessage" -ForegroundColor Red
    try {
        [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } catch {
        Write-Warning "Could not display startup error message box."
    }
}

Write-Host "--- Script Execution Finished ---"

# --- Keep console open when double-clicked ---
# Check if running in console and not ISE or VSCode Integrated Console
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') {
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
}