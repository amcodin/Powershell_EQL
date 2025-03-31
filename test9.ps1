#requires -Version 3.0
#notworking

# --- Load Assemblies FIRST ---
Write-Host "Attempting to load .NET Assemblies..."
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-Host ".NET Assemblies loaded successfully."
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
        "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms", # Quick Access Pinned
        "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt", # Legacy Sticky Notes
        "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite", # Modern Sticky Notes
        "$env:APPDATA\google\googleearth\myplaces.kml", # Google Earth Places
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks", # Chrome Bookmarks
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Bookmarks" # Edge Bookmarks

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
             Write-Host "Specific path not found: $path"
        }
    }

    return $result


function Get-BAUPaths {
    [CmdletBinding()]
    param ()


    # Define specific paths/files to check
    $specificPaths = @(
        "$env:USERNAME\Pictures\" # User Pictures
        "$env:USERNAME\downloads\" # User Downloads
        "$env:USERNAME\videos\" # User Videos
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
             Write-Host "Specific path not found: $path"
        }
    }

    return $result
}


# --- GUI Functions ---

# Show mode selection dialog
function Show-ModeDialog {
    Write-Host "Entering Show-ModeDialog function."
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

    Write-Host "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Host "Mode selection dialog closed with result: $result"

    # Determine mode based on which button was effectively clicked
    $isBackupMode = ($result -eq [System.Windows.Forms.DialogResult]::OK)
    $modeString = if ($isBackupMode) { 'Backup' } else { 'Restore' }
    Write-Host "Determined mode: $modeString" -ForegroundColor Cyan

    # If restore mode, run updates immediately in a non-blocking way
    if (-not $isBackupMode) {
        Write-Host "Restore mode selected. Initiating background system updates job..." -ForegroundColor Yellow
        # Use Start-Job for non-blocking execution
        # Store the job object in a script-scoped variable to retrieve output later
        $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {
            # Functions need to be defined within the job's scope

            # --- Function Definitions INSIDE Job Scope ---
            function Set-GPupdate {
                # Use Write-Output or Write-Verbose inside jobs if you want to capture structured data.
                # Use Write-Host if you specifically want host output (will be captured by Receive-Job).
                Write-Host "JOB: Initiating Group Policy update..." -ForegroundColor Cyan
                try {
                    # Using cmd /c ensures the window closes automatically after gpupdate finishes
                    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
                    if ($process.ExitCode -eq 0) {
                        Write-Host "JOB: Group Policy update completed successfully." -ForegroundColor Green
                    } else {
                        # Use Write-Warning for warnings within the job
                        Write-Warning "JOB: Group Policy update process finished with exit code: $($process.ExitCode)."
                    }
                } catch {
                    # Use Write-Error for errors within the job
                    Write-Error "JOB: Failed to start GPUpdate process: $($_.Exception.Message)"
                }
                Write-Host "JOB: Exiting Set-GPupdate function."
            }

            # --- REFINED Start-ConfigManagerActions Function ---
            function Start-ConfigManagerActions {
                 param() # Keep param block even if empty for consistency
                 Write-Host "JOB: Entering Start-ConfigManagerActions function."

                 Write-Host "JOB: Attempting to trigger Configuration Manager client actions..."
                 $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"
                 $clientSDKNamespace = "root\ccm\clientsdk"
                 $clientClassName = "CCM_ClientUtilities" # Class used for triggering schedules via SDK
                 $scheduleMethodName = "TriggerSchedule"
                 $overallSuccess = $false # Track if *any* method successfully triggers actions
                 $cimAttemptedAndSucceeded = $false # Track if CIM method was tried and worked

                 # Define the actions to trigger
                 $scheduleActions = @(
                     @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'}
                     # Add more schedule IDs as needed
                 )
                 Write-Host "JOB: Defined $($scheduleActions.Count) CM actions to trigger."

                 # --- Check for CM Client Presence (Service Check) ---
                 Write-Host "JOB: Checking for Configuration Manager client service (CcmExec)..."
                 $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
                 if (-not $ccmService) {
                     Write-Warning "JOB: Configuration Manager client service (CcmExec) not found. Skipping CM actions."
                     Write-Host "JOB: Exiting Start-ConfigManagerActions function (CM service not found)."
                     return $false # Indicate failure to trigger actions
                 } elseif ($ccmService.Status -ne 'Running') {
                     Write-Warning "JOB: Configuration Manager client service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping CM actions."
                     Write-Host "JOB: Exiting Start-ConfigManagerActions function (CM service not running)."
                     return $false # Indicate failure to trigger actions
                 } else {
                     Write-Host "JOB: Configuration Manager client service (CcmExec) found and running."
                 }

                 # --- Method 1: Attempt using WMI/CIM (Preferred) ---
                 Write-Host "JOB: Attempting Method 1: Triggering actions via CIM ($clientSDKNamespace -> $clientClassName)..."
                 $cimMethodSuccess = $true # Assume success unless an error occurs or class not found
                 try {
                     Write-Host "JOB: Checking for CIM class '$clientClassName' in namespace '$clientSDKNamespace'."
                     if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                          Write-Host "JOB: CIM Class found. Proceeding to trigger schedules via CIM."
                          foreach ($action in $scheduleActions) {
                             Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                             try {
                                 Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop
                                 Write-Host "JOB:     $($action.Name) triggered successfully via CIM."
                             } catch {
                                 Write-Warning "JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"
                                 $cimMethodSuccess = $false # Mark CIM method as partially/fully failed
                             }
                         }
                         # If loop completed without errors setting $cimMethodSuccess to false, it was successful
                         if ($cimMethodSuccess) {
                             $cimAttemptedAndSucceeded = $true
                             $overallSuccess = $true
                             Write-Host "JOB: All actions successfully triggered via CIM." -ForegroundColor Green
                         } else {
                             Write-Warning "JOB: One or more actions failed to trigger via CIM."
                         }
                     } else {
                          Write-Warning "JOB: CIM Class '$clientClassName' not found in namespace '$clientSDKNamespace'. Cannot use CIM method."
                          $cimMethodSuccess = $false # CIM method cannot be used
                     }
                 } catch {
                     Write-Error "JOB: An unexpected error occurred during CIM attempt: $($_.Exception.Message)"
                     $cimMethodSuccess = $false # Mark CIM method as failed due to unexpected error
                 }

                 # --- Method 2: Fallback to ccmexec.exe (If CIM failed or wasn't fully successful) ---
                 if (-not $cimAttemptedAndSucceeded) {
                     Write-Host "JOB: CIM method did not complete successfully or was not available. Attempting Method 2: Fallback via ccmexec.exe..."

                     # Check if ccmexec.exe exists *now*
                     Write-Host "JOB: Checking for executable: $ccmExecPath"
                     if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                         Write-Host "JOB: Found $ccmExecPath. Proceeding to trigger schedules via executable."
                         $execMethodSuccess = $true # Assume success unless errors occur
                         foreach ($action in $scheduleActions) {
                             Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                             # Note: Triggering via ccmexec.exe is less reliable and provides less feedback
                             try {
                                 # Using -TriggerSchedule requires the GUID format
                                 $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop
                                 if ($process.ExitCode -ne 0) {
                                     Write-Warning "JOB:     $($action.Name) action via ccmexec.exe finished with exit code $($process.ExitCode). (This might still be okay)"
                                     # Don't necessarily mark as failure, as ccmexec often returns non-zero
                                 } else {
                                      Write-Host "JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)."
                                 }
                             } catch {
                                  Write-Warning "JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"
                                  $execMethodSuccess = $false # Mark exec method as failed if execution fails
                             }
                         }
                         # If the loop completed and no execution errors occurred, consider it successful
                         if ($execMethodSuccess) {
                             $overallSuccess = $true # Mark overall success if fallback worked
                             Write-Host "JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green
                         } else {
                             Write-Warning "JOB: One or more actions failed to execute via ccmexec.exe."
                         }
                     } else {
                         Write-Warning "JOB: Fallback executable not found at $ccmExecPath. Cannot use ccmexec.exe method."
                         # $overallSuccess remains false if CIM also failed
                     }
                 } # End Fallback Check

                 # --- Final Status ---
                 if ($overallSuccess) {
                     Write-Host "JOB: Configuration Manager actions attempt finished. At least one method appears to have triggered actions successfully." -ForegroundColor Green
                 } else {
                     Write-Warning "JOB: Configuration Manager actions attempt finished, but neither CIM nor ccmexec.exe methods could be confirmed as fully successful or available."
                 }
                 Write-Host "JOB: Exiting Start-ConfigManagerActions function."
                 return $overallSuccess # Return true if *any* method likely succeeded, false otherwise
            }
            # --- End REFINED Start-ConfigManagerActions Function ---

            # Execute the functions defined above within this job's scope
            Set-GPupdate
            Start-ConfigManagerActions # Call the refined function
            Write-Host "JOB: Background updates finished." # Add completion message inside job
        }
        # Inform user job is running and output will be shown later
        Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown after main window closes. Main window will load now." -ForegroundColor Yellow
    }

    Write-Host "Exiting Show-ModeDialog function."
    return $isBackupMode
}

# Show main window
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [bool]$IsBackup
    )
    $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
    Write-Host "Entering Show-MainWindow function. Mode: $modeString"

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
                                <!-- CheckBox is always enabled -->
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
        Write-Host "Parsing XAML for main window."
        # Parse XAML
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "XAML loaded successfully."

        # Use a PSObject wrapper for the DataContext if needed for other bindings
        Write-Host "Setting Window DataContext."
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }


        # Use a hashtable to store controls for easy access
        Write-Host "Finding controls in main window."
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
        Write-Host "Controls found and stored in hashtable."

        # --- Window Initialization ---
        Write-Host "Initializing window controls based on mode."
        $controls.lblMode.Content = if ($IsBackup) { "Mode: Backup" } else { "Mode: Restore" }
        $controls.btnStart.Content = if ($IsBackup) { "Backup" } else { "Restore" }

        # Add/Remove buttons remain enabled


        # Set default path
        $defaultPath = "C:\LocalData"
        Write-Host "Checking default path: $defaultPath"
        if (-not (Test-Path $defaultPath)) {
            Write-Host "Default path not found. Attempting to create."
            try {
                New-Item -Path $defaultPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Host "Default path created."
            } catch {
                 Write-Warning "Could not create default path: $defaultPath. Please select a location manually."
                 $defaultPath = $env:USERPROFILE # Fallback
                 Write-Host "Using fallback default path: $defaultPath"
            }
        } else { Write-Host "Default path exists."}
        $controls.txtSaveLoc.Text = $defaultPath

        # Load initial items based on mode
        Write-Host "Loading initial items for ListView based on mode."
        if ($IsBackup) {
            Write-Host "Backup Mode: Getting default paths using Get-BackupPaths."
            # Load default backup paths
            $paths = $null # Initialize to null
            try {
                $paths = Get-BackupPaths # Calls the function
                Write-Host "Get-BackupPaths call completed."
            } catch {
                 Write-Error "Error calling Get-BackupPaths: $($_.Exception.Message)"
            }

            if ($paths -ne $null -and $paths.Count -gt 0) {
                 Write-Host "Get-BackupPaths returned $($paths.Count) items. First item type: $($paths[0].GetType().FullName)"
                 Write-Host "Attempting to add 'IsSelected' property..."
                 $pathsWithSelection = $null # Initialize
                 try {
                    $pathsCollection = @($paths)
                    Write-Host "Processing $($pathsCollection.Count) items with Add-Member."
                    $pathsWithSelection = $pathsCollection | ForEach-Object {
                         $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                    }
                    Write-Host "'IsSelected' property added. Type of result: $($pathsWithSelection.GetType().FullName)"
                    $pathsWithSelectionCount = if ($pathsWithSelection -is [array]) { $pathsWithSelection.Count } elseif ($pathsWithSelection -ne $null) { 1 } else { 0 }
                    Write-Host "Count after Add-Member: $pathsWithSelectionCount"
                    if ($pathsWithSelection -eq $null -and $pathsCollection.Count -gt 0) {Write-Warning "pathsWithSelection is null after Add-Member despite input collection!"}

                 } catch {
                     Write-Error "Error during Add-Member: $($_.Exception.Message)"
                 }

                 $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                 Write-Host "Initialized empty itemsList."

                 if ($pathsWithSelection -ne $null) {
                     Write-Host "Attempting to populate itemsList from pathsWithSelection..."
                     try {
                        $selectionCollection = @($pathsWithSelection)
                        Write-Host "Iterating through $($selectionCollection.Count) items to add to itemsList."
                        $selectionCollection | ForEach-Object {
                            if ($_ -ne $null) {
                                $itemsList.Add($_)
                            } else {
                                Write-Warning "Skipping null item found in pathsWithSelection collection."
                            }
                        }
                        Write-Host "Finished populating itemsList. Count: $($itemsList.Count)"
                     } catch {
                         Write-Error "Error populating itemsList: $($_.Exception.Message)"
                         $itemsList.Clear()
                     }
                 } else {
                     Write-Warning "Cannot populate itemsList because pathsWithSelection is null."
                 }

                 Write-Host "Checking if lvwFiles control exists..."
                 if ($controls['lvwFiles'] -ne $null) {
                    Write-Host "Assigning itemsList (Count: $($itemsList.Count)) to ListView ItemsSource."
                    $controls.lvwFiles.ItemsSource = $itemsList
                    Write-Host "Assigned itemsList to ListView ItemsSource."
                 } else {
                     Write-Error "ListView control ('lvwFiles') not found!"
                     throw "ListView control ('lvwFiles') could not be found in the XAML."
                 }

            } else {
                 Write-Warning "Get-BackupPaths returned no items or was null. ListView will be empty."
                 Write-Host "Checking if lvwFiles control exists before assigning empty list..."
                 if ($controls['lvwFiles'] -ne $null) {
                    $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new()
                    Write-Host "Assigned empty list to ListView ItemsSource."
                 } else {
                     Write-Error "ListView control ('lvwFiles') not found when trying to assign empty list!"
                     throw "ListView control ('lvwFiles') could not be found in the XAML."
                 }
            }
        }
        elseif (Test-Path $defaultPath) {
            Write-Host "Restore Mode: Checking for latest backup in $defaultPath."
            # Restore Mode: Look for most recent backup in default path
            $latestBackup = Get-ChildItem -Path $defaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                Write-Host "Found latest backup: $($latestBackup.FullName)"
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
                $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                $backupItems | ForEach-Object { $itemsList.Add($_) }
                $controls.lvwFiles.ItemsSource = $itemsList
                Write-Host "Populated ListView with $($itemsList.Count) items from latest backup."
            } else {
                 Write-Host "No backups found in $defaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $defaultPath. Please browse."
                 $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new() # Ensure it's an empty list
            }
        } else {
             Write-Host "Restore Mode: Default path $defaultPath does not exist."
             $controls.lvwFiles.ItemsSource = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        Write-Host "Finished loading initial items."

        # --- Event Handlers ---
        Write-Host "Assigning event handlers."
        $controls.btnBrowse.Add_Click({
            Write-Host "Browse button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = if ($IsBackup) { "Select location to save backup" } else { "Select backup folder to restore from" }
            if(Test-Path $controls.txtSaveLoc.Text){
                 $dialog.SelectedPath = $controls.txtSaveLoc.Text
            } else {
                 $dialog.SelectedPath = $defaultPath # Fallback if current text isn't valid
            }
            $dialog.ShowNewFolderButton = $IsBackup # Only allow creating new folders in backup mode

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            Write-Host "Showing FolderBrowserDialog."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose() # Dispose the temporary owner form

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath
                Write-Host "Folder selected: $selectedPath"
                $controls.txtSaveLoc.Text = $selectedPath

                if (-not $IsBackup) {
                    Write-Host "Restore Mode: Loading items from selected backup folder."
                    $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                    Write-Host "Checking for log file: $logFilePath"
                    if (Test-Path -Path $logFilePath) {
                         Write-Host "Log file found. Populating ListView."
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
                        $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                        $backupItems | ForEach-Object { $itemsList.Add($_) }
                        $controls.lvwFiles.ItemsSource = $itemsList
                        $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                        Write-Host "ListView updated with $($itemsList.Count) items from selected backup."
                    } else {
                         Write-Warning "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         $controls.lvwFiles.ItemsSource = $null # Clear list view
                         $controls.lblStatus.Content = "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         [System.Windows.MessageBox]::Show("The selected folder does not appear to be a valid backup. It's missing the 'FileList_Backup.csv' log file.", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                } else {
                     $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Corrected Add File/Folder handlers
        $controls.btnAddFile.Add_Click({
            Write-Host "Add File button clicked."
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = "Select File(s) to Add"
            $dialog.Multiselect = $true

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            Write-Host "Showing OpenFileDialog."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($dialog.FileNames.Count) file(s) selected."
                $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    Write-Host "Copying existing items to new list."
                    $controls.lvwFiles.ItemsSource | ForEach-Object {
                        if (-not ($_.PSObject.Properties.Name -contains 'IsSelected')) {
                            $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true
                        }
                        $newItemsList.Add($_)
                    }
                }

                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    if (-not ($newItemsList.Path -contains $file)) {
                        Write-Host "Adding file: $file"
                        $newItemsList.Add([PSCustomObject]@{
                            Name = [System.IO.Path]::GetFileName($file)
                            Type = "File"
                            Path = $file
                            IsSelected = $true # Add IsSelected here too
                        })
                        $addedCount++
                    } else { Write-Host "Skipping duplicate file: $file"}
                }
                $controls.lvwFiles.ItemsSource = $newItemsList # Reassign ItemsSource
                Write-Host "Updated ListView ItemsSource. Added $addedCount new file(s)."
            } else { Write-Host "File selection cancelled."}
        })

        $controls.btnAddFolder.Add_Click({
             Write-Host "Add Folder button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = "Select Folder to Add"
            $dialog.ShowNewFolderButton = $false

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            Write-Host "Showing FolderBrowserDialog for adding folder."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath
                 Write-Host "Folder selected to add: $selectedPath"
                 $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
                 if ($controls.lvwFiles.ItemsSource -ne $null) {
                     Write-Host "Copying existing items to new list."
                     $controls.lvwFiles.ItemsSource | ForEach-Object {
                         if (-not ($_.PSObject.Properties.Name -contains 'IsSelected')) {
                             $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true
                         }
                         $newItemsList.Add($_)
                     }
                 }

                 if (-not ($newItemsList.Path -contains $selectedPath)) {
                    Write-Host "Adding folder: $selectedPath"
                    $newItemsList.Add([PSCustomObject]@{
                        Name = [System.IO.Path]::GetFileName($selectedPath)
                        Type = "Folder"
                        Path = $selectedPath
                        IsSelected = $true # Add IsSelected here too
                    })
                    $controls.lvwFiles.ItemsSource = $newItemsList # Reassign ItemsSource
                    Write-Host "Updated ListView ItemsSource with new folder."
                 } else {
                    Write-Host "Skipping duplicate folder: $selectedPath"
                 }

            } else { Write-Host "Folder selection cancelled."}
        })

        $controls.btnRemove.Add_Click({
            Write-Host "Remove Selected button clicked."
            $selectedObjects = @($controls.lvwFiles.SelectedItems) # Ensure it's an array
            if ($selectedObjects.Count -gt 0) {
                Write-Host "Removing $($selectedObjects.Count) selected item(s)."
                $itemsToKeep = [System.Collections.Generic.List[PSCustomObject]]::new()
                if ($controls.lvwFiles.ItemsSource -ne $null) {
                    $controls.lvwFiles.ItemsSource | Where-Object { $selectedObjects -notcontains $_ } | ForEach-Object {
                        $itemsToKeep.Add($_)
                    }
                }
                $controls.lvwFiles.ItemsSource = $itemsToKeep # Reassign ItemsSource
                Write-Host "ListView ItemsSource updated after removal."
            } else { Write-Host "No items selected to remove."}
        })


        # --- Start Button Logic ---
        $controls.btnStart.Add_Click({
            $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
            Write-Host "Start button clicked. Mode: $modeString"

            $location = $controls.txtSaveLoc.Text
            Write-Host "Selected location: $location"
            if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) { # Check it's a directory
                Write-Warning "Invalid location selected."
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            Write-Host "Disabling UI controls and setting wait cursor."
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
                    Write-Host "--- Starting Backup Operation ---"
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $username = $env:USERNAME -replace '[^a-zA-Z0-9]', '_' # Sanitize username for path
                    $backupRootPath = Join-Path $operationPath "Backup_${username}_$timestamp"
                    Write-Host "Backup root path: $backupRootPath"

                    try {
                        Write-Host "Creating backup directory..."
                        if (-not (Test-Path $backupRootPath)) {
                            New-Item -Path $backupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                            Write-Host "Backup directory created."
                        } else { Write-Host "Backup directory already exists (should not happen with timestamp)."}
                    } catch {
                        throw "Failed to create backup directory: $backupRootPath. Error: $($_.Exception.Message)"
                    }

                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    Write-Host "Creating log file: $csvLogPath"
                    "OriginalFullPath,BackupRelativePath" | Set-Content -Path $csvLogPath -Encoding UTF8

                    # Filter ItemsSource based on IsSelected
                    $itemsToBackup = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected } # Ensure it's an array and filtered
                    if (-not $itemsToBackup -or $itemsToBackup.Count -eq 0) {
                         throw "No items selected (checked) for backup."
                    }
                    Write-Host "Found $($itemsToBackup.Count) CHECKED items in ListView to process for backup."

                    # Estimate total files for progress (based on CHECKED items)
                    Write-Host "Estimating total files for progress bar..."
                    $totalFilesEstimate = 0
                    $itemsToBackup | ForEach-Object { # Iterate only through checked items
                        if ($_.Type -eq 'Folder' -and (Test-Path $_.Path)) {
                            try { $totalFilesEstimate += (Get-ChildItem $_.Path -Recurse -File -Force -ErrorAction SilentlyContinue).Count } catch {}
                        } elseif (Test-Path $_.Path -PathType Leaf) { # Only count existing files
                            $totalFilesEstimate++
                        }
                    }
                    if ($controls.chkNetwork.IsChecked) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked) { $totalFilesEstimate++ }
                    Write-Host "Estimated total files/items: $totalFilesEstimate"

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 } # Avoid max=0
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Process Files/Folders for Backup (using filtered $itemsToBackup)
                    Write-Host "Starting processing of CHECKED files/folders for backup..."
                    foreach ($item in $itemsToBackup) { # Only iterates through checked items now
                        Write-Host "Processing item: $($item.Name) ($($item.Type)) - Path: $($item.Path)"
                        $controls.txtProgress.Text = "Processing: $($item.Name)"
                        $sourcePath = $item.Path

                        if (-not (Test-Path $sourcePath)) {
                            Write-Warning "Source path not found, skipping: $sourcePath"
                            continue
                        }

                        if ($item.Type -eq "Folder") {
                            Write-Host "  Item is a folder. Processing recursively..."
                            try {
                                Get-ChildItem -Path $sourcePath -Recurse -File -Force -ErrorAction Stop | ForEach-Object {
                                    $originalFileFullPath = $_.FullName
                                    $relativeFilePath = $originalFileFullPath.Substring($sourcePath.TrimEnd('\').Length).TrimStart('\')
                                    $backupRelativePath = Join-Path $item.Name $relativeFilePath # Path relative to backup root folder

                                    $targetBackupPath = Join-Path $backupRootPath $backupRelativePath
                                    $targetBackupDir = [System.IO.Path]::GetDirectoryName($targetBackupPath)

                                    if (-not (Test-Path $targetBackupDir)) {
                                        Write-Host "    Creating directory: $targetBackupDir"
                                        New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                    }

                                    Write-Host "    Copying '$($_.Name)' to '$targetBackupPath'"
                                    Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                    Write-Host "    Logging: `"$originalFileFullPath`",`"$backupRelativePath`""
                                    "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                    $filesProcessed++
                                    if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                    $controls.txtProgress.Text = "Backed up: $($_.Name)"
                                }
                                Write-Host "  Finished processing folder: $($item.Name)"
                            } catch {
                                 Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                            }
                        } else { # Single File
                             Write-Host "  Item is a file. Processing..."
                             try {
                                $originalFileFullPath = $sourcePath
                                $backupRelativePath = $item.Name # Store file directly under the backup root
                                $targetBackupPath = Join-Path $backupRootPath $backupRelativePath

                                Write-Host "    Copying '$($item.Name)' to '$targetBackupPath'"
                                Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                Write-Host "    Logging: `"$originalFileFullPath`",`"$backupRelativePath`""
                                "`"$originalFileFullPath`",`"$backupRelativePath`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                $controls.txtProgress.Text = "Backed up: $($item.Name)"
                             } catch {
                                 Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                             }
                        }
                    } # End foreach item
                    Write-Host "Finished processing CHECKED files/folders for backup."

                    # Backup Network Drives
                    if ($controls.chkNetwork.IsChecked) {
                        Write-Host "Processing Network Drives backup..."
                        $controls.txtProgress.Text = "Backing up network drives..."
                        try {
                            Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop |
                                Select-Object Name, ProviderName |
                                Export-Csv -Path (Join-Path $backupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
                            $filesProcessed++
                            if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                            Write-Host "Network drives backed up successfully." -ForegroundColor Green
                        } catch {
                             Write-Warning "Failed to backup network drives: $($_.Exception.Message)"
                        }
                    } else { Write-Host "Skipping Network Drives backup (unchecked)."}

                    # Backup Printers
                    if ($controls.chkPrinters.IsChecked) {
                        Write-Host "Processing Printers backup..."
                        $controls.txtProgress.Text = "Backing up printers..."
                        try {
                            Get-WmiObject -Class Win32_Printer -Filter "Local = False" -ErrorAction Stop |
                                Select-Object -ExpandProperty Name |
                                Set-Content -Path (Join-Path $backupRootPath "Printers.txt") -Encoding UTF8 -ErrorAction Stop
                             $filesProcessed++
                             if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                             Write-Host "Printers backed up successfully." -ForegroundColor Green
                        } catch {
                             Write-Warning "Failed to backup printers: $($_.Exception.Message)"
                        }
                    } else { Write-Host "Skipping Printers backup (unchecked)."}

                    $controls.txtProgress.Text = "Backup completed successfully to: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }
                    Write-Host "--- Backup Operation Finished ---"


                # ===================
                # --- RESTORE Logic ---
                # ===================
                } else {
                    Write-Host "--- Starting Restore Operation ---"
                    $backupRootPath = $operationPath # In restore mode, the selected location IS the backup root
                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    Write-Host "Restore source (backup root): $backupRootPath"
                    Write-Host "Log file path: $csvLogPath"

                    if (-not (Test-Path $csvLogPath -PathType Leaf)) {
                        throw "Backup log file 'FileList_Backup.csv' not found in the selected location: $backupRootPath"
                    }

                    # Restore mode updates (GPUpdate, CM Actions) were initiated earlier via Start-Job
                    Write-Host "Background updates (GPUpdate/CM) should have been initiated earlier (check console output after closing this window)."

                    Write-Host "Importing backup log file..."
                    $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8
                    if (-not $backupLog) {
                         throw "Backup log file is empty or could not be read: $csvLogPath"
                    }
                    Write-Host "Imported $($backupLog.Count) entries from log file."

                    # --- Get Selected Items from ListView ---
                    $listViewItems = @($controls.lvwFiles.ItemsSource)
                    $selectedItemsFromListView = $listViewItems | Where-Object { $_.IsSelected } # Filter based on CheckBox

                    if (-not $selectedItemsFromListView) {
                        throw "No items selected (checked) in the list for restore."
                    }
                    $selectedTopLevelNames = $selectedItemsFromListView | Select-Object -ExpandProperty Name
                    Write-Host "Found $($selectedItemsFromListView.Count) CHECKED items in ListView for restore: $($selectedTopLevelNames -join ', ')"

                    # --- Filter Log Entries Based on Selection ---
                    Write-Host "Filtering log entries based on ListView selection..."
                    $logEntriesToRestore = $backupLog | Where-Object {
                        $topLevelName = ($_.BackupRelativePath -split '[\\/]', 2)[0]
                        $selectedTopLevelNames -contains $topLevelName
                    }

                    if (-not $logEntriesToRestore) {
                        throw "None of the selected (checked) items correspond to entries in the backup log."
                    }
                    Write-Host "Filtered log. $($logEntriesToRestore.Count) log entries will be processed for restore."

                    # Estimate progress based on selected log entries
                    $totalFilesEstimate = $logEntriesToRestore.Count
                    if ($controls.chkNetwork.IsChecked) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked) { $totalFilesEstimate++ }
                    Write-Host "Estimated total items for restore progress: $totalFilesEstimate"

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 } # Avoid max=0
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Restore Files/Folders from Filtered Log
                    Write-Host "Starting restore of files/folders from filtered log..."
                    foreach ($entry in $logEntriesToRestore) {
                        $originalFileFullPath = $entry.OriginalFullPath
                        $backupRelativePath = $entry.BackupRelativePath
                        $sourceBackupPath = Join-Path $backupRootPath $backupRelativePath

                        Write-Host "Processing restore entry: Source='$sourceBackupPath', Target='$originalFileFullPath'"
                        $controls.txtProgress.Text = "Restoring: $(Split-Path $originalFileFullPath -Leaf)"

                        if (Test-Path $sourceBackupPath -PathType Leaf) { # Ensure source exists in backup
                            try {
                                $targetRestoreDir = [System.IO.Path]::GetDirectoryName($originalFileFullPath)
                                if (-not (Test-Path $targetRestoreDir)) {
                                    Write-Host "  Creating target directory: $targetRestoreDir"
                                    New-Item -Path $targetRestoreDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                }

                                Write-Host "  Copying '$sourceBackupPath' to '$originalFileFullPath'"
                                Copy-Item -Path $sourceBackupPath -Destination $originalFileFullPath -Force -ErrorAction Stop

                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }

                            } catch {
                                Write-Warning "Failed to restore '$originalFileFullPath' from '$sourceBackupPath': $($_.Exception.Message)"
                            }
                        } else {
                            Write-Warning "Source file not found in backup, skipping restore: $sourceBackupPath (Expected for: $originalFileFullPath)"
                        }
                    } # End foreach entry
                    Write-Host "Finished restoring files/folders from log."

                    # Restore Network Drives (Not dependent on ListView selection)
                    if ($controls.chkNetwork.IsChecked) {
                        Write-Host "Processing Network Drives restore..."
                        $controls.txtProgress.Text = "Restoring network drives..."
                        $drivesCsvPath = Join-Path $backupRootPath "Drives.csv"
                        if (Test-Path $drivesCsvPath) {
                            Write-Host "Found Drives.csv. Processing mappings..."
                            try {
                                Import-Csv $drivesCsvPath | ForEach-Object {
                                    $driveLetter = $_.Name.TrimEnd(':')
                                    $networkPath = $_.ProviderName
                                    Write-Host "  Checking mapping: $driveLetter -> $networkPath"
                                    if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\' ) {
                                        if (-not (Test-Path -LiteralPath "$($driveLetter):")) {
                                            try {
                                                Write-Host "    Mapping $driveLetter to $networkPath"
                                                New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -ErrorAction Stop
                                            } catch {
                                                 Write-Warning "    Failed to map drive $driveLetter`: $($_.Exception.Message)"
                                            }
                                        } else {
                                             Write-Host "    Drive $driveLetter already exists, skipping."
                                        }
                                    } else {
                                         Write-Warning "    Skipping invalid drive mapping: Name='$($_.Name)', Provider='$networkPath'"
                                    }
                                }
                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                Write-Host "Finished processing network drive mappings."
                            } catch {
                                 Write-Warning "Error processing network drive restorations: $($_.Exception.Message)"
                            }
                        } else { Write-Warning "Network drives backup file (Drives.csv) not found." }
                    } else { Write-Host "Skipping Network Drives restore (unchecked)."}

                    # Restore Printers (Not dependent on ListView selection)
                    if ($controls.chkPrinters.IsChecked) {
                        Write-Host "Processing Printers restore..."
                        $controls.txtProgress.Text = "Restoring printers..."
                        $printersTxtPath = Join-Path $backupRootPath "Printers.txt"
                        if (Test-Path $printersTxtPath) {
                             Write-Host "Found Printers.txt. Processing printers..."
                            try {
                                $wsNet = New-Object -ComObject WScript.Network # Use COM object for broader compatibility
                                Get-Content $printersTxtPath | ForEach-Object {
                                    $printerPath = $_
                                    if (-not ([string]::IsNullOrWhiteSpace($printerPath))) {
                                        Write-Host "  Attempting to add printer: $printerPath"
                                        try {
                                             $wsNet.AddWindowsPrinterConnection($printerPath)
                                             Write-Host "    Added printer connection (or it already existed)."
                                        } catch {
                                             Write-Warning "    Failed to add printer '$printerPath': $($_.Exception.Message)"
                                        }
                                    } else { Write-Host "  Skipping empty line in Printers.txt"}
                                }
                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                Write-Host "Finished processing printers."
                            } catch {
                                 Write-Warning "Error processing printer restorations: $($_.Exception.Message)"
                            }
                        } else { Write-Warning "Printers backup file (Printers.txt) not found." }
                    } else { Write-Host "Skipping Printers restore (unchecked)."}

                    $controls.txtProgress.Text = "Restore completed from: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }
                    Write-Host "--- Restore Operation Finished ---"

                } # End if/else ($IsBackup)

                # --- Operation Completion ---
                Write-Host "Operation completed. Displaying success message." -ForegroundColor Green
                $controls.lblStatus.Content = "Operation completed successfully."
                [System.Windows.MessageBox]::Show("The $($controls.btnStart.Content) operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

            } catch {
                # --- Operation Failure ---
                $errorMessage = "Operation Failed: $($_.Exception.Message)"
                Write-Error $errorMessage
                Write-Host "Operation failed: $errorMessage"
                $controls.lblStatus.Content = "Operation Failed!"
                $controls.txtProgress.Text = $errorMessage
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $false
                [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            } finally {
                Write-Host "Operation finished (finally block). Re-enabling UI controls."
                # Re-enable UI elements
                 $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $true } }
                 # Add/Remove buttons remain enabled

                 $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                 Write-Host "Cursor reset."
            }
        }) # End btnStart.Add_Click

        # --- Show Window ---
        Write-Host "Showing main window."
        $window.ShowDialog() | Out-Null
        Write-Host "Main window closed."

    } catch {
        # --- Window Load Failure ---
        $errorMessage = "Failed to load main window: $($_.Exception.Message)"
        Write-Error $errorMessage
        Write-Host "FATAL ERROR: Failed to load main window: $errorMessage" -ForegroundColor Red
        try {
             [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } catch {
             Write-Warning "Could not display error message box."
        }
    } finally {
         Write-Host "Exiting Show-MainWindow function."
    }
}

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue # Clear previous job variable if it exists

try {
    # Determine mode (Backup = $true, Restore = $false)
    Write-Host "Calling Show-ModeDialog to determine operation mode."
    [bool]$script:isBackupMode = Show-ModeDialog # CALLING the function

    # Show the main window, passing the determined mode
    Write-Host "Calling Show-MainWindow with IsBackup = $script:isBackupMode"
    Show-MainWindow -IsBackup $script:isBackupMode # CALLING the function

} catch {
    # Catch errors during initial mode selection or window loading
    $errorMessage = "An unexpected error occurred during startup: $($_.Exception.Message)"
    Write-Error $errorMessage
    Write-Host "FATAL ERROR during startup: $errorMessage" -ForegroundColor Red
    try {
        [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    } catch {
        Write-Warning "Could not display startup error message box."
    }
} finally {
    # Check for and receive output from the background job AFTER the main window closes
    # Check if the variable exists AND is not null
    if ((Get-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue) -ne $null -and $script:updateJob -ne $null) {
        Write-Host "`n--- Waiting for background update job (GPUpdate/CM Actions) to complete... ---" -ForegroundColor Yellow
        # Wait for the job to finish before trying to receive output
        Wait-Job $script:updateJob | Out-Null
        Write-Host "--- Background Update Job Output (GPUpdate/CM Actions): ---" -ForegroundColor Yellow
        # Receive-Job retrieves output (including Write-Host, Write-Warning, Write-Error)
        Receive-Job $script:updateJob
        Remove-Job $script:updateJob # Clean up the job object
        Write-Host "--- End of Background Update Job Output ---" -ForegroundColor Yellow
    } else {
        Write-Host "`nNo background update job was started (or it was already cleaned up)." -ForegroundColor Gray
    }
}


Write-Host "--- Script Execution Finished ---"

# --- Keep console open when double-clicked ---
# Check if running in console and not ISE or VSCode Integrated Console
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') {
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
}