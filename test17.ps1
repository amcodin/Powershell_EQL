#requires -Version 3.0

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

#region Global Variables & Configuration

# Define path templates using keys for easier management
# Use placeholders like {USERNAME} which will be replaced later
$script:pathTemplates = @{
    Signatures        = "{USERPROFILE}\AppData\Roaming\Microsoft\Signatures"
    # UserFolder        = "{USERPROFILE}" # Be cautious with this - very large! Commented out by default.
    QuickAccess       = "{USERPROFILE}\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"
    # TempFolder        = "C:\Temp" # System-wide, not user-specific usually. Consider if needed.
    StickyNotesLegacy = "{USERPROFILE}\AppData\Roaming\Microsoft\Sticky Notes\StickyNotes.snt"
    StickyNotesModernDB = "{USERPROFILE}\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
    GoogleEarthPlaces = "{USERPROFILE}\AppData\LocalLow\Google\GoogleEarth\myplaces.kml" # Corrected path to LocalLow
    ChromeBookmarks   = "{USERPROFILE}\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
    EdgeBookmarks     = "{USERPROFILE}\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks"
    # Add more templates here
}

#endregion Global Variables & Configuration

#region Functions

# --- Path Resolution Functions ---

# Resolves path templates for a specific user profile path
function Resolve-UserPathTemplates {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserProfilePath, # e.g., C:\Users\jdoe or X:\Users\jdoe
        [Parameter(Mandatory=$true)]
        [hashtable]$Templates
    )

    $resolvedPaths = [System.Collections.Generic.List[PSCustomObject]]::new()
    $userName = Split-Path $UserProfilePath -Leaf # Extract username for logging

    Write-Host "Resolving paths based on profile: $UserProfilePath"

    foreach ($key in $Templates.Keys) {
        $template = $Templates[$key]
        # Replace placeholder with the actual profile path
        # Use -replace with regex escape for backslashes in path if needed, but direct replace often works here.
        $resolvedPath = $template.Replace('{USERPROFILE}', $UserProfilePath.TrimEnd('\'))

        # Check if the resolved path exists
        if (Test-Path -LiteralPath $resolvedPath -ErrorAction SilentlyContinue) {
            $pathType = if (Test-Path -LiteralPath $resolvedPath -PathType Container) { "Folder" } else { "File" }
            Write-Host "  Found: '$key' -> $resolvedPath ($pathType)" -ForegroundColor Green
            $resolvedPaths.Add([PSCustomObject]@{
                Name = $key # Use the template key as the name
                Path = $resolvedPath
                Type = $pathType
                SourceKey = $key # Keep track of the original template key
            })
        } else {
            Write-Host "  Path not found or resolution failed: $resolvedPath (Derived from template key: $key)" -ForegroundColor Yellow
        }
    }
    return $resolvedPaths
}

# --- GUI Functions ---

# Show mode selection dialog (Added Express Mode)
function Show-ModeDialog {
    Write-Host "Entering Show-ModeDialog function."
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation Mode"
    $form.Size = New-Object System.Drawing.Size(350, 180) # Increased size
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.HelpButton = $false

    $lblDescription = New-Object System.Windows.Forms.Label
    $lblDescription.Text = "Choose the operation mode:"
    $lblDescription.Location = New-Object System.Drawing.Point(20, 20)
    $lblDescription.AutoSize = $true
    $form.Controls.Add($lblDescription)

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(40, 50)
    $btnBackup.Size = New-Object System.Drawing.Size(80, 30)
    $btnBackup.Text = "Backup"
    $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::Yes # Using Yes for Backup
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(130, 50)
    $btnRestore.Size = New-Object System.Drawing.Size(80, 30)
    $btnRestore.Text = "Restore"
    $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::No # Using No for Restore
    $form.Controls.Add($btnRestore)

    $btnExpress = New-Object System.Windows.Forms.Button
    $btnExpress.Location = New-Object System.Drawing.Point(220, 50)
    $btnExpress.Size = New-Object System.Drawing.Size(80, 30)
    $btnExpress.Text = "Express"
    $btnExpress.DialogResult = [System.Windows.Forms.DialogResult]::OK # Using OK for Express
    $form.Controls.Add($btnExpress)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(130, 100)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel # Using Cancel to exit
    $form.Controls.Add($btnCancel)

    # Set default button (optional, e.g., Backup)
    $form.AcceptButton = $btnBackup
    # Set cancel button (closes dialog)
    $form.CancelButton = $btnCancel

    Write-Host "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Host "Mode selection dialog closed with result: $result"

    # Determine mode based on DialogResult
    $selectedMode = switch ($result) {
        ([System.Windows.Forms.DialogResult]::Yes) { 'Backup' }
        ([System.Windows.Forms.DialogResult]::No) { 'Restore' }
        ([System.Windows.Forms.DialogResult]::OK) { 'Express' }
        Default { 'Cancel' } # Includes Cancel or closing the dialog
    }

    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan

    # Handle Cancel explicitly
    if ($selectedMode -eq 'Cancel') {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        # Optional: Pause if running directly in console
        if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
        Exit 0 # Exit gracefully
    }

    # If restore mode, run updates immediately in a non-blocking way
    # NOTE: Express mode handles its own updates *after* transfer
    if ($selectedMode -eq 'Restore') {
        Write-Host "Restore mode selected. Initiating background system updates job..." -ForegroundColor Yellow
        # Use Start-Job for non-blocking execution
        # Store the job object in a script-scoped variable to retrieve output later
        $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {
            # Functions need to be defined within the job's scope

            # --- Function Definitions INSIDE Job Scope ---
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
                Write-Host "JOB: Exiting Set-GPupdate function."
            }

            function Start-ConfigManagerActions {
                 param()
                 Write-Host "JOB: Entering Start-ConfigManagerActions function."
                 Write-Host "JOB: Attempting to trigger Configuration Manager client actions..."
                 $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"
                 $clientSDKNamespace = "root\ccm\clientsdk"
                 $clientClassName = "CCM_ClientUtilities"
                 $scheduleMethodName = "TriggerSchedule"
                 $overallSuccess = $false
                 $cimAttemptedAndSucceeded = $false

                 $scheduleActions = @(
                     @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'},
                     @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'}
                 )
                 Write-Host "JOB: Defined $($scheduleActions.Count) CM actions to trigger."

                 Write-Host "JOB: Checking for Configuration Manager client service (CcmExec)..."
                 $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
                 if (-not $ccmService) {
                     Write-Warning "JOB: Configuration Manager client service (CcmExec) not found. Skipping CM actions."
                     return $false
                 } elseif ($ccmService.Status -ne 'Running') {
                     Write-Warning "JOB: Configuration Manager client service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping CM actions."
                     return $false
                 } else {
                     Write-Host "JOB: Configuration Manager client service (CcmExec) found and running."
                 }

                 Write-Host "JOB: Attempting Method 1: Triggering actions via CIM ($clientSDKNamespace -> $clientClassName)..."
                 $cimMethodSuccess = $true
                 try {
                     if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                          Write-Host "JOB: CIM Class found. Proceeding to trigger schedules via CIM."
                          foreach ($action in $scheduleActions) {
                             Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                             try {
                                 Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop
                                 Write-Host "JOB:     $($action.Name) triggered successfully via CIM."
                             } catch {
                                 Write-Warning "JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"
                                 $cimMethodSuccess = $false
                             }
                         }
                         if ($cimMethodSuccess) {
                             $cimAttemptedAndSucceeded = $true
                             $overallSuccess = $true
                             Write-Host "JOB: All actions successfully triggered via CIM." -ForegroundColor Green
                         } else {
                             Write-Warning "JOB: One or more actions failed to trigger via CIM."
                         }
                     } else {
                          Write-Warning "JOB: CIM Class '$clientClassName' not found in namespace '$clientSDKNamespace'. Cannot use CIM method."
                          $cimMethodSuccess = $false
                     }
                 } catch {
                     Write-Error "JOB: An unexpected error occurred during CIM attempt: $($_.Exception.Message)"
                     $cimMethodSuccess = $false
                 }

                 if (-not $cimAttemptedAndSucceeded) {
                     Write-Host "JOB: CIM method did not complete successfully or was not available. Attempting Method 2: Fallback via ccmexec.exe..."
                     Write-Host "JOB: Checking for executable: $ccmExecPath"
                     if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                         Write-Host "JOB: Found $ccmExecPath. Proceeding to trigger schedules via executable."
                         $execMethodSuccess = $true
                         foreach ($action in $scheduleActions) {
                             Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                             try {
                                 $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop
                                 if ($process.ExitCode -ne 0) {
                                     Write-Warning "JOB:     $($action.Name) action via ccmexec.exe finished with exit code $($process.ExitCode). (This might still be okay)"
                                 } else {
                                      Write-Host "JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)."
                                 }
                             } catch {
                                  Write-Warning "JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"
                                  $execMethodSuccess = $false
                             }
                         }
                         if ($execMethodSuccess) {
                             $overallSuccess = $true
                             Write-Host "JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green
                         } else {
                             Write-Warning "JOB: One or more actions failed to execute via ccmexec.exe."
                         }
                     } else {
                         Write-Warning "JOB: Fallback executable not found at $ccmExecPath. Cannot use ccmexec.exe method."
                     }
                 }

                 if ($overallSuccess) {
                     Write-Host "JOB: Configuration Manager actions attempt finished. At least one method appears to have triggered actions successfully." -ForegroundColor Green
                 } else {
                     Write-Warning "JOB: Configuration Manager actions attempt finished, but neither CIM nor ccmexec.exe methods could be confirmed as fully successful or available."
                 }
                 Write-Host "JOB: Exiting Start-ConfigManagerActions function."
                 return $overallSuccess
            }
            # --- End Function Definitions INSIDE Job Scope ---

            # Execute the functions defined above within this job's scope
            Set-GPupdate
            Start-ConfigManagerActions
            Write-Host "JOB: Background updates finished."
        }
        Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown after main window closes. Main window will load now." -ForegroundColor Yellow
    }

    Write-Host "Exiting Show-ModeDialog function."
    return $selectedMode # Return the string 'Backup', 'Restore', or 'Express'
}

# Show main window (Backup/Restore specific)
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Backup', 'Restore')] # Only accepts Backup or Restore
        [string]$Mode
    )
    $IsBackup = ($Mode -eq 'Backup')
    $modeString = $Mode
    Write-Host "Entering Show-MainWindow function. Mode: $modeString"

    # XAML UI Definition (remains largely the same)
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
        <Button Name="btnStart" Content="Start" Width="100" Height="30" Margin="10,0,0,20" VerticalAlignment="Bottom" HorizontalAlignment="Left" IsDefault="True"/>

    </Grid>
</Window>
'@

    try {
        Write-Host "Parsing XAML for main window."
        $reader = New-Object System.Xml.XmlNodeReader $XAML
        $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "XAML loaded successfully."

        Write-Host "Setting Window DataContext."
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }

        Write-Host "Finding controls in main window."
        $controls = @{}
        @(
            'txtSaveLoc', 'btnBrowse', 'btnStart', 'lblMode', 'lblStatus',
            'lvwFiles', 'btnAddFile', 'btnAddFolder', 'btnRemove',
            'chkNetwork', 'chkPrinters', 'prgProgress', 'txtProgress'
        ) | ForEach-Object { $controls[$_] = $window.FindName($_) }
        Write-Host "Controls found and stored in hashtable."

        # --- Window Initialization ---
        Write-Host "Initializing window controls based on mode."
        $controls.lblMode.Content = "Mode: $modeString"
        $controls.btnStart.Content = $modeString # "Backup" or "Restore"

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

        # --- Load initial items based on mode ---
        Write-Host "Loading initial items for ListView based on mode."
        $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()

        if ($IsBackup) {
            Write-Host "Backup Mode: Getting default paths using local user profile."
            $localUserProfile = $env:USERPROFILE
            # Use Resolve-UserPathTemplates with local profile
            $paths = Resolve-UserPathTemplates -UserProfilePath $localUserProfile -Templates $script:pathTemplates

            if ($paths -ne $null -and $paths.Count -gt 0) {
                 Write-Host "Found $($paths.Count) local items. Adding 'IsSelected' property..."
                 $paths | ForEach-Object {
                     $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                     $itemsList.Add($_)
                 }
                 Write-Host "Populated ListView with local items."
            } else {
                 Write-Warning "Could not resolve any default local paths for backup."
            }
        }
        elseif (Test-Path $defaultPath) { # Restore Mode
            Write-Host "Restore Mode: Checking for latest backup in $defaultPath."
            $latestBackup = Get-ChildItem -Path $defaultPath -Directory -Filter "Backup_*" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                Write-Host "Found latest backup: $($latestBackup.FullName)"
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                $logFilePath = Join-Path $latestBackup.FullName "FileList_Backup.csv"
                if (Test-Path $logFilePath) {
                    # Load backup contents into ListView, adding IsSelected property
                    $backupItems = Get-ChildItem -Path $latestBackup.FullName |
                        Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|Drives\.csv|Printers\.txt|TransferLog\.csv)$' } | # Exclude logs
                        ForEach-Object {
                            [PSCustomObject]@{
                                Name = $_.Name # In restore, Name should represent the SourceKey for selection
                                Type = if ($_.PSIsContainer) { "Folder" } else { "File" }
                                Path = $_.FullName # Path within the backup folder
                                IsSelected = $true # Default to selected for restore
                            }
                        }
                    $backupItems | ForEach-Object { $itemsList.Add($_) }
                    Write-Host "Populated ListView with $($itemsList.Count) items from latest backup."
                } else {
                    Write-Warning "Latest backup folder '$($latestBackup.FullName)' is missing log file '$logFilePath'."
                    $controls.lblStatus.Content = "Restore mode: Latest backup folder is invalid. Please browse."
                }
            } else {
                 Write-Host "No backups found in $defaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $defaultPath. Please browse."
            }
        } else { # Restore Mode, default path doesn't exist
             Write-Host "Restore Mode: Default path $defaultPath does not exist."
             $controls.lblStatus.Content = "Restore mode: Default path $defaultPath does not exist. Please browse."
        }

        # Assign items to ListView
        if ($controls['lvwFiles'] -ne $null) {
            $controls.lvwFiles.ItemsSource = $itemsList
            Write-Host "Assigned $($itemsList.Count) items to ListView ItemsSource."
        } else {
            Write-Error "ListView control ('lvwFiles') not found!"
            throw "ListView control ('lvwFiles') could not be found in the XAML."
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

                if (-not $IsBackup) { # Restore Mode - Reload list from selected backup
                    Write-Host "Restore Mode: Loading items from selected backup folder."
                    $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                    Write-Host "Checking for log file: $logFilePath"
                    $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new() # Clear list first
                    if (Test-Path -Path $logFilePath) {
                         Write-Host "Log file found. Populating ListView."
                         # In restore, list items represent the *backed up* items/folders (named by SourceKey)
                         # We need to derive this from the log or the folder structure within the backup
                         # Let's list the top-level items in the backup dir (excluding logs)
                         $backupItems = Get-ChildItem -Path $selectedPath |
                            Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } | # List folders or known settings files
                            Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } | # Exclude logs
                            ForEach-Object {
                                [PSCustomObject]@{
                                    Name = $_.Name # Name represents the SourceKey or setting type
                                    Type = if ($_.PSIsContainer) { "Folder" } else { "Setting" }
                                    Path = $_.FullName
                                    IsSelected = $true # Default to selected
                                }
                            }
                        $backupItems | ForEach-Object { $newItemsList.Add($_) }
                        $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                        Write-Host "ListView updated with $($newItemsList.Count) items/categories from selected backup."
                    } else {
                         Write-Warning "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         $controls.lblStatus.Content = "Selected folder is not a valid backup (missing FileList_Backup.csv)."
                         [System.Windows.MessageBox]::Show("The selected folder does not appear to be a valid backup. It's missing the 'FileList_Backup.csv' log file.", "Invalid Backup Folder", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    $controls.lvwFiles.ItemsSource = $newItemsList # Assign new list (or empty list if invalid)
                } else { # Backup Mode
                     $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Add File handler (Only makes sense in Backup mode really)
        $controls.btnAddFile.Add_Click({
            if (!$IsBackup) {
                [System.Windows.MessageBox]::Show("Adding individual files is only supported in Backup mode.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                return
            }
            Write-Host "Add File button clicked."
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = "Select File(s) to Add for Backup"
            $dialog.Multiselect = $true

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($dialog.FileNames.Count) file(s) selected."
                $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                if ($currentItems -eq $null) { # Initialize if null
                    $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                }

                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    # Check if path already exists in the list
                    if (-not ($currentItems.Path -contains $file)) {
                        Write-Host "Adding file: $file"
                        $fileKey = "ManualFile_" + ([System.IO.Path]::GetFileNameWithoutExtension($file) -replace '[^a-zA-Z0-9_]','_')
                        $currentItems.Add([PSCustomObject]@{
                            Name = $fileKey # Use a generated key for Name/SourceKey
                            Type = "File"
                            Path = $file
                            IsSelected = $true
                            SourceKey = $fileKey
                        })
                        $addedCount++
                    } else { Write-Host "Skipping duplicate file: $file"}
                }
                # No need to reassign if we modified the list in place
                $controls.lvwFiles.Items.Refresh() # Refresh the view
                Write-Host "Updated ListView ItemsSource. Added $addedCount new file(s)."
            } else { Write-Host "File selection cancelled."}
        })

        # Add Folder handler (Only makes sense in Backup mode)
        $controls.btnAddFolder.Add_Click({
            if (!$IsBackup) {
                [System.Windows.MessageBox]::Show("Adding individual folders is only supported in Backup mode.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                return
            }
            Write-Host "Add Folder button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            $dialog.Description = "Select Folder to Add for Backup"
            $dialog.ShowNewFolderButton = $false

            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath
                 Write-Host "Folder selected to add: $selectedPath"
                 $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                 if ($currentItems -eq $null) { # Initialize if null
                    $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new()
                 }

                 if (-not ($currentItems.Path -contains $selectedPath)) {
                    Write-Host "Adding folder: $selectedPath"
                    $folderKey = "ManualFolder_" + ([System.IO.Path]::GetFileName($selectedPath) -replace '[^a-zA-Z0-9_]','_')
                    $currentItems.Add([PSCustomObject]@{
                        Name = $folderKey # Use a generated key for Name/SourceKey
                        Type = "Folder"
                        Path = $selectedPath
                        IsSelected = $true
                        SourceKey = $folderKey
                    })
                    $controls.lvwFiles.Items.Refresh() # Refresh the view
                    Write-Host "Updated ListView ItemsSource with new folder."
                 } else {
                    Write-Host "Skipping duplicate folder: $selectedPath"
                 }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Remove Selected handler
        $controls.btnRemove.Add_Click({
            Write-Host "Remove Selected button clicked."
            $selectedObjects = @($controls.lvwFiles.SelectedItems) # Ensure it's an array
            if ($selectedObjects.Count -gt 0) {
                Write-Host "Removing $($selectedObjects.Count) selected item(s)."
                $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                if ($currentItems -ne $null) {
                    $itemsToRemove = $selectedObjects | ForEach-Object { $_ } # Get actual objects
                    $itemsToRemove | ForEach-Object { $currentItems.Remove($_) } | Out-Null
                    $controls.lvwFiles.Items.Refresh() # Refresh the view
                    Write-Host "ListView ItemsSource updated after removal."
                }
            } else { Write-Host "No items selected to remove."}
        })


        # --- Start Button Logic (Backup/Restore) ---
        $controls.btnStart.Add_Click({
            $modeString = if ($IsBackup) { 'Backup' } else { 'Restore' }
            Write-Host "Start button clicked. Mode: $modeString"

            $location = $controls.txtSaveLoc.Text
            Write-Host "Selected location: $location"
            if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) {
                Write-Warning "Invalid location selected."
                [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            Write-Host "Disabling UI controls and setting wait cursor."
            $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }
            $window.Cursor = [System.Windows.Input.Cursors]::Wait

            try {
                $controls.lblStatus.Content = "Starting $modeString..."
                $controls.txtProgress.Text = "Initializing..."
                $controls.prgProgress.Value = 0
                $controls.prgProgress.IsIndeterminate = $true

                $operationPath = $location

                # ==================
                # --- BACKUP Logic ---
                # ==================
                if ($IsBackup) {
                    Write-Host "--- Starting Backup Operation ---"
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    # Use local username for backup folder name consistency
                    $localUsername = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'
                    $backupRootPath = Join-Path $operationPath "Backup_${localUsername}_$timestamp"
                    Write-Host "Backup root path: $backupRootPath"

                    try {
                        Write-Host "Creating backup directory..."
                        New-Item -Path $backupRootPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        Write-Host "Backup directory created."
                    } catch {
                        throw "Failed to create backup directory: $backupRootPath. Error: $($_.Exception.Message)"
                    }

                    $csvLogPath = Join-Path $backupRootPath "FileList_Backup.csv"
                    Write-Host "Creating log file: $csvLogPath"
                    "OriginalFullPath,BackupRelativePath,SourceKey" | Set-Content -Path $csvLogPath -Encoding UTF8

                    $itemsToBackup = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected }
                    if (-not $itemsToBackup -or $itemsToBackup.Count -eq 0) {
                         throw "No items selected (checked) for backup."
                    }
                    Write-Host "Found $($itemsToBackup.Count) CHECKED items in ListView to process for backup."

                    Write-Host "Estimating total files for progress bar..."
                    $totalFilesEstimate = 0
                    $itemsToBackup | ForEach-Object {
                        if (Test-Path -LiteralPath $_.Path -ErrorAction SilentlyContinue) {
                            if ($_.Type -eq 'Folder') {
                                try { $totalFilesEstimate += (Get-ChildItem $_.Path -Recurse -File -Force -ErrorAction SilentlyContinue).Count } catch {}
                            } else { # File
                                $totalFilesEstimate++
                            }
                        }
                    }
                    if ($controls.chkNetwork.IsChecked) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked) { $totalFilesEstimate++ }
                    Write-Host "Estimated total files/items: $totalFilesEstimate"

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 }
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    Write-Host "Starting processing of CHECKED files/folders for backup..."
                    foreach ($item in $itemsToBackup) {
                        Write-Host "Processing item: $($item.Name) ($($item.Type)) - Path: $($item.Path)"
                        $controls.txtProgress.Text = "Processing: $($item.Name)"
                        $sourcePath = $item.Path
                        $sourceKey = $item.SourceKey # Get the original key

                        if (-not (Test-Path -LiteralPath $sourcePath)) {
                            Write-Warning "Source path not found, skipping: $sourcePath"
                            "`"$sourcePath`",`"SKIPPED_NOT_FOUND`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                            continue
                        }

                        # Use the SourceKey for the top-level folder name in the backup
                        $backupItemRootName = $sourceKey -replace '[^a-zA-Z0-9_]', '_' # Sanitize key for path

                        if ($item.Type -eq "Folder") {
                            Write-Host "  Item is a folder. Processing recursively..."
                            try {
                                # Get source folder info once
                                $sourceFolderInfo = Get-Item -LiteralPath $sourcePath

                                Get-ChildItem -Path $sourcePath -Recurse -File -Force -ErrorAction Stop | ForEach-Object {
                                    $originalFileFullPath = $_.FullName
                                    # Calculate path relative to the *source folder itself*
                                    $relativeFilePath = $originalFileFullPath.Substring($sourceFolderInfo.FullName.Length).TrimStart('\')
                                    # Backup path uses the sanitized SourceKey as the root, then the relative path
                                    $backupRelativePath = Join-Path $backupItemRootName $relativeFilePath

                                    $targetBackupPath = Join-Path $backupRootPath $backupRelativePath
                                    $targetBackupDir = Split-Path $targetBackupPath -Parent

                                    if (-not (Test-Path $targetBackupDir)) {
                                        New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                    }

                                    Write-Host "    Copying '$($_.Name)' to '$targetBackupPath'"
                                    Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                    Write-Host "    Logging: `"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`""
                                    "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                    $filesProcessed++
                                    if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                    $controls.txtProgress.Text = "Backed up: $($_.Name)"
                                }
                                Write-Host "  Finished processing folder: $($item.Name)"
                            } catch {
                                 Write-Warning "Error processing folder '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                                 "`"$sourcePath`",`"ERROR_FOLDER_COPY: $($_.Exception.Message -replace '"', '""')`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                            }
                        } else { # Single File
                             Write-Host "  Item is a file. Processing..."
                             try {
                                $originalFileFullPath = $sourcePath
                                # Place single files directly under a folder named after their SourceKey
                                $backupRelativePath = Join-Path $backupItemRootName (Split-Path $originalFileFullPath -Leaf)
                                $targetBackupPath = Join-Path $backupRootPath $backupRelativePath
                                $targetBackupDir = Split-Path $targetBackupPath -Parent

                                if (-not (Test-Path $targetBackupDir)) {
                                    New-Item -Path $targetBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                                }

                                Write-Host "    Copying '$($item.Name)' to '$targetBackupPath'"
                                Copy-Item -Path $originalFileFullPath -Destination $targetBackupPath -Force -ErrorAction Stop
                                Write-Host "    Logging: `"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`""
                                "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                                $filesProcessed++
                                if ($filesProcessed -le $controls.prgProgress.Maximum) { $controls.prgProgress.Value = $filesProcessed }
                                $controls.txtProgress.Text = "Backed up: $($item.Name)"
                             } catch {
                                 Write-Warning "Error processing file '$($item.Name)' ($sourcePath): $($_.Exception.Message)"
                                 "`"$sourcePath`",`"ERROR_FILE_COPY: $($_.Exception.Message -replace '"', '""')`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
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
                            Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -ErrorAction Stop | # Added Network=True filter
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
                } else { # Restore Mode
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
                    try {
                        $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8 -ErrorAction Stop
                    } catch {
                        throw "Failed to import backup log file '$csvLogPath'. Is it a valid CSV? Error: $($_.Exception.Message)"
                    }
                    if (-not $backupLog) {
                         throw "Backup log file is empty or could not be read: $csvLogPath"
                    }
                    Write-Host "Imported $($backupLog.Count) entries from log file."

                    # --- Get Selected Items from ListView ---
                    $listViewItems = @($controls.lvwFiles.ItemsSource)
                    $selectedItemsFromListView = $listViewItems | Where-Object { $_.IsSelected }

                    if (-not $selectedItemsFromListView) {
                        throw "No items selected (checked) in the list for restore."
                    }
                    # Get the NAMES of the selected items (which correspond to SourceKey in backup or setting file names)
                    $selectedKeysOrFiles = $selectedItemsFromListView | Select-Object -ExpandProperty Name
                    Write-Host "Found $($selectedItemsFromListView.Count) CHECKED items/categories in ListView for restore: $($selectedKeysOrFiles -join ', ')"

                    # --- Filter Log Entries Based on Selection ---
                    Write-Host "Filtering log entries based on ListView selection (matching SourceKey)..."
                    # Ensure the log has the SourceKey column
                    if (-not ($backupLog | Get-Member -Name SourceKey)) {
                        throw "Backup log '$csvLogPath' is missing the required 'SourceKey' column. Cannot perform selective restore of files/folders."
                    }
                    $logEntriesToRestore = $backupLog | Where-Object {
                        $_.SourceKey -in $selectedKeysOrFiles -and
                        $_.BackupRelativePath -notmatch '^(SKIPPED|ERROR)_' # Exclude skipped/error entries
                    }

                    if (-not $logEntriesToRestore) {
                        Write-Warning "No file/folder entries found in the log matching the selected items. Only settings might be restored if selected."
                    } else {
                        Write-Host "Filtered log. $($logEntriesToRestore.Count) file/folder log entries will be processed for restore."
                    }


                    # Estimate progress based on selected log entries AND selected settings
                    $totalFilesEstimate = $logEntriesToRestore.Count
                    if ($controls.chkNetwork.IsChecked -and ($selectedKeysOrFiles -contains "Drives.csv")) { $totalFilesEstimate++ }
                    if ($controls.chkPrinters.IsChecked -and ($selectedKeysOrFiles -contains "Printers.txt")) { $totalFilesEstimate++ }
                    Write-Host "Estimated total items for restore progress: $totalFilesEstimate"

                    $controls.prgProgress.Maximum = if($totalFilesEstimate -gt 0) { $totalFilesEstimate } else { 1 }
                    $controls.prgProgress.IsIndeterminate = $false
                    $controls.prgProgress.Value = 0
                    $filesProcessed = 0

                    # Restore Files/Folders from Filtered Log
                    if($logEntriesToRestore.Count -gt 0) {
                        Write-Host "Starting restore of files/folders from filtered log..."
                        foreach ($entry in $logEntriesToRestore) {
                            $originalFileFullPath = $entry.OriginalFullPath
                            $backupRelativePath = $entry.BackupRelativePath
                            $sourceBackupPath = Join-Path $backupRootPath $backupRelativePath

                            Write-Host "Processing restore entry: Source='$sourceBackupPath', Target='$originalFileFullPath'"
                            $controls.txtProgress.Text = "Restoring: $(Split-Path $originalFileFullPath -Leaf)"

                            if (Test-Path -LiteralPath $sourceBackupPath -PathType Leaf) { # Ensure source exists in backup
                                try {
                                    $targetRestoreDir = Split-Path $originalFileFullPath -Parent
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
                    } else {
                        Write-Host "Skipping file/folder restore as no matching log entries were found for selected items."
                    }


                    # Restore Network Drives (Only if checkbox checked AND "Drives.csv" was selected in list)
                    if ($controls.chkNetwork.IsChecked -and ($selectedKeysOrFiles -contains "Drives.csv")) {
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
                        } else { Write-Warning "Network drives backup file (Drives.csv) not found in backup, although selected." }
                    } else { Write-Host "Skipping Network Drives restore (unchecked or Drives.csv not selected)."}

                    # Restore Printers (Only if checkbox checked AND "Printers.txt" was selected in list)
                    if ($controls.chkPrinters.IsChecked -and ($selectedKeysOrFiles -contains "Printers.txt")) {
                        Write-Host "Processing Printers restore..."
                        $controls.txtProgress.Text = "Restoring printers..."
                        $printersTxtPath = Join-Path $backupRootPath "Printers.txt"
                        if (Test-Path $printersTxtPath) {
                             Write-Host "Found Printers.txt. Processing printers..."
                            try {
                                $wsNet = New-Object -ComObject WScript.Network # Use COM object for broader compatibility
                                Get-Content $printersTxtPath | ForEach-Object {
                                    $printerPath = $_.Trim() # Trim whitespace
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
                        } else { Write-Warning "Printers backup file (Printers.txt) not found in backup, although selected." }
                    } else { Write-Host "Skipping Printers restore (unchecked or Printers.txt not selected)."}

                    $controls.txtProgress.Text = "Restore completed from: $backupRootPath"
                    if ($controls.prgProgress.Maximum -gt 0) {
                        $controls.prgProgress.Value = $controls.prgProgress.Maximum
                    }
                    Write-Host "--- Restore Operation Finished ---"

                } # End if/else ($IsBackup)

                # --- Operation Completion ---
                Write-Host "Operation completed. Displaying success message." -ForegroundColor Green
                $controls.lblStatus.Content = "Operation completed successfully."
                [System.Windows.MessageBox]::Show("The $modeString operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

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

# --- Express Mode Function ---
function Execute-ExpressModeLogic {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Host "Executing Express Mode Logic..."
    $targetDevice = $null
    $mappedDriveLetter = "X" # Or choose another available letter
    $localTempTransferDir = Join-Path $env:TEMP "ExpressTransfer_$(Get-Date -Format 'yyyyMMddHHmmss')"
    $localBackupBaseDir = "C:\LocalData" # Where the final backup copy will reside
    $transferSuccess = $true # Assume success initially
    $transferLog = [System.Collections.Generic.List[string]]::new()
    $transferLog.Add("Timestamp,Action,Status,Details")
    $logFilePath = $null # Initialize log file path variable
    $tempLogPath = $null # Initialize temp log file path variable

    Function LogTransfer ($Action, $Status, $Details) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        # Escape double quotes in details for CSV compatibility
        $safeDetails = $Details -replace '"', '""'
        $logEntry = """$timestamp"",""$Action"",""$Status"",""$safeDetails"""
        $script:transferLog.Add($logEntry)
        Write-Host "LOG: $Action - $Status - $Details"
    }

    try {
        # --- 1. Get Target Device ---
        $targetDevice = Read-Host "Enter the IP address or Hostname of the target remote device"
        if ([string]::IsNullOrWhiteSpace($targetDevice)) {
            throw "Target device cannot be empty."
        }
        LogTransfer "Input" "Info" "Target device specified: $targetDevice"

        # --- 2. Map Network Drive ---
        $uncPath = "\\$targetDevice\c$"
        Write-Host "Attempting to map $uncPath to $mappedDriveLetter`: " -NoNewline
        LogTransfer "Map Drive" "Attempt" "Mapping $uncPath to $mappedDriveLetter`:"

        # Remove existing drive if it exists
        if (Test-Path "${mappedDriveLetter}:") {
            Write-Warning "Drive $mappedDriveLetter`: already exists. Attempting to remove..."
            Remove-PSDrive -Name $mappedDriveLetter -Force -ErrorAction SilentlyContinue
        }

        New-PSDrive -Name $mappedDriveLetter -PSProvider FileSystem -Root $uncPath -Credential $Credential -ErrorAction Stop | Out-Null
        Write-Host "Successfully mapped $uncPath to $mappedDriveLetter`:" -ForegroundColor Green
        LogTransfer "Map Drive" "Success" "Successfully mapped $uncPath to $mappedDriveLetter`:"

        # --- 3. Get Remote Logged-on Username ---
        Write-Host "Attempting to identify logged-on user on '$targetDevice'..."
        LogTransfer "Get Remote User" "Attempt" "Querying Win32_ComputerSystem on $targetDevice"
        $remoteUsername = $null
        try {
            # Use Invoke-Command for reliability, especially across different network/domain setups
            $remoteResult = Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
            } -ErrorAction Stop

            if ($remoteResult -and -not [string]::IsNullOrWhiteSpace($remoteResult)) {
                # Username might be in DOMAIN\User format, we only want the user part
                $remoteUsername = ($remoteResult -split '\\')[-1]
                Write-Host "Identified remote user: '$remoteUsername'" -ForegroundColor Green
                LogTransfer "Get Remote User" "Success" "Identified remote user: $remoteUsername"
            } else {
                throw "Could not retrieve username from Win32_ComputerSystem or it was empty."
            }
        } catch {
            Write-Warning "Failed to automatically identify remote user via Win32_ComputerSystem: $($_.Exception.Message)"
            LogTransfer "Get Remote User" "Warning" "Failed to get user via WMI/Invoke-Command: $($_.Exception.Message)"
            # Fallback: Prompt the user
            $remoteUsername = Read-Host "Could not auto-detect remote user. Please enter the Windows username logged into '$targetDevice'"
            if ([string]::IsNullOrWhiteSpace($remoteUsername)) {
                throw "Remote username is required to proceed."
            }
            LogTransfer "Get Remote User" "Manual Input" "User provided remote username: $remoteUsername"
        }

        # --- 4. Define Remote Paths and Local Destinations ---
        $remoteUserProfile = "$mappedDriveLetter`:\Users\$remoteUsername"
        Write-Host "Remote user profile path set to: $remoteUserProfile"
        LogTransfer "Path Setup" "Info" "Remote user profile path: $remoteUserProfile"

        if (-not (Test-Path -LiteralPath $remoteUserProfile)) {
            throw "Remote user profile path '$remoteUserProfile' not found on mapped drive. Check username and permissions."
        }

        # Resolve paths on the *remote* system using the mapped drive
        Write-Host "Identifying remote files/folders based on templates for user '$remoteUsername'..."
        LogTransfer "Gather Paths" "Attempt" "Resolving templates for $remoteUserProfile"
        $remotePathsToTransfer = Resolve-UserPathTemplates -UserProfilePath $remoteUserProfile -Templates $script:pathTemplates
        LogTransfer "Gather Paths" "Info" "Found $($remotePathsToTransfer.Count) items based on templates."

        if ($remotePathsToTransfer.Count -eq 0) {
            Write-Warning "No files or folders found to transfer based on defined paths for user '$remoteUsername' on '$targetDevice'."
            LogTransfer "Gather Paths" "Warning" "No template paths resolved successfully for $remoteUsername."
            # Decide if this is critical - maybe continue for settings? For now, we'll continue.
        }

        # Create Local Temp Directory
        Write-Host "Creating local temporary transfer directory: $localTempTransferDir"
        LogTransfer "Setup" "Info" "Creating local temp dir: $localTempTransferDir"
        New-Item -Path $localTempTransferDir -ItemType Directory -Force -ErrorAction Stop | Out-Null

        # --- 5. Transfer Files/Folders ---
        Write-Host "Starting file/folder transfer from '$targetDevice' (via $mappedDriveLetter`:) to '$localTempTransferDir'..."
        $totalItems = $remotePathsToTransfer.Count
        $currentItem = 0
        $errorsDuringTransfer = $false

        if ($totalItems -gt 0) {
            foreach ($item in $remotePathsToTransfer) {
                $currentItem++
                $progress = [int](($currentItem / $totalItems) * 50) + 10 # Progress from 10% to 60%
                Write-Progress -Activity "Transferring Files from $targetDevice" -Status "[$progress%]: Copying $($item.Name)" -PercentComplete $progress

                $sourcePath = $item.Path # This is the path on the mapped drive (e.g., X:\Users\...)
                # Destination path uses the SourceKey to maintain structure in the temp dir
                $destRelativePath = $item.SourceKey -replace '[^a-zA-Z0-9_]', '_' # Sanitize key
                if ($item.Type -eq 'File') {
                     $destRelativePath = Join-Path $destRelativePath (Split-Path $sourcePath -Leaf)
                }
                $destinationPath = Join-Path $localTempTransferDir $destRelativePath

                Write-Host "  Copying '$($item.Name)' from '$sourcePath' to '$destinationPath'"
                LogTransfer "File Transfer" "Attempt" "Copying $($item.SourceKey) from $sourcePath"
                try {
                    # Ensure destination directory exists
                    $destDir = Split-Path $destinationPath -Parent
                    if (-not (Test-Path $destDir)) {
                        New-Item -Path $destDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                    }

                    Copy-Item -Path $sourcePath -Destination $destinationPath -Recurse:($item.Type -eq 'Folder') -Force -ErrorAction Stop
                    LogTransfer "File Transfer" "Success" "Copied $($item.SourceKey) to $destinationPath"
                } catch {
                    Write-Warning "  Failed to copy '$($item.Name)' from '$sourcePath': $($_.Exception.Message)"
                    LogTransfer "File Transfer" "Error" "Failed to copy $($item.SourceKey): $($_.Exception.Message)"
                    $errorsDuringTransfer = $true
                    # Continue with next item
                }
            }
            Write-Progress -Activity "Transferring Files from $targetDevice" -Completed

            if ($errorsDuringTransfer) {
                Write-Warning "One or more errors occurred during file/folder transfer. Check log."
            } else {
                Write-Host "File/folder transfer completed." -ForegroundColor Green
            }
        } else {
             Write-Host "Skipping file/folder transfer as no items were resolved."
             Write-Progress -Activity "Transferring Files from $targetDevice" -Status "[60%]: No files to transfer" -PercentComplete 60
        }


        # --- 6. Transfer Settings (Network Drives / Printers) ---
        Write-Progress -Activity "Transferring Settings" -Status "[60%]: Capturing remote network drives..." -PercentComplete 60
        Write-Host "Attempting to get mapped drives from remote computer: $targetDevice"
        LogTransfer "Get Drives" "Attempt" "Querying remote mapped drives on $targetDevice"
        $remoteDrivesCsvPath = Join-Path $localTempTransferDir "Drives.csv"
        try {
            # Use Invoke-Command with the credential to query WMI on the remote machine
            Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop |
                    Select-Object Name, ProviderName
            } -ErrorAction Stop | Export-Csv -Path $remoteDrivesCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            Write-Host "Successfully retrieved and saved remote mapped drives." -ForegroundColor Green
            LogTransfer "Get Drives" "Success" "Saved remote drives to $remoteDrivesCsvPath"
        } catch {
            # FIX: Use ${targetDevice} in the string
            Write-Warning "Failed to get mapped drives from ${targetDevice}: $($_.Exception.Message)"
            LogTransfer "Get Drives" "Error" "Failed: $($_.Exception.Message)"
            # Continue transfer
        }

        Write-Progress -Activity "Transferring Settings" -Status "[75%]: Capturing remote network printers..." -PercentComplete 75
        Write-Host "Attempting to get network printers from remote computer: $targetDevice"
        LogTransfer "Get Printers" "Attempt" "Querying remote network printers on $targetDevice"
        $remotePrintersTxtPath = Join-Path $localTempTransferDir "Printers.txt"
        try {
            # Use Invoke-Command with the credential
            Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -ErrorAction Stop | # Ensure Network=True
                    Select-Object -ExpandProperty Name
            } -ErrorAction Stop | Set-Content -Path $remotePrintersTxtPath -Encoding UTF8 -ErrorAction Stop
            Write-Host "Successfully retrieved and saved remote network printers." -ForegroundColor Green
            LogTransfer "Get Printers" "Success" "Saved remote printers to $remotePrintersTxtPath"
        } catch {
            # FIX: Use ${targetDevice} in the string
            Write-Warning "Failed to get network printers from ${targetDevice}: $($_.Exception.Message)"
            LogTransfer "Get Printers" "Error" "Failed: $($_.Exception.Message)"
            # Continue transfer
        }
        Write-Progress -Activity "Transferring Settings" -Completed

        # --- 7. Post-Transfer Remote Actions (GPUpdate, ConfigMgr) ---
        Write-Host "Executing post-transfer actions (GPUpdate, ConfigMgr Cycles) remotely on '$targetDevice'..."
        LogTransfer "Remote Actions" "Attempt" "Running GPUpdate and ConfigMgr actions on $targetDevice"
        Write-Progress -Activity "Remote Actions" -Status "[85%]: Running GPUpdate/ConfigMgr on $targetDevice..." -PercentComplete 85
        try {
            Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                # --- Function Definitions INSIDE Invoke-Command Scope ---
                # These run ON THE REMOTE machine using the provided credentials

                function Set-GPupdate {
                    Write-Host "REMOTE JOB: Initiating Group Policy update..." -ForegroundColor Cyan
                    try {
                        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
                        if ($process.ExitCode -eq 0) {
                            Write-Host "REMOTE JOB: Group Policy update completed successfully." -ForegroundColor Green
                        } else {
                            Write-Warning "REMOTE JOB: Group Policy update process finished with exit code: $($process.ExitCode)."
                        }
                    } catch {
                        Write-Error "REMOTE JOB: Failed to start GPUpdate process: $($_.Exception.Message)"
                    }
                    Write-Host "REMOTE JOB: Exiting Set-GPupdate function."
                }

                function Start-ConfigManagerActions {
                    param()
                    Write-Host "REMOTE JOB: Entering Start-ConfigManagerActions function."
                    Write-Host "REMOTE JOB: Attempting to trigger Configuration Manager client actions..."
                    $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"
                    $clientSDKNamespace = "root\ccm\clientsdk"
                    $clientClassName = "CCM_ClientUtilities"
                    $scheduleMethodName = "TriggerSchedule"
                    $overallSuccess = $false
                    $cimAttemptedAndSucceeded = $false

                    $scheduleActions = @(
                        @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'},
                        @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'}
                    )
                    Write-Host "REMOTE JOB: Defined $($scheduleActions.Count) CM actions to trigger."

                    Write-Host "REMOTE JOB: Checking for Configuration Manager client service (CcmExec)..."
                    $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
                    if (-not $ccmService) {
                        Write-Warning "REMOTE JOB: Configuration Manager client service (CcmExec) not found. Skipping CM actions."
                        return $false
                    } elseif ($ccmService.Status -ne 'Running') {
                        Write-Warning "REMOTE JOB: Configuration Manager client service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping CM actions."
                        return $false
                    } else {
                        Write-Host "REMOTE JOB: Configuration Manager client service (CcmExec) found and running."
                    }

                    Write-Host "REMOTE JOB: Attempting Method 1: Triggering actions via CIM ($clientSDKNamespace -> $clientClassName)..."
                    $cimMethodSuccess = $true
                    try {
                        if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                            Write-Host "REMOTE JOB: CIM Class found. Proceeding to trigger schedules via CIM."
                            foreach ($action in $scheduleActions) {
                                Write-Host "REMOTE JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                                try {
                                    Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop
                                    Write-Host "REMOTE JOB:     $($action.Name) triggered successfully via CIM."
                                } catch {
                                    Write-Warning "REMOTE JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"
                                    $cimMethodSuccess = $false
                                }
                            }
                            if ($cimMethodSuccess) {
                                $cimAttemptedAndSucceeded = $true
                                $overallSuccess = $true
                                Write-Host "REMOTE JOB: All actions successfully triggered via CIM." -ForegroundColor Green
                            } else {
                                Write-Warning "REMOTE JOB: One or more actions failed to trigger via CIM."
                            }
                        } else {
                            Write-Warning "REMOTE JOB: CIM Class '$clientClassName' not found in namespace '$clientSDKNamespace'. Cannot use CIM method."
                            $cimMethodSuccess = $false
                        }
                    } catch {
                        Write-Error "REMOTE JOB: An unexpected error occurred during CIM attempt: $($_.Exception.Message)"
                        $cimMethodSuccess = $false
                    }

                    if (-not $cimAttemptedAndSucceeded) {
                        Write-Host "REMOTE JOB: CIM method did not complete successfully or was not available. Attempting Method 2: Fallback via ccmexec.exe..."
                        Write-Host "REMOTE JOB: Checking for executable: $ccmExecPath"
                        if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                            Write-Host "REMOTE JOB: Found $ccmExecPath. Proceeding to trigger schedules via executable."
                            $execMethodSuccess = $true
                            foreach ($action in $scheduleActions) {
                                Write-Host "REMOTE JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                                try {
                                    $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop
                                    if ($process.ExitCode -ne 0) {
                                        Write-Warning "REMOTE JOB:     $($action.Name) action via ccmexec.exe finished with exit code $($process.ExitCode). (This might still be okay)"
                                    } else {
                                        Write-Host "REMOTE JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)."
                                    }
                                } catch {
                                    Write-Warning "REMOTE JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"
                                    $execMethodSuccess = $false
                                }
                            }
                            if ($execMethodSuccess) {
                                $overallSuccess = $true
                                Write-Host "REMOTE JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green
                            } else {
                                Write-Warning "REMOTE JOB: One or more actions failed to execute via ccmexec.exe."
                            }
                        } else {
                            Write-Warning "REMOTE JOB: Fallback executable not found at $ccmExecPath. Cannot use ccmexec.exe method."
                        }
                    }

                    if ($overallSuccess) {
                        Write-Host "REMOTE JOB: Configuration Manager actions attempt finished. At least one method appears to have triggered actions successfully." -ForegroundColor Green
                    } else {
                        Write-Warning "REMOTE JOB: Configuration Manager actions attempt finished, but neither CIM nor ccmexec.exe methods could be confirmed as fully successful or available."
                    }
                    Write-Host "REMOTE JOB: Exiting Start-ConfigManagerActions function."
                    return $overallSuccess
                }
                # --- End Function Definitions INSIDE Invoke-Command Scope ---

                # Execute the functions defined above within this remote scope
                Set-GPupdate
                Start-ConfigManagerActions
                Write-Host "REMOTE JOB: Background updates finished."

            } -ErrorAction Stop # End Invoke-Command ScriptBlock
            Write-Host "Successfully executed remote actions on $targetDevice." -ForegroundColor Green
            LogTransfer "Remote Actions" "Success" "Executed GPUpdate and ConfigMgr actions remotely."
        } catch {
            # FIX: Use ${targetDevice} in the string
            Write-Warning "Failed to execute post-transfer actions remotely on ${targetDevice}: $($_.Exception.Message)"
            LogTransfer "Remote Actions" "Error" "Failed: $($_.Exception.Message)"
            # Decide if this constitutes overall failure
            # $transferSuccess = $false
        }
        Write-Progress -Activity "Remote Actions" -Completed

        # --- 8. Create Final Local Backup Copy ---
        Write-Progress -Activity "Creating Local Backup" -Status "[90%]: Copying transferred files to backup folder..." -PercentComplete 90
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        # Use the *remote* username in the backup folder name for clarity
        $finalBackupDir = Join-Path $localBackupBaseDir "TransferBackup_${remoteUsername}_from_($targetDevice -replace '[^a-zA-Z0-9_.-]','_')_$timestamp" # Sanitize target device name for path

        Write-Host "Creating local backup copy of successfully transferred items in '$finalBackupDir'..."
        LogTransfer "Local Backup" "Attempt" "Copying from $localTempTransferDir to $finalBackupDir"

        if ((Get-ChildItem -Path $localTempTransferDir).Count -gt 0) {
            try {
                Copy-Item -Path $localTempTransferDir -Destination $finalBackupDir -Recurse -Force -ErrorAction Stop
                Write-Host "Local backup copy created successfully." -ForegroundColor Green
                LogTransfer "Local Backup" "Success" "Created backup at $finalBackupDir"
            } catch {
                Write-Warning "Failed to create final local backup copy at '$finalBackupDir': $($_.Exception.Message)"
                LogTransfer "Local Backup" "Error" "Failed: $($_.Exception.Message)"
                $transferSuccess = $false # Failure to create final backup is likely critical
            }
        } else {
            Write-Warning "No items were successfully transferred to '$localTempTransferDir' to create a local backup copy."
            LogTransfer "Local Backup" "Warning" "Skipped - No items in temp directory $localTempTransferDir"
            # If nothing was transferred, maybe mark as unsuccessful? Depends on requirements.
        }
        Write-Progress -Activity "Creating Local Backup" -Completed

        # --- 9. Final Log Saving ---
        # Ensure final backup dir exists before trying to save log there
        if (-not ([string]::IsNullOrEmpty($finalBackupDir)) -and -not (Test-Path $finalBackupDir)) {
             New-Item -Path $finalBackupDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        }

        # Set log file path only if final dir seems valid
        if (-not ([string]::IsNullOrEmpty($finalBackupDir)) -and (Test-Path $finalBackupDir)) {
            $logFilePath = Join-Path $finalBackupDir "TransferLog.csv"
        }

        # Try saving log
        try {
            if (-not ([string]::IsNullOrEmpty($logFilePath))) {
                $script:transferLog | Set-Content -Path $logFilePath -Encoding UTF8 -Force
                Write-Host "Transfer log saved to: $logFilePath"
            } else {
                # Fallback to temp if final dir wasn't created/valid
                $tempLogPath = Join-Path $env:TEMP "ExpressTransfer_Log_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
                $script:transferLog | Set-Content -Path $tempLogPath -Encoding UTF8 -Force
                Write-Warning "Could not determine final backup directory. Saved log to temporary location instead: $tempLogPath"
                LogTransfer "Save Log" "Warning" "Saved log to temp: $tempLogPath"
            }
        } catch {
            Write-Error "CRITICAL: Failed to save log. Error: $($_.Exception.Message)"
            LogTransfer "Save Log" "Fatal Error" "Failed to save log: $($_.Exception.Message)"
            # Log might be lost here
        }


        Write-Progress -Activity "Express Transfer Complete" -Status "[100%]: Finished." -PercentComplete 100
        Write-Host "--- REMOTE Transfer Operation Finished ---"

    } catch {
        $errorMessage = "Express Transfer Failed: $($_.Exception.Message)"
        Write-Error $errorMessage
        LogTransfer "Overall Status" "Fatal Error" $errorMessage
        $transferSuccess = $false # Mark as failed on any major exception
    } finally {
        # --- 10. Cleanup ---
        # Unmap drive
        if (Test-Path "${mappedDriveLetter}:") {
            Write-Host "Attempting to remove mapped drive $mappedDriveLetter`:" -NoNewline
            try {
                Remove-PSDrive -Name $mappedDriveLetter -Force -ErrorAction Stop
                Write-Host " Successfully removed mapped drive." -ForegroundColor Green
                LogTransfer "Cleanup" "Success" "Removed mapped drive $mappedDriveLetter`:"
            } catch {
                Write-Warning " Failed to remove mapped drive $mappedDriveLetter`: $($_.Exception.Message)"
                LogTransfer "Cleanup" "Error" "Failed to remove mapped drive $mappedDriveLetter`: $($_.Exception.Message)"
            }
        }

        # Remove local temp directory
        if (Test-Path $localTempTransferDir) {
            Write-Host "Removing local temporary directory: $localTempTransferDir"
            Remove-Item -Path $localTempTransferDir -Recurse -Force -ErrorAction SilentlyContinue
            LogTransfer "Cleanup" "Info" "Removed temp directory $localTempTransferDir"
        }

        # Determine final log path for message (PowerShell v3 compatible)
        $finalLogPathMessage = "Log saving failed"
        if (-not ([string]::IsNullOrEmpty($logFilePath)) -and (Test-Path $logFilePath)) {
            $finalLogPathMessage = $logFilePath
        } elseif (-not ([string]::IsNullOrEmpty($tempLogPath)) -and (Test-Path $tempLogPath)) {
            $finalLogPathMessage = $tempLogPath
        }


        # Final status message
        if ($transferSuccess) {
            Write-Host "Express Transfer process completed." -ForegroundColor Green
            [System.Windows.MessageBox]::Show("Express Transfer completed. Check console output and log for details: `n$finalLogPathMessage", "Express Transfer Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            Write-Error "Express Transfer failed or encountered significant errors. Check console output and log: $finalLogPathMessage"
             [System.Windows.MessageBox]::Show("Express Transfer failed or encountered significant errors. Check console output and log file for details: `n$finalLogPathMessage", "Express Transfer Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
             # Throw an error to potentially stop script execution if called from main flow
             throw "Express Transfer failed or encountered significant errors. Check console output and log: $finalLogPathMessage"
        }
    }
}

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue # Clear previous job variable if it exists
$script:transferLog = $null # Initialize transfer log variable
$operationMode = 'Cancel' # Default to cancel unless changed

try {
    # Determine mode ('Backup', 'Restore', 'Express', or 'Cancel')
    Write-Host "Calling Show-ModeDialog to determine operation mode."
    $operationMode = Show-ModeDialog # Returns the selected mode as a string

    switch ($operationMode) {
        'Backup' {
            Write-Host "Calling Show-MainWindow with Mode = Backup"
            Show-MainWindow -Mode 'Backup'
        }
        'Restore' {
            Write-Host "Calling Show-MainWindow with Mode = Restore"
            Show-MainWindow -Mode 'Restore'
        }
        'Express' {
            Write-Host "Mode selected: Express. Checking elevation..."
            # Check if running as Administrator
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

            if (-not $isAdmin) {
                Write-Warning "Express mode requires administrative privileges to map drives and run remote commands."
                Write-Host "Attempting to relaunch script with elevation..."
                try {
                    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -ErrorAction Stop
                } catch {
                    Write-Error "Failed to relaunch with elevation: $($_.Exception.Message). Please run the script as Administrator manually."
                    [System.Windows.MessageBox]::Show("Failed to relaunch with elevation. Please run the script as Administrator manually.", "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
                Exit # Exit the current non-elevated instance
            } else {
                Write-Host "Already running as Administrator. Proceeding with Express mode."
                # Get credentials for the remote connection
                $credential = $null
                try {
                     $credential = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message "Enter credentials with ADMIN rights on the REMOTE machine for Express Mode"
                } catch {
                    # Catch potential errors if Get-Credential fails (e.g., in non-interactive session)
                     throw "Failed to get credentials. Error: $($_.Exception.Message)"
                }

                if ($credential -eq $null) { throw "Credentials are required for Express mode and were not provided." }

                # Call the Express Mode function
                Execute-ExpressModeLogic -Credential $credential
            }
        }
        'Cancel' {
             Write-Host "Operation Cancelled during mode selection."
             # Script already exited in Show-ModeDialog for Cancel
        }
        Default {
            # Should not happen if Show-ModeDialog is correct
            Write-Error "Invalid operation mode returned: $operationMode"
        }
    }

} catch {
    # Catch errors during initial mode selection, window loading, or Express mode execution
    $errorMessage = "An unexpected error occurred: $($_.Exception.Message)"
    # Check if it's the specific error thrown by Express Mode failure
    if ($_.FullyQualifiedErrorId -match 'Express Transfer failed') {
        # Error already displayed by Execute-ExpressModeLogic, just log here
        Write-Host "Express mode failed. See previous messages." -ForegroundColor Red
    } else {
        # Display general errors
        Write-Error $errorMessage -ErrorAction Continue # Continue to finally block if possible
        Write-Host "FATAL ERROR: $errorMessage" -ForegroundColor Red
        try {
            [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } catch {
            Write-Warning "Could not display startup error message box."
        }
    }
} finally {
    # Check for and receive output from the background job (ONLY for Restore mode)
    if ($operationMode -eq 'Restore' -and (Get-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue) -ne $null -and $script:updateJob -ne $null) {
        Write-Host "`n--- Waiting for background update job (GPUpdate/CM Actions) to complete... ---" -ForegroundColor Yellow
        Wait-Job $script:updateJob | Out-Null
        Write-Host "--- Background Update Job Output (GPUpdate/CM Actions): ---" -ForegroundColor Yellow
        Receive-Job $script:updateJob
        Remove-Job $script:updateJob
        Write-Host "--- End of Background Update Job Output ---" -ForegroundColor Yellow
    } elseif ($operationMode -ne 'Express' -and $operationMode -ne 'Cancel') { # Don't show this for Express or Cancel
        Write-Host "`nNo background update job was started or it was already cleaned up." -ForegroundColor Gray
    }
}


Write-Host "--- Script Execution Finished ---"

# --- Keep console open when double-clicked ---
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') {
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
}