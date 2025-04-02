#requires -Version 3.0

# --- Load Assemblies FIRST ---
Write-Host "Attempting to load .NET Assemblies..."
try {
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop # Added for InputBox
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
$script:DefaultPath = "C:\LocalData" # Define default path globally

#endregion Global Variables & Configuration

#region Functions

# --- Helper Functions ---
function Format-Bytes {
    param([Parameter(Mandatory=$true)][long]$Bytes)
    $suffix = @("B", "KB", "MB", "GB", "TB", "PB"); $index = 0; $value = [double]$Bytes
    while ($value -ge 1024 -and $index -lt ($suffix.Length - 1)) { $value /= 1024; $index++ }
    return "{0:N2} {1}" -f $value, $suffix[$index]
}

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
        $resolvedPath = $template.Replace('{USERPROFILE}', $UserProfilePath.TrimEnd('\'))

        # Check if the resolved path exists
        if (Test-Path -LiteralPath $resolvedPath -ErrorAction SilentlyContinue) {
            $pathType = if (Test-Path -LiteralPath $resolvedPath -PathType Container) { "Folder" } else { "File" }
            Write-Host "  Found: '$key' -> $resolvedPath ($pathType)" -ForegroundColor Green
            $resolvedPaths.Add([PSCustomObject]@{
                Name = $key # Use the template key as the name
                Path = $resolvedPath # This is the SOURCE path (local or remote)
                Type = $pathType
                SourceKey = $key # Keep track of the original template key
            })
        } else {
            Write-Host "  Path not found or resolution failed: $resolvedPath (Derived from template key: $key)" -ForegroundColor Yellow
        }
    }
    return $resolvedPaths
}

# Gets standard user folders (Downloads, Desktop, etc.) - Used for LOCAL Backup Mode
function Get-UserPaths {
    [CmdletBinding()]
    param (
        # No remote parameters needed for local backup context
    )
    $folderNames = @( "Downloads", "Pictures", "Videos" ) # Add Desktop, Documents etc. if needed
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $sourceBasePath = $env:USERPROFILE
    Write-Host "Getting local user folders from: $sourceBasePath"

    foreach ($folderName in $folderNames) {
        $fullPath = Join-Path -Path $sourceBasePath -ChildPath $folderName
        if (Test-Path -LiteralPath $fullPath -PathType Container) {
            $item = Get-Item -LiteralPath $fullPath -ErrorAction SilentlyContinue
             if ($item) {
                 # Create a unique key based on the folder name
                 $folderKey = "UserFolder_" + ($item.Name -replace '[^a-zA-Z0-9_]','_')
                 $result.Add([PSCustomObject]@{
                     Name = $folderKey # Use a generated key for Name/SourceKey
                     Path = $item.FullName; # Source Path (Local)
                     Type = "Folder";
                     SourceKey = $folderKey
                 })
             } else { Write-Host "User folder resolved but Get-Item failed: $fullPath" }
        } else { Write-Host "User folder not found or not a folder: $fullPath" }
    }
    return $result
}

# --- GUI Functions ---

# Show mode selection dialog (Backup/Restore/Express Only)
function Show-ModeDialog {
    Write-Host "Entering Show-ModeDialog function."
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Operation"
    $form.Size = New-Object System.Drawing.Size(400, 150) # Adjusted size
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.HelpButton = $false # Prevent closing via Esc without selection

    # No description label needed if buttons are clear

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Location = New-Object System.Drawing.Point(30, 40) # Adjusted position
    $btnBackup.Size = New-Object System.Drawing.Size(80, 30)
    $btnBackup.Text = "Backup"
    $btnBackup.DialogResult = [System.Windows.Forms.DialogResult]::Yes # Using Yes for Backup
    $form.Controls.Add($btnBackup)

    $btnRestore = New-Object System.Windows.Forms.Button
    $btnRestore.Location = New-Object System.Drawing.Point(140, 40) # Adjusted position
    $btnRestore.Size = New-Object System.Drawing.Size(80, 30)
    $btnRestore.Text = "Restore"
    $btnRestore.DialogResult = [System.Windows.Forms.DialogResult]::No # Using No for Restore
    $form.Controls.Add($btnRestore)

    $btnExpress = New-Object System.Windows.Forms.Button
    $btnExpress.Location = New-Object System.Drawing.Point(250, 40) # Adjusted position
    $btnExpress.Size = New-Object System.Drawing.Size(80, 30)
    $btnExpress.Text = "Express"
    $btnExpress.DialogResult = [System.Windows.Forms.DialogResult]::OK # Using OK for Express
    $form.Controls.Add($btnExpress)

    # No Cancel button

    # Set default button (optional, e.g., Express)
    $form.AcceptButton = $btnExpress
    # Allow closing via Esc key by setting CancelButton (maps to Restore in this case)
    # Or leave it unset if Esc should do nothing / close dialog with Cancel result
    # $form.CancelButton = $btnRestore # Example: Esc = Restore

    Write-Host "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Host "Mode selection dialog closed with result: $result"

    # Determine mode based on DialogResult
    $selectedMode = switch ($result) {
        ([System.Windows.Forms.DialogResult]::Yes) { 'Backup' }
        ([System.Windows.Forms.DialogResult]::No) { 'Restore' }
        ([System.Windows.Forms.DialogResult]::OK) { 'Express' }
        Default { 'Cancel' } # Includes Cancel or closing the dialog via 'X'
    }

    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan

    # Handle Cancel explicitly (if user closes dialog with 'X')
    if ($selectedMode -eq 'Cancel') {
        Write-Host "Operation cancelled by user (dialog closed)." -ForegroundColor Yellow
        # Optional: Pause if running directly in console
        if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
        Exit 0 # Exit gracefully
    }

    # If restore mode, run updates immediately in a non-blocking way
    if ($selectedMode -eq 'Restore') {
        Write-Host "Restore mode selected. Initiating background system updates job..." -ForegroundColor Yellow
        $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {
            # Functions need to be defined within the job's scope
            # --- GPUpdate Function (for Job) ---
            function Set-GPupdate {
                Write-Host "JOB: Initiating Group Policy update..." -ForegroundColor Cyan
                try {
                    $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
                    if ($process.ExitCode -eq 0) { Write-Host "JOB: Group Policy update completed successfully." -ForegroundColor Green }
                    else { Write-Warning "JOB: Group Policy update process finished with exit code: $($process.ExitCode)." }
                } catch { Write-Error "JOB: Failed to start GPUpdate process: $($_.Exception.Message)" }
                Write-Host "JOB: Exiting Set-GPupdate function."
            }
            # --- ConfigMgr Actions Function (for Job) ---
            function Start-ConfigManagerActions {
                 param()
                 Write-Host "JOB: Entering Start-ConfigManagerActions function."
                 $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"; $clientSDKNamespace = "root\ccm\clientsdk"; $clientClassName = "CCM_ClientUtilities"; $scheduleMethodName = "TriggerSchedule"; $overallSuccess = $false; $cimAttemptedAndSucceeded = $false
                 # Shortened IDs for readability in code, full IDs used in Invoke-CimMethod/Start-Process
                 $scheduleActions = @(
                    @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'}
                    @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'}
                    @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'}
                    @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'}
                    @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}
                    @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'} # Duplicate? No, different from HInv Cycle ID
                    @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'}
                    @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'} # Duplicate? No, different from SInv Cycle ID
                 )
                 Write-Host "JOB: Defined $($scheduleActions.Count) CM actions to trigger."
                 $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
                 if (-not $ccmService) { Write-Warning "JOB: CM service (CcmExec) not found. Skipping."; return $false }
                 elseif ($ccmService.Status -ne 'Running') { Write-Warning "JOB: CM service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping."; return $false }
                 else { Write-Host "JOB: CM service (CcmExec) found and running." }
                 Write-Host "JOB: Attempting Method 1: Triggering via CIM ($clientSDKNamespace -> $clientClassName)..."
                 $cimMethodSuccess = $true
                 try {
                     if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                          Write-Host "JOB: CIM Class found."
                          foreach ($action in $scheduleActions) {
                             Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                             try { Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop; Write-Host "JOB:     $($action.Name) triggered successfully via CIM." }
                             catch { Write-Warning "JOB:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"; $cimMethodSuccess = $false }
                          }
                          if ($cimMethodSuccess) { $cimAttemptedAndSucceeded = $true; $overallSuccess = $true; Write-Host "JOB: All actions successfully triggered via CIM." -ForegroundColor Green }
                          else { Write-Warning "JOB: One or more actions failed to trigger via CIM." }
                     } else { Write-Warning "JOB: CIM Class '$clientClassName' not found. Cannot use CIM method."; $cimMethodSuccess = $false }
                 } catch { Write-Error "JOB: Unexpected error during CIM attempt: $($_.Exception.Message)"; $cimMethodSuccess = $false }
                 if (-not $cimAttemptedAndSucceeded) {
                     Write-Host "JOB: CIM failed/unavailable. Attempting Method 2: Fallback via ccmexec.exe..."
                     if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                         Write-Host "JOB: Found $ccmExecPath."
                         $execMethodSuccess = $true
                         foreach ($action in $scheduleActions) {
                             Write-Host "JOB:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                             try { $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop; if ($process.ExitCode -ne 0) { Write-Warning "JOB:     $($action.Name) via ccmexec.exe finished with exit code $($process.ExitCode)." } else { Write-Host "JOB:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)." } }
                             catch { Write-Warning "JOB:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"; $execMethodSuccess = $false }
                         }
                         if ($execMethodSuccess) { $overallSuccess = $true; Write-Host "JOB: Finished attempting actions via ccmexec.exe." -ForegroundColor Green }
                         else { Write-Warning "JOB: One or more actions failed to execute via ccmexec.exe." }
                     } else { Write-Warning "JOB: Fallback executable not found at $ccmExecPath." }
                 }
                 if ($overallSuccess) { Write-Host "JOB: CM actions attempt finished successfully." -ForegroundColor Green }
                 else { Write-Warning "JOB: CM actions attempt finished, but could not be confirmed as fully successful." }
                 Write-Host "JOB: Exiting Start-ConfigManagerActions function."; return $overallSuccess
            }
            # --- Job Execution ---
            Set-GPupdate; Start-ConfigManagerActions; Write-Host "JOB: Background updates finished."
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

    # XAML UI Definition (Added User Folders button and Space Labels)
    [xml]$XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Data Backup/Restore Tool"
    Width="800"
    Height="650"
    WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Location:" Margin="10,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="txtSaveLoc" Width="400" Height="30" Margin="10,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnBrowse" Content="Browse" Width="60" Height="30" Margin="420,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Name="lblMode" Content="" Margin="500,10,10,0" HorizontalAlignment="Right" VerticalAlignment="Top" FontWeight="Bold"/>
        <!-- Added Space Labels -->
        <Label Name="lblFreeSpace" Content="Free Space: -" Margin="500,40,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>
        <Label Name="lblRequiredSpace" Content="Required Space: -" Margin="500,65,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>

        <Label Name="lblStatus" Content="Ready" Margin="10,0,10,10" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontStyle="Italic"/>

        <Label Content="Files/Folders to Process:" Margin="10,90,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,120,200,140" SelectionMode="Extended">
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

        <StackPanel Margin="0,120,10,0" HorizontalAlignment="Right" Width="180">
            <Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/>
            <!-- Added User Folders Button -->
            <Button Name="btnAddBAUPaths" Content="Add User Folders" Width="120" Height="30" Margin="0,0,0,10"/>
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
            'chkNetwork', 'chkPrinters', 'prgProgress', 'txtProgress',
            'btnAddBAUPaths', 'lblFreeSpace', 'lblRequiredSpace' # Added new controls
        ) | ForEach-Object { $controls[$_] = $window.FindName($_) }
        Write-Host "Controls found and stored in hashtable."

        # --- Space Calculation Script Blocks (defined in function scope) ---
        $UpdateFreeSpaceLabel = {
            param($ControlsParam)
            $location = $ControlsParam.txtSaveLoc.Text
            $freeSpaceString = "Free Space: N/A"
            if (-not [string]::IsNullOrEmpty($location)) {
                try {
                    $driveLetter = $null
                    if ($location -match '^[a-zA-Z]:\\') { $driveLetter = $location.Substring(0, 2) }
                    elseif ($location -match '^\\\\[^\\]+\\[^\\]+') { $freeSpaceString = "Free Space: N/A (UNC)" } # Cannot reliably get free space for UNC

                    if ($driveLetter) {
                        $driveInfo = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -ErrorAction SilentlyContinue
                        if ($driveInfo -and $driveInfo.FreeSpace -ne $null) {
                            $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace)"
                        } else { Write-Warning "Could not get free space for drive $driveLetter" }
                    }
                } catch { Write-Warning "Error getting free space for '$location': $($_.Exception.Message)" }
            }
            # Update UI on Dispatcher thread
            $ControlsParam.lblFreeSpace.Dispatcher.InvokeAsync({ $ControlsParam.lblFreeSpace.Content = $freeSpaceString }) | Out-Null
            Write-Host "Updated Free Space Label: $freeSpaceString"
        }

        $UpdateRequiredSpaceLabel = {
            param($ControlsParam)
            $totalSize = 0L
            $requiredSpaceString = "Required Space: Calculating..."
            # Update UI immediately to show "Calculating..."
            $ControlsParam.lblRequiredSpace.Dispatcher.InvokeAsync({ $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString }) | Out-Null

            # Perform calculation (can take time)
            $items = @($ControlsParam.lvwFiles.ItemsSource)
            if ($items -ne $null -and $items.Count -gt 0) {
                $checkedItems = $items | Where-Object { $_.IsSelected }
                Write-Host "Calculating required space for $($checkedItems.Count) checked items..."
                foreach ($item in $checkedItems) {
                    # In Backup mode, path is local source. In Restore mode, path is within the backup folder.
                    $pathToCheck = $item.Path
                    if (Test-Path -LiteralPath $pathToCheck) {
                        try {
                            if ($item.Type -eq 'Folder') {
                                # Measure-Object can be slow on large folders, consider alternatives if performance is critical
                                $folderSize = (Get-ChildItem -LiteralPath $pathToCheck -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                if ($folderSize -ne $null) { $totalSize += $folderSize }
                            } elseif ($item.Type -eq 'File') { # Check if it's explicitly a file
                                $fileSize = (Get-Item -LiteralPath $pathToCheck -Force -ErrorAction SilentlyContinue).Length
                                if ($fileSize -ne $null) { $totalSize += $fileSize }
                            }
                            # Ignore 'Setting' type for size calculation (Drives.csv, Printers.txt)
                        } catch { Write-Warning "Error calculating size for '$($pathToCheck)': $($_.Exception.Message)" }
                    } else { Write-Warning "Checked item path not found, skipping size calculation: $($pathToCheck)" }
                }
                $requiredSpaceString = "Required Space: $(Format-Bytes $totalSize)"
            } else {
                $requiredSpaceString = "Required Space: 0 B"
            }
            # Update UI with final result on Dispatcher thread
            $ControlsParam.lblRequiredSpace.Dispatcher.InvokeAsync({ $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString }) | Out-Null
            Write-Host "Updated Required Space Label: $requiredSpaceString"
        }
        # --- End Space Calculation Script Blocks ---


        # --- Window Initialization ---
        Write-Host "Initializing window controls based on mode."
        $controls.lblMode.Content = "Mode: $modeString"
        $controls.btnStart.Content = $modeString # "Backup" or "Restore"
        $controls.btnAddBAUPaths.IsEnabled = $IsBackup # Only enable "Add User Folders" in Backup mode

        # Set default path
        Write-Host "Checking default path: $script:DefaultPath"
        if (-not (Test-Path $script:DefaultPath)) {
            Write-Host "Default path not found. Attempting to create."
            try {
                New-Item -Path $script:DefaultPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
                Write-Host "Default path created."
            } catch {
                 Write-Warning "Could not create default path: $script:DefaultPath. Please select a location manually."
                 $script:DefaultPath = $env:USERPROFILE # Fallback
                 Write-Host "Using fallback default path: $script:DefaultPath"
            }
        } else { Write-Host "Default path exists."}
        $controls.txtSaveLoc.Text = $script:DefaultPath
        # Update free space label initially
        & $UpdateFreeSpaceLabel -ControlsParam $controls


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
                     # Add IsSelected property for the CheckBox binding
                     $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                     # Add a property to track if it's editable/removable (optional, could be useful)
                     # $_ | Add-Member -MemberType NoteProperty -Name 'IsRemovable' -Value $false
                     $itemsList.Add($_)
                 }
                 Write-Host "Populated ListView with default local items."
            } else {
                 Write-Warning "Could not resolve any default local paths for backup."
            }
            # Optionally add User Folders by default here if desired
            # $userFolders = Get-UserPaths
            # $userFolders | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru; $itemsList.Add($_) }

        }
        elseif (Test-Path $script:DefaultPath) { # Restore Mode
            Write-Host "Restore Mode: Checking for latest backup in $script:DefaultPath."
            $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" | # Standard backup prefix
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if ($latestBackup) {
                Write-Host "Found latest backup: $($latestBackup.FullName)"
                $controls.txtSaveLoc.Text = $latestBackup.FullName
                & $UpdateFreeSpaceLabel -ControlsParam $controls # Update free space for backup location (destination drive)

                # Load items from the backup folder itself for Restore mode
                $backupItems = Get-ChildItem -Path $latestBackup.FullName |
                    Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } | # List folders or known settings files
                    Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } | # Exclude logs
                    ForEach-Object {
                        [PSCustomObject]@{
                            Name = $_.Name # Name represents the SourceKey or setting type
                            Type = if ($_.PSIsContainer) { "Folder" } elseif ($_.Name -match '^(Drives\.csv|Printers\.txt)$') { "Setting" } else { "File" } # Distinguish settings
                            Path = $_.FullName # Path within the backup folder (Source for Restore)
                            IsSelected = $true # Default to selected for restore
                            # SourceKey = $_.Name # Store the name as SourceKey for consistency if needed later
                        }
                    }
                $backupItems | ForEach-Object { $itemsList.Add($_) }
                Write-Host "Populated ListView with $($itemsList.Count) items/categories from latest backup."
                # Check for log file just for validation, not strictly needed to populate list anymore
                $logFilePath = Join-Path $latestBackup.FullName "FileList_Backup.csv"
                if (-not (Test-Path $logFilePath)) {
                    Write-Warning "Latest backup folder '$($latestBackup.FullName)' is missing log file '$logFilePath'. Restore might be incomplete."
                    $controls.lblStatus.Content = "Restore mode: Backup log missing. Restore might be incomplete."
                }

            } else {
                 Write-Host "No backups found in $script:DefaultPath."
                 $controls.lblStatus.Content = "Restore mode: No backups found in $script:DefaultPath. Please browse."
            }
        } else { # Restore Mode, default path doesn't exist
             Write-Host "Restore Mode: Default path $script:DefaultPath does not exist."
             $controls.lblStatus.Content = "Restore mode: Default path $script:DefaultPath does not exist. Please browse."
        }

        # Assign items to ListView
        if ($controls['lvwFiles'] -ne $null) {
            $controls.lvwFiles.ItemsSource = $itemsList
            Write-Host "Assigned $($itemsList.Count) items to ListView ItemsSource."
            # Update required space after loading items
            & $UpdateRequiredSpaceLabel -ControlsParam $controls
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
                 $dialog.SelectedPath = $script:DefaultPath # Fallback if current text isn't valid
            }
            $dialog.ShowNewFolderButton = $IsBackup # Only allow creating new folders in backup mode

            # Use a temporary invisible form as owner for FolderBrowserDialog
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }
            $owner.Show() # Show briefly to establish ownership context
            $owner.Hide() # Hide immediately
            Write-Host "Showing FolderBrowserDialog."
            $result = $dialog.ShowDialog($owner)
            $owner.Dispose() # Dispose the temporary owner form

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                $selectedPath = $dialog.SelectedPath
                Write-Host "Folder selected: $selectedPath"
                $controls.txtSaveLoc.Text = $selectedPath
                & $UpdateFreeSpaceLabel -ControlsParam $controls # Update free space for new location

                if (-not $IsBackup) { # Restore Mode - Reload list from selected backup
                    Write-Host "Restore Mode: Loading items from selected backup folder: $selectedPath"
                    $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new() # Clear list first
                    if (Test-Path $selectedPath -PathType Container) {
                         Write-Host "Populating ListView from selected backup folder."
                         $backupItems = Get-ChildItem -Path $selectedPath |
                            Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } |
                            Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } |
                            ForEach-Object {
                                [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } elseif ($_.Name -match '^(Drives\.csv|Printers\.txt)$') { "Setting" } else { "File" }; Path = $_.FullName; IsSelected = $true }
                            }
                        $backupItems | ForEach-Object { $newItemsList.Add($_) }
                        $controls.lblStatus.Content = "Ready to restore from: $selectedPath"
                        Write-Host "ListView updated with $($newItemsList.Count) items/categories from selected backup."
                        # Validate presence of log file
                        $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"
                        if (-not (Test-Path $logFilePath)) {
                             Write-Warning "Selected folder is missing 'FileList_Backup.csv'. Restore might be incomplete."
                             $controls.lblStatus.Content = "Selected folder is missing log file. Restore might be incomplete."
                             [System.Windows.MessageBox]::Show("The selected folder appears to be a backup, but it's missing the 'FileList_Backup.csv' log file. Restore may not work as expected.", "Missing Log File", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                        }
                    } else {
                         Write-Warning "Selected path '$selectedPath' is not a valid folder."
                         $controls.lblStatus.Content = "Selected path is not a valid folder."
                         [System.Windows.MessageBox]::Show("The selected path is not a valid folder.", "Invalid Path", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                    }
                    $controls.lvwFiles.ItemsSource = $newItemsList # Assign new list (or empty list if invalid)
                    & $UpdateRequiredSpaceLabel -ControlsParam $controls # Update required space
                } else { # Backup Mode
                     $controls.lblStatus.Content = "Backup location set to: $selectedPath"
                     # Required space doesn't change just by changing destination
                }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Add File handler
        $controls.btnAddFile.Add_Click({
            if (!$IsBackup) { [System.Windows.MessageBox]::Show("Adding individual files is only supported in Backup mode.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information); return }
            Write-Host "Add File button clicked."
            $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Title = "Select File(s) to Add for Backup"; $dialog.Multiselect = $true
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $owner.Show(); $owner.Hide(); $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Host "$($dialog.FileNames.Count) file(s) selected."
                $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }
                $addedCount = 0
                foreach ($file in $dialog.FileNames) {
                    if (-not ($currentItems.Path -contains $file)) {
                        # Create a unique key for manually added items
                        $fileKey = "ManualFile_" + ([System.IO.Path]::GetFileNameWithoutExtension($file) -replace '[^a-zA-Z0-9_]','_') + "_" + (Get-Date -Format 'HHmmss')
                        Write-Host "Adding file: $file (Key: $fileKey)"
                        $currentItems.Add([PSCustomObject]@{ Name = $fileKey; Type = "File"; Path = $file; IsSelected = $true; SourceKey = $fileKey }); $addedCount++
                    } else { Write-Host "Skipping duplicate file: $file"}
                }
                if ($addedCount -gt 0) { $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls } # Refresh view and update space
                Write-Host "Updated ListView ItemsSource. Added $addedCount new file(s)."
            } else { Write-Host "File selection cancelled."}
        })

        # Add Folder handler
        $controls.btnAddFolder.Add_Click({
            if (!$IsBackup) { [System.Windows.MessageBox]::Show("Adding individual folders is only supported in Backup mode.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information); return }
            Write-Host "Add Folder button clicked."
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = "Select Folder to Add for Backup"; $dialog.ShowNewFolderButton = $false
            $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $owner.Show(); $owner.Hide(); $result = $dialog.ShowDialog($owner); $owner.Dispose()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
                 $selectedPath = $dialog.SelectedPath; Write-Host "Folder selected to add: $selectedPath"
                 $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                 if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }
                 if (-not ($currentItems.Path -contains $selectedPath)) {
                    # Create a unique key for manually added items
                    $folderKey = "ManualFolder_" + ([System.IO.Path]::GetFileName($selectedPath) -replace '[^a-zA-Z0-9_]','_') + "_" + (Get-Date -Format 'HHmmss')
                    Write-Host "Adding folder: $selectedPath (Key: $folderKey)"
                    $currentItems.Add([PSCustomObject]@{ Name = $folderKey; Type = "Folder"; Path = $selectedPath; IsSelected = $true; SourceKey = $folderKey })
                    $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls # Refresh view and update space
                    Write-Host "Updated ListView ItemsSource with new folder."
                 } else { Write-Host "Skipping duplicate folder: $selectedPath" }
            } else { Write-Host "Folder selection cancelled."}
        })

        # Add User Folders (BAU Paths) handler
        $controls.btnAddBAUPaths.Add_Click({
            if (!$IsBackup) { return } # Should be disabled anyway, but double-check
            Write-Host "Add User Folders button clicked."
            try {
                $userPaths = Get-UserPaths # Gets local user folders like Downloads, Pictures, Videos
                if ($userPaths -and $userPaths.Count -gt 0) {
                    $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                    if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }
                    $addedCount = 0
                    foreach ($pathInfo in $userPaths) {
                        if (-not ($currentItems.Path -contains $pathInfo.Path)) {
                            Write-Host "Adding User Folder: $($pathInfo.Path) (Key: $($pathInfo.SourceKey))"
                            $pathInfo | Add-Member -MemberType NoteProperty -Name 'IsSelected' -Value $true -PassThru
                            $currentItems.Add($pathInfo)
                            $addedCount++
                        } else { Write-Host "Skipping duplicate user folder: $($pathInfo.Path)" }
                    }
                    if ($addedCount -gt 0) { $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls } # Refresh view and update space
                    Write-Host "Added $addedCount new user folder(s)."
                } else { Write-Warning "Could not find any standard user folders (Downloads, Pictures, Videos)." }
            } catch { Write-Error "Error adding user folders: $($_.Exception.Message)" }
        })

        # Remove Selected handler
        $controls.btnRemove.Add_Click({
            Write-Host "Remove Selected button clicked."
            $selectedObjects = @($controls.lvwFiles.SelectedItems)
            if ($selectedObjects.Count -gt 0) {
                Write-Host "Removing $($selectedObjects.Count) selected item(s)."
                $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]
                if ($currentItems -ne $null) {
                    # Create a list of items to remove to avoid modifying the collection while iterating
                    $itemsToRemove = $selectedObjects | ForEach-Object { $_ }
                    $itemsToRemove | ForEach-Object { $currentItems.Remove($_) } | Out-Null
                    $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls # Refresh view and update space
                    Write-Host "ListView ItemsSource updated after removal."
                }
            } else { Write-Host "No items selected to remove."}
        })

        # Event handler for ListView selection change (to update required space)
        $controls.lvwFiles.Add_SelectionChanged({
             Write-Host "ListView Selection Changed - Updating Required Space"
             # Run calculation in background to avoid UI lag if list is huge or paths are slow
             [System.Threading.Tasks.Task]::Run( { & $UpdateRequiredSpaceLabel -ControlsParam $controls } ) | Out-Null
        })
        # Also update space when checkbox is clicked (SelectionChanged doesn't always fire reliably for checkbox clicks)
        $controls.lvwFiles.Add_PreviewMouseLeftButtonUp({
            param($sender, $e)
            # Check if the click was on a CheckBox within the ListViewItem
            $originalSource = $e.OriginalSource
            if ($originalSource -is [System.Windows.Controls.CheckBox]) {
                Write-Host "ListView CheckBox Clicked - Updating Required Space"
                # Use Dispatcher.BeginInvoke to allow the checkbox state to update *before* calculating
                $window.Dispatcher.BeginInvoke([action]{
                    [System.Threading.Tasks.Task]::Run( { & $UpdateRequiredSpaceLabel -ControlsParam $controls } ) | Out-Null
                }, [System.Windows.Threading.DispatcherPriority]::Background) # Background priority is usually sufficient
            }
        })


        # --- Start Button Logic (Backup/Restore) ---
        $controls.btnStart.Add_Click({
            $modeStringLocal = if ($IsBackup) { 'Backup' } else { 'Restore' } # Use local var to avoid scope issues in event handler
            Write-Host "Start button clicked. Mode: $modeStringLocal"

            $location = $controls.txtSaveLoc.Text
            Write-Host "Selected location: $location"
            if ([string]::IsNullOrEmpty($location)) {
                 Write-Warning "Location text box is empty."
                 [System.Windows.MessageBox]::Show("Please select a valid target directory first.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                 return
            }
            # Additional check: Ensure path exists and is a directory
            if (-not (Test-Path $location -PathType Container)) {
                Write-Warning "Invalid location selected: '$location' does not exist or is not a directory."
                [System.Windows.MessageBox]::Show("Please select a valid target directory first. The specified path does not exist or is not a folder: `n$location", "Location Invalid", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Final check/update of required space before starting
            Write-Host "Final check of required space..."
            & $UpdateRequiredSpaceLabel -ControlsParam $controls

            # Get items to process based on selection
            $itemsToProcess = [System.Collections.Generic.List[PSCustomObject]]::new()
            $checkedItems = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected }
            if ($checkedItems) { $checkedItems | ForEach-Object { $itemsToProcess.Add($_) } }

            $doNetwork = $controls.chkNetwork.IsChecked
            $doPrinters = $controls.chkPrinters.IsChecked

            # Check if anything is selected
            if ($itemsToProcess.Count -eq 0 -and !$doNetwork -and !$doPrinters) {
                 Write-Warning "No items or settings selected to process."
                 [System.Windows.MessageBox]::Show("Please select at least one file/folder or check Network Drives/Printers to proceed.", "Nothing Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
                 return
            }


            Write-Host "Disabling UI controls and setting wait cursor."
            $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }
            $window.Cursor = [System.Windows.Input.Cursors]::Wait

            # Define UI Progress Action ScriptBlock
            $uiProgressAction = {
                param($status, $percent, $details)
                # Ensure UI updates happen on the Dispatcher thread
                $window.Dispatcher.InvokeAsync( [action]{
                    $controls.lblStatus.Content = $status
                    $controls.txtProgress.Text = $details
                    if ($percent -ge 0 -and $percent -le 100) { # Normal progress
                        $controls.prgProgress.IsIndeterminate = $false
                        $controls.prgProgress.Value = $percent
                    } elseif ($percent -eq -1) { # Error state
                        $controls.prgProgress.IsIndeterminate = $false
                        $controls.prgProgress.Value = 0
                    } else { # Indeterminate state (e.g., percent = -2 or > 100)
                        $controls.prgProgress.IsIndeterminate = $true
                    }
                }, [System.Windows.Threading.DispatcherPriority]::Background ) | Out-Null
            }

            # Start Operation in a separate task/thread to keep UI responsive
            $operationTask = [System.Threading.Tasks.Task]::Factory.StartNew({
                $success = $false
                try {
                    & $uiProgressAction.Invoke("Initializing...", -2, "Starting operation...") # Indeterminate progress

                    if ($IsBackup) {
                        Write-Host "--- Starting Background Backup Operation ---"
                        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                        $localUsername = $env:USERNAME -replace '[^a-zA-Z0-9]', '_' # Sanitize username for folder name
                        $backupRootPath = Join-Path $location "Backup_${localUsername}_$timestamp"

                        # Call Backup Function (Refactored for background)
                        $success = Invoke-BackupOperation -BackupRootPath $backupRootPath -ItemsToBackup $itemsToProcess -BackupNetworkDrives $doNetwork -BackupPrinters $doPrinters -ProgressAction $uiProgressAction

                    } else { # Restore Mode
                        Write-Host "--- Starting Background Restore Operation ---"
                        $backupRootPath = $location # The selected folder IS the backup root
                        # Get selected keys/files for filtering restore
                        # The 'Name' property in Restore mode corresponds to the folder name or setting file name within the backup dir
                        $selectedKeysOrFilesForRestore = $itemsToProcess | Select-Object -ExpandProperty Name

                        # Call Restore Function (Refactored for background)
                        $success = Invoke-RestoreOperation -BackupRootPath $backupRootPath -SelectedKeysOrFiles $selectedKeysOrFilesForRestore -RestoreNetworkDrives $doNetwork -RestorePrinters $doPrinters -ProgressAction $uiProgressAction
                    }

                    # --- Operation Completion (Background Thread) ---
                    if ($success) {
                        Write-Host "Background operation completed successfully." -ForegroundColor Green
                        & $uiProgressAction.Invoke("Operation Completed", 100, "Successfully finished $modeStringLocal.")
                        # Show message box from UI thread
                        $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show("The $modeStringLocal operation completed successfully!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }) | Out-Null
                    } else {
                        Write-Error "Background operation failed or completed with errors."
                        & $uiProgressAction.Invoke("Operation Failed", -1, "Errors occurred during $modeStringLocal. Check console/log.")
                        $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show("The $modeStringLocal operation failed or completed with errors. Check console output or log files for details.", "Error / Warnings", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) }) | Out-Null # Use Warning icon as it might be partial success
                    }

                } catch {
                    # --- Operation Failure (Background Thread) ---
                    $errorMessage = "Operation Failed (Background Task Exception): $($_.Exception.Message)"
                    Write-Error $errorMessage
                    & $uiProgressAction.Invoke("Operation Failed", -1, $errorMessage)
                    $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }) | Out-Null
                } finally {
                    # --- Re-enable UI (Background Thread -> Dispatcher) ---
                    Write-Host "Operation finished (background finally block). Re-enabling UI controls via Dispatcher."
                    $window.Dispatcher.InvokeAsync({
                        $controls | ForEach-Object {
                            if ($_.Value -is [System.Windows.Controls.Control]) {
                                # Special handling for Add User Folders button (only enable in Backup mode)
                                if ($_.Key -eq 'btnAddBAUPaths') { $_.Value.IsEnabled = $IsBackup }
                                else { $_.Value.IsEnabled = $true }
                            }
                        }
                        $window.Cursor = [System.Windows.Input.Cursors]::Arrow
                        Write-Host "UI Controls re-enabled and cursor reset."
                    }) | Out-Null
                }
            }) # End Task Factory StartNew

            # Optional: Handle task exceptions if needed, though the inner try/catch should handle most
            $operationTask.ContinueWith({
                param($task)
                if ($task.IsFaulted) {
                    $aggEx = $task.Exception.Flatten()
                    $aggEx.InnerExceptions | ForEach-Object { Write-Error "Unhandled Task Exception: $($_.Message)" }
                    # Optionally show a generic error message here too via Dispatcher
                }
            }, [System.Threading.Tasks.TaskContinuationOptions]::OnlyOnFaulted)

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
         # Clean up script-scoped variables if they exist and were created in this function's scope
         Remove-Variable -Name UpdateFreeSpaceLabel -Scope Script -ErrorAction SilentlyContinue
         Remove-Variable -Name UpdateRequiredSpaceLabel -Scope Script -ErrorAction SilentlyContinue
    }
}

# --- Refactored Backup/Restore Functions (Called by UI or Express) ---

function Invoke-BackupOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BackupRootPath,
        [Parameter(Mandatory=$true)] [System.Collections.Generic.List[PSCustomObject]]$ItemsToBackup,
        [Parameter(Mandatory=$true)] [bool]$BackupNetworkDrives,
        [Parameter(Mandatory=$true)] [bool]$BackupPrinters,
        [Parameter(Mandatory=$false)] [scriptblock]$ProgressAction, # Optional: $ProgressAction.Invoke($status, $percent, $details)
        [Parameter(Mandatory=$false)] [hashtable]$OriginalPathOverrides = @{} # Optional: For Express mode final backup logging
    )
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    Write-Host "--- Starting LOCAL Backup Operation to '$BackupRootPath' ---"
    $overallSuccess = $true # Track if any errors occur

    try {
        $UpdateProgress.Invoke("Initializing Backup", 0, "Creating backup directory...")
        if (-not (Test-Path $BackupRootPath)) { New-Item -Path $BackupRootPath -ItemType Directory -Force -EA Stop | Out-Null; Write-Host "Backup directory created." }
        else { Write-Host "Backup directory already exists: $BackupRootPath" }

        # Prepare log file
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"
        "OriginalFullPath,BackupRelativePath,SourceKey" | Set-Content -Path $csvLogPath -Encoding UTF8 -Force # Overwrite if exists

        # Check if there's anything to do
        if (($ItemsToBackup -eq $null -or $ItemsToBackup.Count -eq 0) -and !$BackupNetworkDrives -and !$BackupPrinters) {
            Write-Warning "No items or settings selected for backup."
            # Don't throw, just return true as nothing failed, but nothing was done.
            $UpdateProgress.Invoke("Backup Complete", 100, "No items or settings were selected.")
            return $true
        }

        # Calculate total steps for progress
        $networkDriveCount = $(if ($BackupNetworkDrives) { 1 } else { 0 })
        $printerCount = $(if ($BackupPrinters) { 1 } else { 0 })
        $totalSteps = ($ItemsToBackup | Measure-Object).Count + $networkDriveCount + $printerCount
        $currentStep = 0

        # --- Backup Files/Folders ---
        if ($ItemsToBackup -ne $null -and $ItemsToBackup.Count -gt 0) {
            Write-Host "Processing $($ItemsToBackup.Count) file/folder items..."
            foreach ($item in $ItemsToBackup) {
                $currentStep++
                $percentComplete = if ($totalSteps -gt 0) { [int](($currentStep / $totalSteps) * 90) } else { 0 } # File copy up to 90%
                $statusMessage = "Backing up Item $currentStep of $totalSteps"
                $UpdateProgress.Invoke($statusMessage, $percentComplete, "Processing: $($item.Name)")

                $sourcePath = $item.Path # This is the LOCAL source path
                $sourceKey = $item.SourceKey # The logical name (e.g., Signatures, ManualFolder_MyDocs)

                # Determine the original path for logging (especially for Express mode)
                $originalPathForLog = if ($OriginalPathOverrides.ContainsKey($sourcePath)) { $OriginalPathOverrides[$sourcePath] } else { $sourcePath }

                # Check if source exists
                if (-not (Test-Path -LiteralPath $sourcePath)) {
                    Write-Warning "SKIP: Source path not found: $sourcePath"
                    "`"$originalPathForLog`",`"SKIPPED_NOT_FOUND`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                    $overallSuccess = $false # Mark as not fully successful
                    continue
                }

                # Define backup item name (usually same as SourceKey, sanitized)
                $backupItemRootName = $sourceKey -replace '[^a-zA-Z0-9_.-]', '_' # Allow dots and hyphens too

                try {
                    if ($item.Type -eq "Folder") {
                        $sourceFolderInfo = Get-Item -LiteralPath $sourcePath
                        $targetBackupFolder = Join-Path $BackupRootPath $backupItemRootName
                        Write-Host "  Copying Folder: $sourcePath -> $targetBackupFolder"
                        # Copy the entire folder structure
                        Copy-Item -LiteralPath $sourcePath -Destination $targetBackupFolder -Recurse -Force -ErrorAction Stop

                        # Log all files within the copied folder
                        Get-ChildItem -LiteralPath $targetBackupFolder -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
                            # Reconstruct original path for logging (relative to source folder)
                            $relativeFilePath = $_.FullName.Substring($targetBackupFolder.Length).TrimStart('\')
                            $originalFileFullPath = Join-Path $originalPathForLog $relativeFilePath # Use the potentially overridden original base path
                            $backupRelativePath = Join-Path $backupItemRootName $relativeFilePath
                            "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                        }
                        # Add the root folder itself to the log
                        "`"$originalPathForLog`",`"$backupItemRootName`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8

                    } else { # File
                        $originalFileFullPath = $originalPathForLog # Use the potentially overridden original path
                        $targetBackupDir = Join-Path $BackupRootPath $backupItemRootName # Create a subfolder named after the key
                        $targetBackupPath = Join-Path $targetBackupDir (Split-Path $sourcePath -Leaf) # Place file inside the subfolder
                        Write-Host "  Copying File: $sourcePath -> $targetBackupPath"

                        if (-not (Test-Path $targetBackupDir)) { New-Item -Path $targetBackupDir -ItemType Directory -Force -EA Stop | Out-Null }
                        Copy-Item -LiteralPath $sourcePath -Destination $targetBackupPath -Force -EA Stop

                        # Log the file
                        $backupRelativePath = Join-Path $backupItemRootName (Split-Path $sourcePath -Leaf)
                        "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                    }
                } catch {
                    $errMsg = $_.Exception.Message -replace '"','""' # Escape quotes for CSV
                    Write-Warning "ERROR copying '$($item.Name)' from '$sourcePath': $errMsg"
                    "`"$originalPathForLog`",`"ERROR_COPY: $errMsg`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                    $overallSuccess = $false # Mark as not fully successful
                }
            }
        } else { Write-Host "No file/folder items to backup." }

        # --- Backup Network Drives ---
        if ($BackupNetworkDrives) {
            $currentStep++; $percentComplete = if ($totalSteps -gt 0) { [int](($currentStep / $totalSteps) * 95) } else { 95 }; $UpdateProgress.Invoke("Backing up Settings", $percentComplete, "Network Drives...")
            $driveFile = Join-Path $BackupRootPath "Drives.csv"
            try {
                # Query LOCAL machine for mapped drives
                Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop | Select-Object Name, ProviderName | Export-Csv -Path $driveFile -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
                Write-Host "Network drives backed up to $driveFile."
            }
            catch {
                Write-Warning "Failed to backup network drives: $($_.Exception.Message)"
                $overallSuccess = $false # Mark as not fully successful
            }
        }

        # --- Backup Printers ---
        if ($BackupPrinters) {
            $currentStep++; $percentComplete = if ($totalSteps -gt 0) { [int](($currentStep / $totalSteps) * 100) } else { 100 }; $UpdateProgress.Invoke("Backing up Settings", $percentComplete, "Printers...")
            $printerFile = Join-Path $BackupRootPath "Printers.txt"
            try {
                # Query LOCAL machine for network printers
                Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -ErrorAction Stop | Select-Object -ExpandProperty Name | Set-Content -Path $printerFile -Encoding UTF8 -Force -ErrorAction Stop
                Write-Host "Network printers backed up to $printerFile."
            }
            catch {
                Write-Warning "Failed to backup network printers: $($_.Exception.Message)"
                $overallSuccess = $false # Mark as not fully successful
            }
        }

        # --- Final Status ---
        if ($overallSuccess) {
            $UpdateProgress.Invoke("Backup Complete", 100, "Successfully backed up to: $BackupRootPath")
            Write-Host "--- LOCAL Backup Operation Finished Successfully ---" -ForegroundColor Green
        } else {
            $UpdateProgress.Invoke("Backup Complete (with errors)", 100, "Backup finished with errors. Check log/console.")
            Write-Warning "--- LOCAL Backup Operation Finished with Errors ---"
        }
        return $overallSuccess

    } catch {
        $errorMessage = "FATAL Backup Error: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Backup Failed", -1, $errorMessage) # Use -1 for fatal error state
        return $false
    }
}

function Invoke-RestoreOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BackupRootPath,
        [Parameter(Mandatory=$true)] [string[]]$SelectedKeysOrFiles, # Names from the ListView selection (folder names or Drives.csv/Printers.txt)
        [Parameter(Mandatory=$true)] [bool]$RestoreNetworkDrives,
        [Parameter(Mandatory=$true)] [bool]$RestorePrinters,
        [Parameter(Mandatory=$false)] [scriptblock]$ProgressAction
    )
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    Write-Host "--- Starting LOCAL Restore Operation from '$BackupRootPath' ---"
    $overallSuccess = $true # Track if any errors occur

    try {
        $UpdateProgress.Invoke("Initializing Restore", 0, "Checking backup contents...")
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"
        $drivesCsvPath = Join-Path $BackupRootPath "Drives.csv"
        $printersTxtPath = Join-Path $BackupRootPath "Printers.txt"

        # Validate backup root path
        if (-not (Test-Path $BackupRootPath -PathType Container)) { throw "Backup source path '$BackupRootPath' not found or is not a directory." }

        # Check if anything is selected to restore
        if ($SelectedKeysOrFiles.Count -eq 0 -and !$RestoreNetworkDrives -and !$RestorePrinters) {
            Write-Warning "No items or settings selected for restore."
            $UpdateProgress.Invoke("Restore Complete", 100, "No items or settings were selected.")
            return $true
        }

        # --- Restore Files/Folders ---
        # Filter selected items that are actual folders/files (not Drives.csv/Printers.txt)
        $selectedDataItems = $SelectedKeysOrFiles | Where-Object { $_ -ne "Drives.csv" -and $_ -ne "Printers.txt" }
        $itemsToRestoreFromLog = @() # Default to empty array

        if (Test-Path $csvLogPath -PathType Leaf) {
            Write-Host "Reading backup log: $csvLogPath"
            $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8 -ErrorAction SilentlyContinue
            if ($null -eq $backupLog) {
                Write-Warning "Failed to read or parse backup log '$csvLogPath'. Cannot restore files/folders based on log."
                $overallSuccess = $false
            } elseif (-not ($backupLog | Get-Member -Name SourceKey) -or -not ($backupLog | Get-Member -Name OriginalFullPath) -or -not ($backupLog | Get-Member -Name BackupRelativePath)) {
                Write-Warning "Backup log '$csvLogPath' is missing required columns (SourceKey, OriginalFullPath, BackupRelativePath). Cannot restore files/folders based on log."
                $overallSuccess = $false
            } else {
                # Filter log entries based on the selected SourceKeys (which correspond to the folder names in the backup)
                # We need to restore all files belonging to a selected SourceKey/Folder
                $itemsToRestoreFromLog = $backupLog | Where-Object {
                    $_.SourceKey -in $selectedDataItems -and
                    $_.BackupRelativePath -notmatch '^(SKIPPED|ERROR)_' -and # Don't try to restore skipped/errored items
                    $_.OriginalFullPath -ne $null -and $_.OriginalFullPath -ne "" # Ensure original path exists
                }
                Write-Host "Found $($itemsToRestoreFromLog.Count) file/folder entries in log matching selection."
            }
        } else {
            Write-Warning "Backup log file '$csvLogPath' not found. Cannot restore files/folders."
            # If only settings were selected, this might be okay, otherwise it's an error.
            if ($selectedDataItems.Count -gt 0) { $overallSuccess = $false }
        }

        # Calculate total steps
        $networkDriveCount = $(if ($RestoreNetworkDrives -and ($SelectedKeysOrFiles -contains "Drives.csv")) { 1 } else { 0 })
        $printerCount = $(if ($RestorePrinters -and ($SelectedKeysOrFiles -contains "Printers.txt")) { 1 } else { 0 })
        $totalSteps = $itemsToRestoreFromLog.Count + $networkDriveCount + $printerCount
        $currentStep = 0

        if ($itemsToRestoreFromLog.Count -gt 0) {
            Write-Host "Restoring $($itemsToRestoreFromLog.Count) file/folder entries..."
            foreach ($entry in $itemsToRestoreFromLog) {
                $currentStep++
                $percentComplete = if ($totalSteps -gt 0) { [int](($currentStep / $totalSteps) * 90) } else { 0 } # Restore up to 90%
                $statusMessage = "Restoring Item $currentStep of $totalSteps"
                $originalFileFullPath = $entry.OriginalFullPath
                $sourceBackupPath = Join-Path $BackupRootPath $entry.BackupRelativePath

                # Check if the source item in the backup exists
                $isSourceFile = Test-Path -LiteralPath $sourceBackupPath -PathType Leaf
                $isSourceFolder = Test-Path -LiteralPath $sourceBackupPath -PathType Container

                if ($isSourceFile) {
                    $UpdateProgress.Invoke($statusMessage, $percentComplete, "File: $(Split-Path $originalFileFullPath -Leaf)")
                    try {
                        $targetRestoreDir = Split-Path $originalFileFullPath -Parent
                        # Ensure target directory exists
                        if (-not (Test-Path $targetRestoreDir)) {
                            Write-Host "  Creating target directory: $targetRestoreDir"
                            New-Item -Path $targetRestoreDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
                        }
                        Write-Host "  Copying: $sourceBackupPath -> $originalFileFullPath"
                        Copy-Item -LiteralPath $sourceBackupPath -Destination $originalFileFullPath -Force -ErrorAction Stop
                    } catch {
                        Write-Warning "Failed to restore file '$originalFileFullPath' from '$sourceBackupPath': $($_.Exception.Message)"
                        $overallSuccess = $false
                    }
                } elseif ($isSourceFolder) {
                     # Folders themselves are implicitly created when files are restored into them.
                     # We log the folder in backup, but don't need explicit restore action unless it's empty.
                     # Check if target folder needs creation (might be empty folder backup)
                     $UpdateProgress.Invoke($statusMessage, $percentComplete, "Folder: $originalFileFullPath")
                     if (-not (Test-Path -LiteralPath $originalFileFullPath -PathType Container)) {
                         Write-Host "  Ensuring target directory exists: $originalFileFullPath"
                         try { New-Item -Path $originalFileFullPath -ItemType Directory -Force -ErrorAction Stop | Out-Null }
                         catch { Write-Warning "Failed to create target directory '$originalFileFullPath': $($_.Exception.Message)"; $overallSuccess = $false }
                     } else { Write-Host "  Target directory already exists: $originalFileFullPath" }
                } else {
                    Write-Warning "SKIP Restore: Source item '$sourceBackupPath' not found in backup (for original path '$originalFileFullPath')."
                    # This might happen if the log is inconsistent with the backup contents.
                    $overallSuccess = $false
                }
            }
        } else { Write-Host "No file/folder entries selected or found in log to restore." }

        # --- Restore Network Drives ---
        if ($RestoreNetworkDrives -and ($SelectedKeysOrFiles -contains "Drives.csv")) {
            $currentStep++; $percentComplete = if ($totalSteps -gt 0) { [int](($currentStep / $totalSteps) * 95) } else { 95 }; $UpdateProgress.Invoke("Restoring Settings", $percentComplete, "Network Drives...")
            if (Test-Path $drivesCsvPath -PathType Leaf) {
                Write-Host "Restoring network drives from $drivesCsvPath"
                try {
                    Import-Csv $drivesCsvPath | ForEach-Object {
                        $driveLetter = $_.Name.TrimEnd(':')
                        $networkPath = $_.ProviderName
                        if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') {
                            $drivePath = "$($driveLetter):"
                            if (-not (Test-Path -LiteralPath $drivePath)) {
                                try {
                                    Write-Host "  Mapping $driveLetter -> $networkPath"
                                    # Use New-PSDrive with -Persist for user logon persistence, Scope Global might not be needed/desired unless admin context requires it
                                    New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -ErrorAction Stop # Persist makes it available after script exits
                                } catch {
                                    Write-Warning "  Failed to map drive $driveLetter`: $($_.Exception.Message)"
                                    $overallSuccess = $false
                                }
                            } else {
                                Write-Host "  Skipping drive $driveLetter`: Path already exists."
                            }
                        } else {
                            Write-Warning "  Skipping invalid drive mapping entry: Name='$($_.Name)', ProviderName='$($_.ProviderName)'"
                        }
                    }
                } catch {
                    Write-Warning "Error processing drives file '$drivesCsvPath': $($_.Exception.Message)"
                    $overallSuccess = $false
                }
            } else {
                Write-Warning "Network Drives selected for restore, but 'Drives.csv' not found in backup."
                $overallSuccess = $false
            }
        }

        # --- Restore Printers ---
        if ($RestorePrinters -and ($SelectedKeysOrFiles -contains "Printers.txt")) {
            $currentStep++; $percentComplete = if ($totalSteps -gt 0) { [int](($currentStep / $totalSteps) * 100) } else { 100 }; $UpdateProgress.Invoke("Restoring Settings", $percentComplete, "Printers...")
            if (Test-Path $printersTxtPath -PathType Leaf) {
                 Write-Host "Restoring network printers from $printersTxtPath"
                try {
                    # Use COM object for adding printers - generally reliable
                    $wsNet = New-Object -ComObject WScript.Network
                    Get-Content $printersTxtPath | ForEach-Object {
                        $printerPath = $_.Trim()
                        if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') {
                            try {
                                Write-Host "  Adding printer: $printerPath"
                                $wsNet.AddWindowsPrinterConnection($printerPath)
                                # Optional: Set as default? Requires more logic to identify original default.
                                # $wsNet.SetDefaultPrinter($printerPath)
                            } catch {
                                Write-Warning "  Failed to add printer '$printerPath': $($_.Exception.Message)"
                                $overallSuccess = $false
                            }
                        } else {
                            Write-Warning "  Skipping invalid printer path line: '$_'"
                        }
                    }
                } catch {
                    Write-Warning "Error processing printers file '$printersTxtPath' or using WScript.Network: $($_.Exception.Message)"
                    $overallSuccess = $false
                }
            } else {
                Write-Warning "Printers selected for restore, but 'Printers.txt' not found in backup."
                $overallSuccess = $false
            }
        }

        # --- Final Status ---
        if ($overallSuccess) {
            $UpdateProgress.Invoke("Restore Complete", 100, "Successfully restored from: $BackupRootPath")
            Write-Host "--- LOCAL Restore Operation Finished Successfully ---" -ForegroundColor Green
        } else {
            $UpdateProgress.Invoke("Restore Complete (with errors)", 100, "Restore finished with errors. Check console.")
            Write-Warning "--- LOCAL Restore Operation Finished with Errors ---"
        }
        return $overallSuccess

    } catch {
        $errorMessage = "FATAL Restore Error: $($_.Exception.Message)"
        Write-Error $errorMessage
        $UpdateProgress.Invoke("Restore Failed", -1, $errorMessage) # Use -1 for fatal error state
        return $false
    }
}


# --- Express Mode Function ---
# --- Express Mode Function ---
function Execute-ExpressModeLogic {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential # Admin credential for remote access
    )

    Write-Host "--- Starting Express Mode Logic ---" -ForegroundColor Cyan
    $targetDevice = $null
    $mappedDriveLetter = "X" # Or choose another available letter dynamically
    $localTempTransferDir = Join-Path $env:TEMP "ExpressTransfer_$(Get-Date -Format 'yyyyMMddHHmmss')"
    $localBackupBaseDir = "C:\LocalData" # Where the final backup copy will reside
    $transferSuccess = $true # Assume success initially, set to false on any significant error
    $transferLog = [System.Collections.Generic.List[string]]::new()
    $transferLog.Add("Timestamp,Action,Status,Details")
    $logFilePath = $null # Final log file path
    $tempLogPath = $null # Temp log file path if final dir fails
    $finalBackupDir = $null # Path to the final local backup created at the end
    $remotePathsToTransfer = $null # Store the list of resolved remote paths
    $drivePath = "${mappedDriveLetter}:" # Define drive path with colon here

    # --- Local Logging Function ---
    Function LogTransfer ($Action, $Status, $Details) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        # Sanitize details for CSV logging (basic quote escaping)
        $safeDetails = $Details -replace '"', '""'
        $logEntry = """$timestamp"",""$Action"",""$Status"",""$safeDetails"""
        $transferLog.Add($logEntry)
        # Also write to host based on status
        switch ($Status) {
            "Success" { Write-Host "LOG [Success]: $Action - $Details" -ForegroundColor Green }
            "Info"    { Write-Host "LOG [Info]:    $Action - $Details" }
            "Attempt" { Write-Host "LOG [Attempt]: $Action - $Details" -ForegroundColor Yellow }
            "Warning" { Write-Warning "LOG [Warning]: $Action - $Details" }
            "Error"   { Write-Error   "LOG [Error]:   $Action - $Details" }
            "Fatal Error" { Write-Error "LOG [FATAL]:   $Action - $Details" }
            default   { Write-Host "LOG [$Status]: $Action - $Details" }
        }
    }

    # --- Local Post-Transfer Action Functions (Run on NEW machine) ---
    # Define these locally to ensure they run in the current elevated context
    function Set-LocalGPupdate {
        Write-Host "LOCAL ACTION: Initiating Group Policy update..." -ForegroundColor Cyan
        LogTransfer "Local Action" "Attempt" "Running gpupdate /force locally"
        try {
            $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -WindowStyle Hidden -ErrorAction Stop
            if ($process.ExitCode -eq 0) {
                Write-Host "LOCAL ACTION: Group Policy update completed successfully." -ForegroundColor Green
                LogTransfer "Local Action" "Success" "Local gpupdate completed (Exit Code 0)"
            } else {
                Write-Warning "LOCAL ACTION: Group Policy update process finished with exit code: $($process.ExitCode)."
                LogTransfer "Local Action" "Warning" "Local gpupdate finished (Exit Code $($process.ExitCode))"
            }
        } catch {
            Write-Error "LOCAL ACTION: Failed to start local GPUpdate process: $($_.Exception.Message)"
            LogTransfer "Local Action" "Error" "Failed to start local GPUpdate: $($_.Exception.Message)"
        }
        Write-Host "LOCAL ACTION: Exiting Set-LocalGPupdate function."
    }

    function Start-LocalConfigManagerActions {
         Write-Host "LOCAL ACTION: Entering Start-LocalConfigManagerActions function." -ForegroundColor Cyan
         LogTransfer "Local Action" "Attempt" "Triggering local ConfigMgr client actions"
         $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"; $clientSDKNamespace = "root\ccm\clientsdk"; $clientClassName = "CCM_ClientUtilities"; $scheduleMethodName = "TriggerSchedule"; $overallSuccess = $false; $cimAttemptedAndSucceeded = $false
         $scheduleActions = @(
            @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'}
            @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'}
         )
         Write-Host "LOCAL ACTION: Defined $($scheduleActions.Count) CM actions to trigger locally."
         $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
         if (-not $ccmService) { Write-Warning "LOCAL ACTION: CM service (CcmExec) not found locally. Skipping."; LogTransfer "Local Action" "Warning" "Local CcmExec service not found"; return $false }
         elseif ($ccmService.Status -ne 'Running') { Write-Warning "LOCAL ACTION: CM service (CcmExec) is not running locally (Status: $($ccmService.Status)). Skipping."; LogTransfer "Local Action" "Warning" "Local CcmExec service not running ($($ccmService.Status))"; return $false }
         else { Write-Host "LOCAL ACTION: CM service (CcmExec) found and running locally." }

         Write-Host "LOCAL ACTION: Attempting Method 1: Triggering via CIM (Local)..."
         $cimMethodSuccess = $true
         try {
             if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
                  Write-Host "LOCAL ACTION: CIM Class found locally."
                  foreach ($action in $scheduleActions) {
                     Write-Host "LOCAL ACTION:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                     try { Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop; Write-Host "LOCAL ACTION:     $($action.Name) triggered successfully via CIM." }
                     catch { Write-Warning "LOCAL ACTION:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"; $cimMethodSuccess = $false }
                  }
                  if ($cimMethodSuccess) { $cimAttemptedAndSucceeded = $true; $overallSuccess = $true; Write-Host "LOCAL ACTION: All actions successfully triggered via CIM." -ForegroundColor Green; LogTransfer "Local Action" "Success" "Triggered all CM actions via local CIM" }
                  else { Write-Warning "LOCAL ACTION: One or more actions failed to trigger via CIM."; LogTransfer "Local Action" "Warning" "One or more CM actions failed via local CIM" }
             } else { Write-Warning "LOCAL ACTION: CIM Class '$clientClassName' not found locally. Cannot use CIM method."; $cimMethodSuccess = $false; LogTransfer "Local Action" "Warning" "Local CIM class $clientClassName not found" }
         } catch { Write-Error "LOCAL ACTION: Unexpected error during local CIM attempt: $($_.Exception.Message)"; $cimMethodSuccess = $false; LogTransfer "Local Action" "Error" "Local CIM attempt error: $($_.Exception.Message)" }

         if (-not $cimAttemptedAndSucceeded) {
             Write-Host "LOCAL ACTION: CIM failed/unavailable. Attempting Method 2: Fallback via ccmexec.exe (Local)..."
             if (Test-Path -Path $ccmExecPath -PathType Leaf) {
                 Write-Host "LOCAL ACTION: Found $ccmExecPath locally."
                 $execMethodSuccess = $true
                 foreach ($action in $scheduleActions) {
                     Write-Host "LOCAL ACTION:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                     try { $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop; if ($process.ExitCode -ne 0) { Write-Warning "LOCAL ACTION:     $($action.Name) via ccmexec.exe finished with exit code $($process.ExitCode)." } else { Write-Host "LOCAL ACTION:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)." } }
                     catch { Write-Warning "LOCAL ACTION:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"; $execMethodSuccess = $false }
                 }
                 if ($execMethodSuccess) { $overallSuccess = $true; Write-Host "LOCAL ACTION: Finished attempting actions via ccmexec.exe." -ForegroundColor Green; LogTransfer "Local Action" "Success" "Triggered CM actions via local ccmexec.exe" }
                 else { Write-Warning "LOCAL ACTION: One or more actions failed to execute via ccmexec.exe."; LogTransfer "Local Action" "Warning" "One or more CM actions failed via local ccmexec.exe" }
             } else { Write-Warning "LOCAL ACTION: Fallback executable not found locally at $ccmExecPath."; LogTransfer "Local Action" "Warning" "Local ccmexec.exe not found at $ccmExecPath" }
         }

         if ($overallSuccess) { Write-Host "LOCAL ACTION: CM actions attempt finished successfully." -ForegroundColor Green }
         else { Write-Warning "LOCAL ACTION: CM actions attempt finished, but could not be confirmed as fully successful." } # Don't set $transferSuccess=false here, these are best effort
         Write-Host "LOCAL ACTION: Exiting Start-LocalConfigManagerActions function."; return $overallSuccess
    }
    # --- End Local Post-Transfer Action Functions ---


    try {
        # --- 1. Get Target Device ---
        while ([string]::IsNullOrWhiteSpace($targetDevice)) {
            $targetDevice = Read-Host "Enter the IP address or Hostname of the OLD (source) device"
            if ([string]::IsNullOrWhiteSpace($targetDevice)) { Write-Warning "Target device cannot be empty." }
        }
        LogTransfer "Input" "Info" "Target device specified: $targetDevice"

        # --- 2. Map Network Drive ---
        $uncPath = "\\$targetDevice\c$"
        # $drivePath is defined above as "${mappedDriveLetter}:"
        LogTransfer "Map Drive" "Attempt" "Mapping $uncPath to $drivePath"
        if (Test-Path $drivePath) {
            LogTransfer "Map Drive" "Warning" "Drive $drivePath already exists. Attempting to remove..."
            try { Remove-PSDrive -Name $mappedDriveLetter -Force -ErrorAction Stop }
            catch { throw "Failed to remove existing mapped drive '$drivePath'. Please remove it manually and retry. Error: $($_.Exception.Message)" }
            LogTransfer "Map Drive" "Info" "Removed existing drive $drivePath."
        }
        try {
            New-PSDrive -Name $mappedDriveLetter -PSProvider FileSystem -Root $uncPath -Credential $Credential -ErrorAction Stop | Out-Null
            LogTransfer "Map Drive" "Success" "Successfully mapped $uncPath to $drivePath"
        } catch {
            throw "Failed to map network drive '$uncPath' to '$drivePath'. Verify connectivity, permissions, and credentials. Error: $($_.Exception.Message)"
        }

        # --- 3. Get Remote Logged-on Username ---
        LogTransfer "Get Remote User" "Attempt" "Querying Win32_ComputerSystem on $targetDevice"
        $remoteUsername = $null
        try {
            # Use CIM for modern approach, fallback to WMI if needed, via Invoke-Command
            $remoteResult = Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                try { (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).UserName }
                catch { Write-Warning "CIM failed, trying WMI: $($_.Exception.Message)"; (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop).UserName }
            } -ErrorAction Stop

            if ($remoteResult -and $remoteResult -match '\\') { # Check if it contains a domain\user format
                $remoteUsername = ($remoteResult -split '\\')[-1]
                LogTransfer "Get Remote User" "Success" "Auto-detected remote user: $remoteUsername"
            } elseif ($remoteResult) { # Handle case where it might just return username without domain
                 $remoteUsername = $remoteResult
                 LogTransfer "Get Remote User" "Success" "Auto-detected remote user (no domain): $remoteUsername"
            } else {
                throw "Could not retrieve username from remote machine or it was empty."
            }
        } catch {
            LogTransfer "Get Remote User" "Warning" "Failed to auto-detect remote user: $($_.Exception.Message)"
            while ([string]::IsNullOrWhiteSpace($remoteUsername)) {
                $remoteUsername = Read-Host "Could not auto-detect remote user. Please enter the Windows username logged into '$targetDevice'"
                if ([string]::IsNullOrWhiteSpace($remoteUsername)) { Write-Warning "Remote username cannot be empty." }
            }
            LogTransfer "Get Remote User" "Manual Input" "User provided remote username: $remoteUsername"
        }

        # --- 4. Define Remote Paths and Local Destinations ---
        $remoteUserProfile = "$drivePath\Users\$remoteUsername"
        LogTransfer "Path Setup" "Info" "Expected remote user profile path: $remoteUserProfile"
        if (-not (Test-Path -LiteralPath $remoteUserProfile -PathType Container)) {
            throw "Remote user profile path '$remoteUserProfile' not found on mapped drive. Verify username and profile existence."
        }

        LogTransfer "Gather Paths" "Attempt" "Resolving templates for remote profile '$remoteUserProfile'"
        # Resolve paths on the *remote* machine (via mapped drive)
        $remotePathsToTransfer = Resolve-UserPathTemplates -UserProfilePath $remoteUserProfile -Templates $script:pathTemplates
        if ($remotePathsToTransfer -eq $null -or $remotePathsToTransfer.Count -eq 0) {
            LogTransfer "Gather Paths" "Warning" "No template files/folders found for user '$remoteUsername' on '$targetDevice'."
            # Don't throw, maybe only settings are needed.
        } else {
            LogTransfer "Gather Paths" "Info" "Found $($remotePathsToTransfer.Count) template items on remote machine."
        }

        LogTransfer "Setup" "Info" "Creating local temporary transfer directory: $localTempTransferDir"
        New-Item -Path $localTempTransferDir -ItemType Directory -Force -ErrorAction Stop | Out-Null

        # --- 5. Transfer Files/Folders ---
        Write-Host "`n--- Starting File/Folder Transfer from $targetDevice ---" -ForegroundColor Cyan
        $transferErrors = $false
        $itemsProcessedCount = 0
        if ($remotePathsToTransfer -ne $null -and $remotePathsToTransfer.Count -gt 0) {
            $totalItems = $remotePathsToTransfer.Count
            LogTransfer "File Transfer" "Info" "Attempting to transfer $totalItems items..."
            foreach ($item in $remotePathsToTransfer) {
                $itemsProcessedCount++
                $progress = [int](($itemsProcessedCount / $totalItems) * 50) + 10 # Progress from 10% to 60%
                Write-Progress -Activity "Transferring Files from $targetDevice" -Status "[$progress%]: Copying $($item.Name)" -PercentComplete $progress -Id 1

                $sourcePath = $item.Path # Path on X: drive
                # Destination path uses SourceKey to create a root folder in the temp dir
                $destRelativePathBase = $item.SourceKey -replace '[^a-zA-Z0-9_.-]', '_'
                $destinationPath = Join-Path $localTempTransferDir $destRelativePathBase

                LogTransfer "File Transfer" "Attempt" "Copying [$($item.Type)] '$($item.Name)' from '$sourcePath'"

                try {
                    if ($item.Type -eq 'Folder') {
                        # Copy folder contents into the destination folder named after the key
                        Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Recurse -Force -Container -ErrorAction Stop
                        LogTransfer "File Transfer" "Success" "Copied Folder '$($item.Name)' to '$destinationPath'"
                    } else { # File
                        # Create the key-named directory if it doesn't exist
                        if (-not (Test-Path $destinationPath)) { New-Item -Path $destinationPath -ItemType Directory -Force -EA Stop | Out-Null }
                        # Copy the file into that directory
                        $fileDestPath = Join-Path $destinationPath (Split-Path $sourcePath -Leaf)
                        Copy-Item -LiteralPath $sourcePath -Destination $fileDestPath -Force -ErrorAction Stop
                        LogTransfer "File Transfer" "Success" "Copied File '$($item.Name)' to '$fileDestPath'"
                    }
                } catch {
                    LogTransfer "File Transfer" "Error" "Failed copy '$($item.Name)': $($_.Exception.Message)"
                    $transferErrors = $true
                    $transferSuccess = $false # Mark overall transfer as failed if any copy error occurs
                }
            }
            Write-Progress -Activity "Transferring Files from $targetDevice" -Completed -Id 1
            if ($transferErrors) { LogTransfer "File Transfer" "Warning" "File/folder transfer completed with errors." }
            else { LogTransfer "File Transfer" "Success" "File/folder transfer completed successfully." }
        } else {
            LogTransfer "File Transfer" "Info" "Skipping file/folder transfer - no template items found or resolved."
            Write-Progress -Activity "Transferring Files from $targetDevice" -Status "[60%]: No files/folders found" -PercentComplete 60 -Id 1
        }

        # --- 6. Transfer Settings (Network Drives / Printers) ---
        Write-Host "`n--- Transferring Settings from $targetDevice ---" -ForegroundColor Cyan
        Write-Progress -Activity "Transferring Settings" -Status "[60%]: Retrieving Drives..." -PercentComplete 60 -Id 1
        LogTransfer "Get Drives" "Attempt" "Querying remote mapped drives on $targetDevice"
        $remoteDrivesCsvPath = Join-Path $localTempTransferDir "Drives.csv"
        $drivesRetrieved = $false
        try {
            Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                Get-WmiObject -Class Win32_MappedLogicalDisk -ErrorAction Stop | Select-Object Name, ProviderName
            } -ErrorAction Stop | Export-Csv -Path $remoteDrivesCsvPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
            LogTransfer "Get Drives" "Success" "Saved remote drives list to $remoteDrivesCsvPath"
            $drivesRetrieved = $true
        } catch {
            # *** FIX APPLIED HERE ***
            LogTransfer "Get Drives" "Error" "Failed get drives from ${targetDevice}: $($_.Exception.Message)"
            # Don't set $transferSuccess=false, settings retrieval failure is less critical than file copy failure
        }

        Write-Progress -Activity "Transferring Settings" -Status "[75%]: Retrieving Printers..." -PercentComplete 75 -Id 1
        LogTransfer "Get Printers" "Attempt" "Querying remote network printers on $targetDevice"
        $remotePrintersTxtPath = Join-Path $localTempTransferDir "Printers.txt"
        $printersRetrieved = $false
        try {
            Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock {
                Get-WmiObject -Class Win32_Printer -Filter "Local=False AND Network=True" -ErrorAction Stop | Select-Object -ExpandProperty Name
            } -ErrorAction Stop | Set-Content -Path $remotePrintersTxtPath -Encoding UTF8 -Force -ErrorAction Stop
            LogTransfer "Get Printers" "Success" "Saved remote printers list to $remotePrintersTxtPath"
            $printersRetrieved = $true
        } catch {
            # *** FIX APPLIED HERE ***
            LogTransfer "Get Printers" "Error" "Failed get printers from ${targetDevice}: $($_.Exception.Message)"
        }
        Write-Progress -Activity "Transferring Settings" -Completed -Id 1

        # --- 7. Post-Transfer LOCAL Actions (GPUpdate, ConfigMgr) ---
        # *** MODIFIED: Run these LOCALLY on the NEW machine ***
        Write-Host "`n--- Executing Post-Transfer Actions LOCALLY ---" -ForegroundColor Cyan
        Write-Progress -Activity "Local Actions" -Status "[85%]: Running GPUpdate/ConfigMgr..." -PercentComplete 85 -Id 1
        try {
            Set-LocalGPupdate # Call the locally defined function
            Start-LocalConfigManagerActions # Call the locally defined function
            LogTransfer "Local Actions" "Info" "Finished attempting local GPUpdate/ConfigMgr actions."
        } catch {
            # Catch any unexpected error from the functions themselves (though they have internal try/catch)
            LogTransfer "Local Actions" "Error" "Unexpected error during local actions: $($_.Exception.Message)"
        }
        Write-Progress -Activity "Local Actions" -Completed -Id 1

        # --- 8. Create Final Local Backup Copy ---
        # *** MODIFIED: Replicate Invoke-BackupOperation structure/logging ***
        Write-Host "`n--- Creating Final Local Backup Copy ---" -ForegroundColor Cyan
        Write-Progress -Activity "Creating Local Backup" -Status "[90%]: Preparing backup..." -PercentComplete 90 -Id 1
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        # Sanitize remote device name and username for the backup folder name
        $safeTargetDeviceName = $targetDevice -replace '[^a-zA-Z0-9_.-]', '_'
        $safeRemoteUsername = $remoteUsername -replace '[^a-zA-Z0-9_.-]', '_'
        $finalBackupDir = Join-Path $localBackupBaseDir "ExpressBackup_${safeRemoteUsername}_from_${safeTargetDeviceName}_$timestamp"
        LogTransfer "Local Backup" "Attempt" "Creating final backup structure in '$finalBackupDir'"

        try {
            # Create the root backup directory
            New-Item -Path $finalBackupDir -ItemType Directory -Force -ErrorAction Stop | Out-Null

            # Create the FileList_Backup.csv log
            $finalCsvLogPath = Join-Path $finalBackupDir "FileList_Backup.csv"
            "OriginalFullPath,BackupRelativePath,SourceKey" | Set-Content -Path $finalCsvLogPath -Encoding UTF8 -Force
            LogTransfer "Local Backup" "Info" "Created backup log: $finalCsvLogPath"

            # Process transferred files/folders
            if ($remotePathsToTransfer -ne $null -and $remotePathsToTransfer.Count -gt 0) {
                LogTransfer "Local Backup" "Info" "Processing $($remotePathsToTransfer.Count) transferred items for final backup..."
                foreach ($item in $remotePathsToTransfer) {
                    $originalRemotePath = $item.Path # The path on X:
                    $sourceKey = $item.SourceKey
                    $backupItemName = $sourceKey -replace '[^a-zA-Z0-9_.-]', '_' # Relative path base in backup

                    # Path where the item *should* be in the temp directory
                    $sourcePathInTempBase = Join-Path $localTempTransferDir $backupItemName

                    # Check if the item actually exists in the temp directory (i.e., transfer succeeded)
                    $tempItemExists = Test-Path -LiteralPath $sourcePathInTempBase #-ErrorAction SilentlyContinue

                    if ($tempItemExists) {
                        $destinationPathInFinal = Join-Path $finalBackupDir $backupItemName
                        Write-Progress -Activity "Creating Local Backup" -Status "[90%]: Copying $($item.Name)..." -PercentComplete 90 -Id 1
                        LogTransfer "Local Backup" "Attempt" "Copying '$($item.Name)' from temp to final backup '$destinationPathInFinal'"
                        try {
                            # Copy from Temp to Final Backup directory
                            Copy-Item -LiteralPath $sourcePathInTempBase -Destination $destinationPathInFinal -Recurse -Force -ErrorAction Stop

                            # Log success in FileList_Backup.csv
                            # If it was a folder, log the base folder entry
                            if ($item.Type -eq "Folder") {
                                "`"$originalRemotePath`",`"$backupItemName`",`"$sourceKey`"" | Add-Content -Path $finalCsvLogPath -Encoding UTF8
                                # Log files within the folder as well
                                Get-ChildItem -LiteralPath $destinationPathInFinal -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
                                    $relativeFilePath = $_.FullName.Substring($destinationPathInFinal.Length).TrimStart('\')
                                    $originalFileFullPath = "$originalRemotePath\$relativeFilePath" # Append relative path to original remote base
                                    $backupRelativePath = "$backupItemName\$relativeFilePath"
                                    "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $finalCsvLogPath -Encoding UTF8
                                }
                            } else { # File
                                # The file was copied into a folder named $backupItemName, so the relative path includes that folder.
                                $backupRelativePath = Join-Path $backupItemName (Split-Path $originalRemotePath -Leaf)
                                "`"$originalRemotePath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $finalCsvLogPath -Encoding UTF8
                            }
                            LogTransfer "Local Backup" "Success" "Copied '$($item.Name)' to final backup."

                        } catch {
                            LogTransfer "Local Backup" "Error" "Failed to copy '$($item.Name)' from temp to final backup: $($_.Exception.Message)"
                            # Log error in FileList_Backup.csv
                            "`"$originalRemotePath`",`"ERROR_FINAL_COPY: $($_.Exception.Message -replace '"','""')`",`"$sourceKey`"" | Add-Content -Path $finalCsvLogPath -Encoding UTF8
                            $transferSuccess = $false # Failed to create proper backup
                        }
                    } else {
                        LogTransfer "Local Backup" "Warning" "Item '$($item.Name)' not found in temp directory '$sourcePathInTempBase'. Skipping final backup copy."
                        # Log skipped item in FileList_Backup.csv
                        "`"$originalRemotePath`",`"SKIPPED_NOT_TRANSFERRED`",`"$sourceKey`"" | Add-Content -Path $finalCsvLogPath -Encoding UTF8
                        # $transferSuccess = false # Don't mark as failure if initial transfer already logged error
                    }
                }
            } else { LogTransfer "Local Backup" "Info" "No file/folder items were processed in the initial transfer." }

            # Copy retrieved settings files
            Write-Progress -Activity "Creating Local Backup" -Status "[95%]: Copying Settings..." -PercentComplete 95 -Id 1
            if ($drivesRetrieved -and (Test-Path $remoteDrivesCsvPath)) {
                try { Copy-Item -LiteralPath $remoteDrivesCsvPath -Destination (Join-Path $finalBackupDir "Drives.csv") -Force -ErrorAction Stop; LogTransfer "Local Backup" "Success" "Copied Drives.csv to final backup." }
                catch { LogTransfer "Local Backup" "Error" "Failed copy Drives.csv to final backup: $($_.Exception.Message)"; $transferSuccess = $false }
            } elseif ($drivesRetrieved) { LogTransfer "Local Backup" "Warning" "Drives were marked as retrieved, but temp file '$remoteDrivesCsvPath' not found." }
            else { LogTransfer "Local Backup" "Info" "Drives were not retrieved, skipping Drives.csv copy." }

            if ($printersRetrieved -and (Test-Path $remotePrintersTxtPath)) {
                try { Copy-Item -LiteralPath $remotePrintersTxtPath -Destination (Join-Path $finalBackupDir "Printers.txt") -Force -ErrorAction Stop; LogTransfer "Local Backup" "Success" "Copied Printers.txt to final backup." }
                catch { LogTransfer "Local Backup" "Error" "Failed copy Printers.txt to final backup: $($_.Exception.Message)"; $transferSuccess = $false }
            } elseif ($printersRetrieved) { LogTransfer "Local Backup" "Warning" "Printers were marked as retrieved, but temp file '$remotePrintersTxtPath' not found." }
            else { LogTransfer "Local Backup" "Info" "Printers were not retrieved, skipping Printers.txt copy." }

            LogTransfer "Local Backup" "Success" "Finished creating local backup structure at '$finalBackupDir'"

        } catch {
            LogTransfer "Local Backup" "Fatal Error" "Failed to create final backup directory or initial log: $($_.Exception.Message)"
            $transferSuccess = $false
            # Attempt to save log to temp if backup dir creation failed
            $finalBackupDir = $null # Ensure log doesn't try to save here
        }
        Write-Progress -Activity "Creating Local Backup" -Completed -Id 1


        # --- 9. Final Log Saving ---
        Write-Host "`n--- Finalizing Operation ---" -ForegroundColor Cyan
        # Determine final log path
        if (-not ([string]::IsNullOrEmpty($finalBackupDir)) -and (Test-Path $finalBackupDir)) {
            $logFilePath = Join-Path $finalBackupDir "ExpressTransferLog.csv" # Use a distinct name
        } else {
            $logFilePath = Join-Path $env:TEMP "ExpressTransfer_Log_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            LogTransfer "Save Log" "Warning" "Final backup directory '$finalBackupDir' not available. Saving main log to temp: $logFilePath"
        }

        try {
            LogTransfer "Save Log" "Attempt" "Saving main transfer log to '$logFilePath'"
            $transferLog | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Set-Content -Path $logFilePath -Encoding UTF8 -Force
            LogTransfer "Save Log" "Success" "Main transfer log saved successfully."
        } catch {
            Write-Error "CRITICAL: Failed to save main transfer log to '$logFilePath': $($_.Exception.Message)"
            LogTransfer "Save Log" "Fatal Error" "Failed save main log: $($_.Exception.Message)"
            # Try saving to an alternate temp location as a last resort
            $altTempLog = Join-Path $env:TEMP "ExpressTransfer_Log_FAILED_SAVE_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            try { $transferLog | ConvertTo-Csv -NoTypeInformation -Delimiter ',' | Set-Content -Path $altTempLog -Encoding UTF8 -Force; Write-Warning "Saved log to alternate temp location: $altTempLog" } catch {}
            $logFilePath = $altTempLog # Update path for final message
        }

        Write-Progress -Activity "Express Transfer" -Status "[100%]: Finished." -PercentComplete 100 -Id 1
        Write-Host "--- Express Transfer Operation Finished ---" -ForegroundColor Cyan

    } catch {
        # Catch major failures like mapping drive, getting username, creating temp dir etc.
        $errorMessage = "FATAL Express Transfer Error: $($_.Exception.Message)"
        Write-Error $errorMessage
        LogTransfer "Overall Status" "Fatal Error" $errorMessage
        $transferSuccess = $false
        # Ensure progress bar is stopped
        Write-Progress -Activity "Express Transfer" -Completed -Id 1 -ErrorAction SilentlyContinue
    } finally {
        # --- 10. Cleanup ---
        Write-Host "`n--- Cleaning Up ---" -ForegroundColor Cyan
        # Remove mapped drive
        if (Test-Path $drivePath) {
            LogTransfer "Cleanup" "Attempt" "Removing mapped drive $drivePath"
            try { Remove-PSDrive -Name $mappedDriveLetter -Force -ErrorAction Stop; LogTransfer "Cleanup" "Success" "Removed drive $drivePath" }
            catch {
                # *** FIX APPLIED HERE ***
                LogTransfer "Cleanup" "Error" "Failed remove drive ${drivePath}: $($_.Exception.Message)"
            }
        } else { LogTransfer "Cleanup" "Info" "Mapped drive $drivePath already removed or not mapped."}

        # Remove local temp directory
        if (Test-Path $localTempTransferDir) {
            LogTransfer "Cleanup" "Attempt" "Removing temp directory: $localTempTransferDir"
            try { Remove-Item -Path $localTempTransferDir -Recurse -Force -ErrorAction Stop; LogTransfer "Cleanup" "Success" "Removed temp dir $localTempTransferDir" }
            catch {
                # *** FIX APPLIED HERE ***
                LogTransfer "Cleanup" "Error" "Failed remove temp dir ${localTempTransferDir}: $($_.Exception.Message)"
            }
        } else { LogTransfer "Cleanup" "Info" "Temp directory '$localTempTransferDir' not found for removal."}

        # --- Final User Message ---
        $finalLogPathMessage = if ($logFilePath -and (Test-Path $logFilePath)) { $logFilePath } else { "Log saving failed. Check console output." }

        if ($transferSuccess) {
            Write-Host "Express Transfer process completed successfully." -ForegroundColor Green
            LogTransfer "Overall Status" "Success" "Express Transfer Completed Successfully. Log: $finalLogPathMessage"
            [System.Windows.MessageBox]::Show("Express Transfer completed successfully.`n`nFinal Backup Location: $finalBackupDir`nLog File: $finalLogPathMessage", "Express Transfer Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } else {
            Write-Error "Express Transfer failed or completed with errors. Please review the log file."
            LogTransfer "Overall Status" "Failure" "Express Transfer Failed or had errors. Log: $finalLogPathMessage"
            [System.Windows.MessageBox]::Show("Express Transfer failed or completed with errors.`n`nPlease review the log file for details:`n$finalLogPathMessage", "Express Transfer Failed / Errors", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            # Optional: Throw an error to halt script execution in console if desired
            # throw "Express Transfer failed. Check log: $finalLogPathMessage"
        }
    }
}

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---" -ForegroundColor Magenta
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue
$operationMode = 'Cancel' # Default

# Ensure Default Path Exists (used by Backup/Restore GUI and Express final backup)
if (-not (Test-Path $script:DefaultPath -PathType Container)) {
    Write-Host "Default path '$($script:DefaultPath)' not found or is not a directory. Attempting to create." -ForegroundColor Yellow
    try {
        New-Item -Path $script:DefaultPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Default path created." -ForegroundColor Green
    } catch {
        Write-Warning "Could not create default path: $($script:DefaultPath). Express mode backups may fail if C:\LocalData cannot be created. Error: $($_.Exception.Message)"
    }
}

try {
    Write-Host "Calling Show-ModeDialog..."
    $operationMode = Show-ModeDialog

    switch ($operationMode) {
        'Backup' {
            Write-Host "Mode Selected: Backup" -ForegroundColor Cyan
            Show-MainWindow -Mode 'Backup'
        }
        'Restore' {
            Write-Host "Mode Selected: Restore" -ForegroundColor Cyan
            # Background job for GPUpdate/SCCM was started in Show-ModeDialog
            Show-MainWindow -Mode 'Restore'
        }
        'Express' {
            Write-Host "Mode Selected: Express" -ForegroundColor Cyan
            Write-Host "Checking for Administrator privileges..."
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

            if (-not $isAdmin) {
                Write-Warning "Express mode requires Administrator privileges. Attempting to relaunch..."
                LogTransfer "Elevation" "Attempt" "Relaunching script as Administrator" # Use LogTransfer if defined globally, otherwise Write-Host
                try {
                    # Construct arguments carefully, especially paths with spaces
                    $powershellArgs = "-NoProfile -ExecutionPolicy Bypass -File ""$PSCommandPath"""
                    Start-Process powershell.exe -ArgumentList $powershellArgs -Verb RunAs -ErrorAction Stop
                    Write-Host "Relaunch request sent. Exiting current instance."
                } catch {
                    $errMsg = "Failed to automatically relaunch as Administrator: $($_.Exception.Message). Please run the script manually using 'Run as Administrator'."
                    Write-Error $errMsg
                    LogTransfer "Elevation" "Fatal Error" $errMsg # Use LogTransfer if defined globally
                    [System.Windows.MessageBox]::Show($errMsg, "Elevation Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                }
                Exit # Exit the non-elevated instance
            } else {
                Write-Host "Running with Administrator privileges." -ForegroundColor Green
                LogTransfer "Elevation" "Success" "Script is running as Administrator" # Use LogTransfer if defined globally

                # Get credentials for the *remote* machine access
                $credential = $null
                try {
                    # Suggest current user, but allow changing it
                    $currentUser = "$env:USERDOMAIN\$env:USERNAME"
                    $credential = Get-Credential -UserName $currentUser -Message "Enter Administrator credentials for the OLD (source) computer '$targetDevice'" -ErrorAction Stop
                } catch {
                    # Handle user cancelling the credential prompt
                    if ($_.Exception.Message -match 'Operation canceled by user') {
                        throw "Credential input cancelled by user. Express mode cannot continue without credentials."
                    } else {
                        throw "Failed to get credentials: $($_.Exception.Message)"
                    }
                }

                if ($credential -eq $null) { # Should be caught by throw above, but double-check
                    throw "Credentials are required for Express mode."
                }

                # Execute the main logic for Express mode
                Execute-ExpressModeLogic -Credential $credential
            }
        }
        'Cancel' {
            Write-Host "Operation Cancelled by user." -ForegroundColor Yellow
        }
        Default {
            # This case should not happen with the current Show-ModeDialog logic
            Write-Error "Invalid operation mode returned: '$operationMode'"
        }
    }
} catch {
    # Catch errors from the main switch statement or credential gathering
    $errorMessage = "FATAL SCRIPT ERROR: $($_.Exception.Message)"
    # Avoid showing duplicate message boxes if Execute-ExpressModeLogic already showed one
    if ($_.FullyQualifiedErrorId -notmatch 'Express Transfer Failed') {
        Write-Error $errorMessage -ErrorAction Continue
        try {
            [System.Windows.MessageBox]::Show($errorMessage, "Fatal Script Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        } catch {
            Write-Warning "Could not display final error message box."
        }
    } else {
         Write-Host "Express mode failed. See previous messages and log file." -ForegroundColor Red
    }
    # Exit with a non-zero code to indicate failure
    Exit 1
} finally {
    # Wait for Restore mode's background job if it was started
    if ($operationMode -eq 'Restore' -and (Get-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue) -ne $null -and $script:updateJob -ne $null) {
        if ($script:updateJob.State -eq 'Running') {
            Write-Host "`n--- Waiting for background Restore updates job (GPUpdate/SCCM)... ---" -ForegroundColor Yellow
            Wait-Job $script:updateJob | Out-Null
        }
        Write-Host "--- Background Job Output (Restore Updates): ---" -ForegroundColor Yellow
        Receive-Job $script:updateJob
        Remove-Job $script:updateJob
        Write-Host "--- End Background Job Output ---" -ForegroundColor Yellow
    } elseif ($operationMode -ne 'Express' -and $operationMode -ne 'Cancel') {
        # Only log this if a mode was actually run (not cancelled)
        Write-Host "`nNo background update job was started for $operationMode mode." -ForegroundColor Gray
    }
}

Write-Host "--- Script Execution Finished ---" -ForegroundColor Magenta
# Pause only if running in console and not in ISE/VSCode terminal
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') {
    Write-Host "Press Enter to exit..." -ForegroundColor Yellow
    Read-Host
}