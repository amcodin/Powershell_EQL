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
        [hashtable]$Templates,
        [switch]$LocalResolve # Add switch to resolve against local environment
    )

    $resolvedPaths = [System.Collections.Generic.List[PSCustomObject]]::new()
    $userName = Split-Path $UserProfilePath -Leaf # Extract username for logging

    Write-Host "Resolving paths based on profile: $UserProfilePath (LocalResolve: $LocalResolve)"

    foreach ($key in $Templates.Keys) {
        $template = $Templates[$key]
        $resolvedPath = $null

        if ($LocalResolve) {
            # Resolve against the *local* machine's environment, substituting the *local* user profile
            $localUserProfile = $env:USERPROFILE
            $resolvedPath = $template.Replace('{USERPROFILE}', $localUserProfile.TrimEnd('\'))
            # No Test-Path here, we just want the target path string
             Write-Host "  Local Target Path for '$key': $resolvedPath"
             $resolvedPaths.Add([PSCustomObject]@{
                Name = $key
                Path = $resolvedPath # This is the *intended local destination*
                Type = $null # Type isn't relevant for destination-only resolution
                SourceKey = $key
            })
        } else {
            # Resolve against the provided (likely remote mapped) path
            $resolvedPath = $template.Replace('{USERPROFILE}', $UserProfilePath.TrimEnd('\'))
            # Check if the resolved path exists on the source
            if (Test-Path -LiteralPath $resolvedPath -ErrorAction SilentlyContinue) {
                $pathType = if (Test-Path -LiteralPath $resolvedPath -PathType Container) { "Folder" } else { "File" }
                Write-Host "  Found Remote Source: '$key' -> $resolvedPath ($pathType)" -ForegroundColor Green
                $resolvedPaths.Add([PSCustomObject]@{
                    Name = $key # Use the template key as the name
                    Path = $resolvedPath # This is the *remote source path*
                    Type = $pathType
                    SourceKey = $key # Keep track of the original template key
                })
            } else {
                Write-Host "  Remote Source Path not found: $resolvedPath (Derived from template key: $key)" -ForegroundColor Yellow
            }
        }
    }
    return $resolvedPaths
}


# Gets standard user folders (Downloads, Desktop, etc.) - FOR LOCAL USE ONLY (GUI)
function Get-UserPaths {
    [CmdletBinding()]
    param () # No parameters needed for local context
    $folderNames = @( "Downloads", "Pictures", "Videos" ) # Removed Desktop/Documents based on previous version
    $result = [System.Collections.Generic.List[PSCustomObject]]::new()
    $sourceBasePath = $env:USERPROFILE
    Write-Host "Getting local user folders from: $sourceBasePath"

    foreach ($folderName in $folderNames) {
        $fullPath = Join-Path -Path $sourceBasePath -ChildPath $folderName
        if (Test-Path -LiteralPath $fullPath -PathType Container) {
            $item = Get-Item -LiteralPath $fullPath -ErrorAction SilentlyContinue
             if ($item) {
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

    $form.AcceptButton = $btnExpress

    Write-Host "Showing mode selection dialog."
    $result = $form.ShowDialog()
    $form.Dispose()
    Write-Host "Mode selection dialog closed with result: $result"

    $selectedMode = switch ($result) {
        ([System.Windows.Forms.DialogResult]::Yes) { 'Backup' }
        ([System.Windows.Forms.DialogResult]::No) { 'Restore' }
        ([System.Windows.Forms.DialogResult]::OK) { 'Express' }
        Default { 'Cancel' } # Includes Cancel or closing the dialog via 'X'
    }

    Write-Host "Determined mode: $selectedMode" -ForegroundColor Cyan

    if ($selectedMode -eq 'Cancel') {
        Write-Host "Operation cancelled by user (dialog closed)." -ForegroundColor Yellow
        if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Read-Host "Press Enter to exit" }
        Exit 0
    }

    # Start background job ONLY for Restore mode
    if ($selectedMode -eq 'Restore') {
        Write-Host "Restore mode selected. Initiating background system updates job..." -ForegroundColor Yellow
        $script:updateJob = Start-Job -Name "BackgroundUpdates" -ScriptBlock {
            # --- Function Definitions INSIDE Job Scope ---
            function Set-GPupdate { Write-Host "JOB: GPUpdate..."; try { $p = Start-Process cmd.exe '/c gpupdate /force' -PassThru -Wait -EA Stop; if ($p.ExitCode -eq 0) { Write-Host "JOB: GPUpdate OK." -FG Green } else { Write-Warning "JOB: GPUpdate ExitCode $($p.ExitCode)." } } catch { Write-Error "JOB: GPUpdate failed: $($_.Exception.Message)" } }
            function Start-ConfigManagerActions { Write-Host "JOB: ConfigMgr Actions..."; $ccmExec="C:\Windows\CCM\ccmexec.exe"; $ns="root\ccm\clientsdk"; $cl="CCM_ClientUtilities"; $m="TriggerSchedule"; $ok=$false; $cimOk=$false; $acts=@(@{ID='{...21}';N='MachPol'},@{ID='{...22}';N='UserPol'},@{ID='{...01}';N='HInv'},@{ID='{...02}';N='SInv'},@{ID='{...113}';N='SUScan'},@{ID='{...101}';N='HInvCol'},@{ID='{...108}';N='SUAssEval'},@{ID='{...102}';N='SInvCol'}); $svc=Get-Service CcmExec -EA SilentlyContinue; if(!$svc){Write-Warning "JOB: CcmExec Svc N/F"; return $false}; if($svc.Status -ne 'Running'){Write-Warning "JOB: CcmExec Svc not running"; return $false}; Write-Host "JOB: CcmExec Svc OK."; Write-Host "JOB: Try CIM..."; $cimMethOk=$true; try { if(Get-CimClass -N $ns -C $cl -EA SilentlyContinue){ foreach($a in $acts){ Write-Host "JOB: . CIM $($a.N)"; try{Invoke-CimMethod -N $ns -C $cl -M $m -Args @{sScheduleID=$a.ID} -EA Stop}catch{Write-Warning "JOB: . CIM $($a.N) failed: $($_.Exception.Message)";$cimMethOk=$false} } if($cimMethOk){$cimOk=$true;$ok=$true;Write-Host "JOB: CIM OK." -FG Green}else{Write-Warning "JOB: CIM failed some."} }else{Write-Warning "JOB: CIM class N/F.";$cimMethOk=$false} }catch{Write-Error "JOB: CIM error: $($_.Exception.Message)";$cimMethOk=$false}; if(!$cimOk){ Write-Host "JOB: Try ccmexec.exe..."; if(Test-Path $ccmExec){ $exeOk=$true; foreach($a in $acts){ Write-Host "JOB: . EXE $($a.N)"; try{$p=Start-Process $ccmExec "-TriggerSchedule $($a.ID)" -NoNewWindow -PassThru -Wait -EA Stop;if($p.ExitCode -ne 0){Write-Warning "JOB: . EXE $($a.N) ExitCode $($p.ExitCode)."}}catch{Write-Warning "JOB: . EXE $($a.N) failed: $($_.Exception.Message)";$exeOk=$false} } if($exeOk){$ok=$true;Write-Host "JOB: EXE OK." -FG Green}else{Write-Warning "JOB: EXE failed some."} }else{Write-Warning "JOB: ccmexec.exe N/F."} }; if($ok){Write-Host "JOB: ConfigMgr Actions OK." -FG Green}else{Write-Warning "JOB: ConfigMgr Actions Failed."}; return $ok }
            Set-GPupdate; Start-ConfigManagerActions; Write-Host "JOB: Background updates finished."
        }
        Write-Host "Background update job started (ID: $($script:updateJob.Id)). Output will be shown after main window closes." -ForegroundColor Yellow
    }

    Write-Host "Exiting Show-ModeDialog function."
    return $selectedMode
}

# Show main window (Backup/Restore specific)
function Show-MainWindow {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Backup', 'Restore')]
        [string]$Mode
    )
    $IsBackup = ($Mode -eq 'Backup')
    $modeString = $Mode
    Write-Host "Entering Show-MainWindow function. Mode: $modeString"

    [xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="User Data Backup/Restore Tool" Width="800" Height="650" WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Location:" Margin="10,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="txtSaveLoc" Width="400" Height="30" Margin="10,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" IsReadOnly="True"/>
        <Button Name="btnBrowse" Content="Browse" Width="60" Height="30" Margin="420,40,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Name="lblMode" Content="" Margin="500,10,10,0" HorizontalAlignment="Right" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Name="lblFreeSpace" Content="Free Space: -" Margin="500,40,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>
        <Label Name="lblRequiredSpace" Content="Required Space: -" Margin="500,65,10,0" HorizontalAlignment="Right" VerticalAlignment="Top"/>
        <Label Name="lblStatus" Content="Ready" Margin="10,0,10,10" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontStyle="Italic"/>
        <Label Content="Files/Folders to Process:" Margin="10,90,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <ListView Name="lvwFiles" Margin="10,120,200,140" SelectionMode="Extended">
             <ListView.View>
                <GridView>
                    <GridViewColumn Width="30"><GridViewColumn.CellTemplate><DataTemplate><CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay}" /></DataTemplate></GridViewColumn.CellTemplate></GridViewColumn>
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="180"/>
                    <GridViewColumn Header="Type" DisplayMemberBinding="{Binding Type}" Width="70"/>
                    <GridViewColumn Header="Path" DisplayMemberBinding="{Binding Path}" Width="280"/>
                </GridView>
            </ListView.View>
        </ListView>
        <StackPanel Margin="0,120,10,0" HorizontalAlignment="Right" Width="180">
            <Button Name="btnAddFile" Content="Add File" Width="120" Height="30" Margin="0,0,0,10"/>
            <Button Name="btnAddFolder" Content="Add Folder" Width="120" Height="30" Margin="0,0,0,10"/>
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
        Write-Host "Parsing XAML..."
        $reader = New-Object System.Xml.XmlNodeReader $XAML; $window = [Windows.Markup.XamlReader]::Load($reader)
        Write-Host "Setting DataContext..."
        $window.DataContext = [PSCustomObject]@{ IsRestoreMode = (-not $IsBackup) }
        Write-Host "Finding controls..."
        $controls = @{}; @('txtSaveLoc','btnBrowse','btnStart','lblMode','lblStatus','lvwFiles','btnAddFile','btnAddFolder','btnRemove','chkNetwork','chkPrinters','prgProgress','txtProgress','btnAddBAUPaths','lblFreeSpace','lblRequiredSpace') | ForEach-Object { $controls[$_] = $window.FindName($_) }

        # --- Space Calculation Script Blocks ---
        $UpdateFreeSpaceLabel = { param($ControlsParam) $location = $ControlsParam.txtSaveLoc.Text; $freeSpaceString = "Free Space: N/A"; if (-not [string]::IsNullOrEmpty($location)) { try { $driveLetter = $null; if ($location -match '^[a-zA-Z]:\\') { $driveLetter = $location.Substring(0, 2) } elseif ($location -match '^\\\\[^\\]+\\[^\\]+') { $freeSpaceString = "Free Space: N/A (UNC)" }; if ($driveLetter) { $driveInfo = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$driveLetter'" -EA SilentlyContinue; if ($driveInfo -and $driveInfo.FreeSpace -ne $null) { $freeSpaceString = "Free Space: $(Format-Bytes $driveInfo.FreeSpace)" } } } catch { Write-Warning "Err Free Space: $($_.Exception.Message)" } }; Write-Host "Free Space: $freeSpaceString"; $window.Dispatcher.InvokeAsync({ $ControlsParam.lblFreeSpace.Content = $freeSpaceString }) | Out-Null }
        $UpdateRequiredSpaceLabel = { param($ControlsParam) $window.Dispatcher.InvokeAsync({ $ControlsParam.lblRequiredSpace.Content = "Required Space: Calculating..." }) | Out-Null; [System.Threading.Tasks.Task]::Run({ $totalSize = 0L; $requiredSpaceString = "Required Space: Error"; try { $items = @($ControlsParam.lvwFiles.ItemsSource); if ($items -ne $null -and $items.Count -gt 0) { $checkedItems = $items | Where-Object { $_.IsSelected }; Write-Host "Calc Req Space for $($checkedItems.Count) items..."; foreach ($item in $checkedItems) { if (Test-Path -LiteralPath $item.Path -EA SilentlyContinue) { try { if ($item.Type -eq 'Folder') { $folderSize = (Get-ChildItem -LiteralPath $item.Path -Recurse -File -Force -EA SilentlyContinue | Measure-Object Length -Sum -EA SilentlyContinue).Sum; if ($folderSize -ne $null) { $totalSize += $folderSize } } else { $fileSize = (Get-Item -LiteralPath $item.Path -Force -EA SilentlyContinue).Length; if ($fileSize -ne $null) { $totalSize += $fileSize } } } catch { Write-Warning "Size calc err '$($item.Path)': $($_.Exception.Message)" } } else { Write-Warning "Path N/F for size: $($item.Path)" } }; $requiredSpaceString = "Required Space: $(Format-Bytes $totalSize)" } else { $requiredSpaceString = "Required Space: 0 B" } } catch { Write-Error "Req space task err: $($_.Exception.Message)" } finally { Write-Host "Req Space: $requiredSpaceString"; $window.Dispatcher.InvokeAsync({ $ControlsParam.lblRequiredSpace.Content = $requiredSpaceString }) | Out-Null } }) | Out-Null }

        # --- Window Initialization ---
        Write-Host "Initializing controls..."
        $controls.lblMode.Content = "Mode: $modeString"; $controls.btnStart.Content = $modeString; $controls.btnAddBAUPaths.IsEnabled = $IsBackup
        if (-not (Test-Path $script:DefaultPath)) { try { New-Item -Path $script:DefaultPath -ItemType Directory -Force -EA Stop | Out-Null } catch { Write-Warning "Could not create default path: $script:DefaultPath."; $script:DefaultPath = $env:USERPROFILE } }; $controls.txtSaveLoc.Text = $script:DefaultPath; & $UpdateFreeSpaceLabel -ControlsParam $controls

        # --- Load initial items ---
        Write-Host "Loading initial items..."
        $itemsList = [System.Collections.Generic.List[PSCustomObject]]::new()
        if ($IsBackup) { $localUserProfile = $env:USERPROFILE; $paths = Resolve-UserPathTemplates -UserProfilePath $localUserProfile -Templates $script:pathTemplates; if ($paths) { $paths | ForEach-Object { $_ | Add-Member NoteProperty IsSelected $true -PassThru; $itemsList.Add($_) } } }
        elseif (Test-Path $script:DefaultPath) { $latestBackup = Get-ChildItem -Path $script:DefaultPath -Directory -Filter "Backup_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($latestBackup) { $controls.txtSaveLoc.Text = $latestBackup.FullName; & $UpdateFreeSpaceLabel -ControlsParam $controls; $logFilePath = Join-Path $latestBackup.FullName "FileList_Backup.csv"; if (Test-Path $logFilePath) { $backupItems = Get-ChildItem -Path $latestBackup.FullName | Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } | Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } else { "Setting" }; Path = $_.FullName; IsSelected = $true } }; $backupItems | ForEach-Object { $itemsList.Add($_) } } else { $controls.lblStatus.Content = "Restore: Backup invalid (no log)." } } else { $controls.lblStatus.Content = "Restore: No backups found." } }
        else { $controls.lblStatus.Content = "Restore: Default path missing." }
        if ($controls['lvwFiles'] -ne $null) { $controls.lvwFiles.ItemsSource = $itemsList; & $UpdateRequiredSpaceLabel -ControlsParam $controls } else { throw "ListView not found." }
        Write-Host "Finished loading initial items."

        # --- Event Handlers ---
        Write-Host "Assigning event handlers..."
        $controls.btnBrowse.Add_Click({ Write-Host "Browse clicked."; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = if ($IsBackup) { "Select save location" } else { "Select backup folder" }; if(Test-Path $controls.txtSaveLoc.Text){ $dialog.SelectedPath = $controls.txtSaveLoc.Text } else { $dialog.SelectedPath = $script:DefaultPath }; $dialog.ShowNewFolderButton = $IsBackup; $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $result = $dialog.ShowDialog($owner); $owner.Dispose(); if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $selectedPath = $dialog.SelectedPath; $controls.txtSaveLoc.Text = $selectedPath; & $UpdateFreeSpaceLabel -ControlsParam $controls; if (-not $IsBackup) { $logFilePath = Join-Path $selectedPath "FileList_Backup.csv"; $newItemsList = [System.Collections.Generic.List[PSCustomObject]]::new(); if (Test-Path $logFilePath) { $backupItems = Get-ChildItem $selectedPath | Where-Object { $_.PSIsContainer -or $_.Name -match '^(Drives\.csv|Printers\.txt)$' } | Where-Object { $_.Name -notmatch '^(FileList_.*\.csv|TransferLog\.csv)$' } | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Type = if ($_.PSIsContainer) { "Folder" } else { "Setting" }; Path = $_.FullName; IsSelected = $true } }; $backupItems | ForEach-Object { $newItemsList.Add($_) }; $controls.lblStatus.Content = "Ready: $selectedPath" } else { $controls.lblStatus.Content = "Selected folder invalid."; [System.Windows.MessageBox]::Show("Missing 'FileList_Backup.csv'.", "Invalid Backup", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) }; $controls.lvwFiles.ItemsSource = $newItemsList; & $UpdateRequiredSpaceLabel -ControlsParam $controls } else { $controls.lblStatus.Content = "Backup location set." } } else { Write-Host "Browse cancelled."} })
        $controls.btnAddFile.Add_Click({ if (!$IsBackup) { return }; Write-Host "Add File clicked."; $dialog = New-Object System.Windows.Forms.OpenFileDialog; $dialog.Title = "Select File(s)"; $dialog.Multiselect = $true; $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $result = $dialog.ShowDialog($owner); $owner.Dispose(); if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]; if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }; $addedCount = 0; foreach ($file in $dialog.FileNames) { if (-not ($currentItems.Path -contains $file)) { $fileKey = "ManualFile_" + ([System.IO.Path]::GetFileNameWithoutExtension($file) -replace '[^a-zA-Z0-9_]','_'); $currentItems.Add([PSCustomObject]@{ Name = $fileKey; Type = "File"; Path = $file; IsSelected = $true; SourceKey = $fileKey }); $addedCount++ } }; if ($addedCount -gt 0) { $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls }; Write-Host "Added $addedCount file(s)." } })
        $controls.btnAddFolder.Add_Click({ if (!$IsBackup) { return }; Write-Host "Add Folder clicked."; $dialog = New-Object System.Windows.Forms.FolderBrowserDialog; $dialog.Description = "Select Folder"; $dialog.ShowNewFolderButton = $false; $owner = New-Object System.Windows.Forms.Form -Property @{ ShowInTaskbar = $false; WindowState = 'Minimized' }; $result = $dialog.ShowDialog($owner); $owner.Dispose(); if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $selectedPath = $dialog.SelectedPath; $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]; if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }; if (-not ($currentItems.Path -contains $selectedPath)) { $folderKey = "ManualFolder_" + ([System.IO.Path]::GetFileName($selectedPath) -replace '[^a-zA-Z0-9_]','_'); $currentItems.Add([PSCustomObject]@{ Name = $folderKey; Type = "Folder"; Path = $selectedPath; IsSelected = $true; SourceKey = $folderKey }); $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls; Write-Host "Added folder: $selectedPath" } else { Write-Host "Skip duplicate folder." } } })
        $controls.btnAddBAUPaths.Add_Click({ if (!$IsBackup) { return }; Write-Host "Add User Folders clicked."; try { $userPaths = Get-UserPaths; if ($userPaths) { $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]; if ($currentItems -eq $null) { $currentItems = [System.Collections.Generic.List[PSCustomObject]]::new(); $controls.lvwFiles.ItemsSource = $currentItems }; $addedCount = 0; foreach ($pathInfo in $userPaths) { if (-not ($currentItems.Path -contains $pathInfo.Path)) { $pathInfo | Add-Member NoteProperty IsSelected $true -PassThru; $currentItems.Add($pathInfo); $addedCount++ } }; if ($addedCount -gt 0) { $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls }; Write-Host "Added $addedCount user folder(s)." } else { Write-Warning "No user folders found." } } catch { Write-Error "Error adding user folders: $($_.Exception.Message)" } })
        $controls.btnRemove.Add_Click({ Write-Host "Remove clicked."; $selectedObjects = @($controls.lvwFiles.SelectedItems); if ($selectedObjects.Count -gt 0) { $currentItems = $controls.lvwFiles.ItemsSource -as [System.Collections.Generic.List[PSCustomObject]]; if ($currentItems -ne $null) { $itemsToRemove = $selectedObjects | ForEach-Object { $_ }; $itemsToRemove | ForEach-Object { $currentItems.Remove($_) } | Out-Null; $controls.lvwFiles.Items.Refresh(); & $UpdateRequiredSpaceLabel -ControlsParam $controls; Write-Host "Removed $($selectedObjects.Count) item(s)." } } else { Write-Host "No items selected."} })
        $controls.lvwFiles.Add_SelectionChanged({ Write-Host "Selection Changed."; & $UpdateRequiredSpaceLabel -ControlsParam $controls })
        $controls.lvwFiles.Add_PreviewMouseLeftButtonUp({ param($s, $e) if ($e.OriginalSource -is [System.Windows.Controls.CheckBox]) { Write-Host "Checkbox clicked."; $window.Dispatcher.BeginInvoke([action]{ & $UpdateRequiredSpaceLabel -ControlsParam $controls }, [System.Windows.Threading.DispatcherPriority]::Background) } })

        # --- Start Button Logic ---
        $controls.btnStart.Add_Click({
            $modeStringLocal = if ($IsBackup) { 'Backup' } else { 'Restore' }; Write-Host "Start clicked ($modeStringLocal)."
            $location = $controls.txtSaveLoc.Text; if ([string]::IsNullOrEmpty($location) -or -not (Test-Path $location -PathType Container)) { [System.Windows.MessageBox]::Show("Select valid directory.", "Location Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return }
            Write-Host "Final space check..."; & $UpdateRequiredSpaceLabel -ControlsParam $controls
            $itemsToProcess = [System.Collections.Generic.List[PSCustomObject]]::new(); $checkedItems = @($controls.lvwFiles.ItemsSource) | Where-Object { $_.IsSelected }; if ($checkedItems) { $checkedItems | ForEach-Object { $itemsToProcess.Add($_) } }
            $doNetwork = $controls.chkNetwork.IsChecked; $doPrinters = $controls.chkPrinters.IsChecked
            Write-Host "Disabling UI..."; $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { $_.Value.IsEnabled = $false } }; $window.Cursor = [System.Windows.Input.Cursors]::Wait
            $uiProgressAction = { param($status, $percent, $details) $window.Dispatcher.InvokeAsync( [action]{ $controls.lblStatus.Content = $status; $controls.txtProgress.Text = $details; if ($percent -ge 0) { $controls.prgProgress.IsIndeterminate = $false; $controls.prgProgress.Value = $percent } elseif ($percent -eq -1) { $controls.prgProgress.IsIndeterminate = $false; $controls.prgProgress.Value = 0 } else { $controls.prgProgress.IsIndeterminate = $true } }, [System.Windows.Threading.DispatcherPriority]::Background ) | Out-Null }

            $operationTask = [System.Threading.Tasks.Task]::Factory.StartNew({
                $success = $false
                try {
                    & $uiProgressAction.Invoke("Initializing...", -2, "Starting...")
                    if ($IsBackup) { Write-Host "BG: Starting Backup..."; $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"; $localUsername = $env:USERNAME -replace '[^a-zA-Z0-9]', '_'; $backupRootPath = Join-Path $location "Backup_${localUsername}_$timestamp"; if ($itemsToProcess.Count -eq 0 -and !$doNetwork -and !$doPrinters) { throw "Nothing selected." }; $success = Invoke-BackupOperation -BackupRootPath $backupRootPath -ItemsToBackup $itemsToProcess -BackupNetworkDrives $doNetwork -BackupPrinters $doPrinters -ProgressAction $uiProgressAction }
                    else { Write-Host "BG: Starting Restore..."; $backupRootPath = $location; $selectedKeysOrFilesForRestore = $itemsToProcess | Select-Object -ExpandProperty Name; $success = Invoke-RestoreOperation -BackupRootPath $backupRootPath -SelectedKeysOrFiles $selectedKeysOrFilesForRestore -RestoreNetworkDrives $doNetwork -RestorePrinters $doPrinters -ProgressAction $uiProgressAction }
                    if ($success) { Write-Host "BG: Success." -FG Green; & $uiProgressAction.Invoke("Completed", 100, "Finished $modeStringLocal."); $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show("Operation completed!", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }) | Out-Null }
                    else { Write-Error "BG: Failed."; & $uiProgressAction.Invoke("Failed", -1, "Errors occurred."); $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show("Operation failed. Check console.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }) | Out-Null }
                } catch { $errorMessage = "BG Task Error: $($_.Exception.Message)"; Write-Error $errorMessage; & $uiProgressAction.Invoke("Failed", -1, $errorMessage); $window.Dispatcher.InvokeAsync({ [System.Windows.MessageBox]::Show($errorMessage, "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }) | Out-Null }
                finally { Write-Host "BG: Re-enabling UI..."; $window.Dispatcher.InvokeAsync({ $controls | ForEach-Object { if ($_.Value -is [System.Windows.Controls.Control]) { if ($_.Key -eq 'btnAddBAUPaths') { $_.Value.IsEnabled = $IsBackup } else { $_.Value.IsEnabled = $true } } }; $window.Cursor = [System.Windows.Input.Cursors]::Arrow; Write-Host "BG: UI re-enabled." }) | Out-Null }
            })
            $operationTask.ContinueWith({ param($task) if ($task.IsFaulted) { $task.Exception.Flatten().InnerExceptions | ForEach-Object { Write-Error "Unhandled Task Ex: $($_.Message)" } } }, [System.Threading.Tasks.TaskContinuationOptions]::OnlyOnFaulted)
        }) # End btnStart.Add_Click

        Write-Host "Showing main window."; $window.ShowDialog() | Out-Null; Write-Host "Main window closed."
    } catch { $errorMessage = "Failed load main window: $($_.Exception.Message)"; Write-Error $errorMessage; try { [System.Windows.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {} }
    finally { Write-Host "Exiting Show-MainWindow function." }
}

# --- Refactored Backup/Restore Functions (Called by UI or Express) ---

function Invoke-BackupOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BackupRootPath,
        [Parameter(Mandatory=$true)] [System.Collections.Generic.List[PSCustomObject]]$ItemsToBackup,
        [Parameter(Mandatory=$true)] [bool]$BackupNetworkDrives,
        [Parameter(Mandatory=$true)] [bool]$BackupPrinters,
        [Parameter(Mandatory=$false)] [scriptblock]$ProgressAction # Optional: $ProgressAction.Invoke($status, $percent, $details)
    )
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    Write-Host "--- Starting LOCAL Backup Operation to '$BackupRootPath' ---"
    try {
        $UpdateProgress.Invoke("Initializing Backup", 0, "Creating backup directory...")
        if (-not (Test-Path $BackupRootPath)) { New-Item -Path $BackupRootPath -ItemType Directory -Force -EA Stop | Out-Null; Write-Host "Backup directory created." }
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"; "OriginalFullPath,BackupRelativePath,SourceKey" | Set-Content -Path $csvLogPath -Encoding UTF8
        if (-not $ItemsToBackup -or $ItemsToBackup.Count -eq 0) { throw "No items specified for backup." }
        $networkDriveCount = $(if ($BackupNetworkDrives) { 1 } else { 0 }); $printerCount = $(if ($BackupPrinters) { 1 } else { 0 })
        $totalItems = $ItemsToBackup.Count + $networkDriveCount + $printerCount; $currentItemIndex = 0
        foreach ($item in $ItemsToBackup) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 90) } else { 0 } # File copy up to 90%
            $statusMessage = "Backing up Item $currentItemIndex of $totalItems"; $UpdateProgress.Invoke($statusMessage, $percentComplete, "Processing: $($item.Name)")
            $sourcePath = $item.Path; $sourceKey = $item.SourceKey
            if (-not (Test-Path -LiteralPath $sourcePath)) { Write-Warning "SKIP: $sourcePath not found."; "`"$sourcePath`",`"SKIPPED_NOT_FOUND`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8; continue }
            $backupItemRootName = $sourceKey -replace '[^a-zA-Z0-9_]', '_'
            try {
                if ($item.Type -eq "Folder") {
                    $sourceFolderInfo = Get-Item -LiteralPath $sourcePath
                    Get-ChildItem -LiteralPath $sourcePath -Recurse -File -Force -EA Stop | ForEach-Object {
                        $originalFileFullPath = $_.FullName; $relativeFilePath = $originalFileFullPath.Substring($sourceFolderInfo.FullName.Length).TrimStart('\'); $backupRelativePath = Join-Path $backupItemRootName $relativeFilePath
                        $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath; $targetBackupDir = Split-Path $targetBackupPath -Parent
                        if (-not (Test-Path $targetBackupDir)) { New-Item -Path $targetBackupDir -ItemType Directory -Force -EA Stop | Out-Null }
                        Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -EA Stop
                        "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                        $UpdateProgress.Invoke($statusMessage, $percentComplete, "Copied: $($_.Name)")
                    }
                } else { # File
                    $originalFileFullPath = $sourcePath; $backupRelativePath = Join-Path $backupItemRootName (Split-Path $originalFileFullPath -Leaf); $targetBackupPath = Join-Path $BackupRootPath $backupRelativePath; $targetBackupDir = Split-Path $targetBackupPath -Parent
                    if (-not (Test-Path $targetBackupDir)) { New-Item -Path $targetBackupDir -ItemType Directory -Force -EA Stop | Out-Null }
                    Copy-Item -LiteralPath $originalFileFullPath -Destination $targetBackupPath -Force -EA Stop
                    "`"$originalFileFullPath`",`"$backupRelativePath`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8
                }
            } catch { $errMsg = $_.Exception.Message -replace '"','""'; Write-Warning "ERROR copying '$($item.Name)': $errMsg"; "`"$sourcePath`",`"ERROR_COPY: $errMsg`",`"$sourceKey`"" | Add-Content -Path $csvLogPath -Encoding UTF8 }
        }
        if ($BackupNetworkDrives) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 95) } else { 0 }; $UpdateProgress.Invoke("Backing up Settings", $percentComplete, "Network Drives...")
            try { Get-WmiObject -Class Win32_MappedLogicalDisk -EA Stop | Select-Object Name, ProviderName | Export-Csv -Path (Join-Path $BackupRootPath "Drives.csv") -NoTypeInformation -Encoding UTF8 -EA Stop; Write-Host "Drives backed up." }
            catch { Write-Warning "Failed to backup drives: $($_.Exception.Message)" }
        }
        if ($BackupPrinters) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }; $UpdateProgress.Invoke("Backing up Settings", $percentComplete, "Printers...")
            try { Get-WmiObject -Class Win32_Printer -Filter "Local = False AND Network = True" -EA Stop | Select-Object -ExpandProperty Name | Set-Content -Path (Join-Path $BackupRootPath "Printers.txt") -Encoding UTF8 -EA Stop; Write-Host "Printers backed up." }
            catch { Write-Warning "Failed to backup printers: $($_.Exception.Message)" }
        }
        $UpdateProgress.Invoke("Backup Complete", 100, "Successfully backed up to: $BackupRootPath"); Write-Host "--- LOCAL Backup Operation Finished ---"; return $true
    } catch { $errorMessage = "LOCAL Backup Failed: $($_.Exception.Message)"; Write-Error $errorMessage; $UpdateProgress.Invoke("Backup Failed", -1, $errorMessage); return $false }
}

function Invoke-RestoreOperation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$BackupRootPath,
        [Parameter(Mandatory=$true)] [string[]]$SelectedKeysOrFiles, # Names from the ListView selection
        [Parameter(Mandatory=$true)] [bool]$RestoreNetworkDrives,
        [Parameter(Mandatory=$true)] [bool]$RestorePrinters,
        [Parameter(Mandatory=$false)] [scriptblock]$ProgressAction
    )
    $UpdateProgress = { if ($ProgressAction) { $ProgressAction.Invoke($args[0], $args[1], $args[2]) } else { Write-Host "$($args[0]) ($($args[1])%) - $($args[2])" } }
    Write-Host "--- Starting LOCAL Restore Operation from '$BackupRootPath' ---"
    try {
        $UpdateProgress.Invoke("Initializing Restore", 0, "Checking backup contents...")
        $csvLogPath = Join-Path $BackupRootPath "FileList_Backup.csv"; $drivesCsvPath = Join-Path $BackupRootPath "Drives.csv"; $printersTxtPath = Join-Path $BackupRootPath "Printers.txt"
        if (-not (Test-Path $csvLogPath -PathType Leaf)) { throw "Backup log file 'FileList_Backup.csv' not found." }
        $backupLog = Import-Csv -Path $csvLogPath -Encoding UTF8; if (-not $backupLog) { throw "Backup log is empty." }
        if (-not ($backupLog | Get-Member -Name SourceKey)) { throw "Backup log missing 'SourceKey' column." }
        $logEntriesToRestore = $backupLog | Where-Object { $_.SourceKey -in $SelectedKeysOrFiles -and $_.BackupRelativePath -notmatch '^(SKIPPED|ERROR)_' }
        if (-not $logEntriesToRestore) { Write-Warning "No valid file/folder entries in log match selection." }
        $networkDriveCount = $(if ($RestoreNetworkDrives -and ($SelectedKeysOrFiles -contains "Drives.csv")) { 1 } else { 0 })
        $printerCount = $(if ($RestorePrinters -and ($SelectedKeysOrFiles -contains "Printers.txt")) { 1 } else { 0 })
        $totalItems = $logEntriesToRestore.Count + $networkDriveCount + $printerCount; $currentItemIndex = 0
        if ($logEntriesToRestore.Count -gt 0) {
            Write-Host "Restoring $($logEntriesToRestore.Count) file/folder entries..."
            foreach ($entry in $logEntriesToRestore) {
                $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 90) } else { 0 } # Restore up to 90%
                $statusMessage = "Restoring Item $currentItemIndex of $totalItems"; $originalFileFullPath = $entry.OriginalFullPath; $sourceBackupPath = Join-Path $BackupRootPath $entry.BackupRelativePath
                $UpdateProgress.Invoke($statusMessage, $percentComplete, "Restoring: $(Split-Path $originalFileFullPath -Leaf)")
                if (Test-Path -LiteralPath $sourceBackupPath -PathType Leaf) {
                    try { $targetRestoreDir = Split-Path $originalFileFullPath -Parent; if (-not (Test-Path $targetRestoreDir)) { New-Item -Path $targetRestoreDir -ItemType Directory -Force -EA Stop | Out-Null }; Copy-Item -LiteralPath $sourceBackupPath -Destination $originalFileFullPath -Force -EA Stop }
                    catch { Write-Warning "Failed to restore '$originalFileFullPath': $($_.Exception.Message)" }
                } else { Write-Warning "SKIP: Source '$sourceBackupPath' not found." }
            }
        }
        if ($RestoreNetworkDrives -and ($SelectedKeysOrFiles -contains "Drives.csv")) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 95) } else { 0 }; $UpdateProgress.Invoke("Restoring Settings", $percentComplete, "Network Drives...")
            if (Test-Path $drivesCsvPath) {
                try { Import-Csv $drivesCsvPath | ForEach-Object { $driveLetter = $_.Name.TrimEnd(':'); $networkPath = $_.ProviderName; if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') { if (-not (Test-Path -LiteralPath "$($driveLetter):")) { try { Write-Host "  Mapping $driveLetter -> $networkPath"; New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -EA Stop } catch { Write-Warning "  Failed map $driveLetter`: $($_.Exception.Message)" } } else { Write-Host "  Skip $driveLetter`: exists." } } else { Write-Warning "  Skip invalid map: $($_.Name)" } } }
                catch { Write-Warning "Error restoring drives: $($_.Exception.Message)" }
            } else { Write-Warning "Drives.csv selected but not found." }
        }
        if ($RestorePrinters -and ($SelectedKeysOrFiles -contains "Printers.txt")) {
            $currentItemIndex++; $percentComplete = if ($totalItems -gt 0) { [int](($currentItemIndex / $totalItems) * 100) } else { 0 }; $UpdateProgress.Invoke("Restoring Settings", $percentComplete, "Printers...")
            if (Test-Path $printersTxtPath) {
                try { $wsNet = New-Object -ComObject WScript.Network; Get-Content $printersTxtPath | ForEach-Object { $printerPath = $_.Trim(); if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') { try { Write-Host "  Adding printer: $printerPath"; $wsNet.AddWindowsPrinterConnection($printerPath) } catch { Write-Warning "  Failed add printer '$printerPath': $($_.Exception.Message)" } } else { Write-Warning "  Skip invalid printer line: '$_'" } } }
                catch { Write-Warning "Error restoring printers: $($_.Exception.Message)" }
            } else { Write-Warning "Printers.txt selected but not found." }
        }
        $UpdateProgress.Invoke("Restore Complete", 100, "Successfully restored from: $BackupRootPath"); Write-Host "--- LOCAL Restore Operation Finished ---"; return $true
    } catch { $errorMessage = "LOCAL Restore Failed: $($_.Exception.Message)"; Write-Error $errorMessage; $UpdateProgress.Invoke("Restore Failed", -1, $errorMessage); return $false }
}

# --- LOCAL Action Functions (Called by Express Mode) ---
function Set-LocalGPUpdate {
    Write-Host "LOCAL ACTION: Initiating Group Policy update..." -ForegroundColor Cyan
    try {
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c gpupdate /force" -PassThru -Wait -ErrorAction Stop
        if ($process.ExitCode -eq 0) { Write-Host "LOCAL ACTION: Group Policy update completed successfully." -ForegroundColor Green }
        else { Write-Warning "LOCAL ACTION: Group Policy update process finished with exit code: $($process.ExitCode)." }
    } catch { Write-Error "LOCAL ACTION: Failed to start GPUpdate process: $($_.Exception.Message)" }
    Write-Host "LOCAL ACTION: Exiting Set-LocalGPUpdate function."
}

function Start-LocalConfigManagerActions {
     param()
     Write-Host "LOCAL ACTION: Entering Start-LocalConfigManagerActions function."
     $ccmExecPath = "C:\Windows\CCM\ccmexec.exe"; $clientSDKNamespace = "root\ccm\clientsdk"; $clientClassName = "CCM_ClientUtilities"; $scheduleMethodName = "TriggerSchedule"; $overallSuccess = $false; $cimAttemptedAndSucceeded = $false
     $scheduleActions = @( @{ID = '{00000000-0000-0000-0000-000000000021}'; Name = 'Machine Policy Retrieval & Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000022}'; Name = 'User Policy Retrieval & Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000001}'; Name = 'Hardware Inventory Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000002}'; Name = 'Software Inventory Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000113}'; Name = 'Software Updates Scan Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000101}'; Name = 'Hardware Inventory Collection Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000108}'; Name = 'Software Updates Assignments Evaluation Cycle'}, @{ID = '{00000000-0000-0000-0000-000000000102}'; Name = 'Software Inventory Collection Cycle'} )
     Write-Host "LOCAL ACTION: Defined $($scheduleActions.Count) CM actions to trigger."
     $ccmService = Get-Service -Name CcmExec -ErrorAction SilentlyContinue
     if (-not $ccmService) { Write-Warning "LOCAL ACTION: CM service (CcmExec) not found. Skipping."; return $false }
     elseif ($ccmService.Status -ne 'Running') { Write-Warning "LOCAL ACTION: CM service (CcmExec) is not running (Status: $($ccmService.Status)). Skipping."; return $false }
     else { Write-Host "LOCAL ACTION: CM service (CcmExec) found and running." }
     Write-Host "LOCAL ACTION: Attempting Method 1: Triggering via CIM ($clientSDKNamespace -> $clientClassName)..."
     $cimMethodSuccess = $true
     try {
         if (Get-CimClass -Namespace $clientSDKNamespace -ClassName $clientClassName -ErrorAction SilentlyContinue) {
              Write-Host "LOCAL ACTION: CIM Class found."
              foreach ($action in $scheduleActions) {
                 Write-Host "LOCAL ACTION:   Triggering $($action.Name) (ID: $($action.ID)) via CIM."
                 try { Invoke-CimMethod -Namespace $clientSDKNamespace -ClassName $clientClassName -MethodName $scheduleMethodName -Arguments @{sScheduleID = $action.ID} -ErrorAction Stop; Write-Host "LOCAL ACTION:     $($action.Name) triggered successfully via CIM." }
                 catch { Write-Warning "LOCAL ACTION:     Failed to trigger $($action.Name) via CIM: $($_.Exception.Message)"; $cimMethodSuccess = $false }
              }
              if ($cimMethodSuccess) { $cimAttemptedAndSucceeded = $true; $overallSuccess = $true; Write-Host "LOCAL ACTION: All actions successfully triggered via CIM." -ForegroundColor Green }
              else { Write-Warning "LOCAL ACTION: One or more actions failed to trigger via CIM." }
         } else { Write-Warning "LOCAL ACTION: CIM Class '$clientClassName' not found. Cannot use CIM method."; $cimMethodSuccess = $false }
     } catch { Write-Error "LOCAL ACTION: Unexpected error during CIM attempt: $($_.Exception.Message)"; $cimMethodSuccess = $false }
     if (-not $cimAttemptedAndSucceeded) {
         Write-Host "LOCAL ACTION: CIM failed/unavailable. Attempting Method 2: Fallback via ccmexec.exe..."
         if (Test-Path -Path $ccmExecPath -PathType Leaf) {
             Write-Host "LOCAL ACTION: Found $ccmExecPath."
             $execMethodSuccess = $true
             foreach ($action in $scheduleActions) {
                 Write-Host "LOCAL ACTION:   Triggering $($action.Name) (ID: $($action.ID)) via ccmexec.exe."
                 try { $process = Start-Process -FilePath $ccmExecPath -ArgumentList "-TriggerSchedule $($action.ID)" -NoNewWindow -PassThru -Wait -ErrorAction Stop; if ($process.ExitCode -ne 0) { Write-Warning "LOCAL ACTION:     $($action.Name) via ccmexec.exe finished with exit code $($process.ExitCode)." } else { Write-Host "LOCAL ACTION:     $($action.Name) triggered via ccmexec.exe (Exit Code 0)." } }
                 catch { Write-Warning "LOCAL ACTION:     Failed to execute ccmexec.exe for $($action.Name): $($_.Exception.Message)"; $execMethodSuccess = $false }
             }
             if ($execMethodSuccess) { $overallSuccess = $true; Write-Host "LOCAL ACTION: Finished attempting actions via ccmexec.exe." -ForegroundColor Green }
             else { Write-Warning "LOCAL ACTION: One or more actions failed to execute via ccmexec.exe." }
         } else { Write-Warning "LOCAL ACTION: Fallback executable not found at $ccmExecPath." }
     }
     if ($overallSuccess) { Write-Host "LOCAL ACTION: CM actions attempt finished successfully." -ForegroundColor Green }
     else { Write-Warning "LOCAL ACTION: CM actions attempt finished, but could not be confirmed as fully successful." }
     Write-Host "LOCAL ACTION: Exiting Start-LocalConfigManagerActions function."; return $overallSuccess
}


# --- Express Mode Function ---
function Execute-ExpressModeLogic {
    param(
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Write-Host "Executing Express Mode Logic..."
    $targetDevice = $null
    $mappedDriveLetter = "X"
    $localBackupBaseDir = "C:\LocalData" # Base for the backup copy
    $transferSuccess = $true # Assume success initially
    $transferLog = [System.Collections.Generic.List[string]]::new()
    $transferLog.Add("Timestamp,Action,Status,Details")
    $logFilePath = $null
    $finalBackupDir = $null # Initialize backup dir variable

    Function LogTransfer ($Action, $Status, $Details) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $safeDetails = $Details -replace '"', '""'
        $logEntry = """$timestamp"",""$Action"",""$Status"",""$safeDetails"""
        $transferLog.Add($logEntry)
        Write-Host "LOG: $Action - $Status - $Details"
    }

    try {
        # --- 1. Get Target Device ---
        $targetDevice = Read-Host "Enter the IP address or Hostname of the target remote device"
        if ([string]::IsNullOrWhiteSpace($targetDevice)) { throw "Target device cannot be empty." }
        LogTransfer "Input" "Info" "Target device specified: $targetDevice"

        # --- 2. Map Network Drive ---
        $uncPath = "\\$targetDevice\c$"
        Write-Host "Attempting to map $uncPath to $mappedDriveLetter`: " -NoNewline
        LogTransfer "Map Drive" "Attempt" "Mapping $uncPath to $mappedDriveLetter`:"
        if (Test-Path "${mappedDriveLetter}:") { Write-Warning "Drive $mappedDriveLetter`: exists. Removing..."; Remove-PSDrive -Name $mappedDriveLetter -Force -EA SilentlyContinue }
        New-PSDrive -Name $mappedDriveLetter -PSProvider FileSystem -Root $uncPath -Credential $Credential -ErrorAction Stop | Out-Null
        Write-Host "Successfully mapped $uncPath to $mappedDriveLetter`:" -ForegroundColor Green
        LogTransfer "Map Drive" "Success" "Successfully mapped $uncPath to $mappedDriveLetter`:"

        # --- 3. Get Remote Logged-on Username ---
        Write-Host "Attempting to identify logged-on user on '$targetDevice'..."
        LogTransfer "Get Remote User" "Attempt" "Querying Win32_ComputerSystem on $targetDevice"
        $remoteUsername = $null
        try {
            $remoteResult = Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock { (Get-CimInstance -ClassName Win32_ComputerSystem).UserName } -ErrorAction Stop
            if ($remoteResult -and -not [string]::IsNullOrWhiteSpace($remoteResult)) { $remoteUsername = ($remoteResult -split '\\')[-1]; Write-Host "Identified remote user: '$remoteUsername'" -ForegroundColor Green; LogTransfer "Get Remote User" "Success" "Identified remote user: $remoteUsername" }
            else { throw "Could not retrieve username from Win32_ComputerSystem or it was empty." }
        } catch {
            Write-Warning "Failed auto-detect remote user: $($_.Exception.Message)"; LogTransfer "Get Remote User" "Warning" "Failed WMI/Invoke: $($_.Exception.Message)"
            $remoteUsername = Read-Host "Could not auto-detect remote user. Please enter the Windows username logged into '$targetDevice'"; if ([string]::IsNullOrWhiteSpace($remoteUsername)) { throw "Remote username required." }; LogTransfer "Get Remote User" "Manual Input" "User provided: $remoteUsername"
        }

        # --- 4. Define Remote Paths and Local Backup Destination ---
        $remoteUserProfile = "$mappedDriveLetter`:\Users\$remoteUsername"; Write-Host "Remote user profile path: $remoteUserProfile"; LogTransfer "Path Setup" "Info" "Remote profile path: $remoteUserProfile"
        if (-not (Test-Path -LiteralPath $remoteUserProfile)) { throw "Remote user profile path '$remoteUserProfile' not found." }

        # Create the final local backup directory FIRST
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"; $sanitizedTarget = $targetDevice -replace '[^a-zA-Z0-9_.-]','_'
        $finalBackupDir = Join-Path $localBackupBaseDir "TransferBackup_${remoteUsername}_from_${sanitizedTarget}_$timestamp"
        Write-Host "Creating local backup directory: $finalBackupDir"; LogTransfer "Setup" "Info" "Creating backup dir: $finalBackupDir"; New-Item -Path $finalBackupDir -ItemType Directory -Force -EA Stop | Out-Null

        # Resolve *remote* source paths using templates
        Write-Host "Identifying remote files/folders for user '$remoteUsername' (Templates Only)..."; LogTransfer "Gather Paths" "Attempt" "Resolving templates for $remoteUserProfile"
        $remotePathsToTransfer = Resolve-UserPathTemplates -UserProfilePath $remoteUserProfile -Templates $script:pathTemplates -LocalResolve:$false # Ensure LocalResolve is false
        LogTransfer "Gather Paths" "Info" "Found $($remotePathsToTransfer.Count) template items on remote source."
        if ($remotePathsToTransfer.Count -eq 0) { Write-Warning "No template files/folders found for '$remoteUsername' on '$targetDevice'."; LogTransfer "Gather Paths" "Warning" "No template paths resolved on remote." }

        # --- 5. Transfer Files/Folders -> LOCAL BACKUP FOLDER ---
        Write-Host "Starting transfer from '$targetDevice' to LOCAL BACKUP '$finalBackupDir'..."
        $totalItems = $remotePathsToTransfer.Count; $currentItem = 0; $errorsDuringBackupTransfer = $false
        $successfullyBackedUpItems = [System.Collections.Generic.List[PSCustomObject]]::new() # Track items successfully copied to backup

        if ($totalItems -gt 0) {
            foreach ($item in $remotePathsToTransfer) {
                $currentItem++; $progress = [int](($currentItem / $totalItems) * 50) # Transfer to backup is ~50%
                Write-Progress -Activity "Transferring to Backup" -Status "[$progress%]: Copying $($item.Name)" -PercentComplete $progress
                $sourcePath = $item.Path # Remote source (e.g., X:\Users\...)
                # Construct destination path within the *local backup folder*
                $backupDestRelativePath = $item.SourceKey -replace '[^a-zA-Z0-9_]', '_'; if ($item.Type -eq 'File') { $backupDestRelativePath = Join-Path $backupDestRelativePath (Split-Path $sourcePath -Leaf) }
                $backupDestinationPath = Join-Path $finalBackupDir $backupDestRelativePath
                Write-Host "  Copying (Remote -> Backup): '$($item.Name)' from '$sourcePath' to '$backupDestinationPath'"; LogTransfer "Transfer to Backup" "Attempt" "Copying $($item.SourceKey) from $sourcePath"
                try {
                    $backupDestDir = Split-Path $backupDestinationPath -Parent; if (-not (Test-Path $backupDestDir)) { New-Item -Path $backupDestDir -ItemType Directory -Force -EA Stop | Out-Null }
                    Copy-Item -Path $sourcePath -Destination $backupDestinationPath -Recurse:($item.Type -eq 'Folder') -Force -EA Stop
                    LogTransfer "Transfer to Backup" "Success" "Copied $($item.SourceKey) to $backupDestinationPath"
                    # Add item details needed for the next step (restoring to local profile)
                    $successfullyBackedUpItems.Add([PSCustomObject]@{ SourceKey = $item.SourceKey; BackupPath = $backupDestinationPath; Type = $item.Type })
                } catch { Write-Warning "  Failed copy (Remote -> Backup) '$($item.Name)': $($_.Exception.Message)"; LogTransfer "Transfer to Backup" "Error" "Failed copy $($item.SourceKey): $($_.Exception.Message)"; $errorsDuringBackupTransfer = $true }
            }
            Write-Progress -Activity "Transferring to Backup" -Completed
            if ($errorsDuringBackupTransfer) { Write-Warning "Errors occurred during transfer to backup folder." } else { Write-Host "Transfer to backup folder completed." -ForegroundColor Green }
        } else { Write-Host "Skipping transfer to backup - no template items found on remote."; Write-Progress -Activity "Transferring to Backup" -Status "[50%]: No files" -PercentComplete 50 }

        # --- 6. Transfer Settings (Drives/Printers) -> LOCAL BACKUP FOLDER ---
        Write-Progress -Activity "Capturing Settings" -Status "[55%]: Drives..." -PercentComplete 55; Write-Host "Getting mapped drives from $targetDevice"; LogTransfer "Get Drives" "Attempt" "Querying remote drives"
        $drivesCsvBackupPath = Join-Path $finalBackupDir "Drives.csv" # Save directly to backup folder
        try { Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock { Get-WmiObject -Class Win32_MappedLogicalDisk -EA Stop | Select-Object Name, ProviderName } -EA Stop | Export-Csv -Path $drivesCsvBackupPath -NoTypeInformation -Encoding UTF8 -EA Stop; Write-Host "Saved remote drives list to backup." -FG Green; LogTransfer "Get Drives" "Success" "Saved to $drivesCsvBackupPath" }
        catch { Write-Warning "Failed get drives from ${targetDevice}: $($_.Exception.Message)"; LogTransfer "Get Drives" "Error" "Failed: $($_.Exception.Message)" }
        Write-Progress -Activity "Capturing Settings" -Status "[60%]: Printers..." -PercentComplete 60; Write-Host "Getting network printers from $targetDevice"; LogTransfer "Get Printers" "Attempt" "Querying remote printers"
        $printersTxtBackupPath = Join-Path $finalBackupDir "Printers.txt" # Save directly to backup folder
        try { Invoke-Command -ComputerName $targetDevice -Credential $Credential -ScriptBlock { Get-WmiObject -Class Win32_Printer -Filter "Local=False AND Network=True" -EA Stop | Select-Object -ExpandProperty Name } -EA Stop | Set-Content -Path $printersTxtBackupPath -Encoding UTF8 -EA Stop; Write-Host "Saved remote printers list to backup." -FG Green; LogTransfer "Get Printers" "Success" "Saved to $printersTxtBackupPath" }
        catch { Write-Warning "Failed get printers from ${targetDevice}: $($_.Exception.Message)"; LogTransfer "Get Printers" "Error" "Failed: $($_.Exception.Message)" }
        Write-Progress -Activity "Capturing Settings" -Completed

        # --- 7. Restore Files/Folders -> LOCAL PROFILE ---
        Write-Host "Starting restore from LOCAL BACKUP '$finalBackupDir' to LOCAL PROFILE..."
        $totalRestoreItems = $successfullyBackedUpItems.Count; $currentRestoreItem = 0; $errorsDuringRestore = $false
        # Get local destination paths based on SourceKey
        $localDestinationMap = @{}
        Resolve-UserPathTemplates -UserProfilePath $env:USERPROFILE -Templates $script:pathTemplates -LocalResolve | ForEach-Object { $localDestinationMap[$_.SourceKey] = $_.Path }

        if ($totalRestoreItems -gt 0) {
            foreach ($item in $successfullyBackedUpItems) {
                $currentRestoreItem++; $progress = [int](($currentRestoreItem / $totalRestoreItems) * 30) + 60 # Restore is 60% -> 90%
                Write-Progress -Activity "Restoring to Local Profile" -Status "[$progress%]: Restoring $($item.SourceKey)" -PercentComplete $progress
                $sourceBackupPath = $item.BackupPath # Source is now the item in our backup folder
                $localDestinationPath = $localDestinationMap[$item.SourceKey] # Target is the actual local profile path

                if (-not $localDestinationPath) { Write-Warning "  Cannot determine local destination for SourceKey '$($item.SourceKey)'. Skipping restore."; LogTransfer "Restore to Profile" "Warning" "No local dest for $($item.SourceKey)"; continue }
                Write-Host "  Copying (Backup -> Local): '$($item.SourceKey)' from '$sourceBackupPath' to '$localDestinationPath'"; LogTransfer "Restore to Profile" "Attempt" "Copying $($item.SourceKey) from backup to $localDestinationPath"
                try {
                    $localDestDir = Split-Path $localDestinationPath -Parent; if (-not (Test-Path $localDestDir)) { New-Item -Path $localDestDir -ItemType Directory -Force -EA Stop | Out-Null }
                    Copy-Item -Path $sourceBackupPath -Destination $localDestinationPath -Recurse:($item.Type -eq 'Folder') -Force -EA Stop
                    LogTransfer "Restore to Profile" "Success" "Restored $($item.SourceKey) to $localDestinationPath"
                } catch { Write-Warning "  Failed copy (Backup -> Local) '$($item.SourceKey)': $($_.Exception.Message)"; LogTransfer "Restore to Profile" "Error" "Failed restore $($item.SourceKey): $($_.Exception.Message)"; $errorsDuringRestore = $true }
            }
            Write-Progress -Activity "Restoring to Local Profile" -Completed
            if ($errorsDuringRestore) { Write-Warning "Errors occurred during restore to local profile." } else { Write-Host "Restore to local profile completed." -ForegroundColor Green }
        } else { Write-Host "Skipping restore to profile - no items successfully backed up."; Write-Progress -Activity "Restoring to Local Profile" -Status "[90%]: No files" -PercentComplete 90 }

        # --- 8. Restore Settings Locally ---
        Write-Progress -Activity "Restoring Settings Locally" -Status "[90%]: Drives..." -PercentComplete 90; Write-Host "Restoring Network Drives Locally..."; LogTransfer "Restore Drives" "Attempt" "Reading $drivesCsvBackupPath"
        if (Test-Path $drivesCsvBackupPath) {
            try { Import-Csv $drivesCsvBackupPath | ForEach-Object { $driveLetter = $_.Name.TrimEnd(':'); $networkPath = $_.ProviderName; if ($driveLetter -match '^[A-Z]$' -and $networkPath -match '^\\\\') { if (-not (Test-Path -LiteralPath "$($driveLetter):")) { try { Write-Host "  Mapping $driveLetter -> $networkPath"; New-PSDrive -Name $driveLetter -PSProvider FileSystem -Root $networkPath -Persist -Scope Global -EA Stop; LogTransfer "Restore Drives" "Success" "Mapped $driveLetter to $networkPath" } catch { Write-Warning "  Failed map $driveLetter`: $($_.Exception.Message)"; LogTransfer "Restore Drives" "Error" "Failed map $driveLetter`: $($_.Exception.Message)" } } else { Write-Host "  Skip $driveLetter`: exists."; LogTransfer "Restore Drives" "Skipped" "$driveLetter exists" } } else { Write-Warning "  Skip invalid map: $($_.Name)"; LogTransfer "Restore Drives" "Skipped" "Invalid line: $($_.Name)" } } }
            catch { Write-Warning "Error restoring drives: $($_.Exception.Message)"; LogTransfer "Restore Drives" "Error" "Processing CSV failed: $($_.Exception.Message)" }
        } else { Write-Warning "Drives.csv not found in backup folder '$finalBackupDir'."; LogTransfer "Restore Drives" "Skipped" "Drives.csv not found" }
        Write-Progress -Activity "Restoring Settings Locally" -Status "[95%]: Printers..." -PercentComplete 95; Write-Host "Restoring Network Printers Locally..."; LogTransfer "Restore Printers" "Attempt" "Reading $printersTxtBackupPath"
        if (Test-Path $printersTxtBackupPath) {
            try { $wsNet = New-Object -ComObject WScript.Network; Get-Content $printersTxtBackupPath | ForEach-Object { $printerPath = $_.Trim(); if (-not ([string]::IsNullOrWhiteSpace($printerPath)) -and $printerPath -match '^\\\\') { try { Write-Host "  Adding printer: $printerPath"; $wsNet.AddWindowsPrinterConnection($printerPath); LogTransfer "Restore Printers" "Success" "Added $printerPath" } catch { Write-Warning "  Failed add printer '$printerPath': $($_.Exception.Message)"; LogTransfer "Restore Printers" "Error" "Failed add $printerPath: $($_.Exception.Message)" } } else { Write-Warning "  Skip invalid printer line: '$_'"; LogTransfer "Restore Printers" "Skipped" "Invalid line: $_" } } }
            catch { Write-Warning "Error restoring printers: $($_.Exception.Message)"; LogTransfer "Restore Printers" "Error" "Processing TXT failed: $($_.Exception.Message)" }
        } else { Write-Warning "Printers.txt not found in backup folder '$finalBackupDir'."; LogTransfer "Restore Printers" "Skipped" "Printers.txt not found" }
        Write-Progress -Activity "Restoring Settings Locally" -Completed

        # --- 9. Run Local Actions ---
        Write-Progress -Activity "Running Local Actions" -Status "[98%]: GPUpdate/ConfigMgr..." -PercentComplete 98
        LogTransfer "Local Actions" "Attempt" "Running GPUpdate and ConfigMgr Actions Locally"
        Set-LocalGPUpdate
        Start-LocalConfigManagerActions
        LogTransfer "Local Actions" "Info" "Finished running local actions"
        Write-Progress -Activity "Running Local Actions" -Completed

        # --- 10. Final Log Saving ---
        $logFilePath = Join-Path $finalBackupDir "TransferLog.csv" # Log saved in backup dir
        try { $transferLog | Set-Content -Path $logFilePath -Encoding UTF8 -Force; Write-Host "Transfer log saved: $logFilePath" }
        catch { Write-Error "CRITICAL: Failed save log: $($_.Exception.Message)"; LogTransfer "Save Log" "Fatal Error" "Failed save log: $($_.Exception.Message)" }
        Write-Progress -Activity "Express Transfer Complete" -Status "[100%]: Finished." -PercentComplete 100; Write-Host "--- Express Transfer Operation Finished ---"

    } catch { $errorMessage = "Express Transfer Failed: $($_.Exception.Message)"; Write-Error $errorMessage; LogTransfer "Overall Status" "Fatal Error" $errorMessage; $transferSuccess = $false }
    finally {
        # --- 11. Cleanup ---
        if (Test-Path "${mappedDriveLetter}:") { Write-Host "Removing mapped drive $mappedDriveLetter`:" -NoNewline; try { Remove-PSDrive -Name $mappedDriveLetter -Force -EA Stop; Write-Host " OK." -FG Green; LogTransfer "Cleanup" "Success" "Removed drive $mappedDriveLetter`:" } catch { Write-Warning " FAILED remove drive $mappedDriveLetter`: $($_.Exception.Message)"; LogTransfer "Cleanup" "Error" "Failed remove drive $mappedDriveLetter`: $($_.Exception.Message)" } }
        # DO NOT remove $finalBackupDir
        $finalLogPathMessage = "Log saving failed"; if (-not ([string]::IsNullOrEmpty($logFilePath)) -and (Test-Path $logFilePath)) { $finalLogPathMessage = $logFilePath }
        if ($transferSuccess) { Write-Host "Express Transfer process completed." -FG Green; [System.Windows.MessageBox]::Show("Express Transfer completed.`nData restored locally.`nBackup copy in: $finalBackupDir`nLog: $finalLogPathMessage", "Express Transfer Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) }
        else { Write-Error "Express Transfer failed. Check log: $finalLogPathMessage"; [System.Windows.MessageBox]::Show("Express Transfer failed. Check log: `n$finalLogPathMessage", "Express Transfer Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error); throw "Express Transfer failed. Check log: $finalLogPathMessage" }
    }
}

#endregion Functions

# --- Main Execution ---
Write-Host "--- Script Starting ---"
Clear-Variable -Name updateJob -Scope Script -ErrorAction SilentlyContinue
$operationMode = 'Cancel' # Default

# Ensure Default Path Exists
if (-not (Test-Path $script:DefaultPath)) { Write-Host "Default path '$($script:DefaultPath)' not found. Creating."; try { New-Item -Path $script:DefaultPath -ItemType Directory -Force -EA Stop | Out-Null; Write-Host "Default path created." } catch { Write-Warning "Could not create default path: $($script:DefaultPath)." } }

try {
    Write-Host "Calling Show-ModeDialog..."
    $operationMode = Show-ModeDialog

    switch ($operationMode) {
        'Backup' { Write-Host "Mode: Backup"; Show-MainWindow -Mode 'Backup' }
        'Restore' { Write-Host "Mode: Restore"; Show-MainWindow -Mode 'Restore' } # Background job started in Show-ModeDialog
        'Express' {
            Write-Host "Mode: Express. Checking elevation..."
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $isAdmin) {
                Write-Warning "Elevation required. Relaunching..."
                try { Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs -EA Stop }
                catch { Write-Error "Failed relaunch: $($_.Exception.Message). Run as Admin manually."; [System.Windows.MessageBox]::Show("Failed relaunch. Run as Admin manually.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) }
                Exit
            } else {
                Write-Host "Already Admin. Getting credentials..."
                $credential = $null; try { $credential = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message "Enter ADMIN credentials for REMOTE machine" } catch { throw "Failed get credentials: $($_.Exception.Message)" }
                if ($credential -eq $null) { throw "Credentials required for Express mode." }
                Execute-ExpressModeLogic -Credential $credential
            }
        }
        'Cancel' { Write-Host "Operation Cancelled." }
        Default { Write-Error "Invalid mode: $operationMode" }
    }
} catch {
    $errorMessage = "FATAL ERROR: $($_.Exception.Message)"
    if ($_.FullyQualifiedErrorId -match 'Express Transfer failed') { Write-Host "Express mode failed. See previous messages." -FG Red }
    else { Write-Error $errorMessage -EA Continue; try { [System.Windows.MessageBox]::Show($errorMessage, "Fatal Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) } catch {} }
} finally {
    if ($operationMode -eq 'Restore' -and (Get-Variable -Name updateJob -Scope Script -EA SilentlyContinue) -ne $null -and $script:updateJob -ne $null) {
        Write-Host "`n--- Waiting for background Restore updates job... ---" -FG Yellow; Wait-Job $script:updateJob | Out-Null; Write-Host "--- Background Job Output: ---" -FG Yellow; Receive-Job $script:updateJob; Remove-Job $script:updateJob; Write-Host "--- End Background Job Output ---" -FG Yellow
    } elseif ($operationMode -ne 'Express' -and $operationMode -ne 'Cancel') { Write-Host "`nNo background update job started/needed." -FG Gray }
}

Write-Host "--- Script Execution Finished ---"
if ($Host.Name -eq 'ConsoleHost' -and -not $psISE -and $env:TERM_PROGRAM -ne 'vscode') { Write-Host "Press Enter to exit..." -FG Yellow; Read-Host }