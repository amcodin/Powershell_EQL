# Project Overview
A PowerShell-based utility for comprehensive user data backup and restore operations. Designed for IT administrators and support staff to manage user data migration and backup processes.  This should work with powershell version 1.

This has a GUI for users to interact with when selecting options.

New functionality needs to added and updated specifically;

# THe backup process
Function Set-InitialStateAdminBackup
Function Start-Backup

# The restore process:

Function Set-InitialStateRestore 
Function Start-Restore

Function Set-GPupdate
launches a new commandprompt window to start gpupdate /force

# front end
function Initialize-Form
[xml]$XAML =
Function Show-ProgressBar
Function Update-ProgressBar
Function Show-UserPrompt
Function Set-InitialStateBackup

# Processing
Format-RegexSafe


## To-do
- These locations where to look for more files to backup
$Paths = #default paths we want to add if they exist.
Function Get-FilePath 
Add default folder as C:\local for backup
add default folder as C:\local\backupfolder for restore

on-backup
grab quick access links
grab outlook signatures - already working
export edge favourites - export edge favorites
export chrome favourites, import chrome favorites to edge
grab sticky notes data - read @stickynotes.md
grab onenote mapped books
grab network drives, capture all drives, even if they're disconnected
grab printers
alert about .pst files- check common PST file locations


on-restore
the Set-GPupdate should only trigger on restore, and should gpupdate /force in a separate cmd instance window
run all config manager actions - read @config-manager-actions.md
request manage user certificates for wifi - read @manage-users-certification-wifi.md
disable nvidia container service - read @disable-service.md
if an existing signature, replace the old one with the most recent timestamp