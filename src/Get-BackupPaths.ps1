function Get-BackupPaths {
    [CmdletBinding()]
    param ()
    
    $pathsToCheck = @(
        @{Path = "$env:APPDATA\Microsoft\Signatures"; Name = "Outlook Signatures"; Type = "Folder"}
        @{Path = "$env:SystemDrive\User"; Name = "User Directory"; Type = "Folder"}
        @{Path = "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms"; Name = "Quick Access"; Type = "File"}
        @{Path = "$env:SystemDrive\Temp"; Name = "Temp Directory"; Type = "Folder"}
        @{Path = "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt"; Name = "Sticky Notes (Legacy)"; Type = "File"}
        @{Path = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"; Name = "Sticky Notes"; Type = "File"}
        @{Path = "$env:APPDATA\google\googleearth\myplaces.kml"; Name = "Google Earth Places"; Type = "File"}
    )

    $result = @()

    foreach ($item in $pathsToCheck) {
        if (Test-Path -Path $item.Path) {
            $result += [PSCustomObject]@{
                Name = $item.Name
                Path = $item.Path
                Type = $item.Type
            }
        }
    }

    # Add Chrome bookmarks if accessible
    $chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
    try {
        if (Test-Path $chromePath) {
            Get-Content $chromePath -ErrorAction Stop | Out-Null
            $result += [PSCustomObject]@{
                Name = "Chrome Bookmarks"
                Path = $chromePath
                Type = "File"
            }
        }
    } catch {
        Write-Warning "Chrome bookmarks file not accessible"
    }

    return $result
}
