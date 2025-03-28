function Get-BackupPaths {
    [CmdletBinding()]
    param ()
    
    $pathsToCheck = @(
        "$env:APPDATA\Microsoft\Signatures",
        "$env:SystemDrive\User",
        "$env:APPDATA\Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms",
        "$env:SystemDrive\Temp",
        "$env:APPDATA\Microsoft\Sticky Notes\StickyNotes.snt",
        "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite",
        "$env:APPDATA\google\googleearth\myplaces.kml"
    )

    $result = @()

    foreach ($path in $pathsToCheck) {
        if (Test-Path -Path $path) {
            $result += [PSCustomObject]@{
                Name = Split-Path $path -Leaf
                Path = $path
                Type = if (Test-Path -Path $path -PathType Container) { "Folder" } else { "File" }
            }
        }
    }

    # Add Chrome bookmarks if accessible
    $chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"
    try {
        if (Test-Path $chromePath) {
            Get-Content $chromePath -ErrorAction Stop | Out-Null
            $result += [PSCustomObject]@{
                Name = "Bookmarks"  # Chrome's actual bookmark file name
                Path = $chromePath
                Type = "File"
            }
        }
    } catch {
        Write-Warning "Chrome bookmarks file not accessible"
    }

    return $result
}
