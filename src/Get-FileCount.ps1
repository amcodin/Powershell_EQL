function Get-FileCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    try {
        # Get all child items recursively
        $items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction Stop
        
        # Initialize counters
        $fileCount = 0
        $folderCount = 0
        
        foreach ($item in $items) {
            if ($item -is [System.IO.FileInfo]) {
                $fileCount++
            } elseif ($item -is [System.IO.DirectoryInfo]) {
                $folderCount++
            }
        }
        
        # Create output object with counts
        $output = [PSCustomObject]@{
            Files = $fileCount
            Folders = $folderCount
            Total = $fileCount + $folderCount
        }
        
        return $output
    }
    catch {
        Write-Error "Error counting files in $Path : $($_.Exception.Message)"
        return $null
    }
}
