function Get-FileCount {
    param([string]$Path)
    $count = 0
    Get-ChildItem -Path $Path -Recurse -File | ForEach-Object { $count++ }
    return $count
}
