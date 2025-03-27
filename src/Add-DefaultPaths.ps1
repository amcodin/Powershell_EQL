function Add-DefaultPaths {
    foreach ($path in $script:DefaultPaths) {
        if (Test-Path -Path $path.Path) {
            $script:lvwFiles.Items.Add([PSCustomObject]@{
                Name = $path.Name
                Type = $path.Type
                Path = $path.Path
            })
        }
    }
}
