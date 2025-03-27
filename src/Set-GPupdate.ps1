function Set-GPupdate {
    Write-Host "Initiating Group Policy update..." -ForegroundColor Cyan
    Start-Process -FilePath "cmd.exe" -ArgumentList "/k gpupdate /force" -PassThru | Out-Null
    Write-Host "Group Policy update initiated" -ForegroundColor Green
}
