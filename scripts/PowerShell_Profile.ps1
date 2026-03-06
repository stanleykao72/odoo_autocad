# PowerShell Profile to auto-load environment variables
# This ensures that new PowerShell sessions have the correct PATH

# Refresh PATH from registry
function Update-PathFromRegistry {
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine)
    $userPath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
    $env:PATH = $machinePath + ";" + $userPath
}

# Auto-update PATH when profile loads
Update-PathFromRegistry

# Optional: Add alias for easy PATH refresh
Set-Alias -Name refreshpath -Value Update-PathFromRegistry

Write-Host "PowerShell Profile loaded - PATH updated!" -ForegroundColor Green
