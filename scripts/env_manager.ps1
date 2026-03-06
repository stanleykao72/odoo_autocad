# Environment Variable Management Script
# Simple version for managing PATH environment variables

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("show", "add", "remove", "clean", "refresh")]
    [string]$Action = "show",
    
    [Parameter(Mandatory=$false)]
    [string]$Path = "",
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("User", "Machine")]
    [string]$Scope = "User"
)

function Show-PathVariable {
    param([string]$VarScope)
    
    Write-Host "=== $VarScope level PATH environment variable ===" -ForegroundColor Green
    $envVar = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::$VarScope)
    if ($envVar) {
        $paths = $envVar -split ';' | Where-Object { $_ -ne "" }
        for ($i = 0; $i -lt $paths.Count; $i++) {
            $exists = Test-Path $paths[$i]
            $status = if ($exists) { "[OK]" } else { "[MISSING]" }
            $color = if ($exists) { "Green" } else { "Red" }
            Write-Host "$($i+1). $status $($paths[$i])" -ForegroundColor $color
        }
    } else {
        Write-Host "Environment variable is empty" -ForegroundColor Yellow
    }
    Write-Host ""
}

function Add-PathVariable {
    param([string]$NewPath, [string]$VarScope)
    
    if (-not (Test-Path $NewPath)) {
        Write-Host "Warning: Path '$NewPath' does not exist" -ForegroundColor Yellow
        $confirm = Read-Host "Do you want to add it anyway? (y/N)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            return
        }
    }
    
    $currentVar = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::$VarScope)
    $paths = if ($currentVar) { $currentVar -split ';' | Where-Object { $_ -ne "" } } else { @() }
    
    if ($paths -contains $NewPath) {
        Write-Host "Path already exists: $NewPath" -ForegroundColor Yellow
        return
    }
    
    $newVar = ($paths + $NewPath) -join ';'
    [Environment]::SetEnvironmentVariable("PATH", $newVar, [EnvironmentVariableTarget]::$VarScope)
    Write-Host "Added path: $NewPath" -ForegroundColor Green
    
    # Update current session
    Update-SessionPath
}

function Remove-PathVariable {
    param([string]$PathToRemove, [string]$VarScope)
    
    $currentVar = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::$VarScope)
    if (-not $currentVar) {
        Write-Host "PATH variable is empty" -ForegroundColor Yellow
        return
    }
    
    $paths = $currentVar -split ';' | Where-Object { $_ -ne "" -and $_ -ne $PathToRemove }
    
    $newVar = $paths -join ';'
    [Environment]::SetEnvironmentVariable("PATH", $newVar, [EnvironmentVariableTarget]::$VarScope)
    Write-Host "Removed path: $PathToRemove" -ForegroundColor Green
    
    # Update current session
    Update-SessionPath
}

function Clean-PathVariable {
    param([string]$VarScope)
    
    $currentVar = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::$VarScope)
    if (-not $currentVar) {
        Write-Host "PATH variable is empty" -ForegroundColor Yellow
        return
    }
    
    $paths = $currentVar -split ';' | Where-Object { 
        $_ -ne "" -and (Test-Path $_)
    } | Sort-Object -Unique
    
    $newVar = $paths -join ';'
    [Environment]::SetEnvironmentVariable("PATH", $newVar, [EnvironmentVariableTarget]::$VarScope)
    Write-Host "Cleaned duplicate and invalid paths" -ForegroundColor Green
    
    # Update current session
    Update-SessionPath
}

function Update-SessionPath {
    Write-Host "Refreshing environment variables..." -ForegroundColor Yellow
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::Machine)
    $userPath = [Environment]::GetEnvironmentVariable("PATH", [EnvironmentVariableTarget]::User)
    $env:PATH = $machinePath + ";" + $userPath
    Write-Host "Environment variables refreshed" -ForegroundColor Green
}

# Main logic
switch ($Action.ToLower()) {
    "show" {
        Show-PathVariable -VarScope "User"
        Show-PathVariable -VarScope "Machine"
    }
    "add" {
        if ($Path -eq "") {
            $Path = Read-Host "Enter the path to add"
        }
        Add-PathVariable -NewPath $Path -VarScope $Scope
    }
    "remove" {
        if ($Path -eq "") {
            Show-PathVariable -VarScope $Scope
            $Path = Read-Host "Enter the path to remove"
        }
        Remove-PathVariable -PathToRemove $Path -VarScope $Scope
    }
    "clean" {
        Write-Host "Cleaning $Scope level PATH..." -ForegroundColor Yellow
        Clean-PathVariable -VarScope $Scope
    }
    "refresh" {
        Update-SessionPath
    }
}

# Show usage if no valid action
if ($Action -notin @("show", "add", "remove", "clean", "refresh")) {
    Write-Host @"
Environment Variable Management Script

Usage:
  .\env_manager.ps1 [parameters]

Parameters:
  -Action     Operation type (show|add|remove|clean|refresh)
  -Path       Path (for add/remove operations)
  -Scope      Scope (User|Machine, default: User)

Examples:
  # Show current PATH
  .\env_manager.ps1 -Action show

  # Add path to user-level PATH
  .\env_manager.ps1 -Action add -Path "C:\MyTools"

  # Remove path from system-level PATH
  .\env_manager.ps1 -Action remove -Path "C:\OldTool" -Scope Machine

  # Clean duplicate and invalid paths
  .\env_manager.ps1 -Action clean

  # Refresh current session environment variables
  .\env_manager.ps1 -Action refresh
"@ -ForegroundColor Cyan
}
