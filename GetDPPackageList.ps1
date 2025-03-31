<# 
.SYNOPSIS
  Lists packages currently stored on a given distribution point.
.DESCRIPTION
  Takes in the hostname of a Distribution Point and lists out the package names of packaages currently stored on that distribution point helpful
  for working out Distribution point content mismatch errors in the MECM console.
#>

$InputComputer = Read-Host "Input Hostname of Distribution Point"
Get-ChildItem -Path \\$InputComputer\c$\SCCMContentLib\PkgLib -Name | ForEach-Object -process {[system.io.path]::GetFileNameWithoutExtension($_)}
