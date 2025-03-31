<# 
.SYNOPSIS
  Lists package names of all the packages stored in WMI on a given distribution point

.DESCRIPTION
  Takes the hostname of a distribution point and lists out all of the packages currently stored in WMI on this distibution point. This can be useful for working
  out ditribution point mismatch errors in the MECM console.
#>
$InputComputer = Read-Host "Enter Hostname of Distribution Point"
Get-WMIObject -ComputerName $InputComputer -Namespace "root\sccmdp" -Query "Select * from SMS_PackagesInContLib" | Select PackageID | ft -HideTableHeaders
