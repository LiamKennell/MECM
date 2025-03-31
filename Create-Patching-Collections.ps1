<#
.SYNOPSIS
  Script to create patching collections in MECM for monthly OS Updates

.DESCRIPTION
  Requires you to enter 3 digit site code of MECM. Script then works out when patch Tuesday is for the month when the script is 
  being run and works out the dates for the collections based on this date.

  Modifications will be required for specific use cases i.e. Limited Collection IDs and Collection Queries
#>


import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

$MECMSiteCode = Read-Host "Enter MECM Site Code"
$MECMSiteCode +=":"

CD $MECMSiteCode

#Stores current date in various formats for use in Collection Names

$yyyyMM = (get-date -f "yyyy-MM")
$MM = (get-date -f "MM")
$YY = (get-date -f "YY")
$YYYY =(get-date -f "yyyy")
$ShortMonth = (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($MM)

#Find Patch Tuesday (13th always in same week as 2nd Tuesday so find 13th then find 2nd day of that week)

$Thirteenth = (get-date -Day 13).Date
#This line finds out which numbered day of the week the 13th is and Subtracts this from the second day which is Tuesday and passes this to .AddDays to the 13th 
#(i.e. if the 13th is a Saturday - day 6, then 2-6 = -4 so 13th -4 =  Tuesday 9th)
$PatchTuesday = $Thirteenth.AddDays( 2 - [int]$Thirteenth.DayOfWeek )

#Step 1 - Creates New Patching folder for the current month

New-CMFolder -Name "Windows 10 - $yyyyMM $ShortMonth" -ParentFolderPath "PR0:\DeviceCollection\Stores\Patching"

#Step 2 - Creates Patching collection in root folder in SCCM

$Devcollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | Dev - Install Now" -LimitingCollectionId "PR000200"
$Pilotcollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | $($PatchTuesday.AddDays(5).ToString("MMdd")) Pilot" -LimitingCollectionId "PR0000A0"
$PMcollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | $($PatchTuesday.AddDays(6).ToString("MMdd")) PM Machines" -LimitingCollectionId "PR0004BB"
$PScollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | $($PatchTuesday.AddDays(7).ToString("MMdd")) PS Machines" -LimitingCollectionId "PR0004BB"
$PNcollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | $($PatchTuesday.AddDays(12).ToString("MMdd")) PN Machines" -LimitingCollectionId "PR0000A7"
$PrimCADcollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | $($PatchTuesday.AddDays(12).ToString("MMdd")) Primary CAD" -LimitingCollectionId "PR0004BB"
$NonPrimCADcollection = New-CMDeviceCollection -Name "Windows 10 | $ShortMonth $YYYY Store Patching | $($PatchTuesday.AddDays(7).ToString("MMdd")) Non-Primary CAD" -LimitingCollectionId "PR0004BB"

#Step 3 - Moves new collections to folder created in step 1

Move-CMObject -InputObject $Devcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $Pilotcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PMcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PScollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PNcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PrimCADcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $NonPrimCADcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"

#Step 4 - Add Query Membership Rules

#Pilot
Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $Pilotcollection.CollectionID -IncludeCollectionId PR00044E
#PM
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PMcollection.CollectionID -RuleName "PM Patching $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PM%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='19045' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PMcollection.CollectionID -ExcludeCollectionId $Devcollection.CollectionID
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PMcollection.CollectionID -ExcludeCollectionId $Pilotcollection.CollectionID
#PS
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PScollection.CollectionID -RuleName "PS Patching $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PS%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='19045' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PScollection.CollectionID -ExcludeCollectionId $Devcollection.CollectionID
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PScollection.CollectionID -ExcludeCollectionId $Pilotcollection.CollectionID
#PN
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PNcollection.CollectionID -RuleName "PN Patching $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PN%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='19045' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PNcollection.CollectionID -ExcludeCollectionId $Devcollection.CollectionID
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PNcollection.CollectionID -ExcludeCollectionId $Pilotcollection.CollectionID
#Primary CAD
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PrimCADcollection.CollectionID -RuleName "Primary CAD Patching $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PK%1%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='19045' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PrimCADcollection.CollectionID -ExcludeCollectionId $Devcollection.CollectionID
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PrimCADcollection.CollectionID -ExcludeCollectionId $Pilotcollection.CollectionID
#Non-Primary CAD
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $NonPrimCADcollection.CollectionID -RuleName "Non-Primary CAD Patching $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PK%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='19045' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $NonPrimCADcollection.CollectionID -ExcludeCollectionId $Devcollection.CollectionID
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $NonPrimCADcollection.CollectionID -ExcludeCollectionId $Pilotcollection.CollectionID
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $NonPrimCADcollection.CollectionID -ExcludeCollectionId $PrimCADcollection.CollectionID


#If Month is Apr, Aug or Dec then also create collections for a till patch cycle

if ($MM -eq "04" -or $MM -eq "08" -or $MM -eq "12"){


#Step 1 (Tills) - Find first Sunday of next month as this will be the start of the rollout for tills

$FirstofNextMonth = (Get-Date -Day 1).AddMonths(+1)

while (($FirstofNextMonth).DayOfWeek -ne "Sunday") {
$FirstofNextMonth = $FirstofNextMonth.AddDays(1)
}
#After this while loop the FirstofNextMonth variable now contains the first Sunday of Next Month


#Step 2 (Tills) - Create Till Patching collections in root folder of SCCM

$TillDevcollection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | Dev - Install Now" -LimitingCollectionId "PR000200"
$TillPilot1collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.ToString("MMdd")) Pilot 1" -LimitingCollectionId "PR0007DB"
$TillPilot2collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.AddDays(14).ToString("MMdd")) Pilot 2" -LimitingCollectionId "PR0007DB"
$TillPilot3collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.AddDays(21).ToString("MMdd")) Pilot 3" -LimitingCollectionId "PR000491"
$PT20collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.AddDays(28).ToString("MMdd")) Till 20's" -LimitingCollectionId "PR000491"
$PT1collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.AddDays(30).ToString("MMdd")) Till 1's" -LimitingCollectionId "PR000491"
$PT3collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.AddDays(35).ToString("MMdd")) Till 3's" -LimitingCollectionId "PR000491"
$PT2collection = New-CMDeviceCollection -Name "Windows 10 LTSC | $ShortMonth $YYYY Store Till Patching | $($FirstofNextMonth.AddDays(37).ToString("MMdd")) Till 2's" -LimitingCollectionId "PR000491"

#Step 3 (Tills) - Move newly created collections to the correct folder in SCCM

Move-CMObject -InputObject $TillDevcollection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $TillPilot1collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $TillPilot2collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $TillPilot3collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PT20collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PT1collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PT3collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"
Move-CMObject -InputObject $PT2collection -FolderPath "$MECMSiteCode\DeviceCollection\Stores\Patching\Windows 10 - $yyyyMM $ShortMonth"


#Step 4 (Tills) - Add Query Membership Rules

#Pilot1 - Pilot Store Till 20's
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $TillPilot1collection.CollectionID -RuleName "Till Patching Pilot1 $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PT20%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='17763' "
#Pilot2 - Pilot Store Remaining Tills
Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $TillPilot2collection.CollectionID -IncludeCollectionId PR0007DB
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $TillPilot2collection.CollectionID -ExcludeCollectionId $TillPilot1collection.CollectionID
#Pilot3 - All Stores Till 21's and Till 4's (limited numbers of these)
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $TillPilot3collection.CollectionID -RuleName "Till Patching Pilot3 $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PT21%-02' or SMS_R_System.Name like 'S0%PT%25%-02'or SMS_R_System.Name like 'S0%PT%04%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='17763' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $TillPilot3collection.CollectionID -ExcludeCollectionId PR0007DB
#Till 20's
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PT20collection.CollectionID -RuleName "Till Patching Till 20s $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PT20%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='17763' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PT20collection.CollectionID -ExcludeCollectionId PR0007DB
#Till 1's
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PT1collection.CollectionID -RuleName "Till Patching Till 1s $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PT01%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='17763' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PT1collection.CollectionID -ExcludeCollectionId PR0007DB
#Till 3's
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PT3collection.CollectionID -RuleName "Till Patching Till 3s $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PT03%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='17763' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PT3collection.CollectionID -ExcludeCollectionId PR0007DB
#Till 2's
Add-CMDeviceCollectionQueryMembershipRule -CollectionId $PT2collection.CollectionID -RuleName "Till Patching Till 2s $yyyyMM" -QueryExpression "select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_OPERATING_SYSTEM on SMS_G_System_OPERATING_SYSTEM.ResourceID = SMS_R_System.ResourceId where SMS_R_System.Name like 'S0%PT02%-02' and SMS_G_System_OPERATING_SYSTEM.BuildNumber ='17763' "
Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $PT2collection.CollectionID -ExcludeCollectionId PR0007DB
}
