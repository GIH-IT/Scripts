#===============================================================================================================================================================================================
# GetEmailAdressFromADGroups v1.0 by Charlie Gustav Skog - https://github.com/GIH-IT/Scripts/blob/master/GetEmailAdressFromDistributionGroups/GetEmailAdressFromDistributionGroups.ps1
# Script to get groups from an OU and then check another OU for users that are members of those groups and compile a list for each group.
# GetEmailAdressFromADGroups.ps1 -DCTarget <hostname> -GroupsOU <path> -UsersOU <path> -ExportFolderPath <path>
#===============================================================================================================================================================================================
### Parameters
Param(
  [Parameter(Mandatory=$false, Position=0)]
  [string]$DCTarget = "gihdc03.ihs.se"
  [Parameter(Mandatory=$false, Position=1)]
  [string]$GroupsOU = "OU=DistributionGroupsStudents,OU=Groups,OU=GIH,DC=ihs,DC=se"
  [Parameter(Mandatory=$false, Position=2)]
  [string]$UsersOU = "OU=StudentAccounts,OU=Users,OU=GIH,DC=ihs,DC=se"
  [Parameter(Mandatory=$false, Position=3)]
  [string]$ExportFolderPath = $PSScriptRoot + "groups\"
)

### Script title
$host.ui.RawUI.WindowTitle = "GetEmailAdressFromDistributionGroups v1.0 by Charlie Gustav Skog"


### Modules
# Load the Active Directory module.
Import-Module ActiveDirectory


### Variables
$DCTarget = "gihdc03.ihs.se"
$Credentials = Get-Credential
$Groups = Get-ADGroup -Server $DCTarget -Credential $Credentials -Filter * -SearchBase $GroupsOU


### Script
# Get users that are members of the current group in the loop and export them to "<ScriptRoot>\groups\<GroupName>.csv".
ForEach($Group in $Groups){
  $SearchFilter = "CN=" + $Group.Name + "," + $GroupsOU
  $Users = Get-ADUser -Server $DCTarget -Credential $Credentials -Filter {memberOf -eq $SearchFilter} -SearchBase $UsersOU -Properties mail | Select mail | Export-CSV ($ExportFolderPath + $Group.Name + ".csv") -Encoding UTF8
}
