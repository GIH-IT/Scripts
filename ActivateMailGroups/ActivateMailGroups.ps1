#=========================================================================================================================================
# ActivateMailGroups v1.0 by Charlie Gustav Skog - https://github.com/GIH-IT/Scripts/blob/master/ActivateMailGroups/ActivateMailGroups.ps1
# Script to activate the mailgroups with a set prefix.
# ActivateMailGroups.ps1 -GroupPrefix <string> -GroupsOU <path>
# GroupPrefix is mandatory.
#=========================================================================================================================================
### Parameters, Title, Checks, Variables and PSSessions.
# Set parameters
Param(
  [Parameter(Mandatory=$true, Position=0)]
  [string]$GroupPrefix,
  [Parameter(Mandatory=$false, Position=1)]
  [string]$GroupsOU = "OU=DistributionGroupsStudents,OU=Groups,OU=GIH,DC=ihs,DC=se"
)

# Set PowerShell title.
$host.ui.RawUI.WindowTitle = "ActivateMailGroups v1.0 by Charlie Gustav Skog"

# Default variables.
$ScriptCredentials = Get-Credential
$DCTarget = "gihdc03.ihs.se"
$MailServerTarget = "gihex02.ihs.se"
$MailServerTargetURI = "http://" + $MailServerTarget + "/powershell/"

# Check connection to Exchange server, if down exit script.
If (Test-Connection -ComputerName $MailServerTarget -Count 1 -ErrorAction SilentlyContinue) {}
Else {
  Write-Host "Connection to" $MailServerTarget "is down."
  exit
}

# Check connection to Domain Controller, if down exit script.
If (Test-Connection -ComputerName $DCTarget -Count 1 -ErrorAction SilentlyContinue) {}
Else {
  Write-Host "Connection to" $DCTarget "is down."
  exit
}

# Set up PSSessions to Exchange Server, Domain Controller and import PSSessions with the needed commands.
$MailServerPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $MailServerTargetURI -Credential $ScriptCredentials -Authentication Kerberos -AllowRedirection
$DCPSSession = New-PSSession -ComputerName $DCTarget -Credential $ScriptCredentials -Authentication Kerberos
Import-PSSession $MailServerPSSession -CommandName Enable-DistributionGroup
Import-PSSession $MailServerPSSession -CommandName Set-DistributionGroup
Import-PSSession $DCPSSession -CommandName Get-ADGroup

# Get Groups from Active Directory.
$Groups = Get-ADGroup -Server $DCTarget -Credential $ScriptCredentials -LDAPFilter "(name=$GroupPrefix*)" -SearchBase $GroupsOU -Properties *

### Script
ForEach ($Group in $Groups) {
  Enable-DistributionGroup -Identity $Group.Name
  Set-DistributionGroup -Identity $Group.Name -ForceUpgrade
  Write-Host $Group.Name "Activated"
}
