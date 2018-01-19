#=========================================================================================================================================================================================
# ActivateMailboxAndGeneratePassword v1.0 by Charlie Gustav Skog - https://github.com/GIH-IT/Scripts/blob/master/ActivateMailboxAndGeneratePassword/ActivateMailboxAndGeneratePassword.ps1
# Script to activate the mailbox of newly created students and change their password.
# ActivateMailboxAndGeneratePassword.ps1 -InDataPath <path> -ResultPath <path>
# InDataPath is mandatory.
#=========================================================================================================================================================================================
### Parameters, Title, Checks, Variables and PSSessions.
# Set parameters
Param(
  [Parameter(Mandatory=$true, Position=0)]
  [string]$InDataPath,
  [Parameter(Mandatory=$false, Position=1)]
  [string]$ResultPath = $PSScriptRoot + "\" + $(Get-Item $InDataPath).BaseName + "-log.csv"
)

# Set PowerShell title.
$host.ui.RawUI.WindowTitle = "ActivateMailboxAndGeneratePassword v1.0 by Charlie Gustav Skog"

# Check if input file exists. If it doesn't, exit script.
If (Test-Path $InDataPath) {}
Else {
  Write-Host "Input file not found."
  exit
}

# Check if result file already exists. If it does, exit script.
If (Test-Path $ResultPath) {
  Write-Host $ResultPath "already exists."
  exit
}

# Default variables.
$ScriptCredentials = Get-Credential
$DCTarget = "gihdc03.ihs.se"
$MailServerTarget = "gihex02.ihs.se"
$MailServerTargetURI = "http://" + $MailServerTarget + "/powershell/"
$InData = ConvertFrom-CSV -Delimiter ";" $(Get-Content $InDataPath)
$StudentMailDatabase = "GIH-STUD01"
$StudentOU = "OU=StudentAccounts,OU=Users,OU=GIH,DC=ihs,DC=se"
$ResultHeaders = "StudentName;StudentDisplayName;StudentSAMAccountName;StudentMail;StudentPNR;StudentPassword"

# Create result file and write headers.
New-Item $ResultPath -Force
$ResultHeaders >> $ResultPath

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
Import-PSSession $MailServerPSSession -CommandName Enable-Mailbox
Import-PSSession $MailServerPSSession -CommandName Set-MailUser
Import-PSSession $DCPSSession -CommandName Get-ADUser
Import-PSSession $DCPSSession -CommandName Set-ADUser
Import-PSSession $DCPSSession -CommandName Set-ADAccountPassword


### Functions
# Function to convert password into a secure string.
Function New-SecureString() {
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$PlainText
  )
  $SecureString = new-object System.Security.SecureString
  ForEach($char in $PlainText.ToCharArray()) {
    $SecureString.AppendChar($char)
  }
  Return $SecureString
}

# Function to generate a random password.
Function Get-RandomPassword() {
  Param(
    [Parameter(Mandatory=$false, Position=0)]
    [int]$length=10
  )
  $AlphabetUpper = $NULL;For ($a=65; $a -le 90; $a++){$AlphabetUpper+=,[char][byte]$a}
  $AlphabetLower = $NULL;For ($a=97; $a -le 122; $a++){$AlphabetLower+=,[char][byte]$a}
  $Numerics = $NULL;For ($a=48; $a -le 57; $a++){$Numerics+=,[char][byte]$a}
  For ($loop=1; $loop -le $length; $loop++) {
    $TempPassword+=($AlphabetUpper + $AlphabetLower + $Numerics | Get-Random)
  }
  Return $TempPassword
}

# Function to write output to result file.
Function Write-Result() {
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$path,
    [Parameter(Mandatory=$true, Position=1)]
    [string]$data
  )
  $data >> $path
}


### Script
ForEach ($Student in $InData) {
  # Get Student object.
  $StudentUser = Get-ADUser -Identity $Student.posix_namn -Credential $ScriptCredentials -Server $DCTarget -Properties *

  # Student variables.
  $StudentRandomPassword = Get-RandomPassword
  $StudentRandomPasswordSecure = New-SecureString $StudentRandomPassword
  $StudentName = $StudentUser.Name
  $StudentDisplayName = $StudentUser.GivenName + " " + $StudentUser.Surname
  $StudentSAM = $StudentUser.SamAccountName
  $StudentMail = $StudentUser.mail
  $StudentPNR = $StudentUser.EmployeeID

  # Clear ProxyAddresses
  Set-AdUser -Identity $StudentSAM -Clear ProxyAddresses

  # Set ExternalEmailAddress
  Set-MailUser -Identity $StudentSAM -ExternalEmailAddress $StudentMail

  # Activate the Student mailbox.
  Enable-Mailbox -Identity $StudentSAM -Database $StudentMailDatabase

  # Change the Student password.
  Set-ADAccountPassword -Identity $StudentSAM -NewPassword $StudentRandomPasswordSecure

  # Write out to result file.
  Write-Result -path $ResultPath -data "$StudentName;$StudentDisplayName;$StudentSAM;$StudentMail;$StudentPNR;$StudentRandomPassword"
}
