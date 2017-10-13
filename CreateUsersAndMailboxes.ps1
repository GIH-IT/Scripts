#=================================================================================================================
# CreateUserAndMailboxes v1.0 by Charlie Gustav Skog - https://www.github.com/neotheone/CreateUserAndMailboxes
# Script to create student accounts from CSV file (pnr,FirstName,LastName).
# Make sure you have the Active Directory PowerShell module installed and are running this in an elevated terminal
# CreateUserAndMailboxes.ps1 -InDataPath <path> $ResultPath <path>
# ResultPath is not mandatory.
#=================================================================================================================
### Parameters
Param(
  [Parameter(Mandatory=$true, Position=0)]
  [string]$InDataPath,
  [Parameter(Mandatory=$false, Position=1)]
  [string]$ResultPath = $PSScriptRoot + $(Get-Item $InDataPath).BaseName + "-log.csv"
)


### Script title, result directory and usage
$host.ui.RawUI.WindowTitle = "CreateUserAndMailboxes v1.0 by Charlie Gustav Skog"


### Modules
# Load the Active Directory module
Import-Module ActiveDirectory


### Checks and Variables
# Check if input file exists.
If (Test-Path $InDataPath) {}
Else {
  Write-Host "Input file not found."
  exit
}

# Check if result file already exists. If it does, delete it.
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
$StudentFQDN = "student.gih.se"
$StudentMailDatabase = "GIH-STUD01"
$StudentOU = "OU=ScriptTemp,OU=ManuallyCreated,OU=StudentAccounts,OU=Users,OU=GIH,DC=ihs,DC=se"
$StudentDescription = "Created by script CreateUserAndMailboxes V1.0 run by " + $ScriptCredentials.UserName
$StudentHomeFolder = "\\gihfile02.ihs.se\home$\"
$StudentProfileFolder = "\\gihfile02.ihs.se\profiles$\"
$StudentLogonScript = "kix32.exe studentlogon.kix"
$ResultHeaders = "StudentName;StudentDisplayName;StudentSAMAccountName;StudentUserPrincipalName;StudentPNR;StudentPassword;StudentCreated;StudentNotCreatedReason"

# Create result file and write headers.
New-Item $ResultPath -Force
$ResultHeaders >> $ResultPath

# Check connection to Exchange server and load Exchange PowerShell module, if down exit script.
If (Test-Connection -ComputerName $MailServerTarget -Count 1 -ErrorAction SilentlyContinue) {}
Else {
  Write-Host "Connection to" $MailServerTarget "is down."
  exit
}

# Check connection to Domain Controller from Exchange server, if down exit script.
If (Test-Connection -ComputerName $DCTarget -Count 1 -ErrorAction SilentlyContinue) {}
Else {
  Write-Host "Connection to" $DCTarget "is down."
  exit
}

# Set up PSSession
$MailServerPSSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $MailServerTargetURI -Credential $ScriptCredentials -Authentication Kerberos -AllowRedirection
Import-PSSession $MailServerPSSession -CommandName New-Mailbox

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

# Function to generate a four letter number to append to the username.
Function Get-RandomNumber() {
  Param(
    [Parameter(Mandatory=$false, Position=0)]
    [int]$length=4
  )
  $Numerics = $NULL;For ($a=48; $a -le 57; $a++){$Numerics+=,[char][byte]$a}
  For ($loop=1; $loop -le $length; $loop++) {
    $TempNumber+=($Numerics | Get-Random)
  }
  Return $TempNumber
}

# Convert special characters
Function Convert-ToLatinCharacters() {
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$inputString
  )
  [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($inputString))
}

# Function to write output to result file
Function Write-Result() {
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$path,
    [Parameter(Mandatory=$true, Position=1)]
    [string]$data
  )
  $data >> $path
}

# Function to set permissions of Student owned folders.
Function Set-Permissions() {
  Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$path,
    [Parameter(Mandatory=$true, Position=1)]
    [string]$user
  )
  icacls $path /grant $user":(OI)(CI)F" /T
  icacls $path /setowner $user /c /t
}


### Create Students loop
ForEach ($Student in $InData) {
  # Student variables.
  $StudentRandomPassword = Get-RandomPassword
  $StudentRandomPasswordSecure = New-SecureString $StudentRandomPassword
  $StudentRandomNumber = Get-RandomNumber
  $StudentFirstName = $Student.FirstName.Trim()
  $StudentLastName = $Student.LastName.Trim()
  $StudentFirstNameFixed = $(Convert-ToLatinCharacters $StudentFirstName)
  $StudentLastNameFixed = $(Convert-ToLatinCharacters $StudentLastName)
  $StudentAlias = $StudentFirstNameFixed.ToLower().replace(' ', '_') + "." + $StudentLastNameFixed.ToLower().replace(' ', '_') + "$StudentRandomNumber"
  $StudentSAM = $StudentFirstNameFixed.ToLower().SubString(0,2) + "$StudentRandomNumber" + $StudentLastNameFixed.ToLower().SubString(0,2)
  $StudentUPN = $StudentAlias + "@" + $StudentFQDN
  $StudentName= $StudentFirstName + ' ' + $StudentLastName + ' ' + "$StudentRandomNumber"
  $StudentDisplayName = $StudentFirstName + ' ' + $StudentLastName
  $StudentPNR = $Student.pnr.Trim().SubString(0,6) + $Student.pnr.Trim().SubString(7,4)
  $StudentHomeFolderPath = $StudentHomeFolder + $StudentSAM
  $StudentProfileFolderPath = $StudentProfileFolder + $StudentSAM
  $StudentProfilePath = $StudentProfileFolder + $StudentSAM + "\profile"
  $StudentCreated = "No"
  $StudentNotCreatedReason = ""

  # Validate that no other Student in the Active Directory has the same information, if there is export information and exit script.
  $ExistingStudent = Get-ADUser -filter {(SamAccountName -eq $StudentSAM) -or (UserPrincipalName -eq $StudentUPN) -or (EmployeeID -eq $StudentPNR) -or (HomeDirectory -eq $StudentHomeFolderPath)}
  If (!$ExistingStudent) {}
  Else {
    $StudentNotCreatedReason = "already exist"
    Write-Host $StudentDisplayName $StudentNotCreatedReason"."
    Write-Result -path $ResultPath -data "$StudentName;$StudentDisplayName;$StudentSAM;$StudentUPN;$StudentPNR;$StudentRandomPassword;$StudentCreated;$StudentNotCreatedReason"
    Continue
  }

  # Create the Student by connecting through the PowerShell session to a mail server and running New-Mailbox which creates a corresponding Active Directory account.
  New-Mailbox -DomainController $DCTarget -ResetPasswordOnNextLogon $true -Password $StudentRandomPasswordSecure -Database $StudentMailDatabase -UserPrincipalName $StudentUPN -SamAccountName $StudentSAM -Name $StudentName -OrganizationalUnit $StudentOU -FirstName $StudentFirstName -LastName $StudentLastName

  # Check if Student was created, if not export information and exit script.
  $StudentUser = Get-ADUser -Identity $StudentSAM -Credential $ScriptCredentials -Server $DCTarget
  If (!$StudentUser) {
    $StudentNotCreatedReason = "was not created"
    Write-Host $StudentDisplayName $StudentNotCreatedReason"."
    Write-Host "Possible issue could be the credentials does not have access to the Exchange server."
    Write-Result -path $ResultPath -data "$StudentName;$StudentDisplayName;$StudentSAM;$StudentUPN;$StudentPNR;$StudentRandomPassword;$StudentCreated;$StudentNotCreatedReason"
    Continue
  }
  Else {
    $StudentCreated = "Yes"
  }

  # Create home and profile folders, add the students permissions and set as owner. This is not run as $Credentials.
  New-Item -path $StudentHomeFolderPath -type Directory
  New-Item -path $StudentProfileFolderPath -type Directory
  New-Item -path $StudentProfilePath -type Directory
  Set-Permissions -path $StudentHomeFolderPath -user $StudentSAM
  Set-Permissions -path $StudentProfileFolderPath -user $StudentSAM

  # Add home directory path, profile path, logonscript path, PNR, AllStudents group to the Student user and export information.
  Set-ADUser -Identity $StudentSAM -DisplayName $StudentDisplayName -Description $StudentDescription -HomeDirectory $StudentHomeFolderPath -ProfilePath $StudentProfilePath -HomeDrive K: -ScriptPath $StudentLogonScript -employeeID $StudentPNR -Credential $ScriptCredentials -Server $DCTarget
  Add-ADGroupMember -Identity AllStudents -Members $StudentUser
  Write-Result -path $ResultPath -data "$StudentName;$StudentDisplayName;$StudentSAM;$StudentUPN;$StudentPNR;$StudentRandomPassword;$StudentCreated;$StudentNotCreatedReason"
}
