#============================================================================================================================================================================================================================================
# ActivateMailboxAndGeneratePassword v1.0 by Charlie Gustav Skog - https://github.com/GIH-IT/Scripts/blob/master/ActivateMailboxAndGeneratePassword/ActivateMailboxAndGeneratePassword.ps1
# Script to activate the mailbox of Students from a PDB User Create log(change_id;fkey_pdb_id;id_type;uuid;posix_namn;comment;change_dict;change_type;changed_at;changed_by_system;changed_by_user;logger_version) and change their password.
# ActivateMailboxAndGeneratePassword.ps1 -InDataFile <path> -ResultPath <path>
# InDataPath is mandatory.
#============================================================================================================================================================================================================================================
### Parameters, Title, Checks, Variables, Modules and PSSessions.
# Set parameters
Param(
  [Parameter(Mandatory=$true, Position=0)]
  [string]$InDataFile,
  [Parameter(Mandatory=$false, Position=1)]
  [string]$OutputPath = $PSScriptRoot + "\Output\",
  [Parameter(Mandatory=$false, Position=2)]
  [string]$PDFModule = $PSScriptRoot + "\PDF-Form.psm1",
  [Parameter(Mandatory=$false, Position=3)]
  [string]$iTextSharp = $PSScriptRoot + "\itextsharp.dll",
  [Parameter(Mandatory=$false, Position=4)]
  [string]$PDFTemplate = $PSScriptRoot + "\PDFTemplate.pdf"
)

# Set PowerShell title.
$host.ui.RawUI.WindowTitle = "ActivateMailboxAndGeneratePassword v2.0 by Charlie Gustav Skog"

# Check if input files exists. If they don't, exit script.
If (Test-Path $InDataFile) {}
Else {
  Write-Host "Data file not found, exiting script."
  exit
}
If (Test-Path $OutputPath) {}
Else {
  Write-Host "Output path not found, exiting script."
  exit
}
If (Test-Path $PDFModule) {}
Else {
  Write-Host "iTextSharp not found, exiting script."
  exit
}
If (Test-Path $iTextSharp) {}
Else {
  Write-Host "iTextSharp not found, exiting script."
  exit
}
If (Test-Path $PDFTemplate) {}
Else {
  Write-Host "PDF template not found, exiting script."
  exit
}

# Default variables.
$ScriptCredentials = Get-Credential
$DCTarget = "gihdc03.ihs.se"
$MailServerTarget = "gihex02.ihs.se"
$MailServerTargetURI = "http://" + $MailServerTarget + "/powershell/"
$InData = ConvertFrom-CSV -Delimiter ";" $(Get-Content $InDataFile)
$StudentMailDatabase = "GIH-STUD01"
$StudentOU = "OU=StudentAccounts,OU=Users,OU=GIH,DC=ihs,DC=se"
$ResultFile = $OutputPath + $(Get-Item $InDataFile).BaseName + "\" + $(Get-Item $InDataFile).BaseName + "-log.csv"
$ResultHeaders = "StudentName;StudentDisplayName;StudentSAMAccountName;StudentMail;StudentPNR;StudentPassword"

# Modules
Import-Module $PDFModule

# Create result file and write headers.
New-Item $ResultFile -Force
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
    [string]$file,
    [Parameter(Mandatory=$true, Position=1)]
    [string]$data
  )
  $data >> $file
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
  $StudentMail = $StudentUser.UserPrincipalName
  $StudentPNR = $StudentUser.EmployeeID
  $OutputPDF = $OutputPath + $(Get-Item $InDataFile).BaseName + "\" + $StudentPNR + ".pdf"

  # Activate the Student mailbox.
  Enable-Mailbox -Identity $StudentSAM -Database $StudentMailDatabase

  # Change the Student password.
  Set-ADAccountPassword -Identity $StudentSAM -NewPassword $StudentRandomPasswordSecure

  # Write out to result file.
  Write-Result -file $ResultFile -data "$StudentName;$StudentDisplayName;$StudentSAM;$StudentMail;$StudentPNR;$StudentRandomPassword"

  # Show PDF form fields.
  Get-PdfFieldNames -FilePath $PDFTemplate -ITextLibraryPath $iTextSharp

  # Generate PDF.
  Save-PdfField -Fields @{'name' = "$StudentDisplayName";'studentpnr' = "$StudentPNR";'mail' = "$StudentMail";'username' = "$StudentSAM";'password' = "$StudentRandomPassword"} -InputPdfFilePath $PDFTemplate -ITextSharpLibrary $iTextSharp -OutputPdfFilePath $OutputPDF
}
