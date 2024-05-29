<#
.SYNOPSIS
  Name: Export-MyNspCertificates.ps1
  Exports all own valid certifcifates as PFX with a random password.

.DESCRIPTION
  This script will export all own valid certificates including the private key.
  A random password is generated for each certificate and will be safed in a seperate text file.

.PARAMETER CertificateFolder
  Mandatory parameter
  Input the full path to a folder where the certificates should be saved.

.PARAMETER PasswordFolder
  Mandatory parameter
  Input the full path to a folder where the passwords for the certificates should be saved.

.PARAMETER TenantPrimaryDomain
 Used to login into the desired NoSpamProxy tenant to run this script on.
 Only required if NoSpamProxy v14 is used in provider mode. 

.OUTPUTS
  Certificates are saved under the given path in <CertificateFolder> parameter.
  Passwords are saved under the given path in <PasswordFolder> parameter.
  A file called "ListOfThumbprints.txt" will be saved into the <CertificateFolder> path. 
  It contains all exported thumbprints and is used to prevent to export all certificates at every time.

.NOTES
  Version:        1.0.3
  Author:         Jan Jaeschke
  Creation Date:  2024-05-29
  Purpose/Change: added v14 compatibility
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.com
  https://github.com/noSpamProxy

.EXAMPLE
  .\Export-MyNspCertificates.ps1 -CertificateFolder "C:\certs" -PasswordFolder "C:\pw"
#>
param (
  [Parameter(Mandatory = $true)][string] $CertificateFolder,
  [Parameter(Mandatory = $true)][string] $PasswordFolder,
  # only needed for v14 with enabled provider mode
  [Parameter(Mandatory = $false)][string] $TenantPrimaryDomain
)

function processCertificates($myCertificates) { 
  # import list of already exported certificates
  if (Test-Path $thumbprintList) {
    $savedCertificates = Get-Content "$thumbprintList" 
  }
  else {
    $savedCertificates = ''
  }

  # export certificate as pfx if it is not already in the lsit of exported certificates
  foreach ($certificate in $myCertificates) {
    $certThumbprint = $certificate.Thumbprint
    $certMail = $certificate.MailAddresses[0].Address
    if ($savedCertificates -match $certThumbprint) {
      Continue
    }
    else {
      $pfxFile = "$CertificateFolder\$certMail-$certThumbprint.p12"
      $pwFile = "$PasswordFolder\$certMail-$certThumbprint.txt"
      # generates a random 16 char password, 4 chars have to be not alphanumeric
      $pfxPassword = [System.Web.Security.Membership]::GeneratePassword(16, 4)
      $bytetCertificate = (Export-NspCertificate -Thumbprint $certThumbprint -PrivateKeyPassword $pfxPassword).Certificate
      [io.file]::WriteAllBytes("$pfxFile", $bytetCertificate)
      $pfxPassword | Out-File -FilePath "$pwFile"
      $certThumbprint | Out-File -Append -FilePath $thumbprintList
    }
  }
}

$nspVersion = (Get-ItemProperty -Path HKLM:\SOFTWARE\NoSpamProxy\Components -ErrorAction SilentlyContinue).'Intranet Role'
if ($nspVersion -gt '14.0') {
  try {
    Connect-Nsp -IgnoreServerCertificateErrors -ErrorAction Stop
  }
  catch {
    $e = $_
    Write-Warning "Not possible to connect with the NoSpamProxy. Please check the error message below."
    $e | Format-List * -Force
    EXIT
  }
  if ($(Get-NspIsProviderModeEnabled) -eq $true) {
    if ($null -eq $TenantPrimaryDomain -OR $TenantPrimaryDomain -eq "") {
      Write-Host "Please provide a TenantPrimaryDomain to run this script with NoSpamProxy v14 in provider mode."
      EXIT
    }
    else {
      # NSP v14 has a new authentication mechanism, Connect-Nsp is required to authenticate properly
      # -IgnoreServerCertificateErrors allows the usage of self-signed certificates
      Connect-Nsp -IgnoreServerCertificateErrors -PrimaryDomain $TenantPrimaryDomain
    }
  }
}

$thumbprintList = "$CertificateFolder\ListOfThumbprints.txt"
$getMessageTracks = $true
$skipMessageTracks = 0

Write-Host "Start exporting private certificates ..."
while ($getMessageTracks -eq $true) {	
  if ($skipMessageTracks -eq 0) {
    # get all own certificate thubmbprints
    $myCertificates = (Get-NspCertificate -StoreIds My -KeyValidity Valid -First 100)
  }
  else {
    $myCertificates = (Get-NspCertificate -StoreIds My -KeyValidity Valid -First 100 -Skip $skipMessageTracks)
  }

  processCertificates $myCertificates

  # exit condition
  if ($messageTracks) {
    $skipMessageTracks = $skipMessageTracks + 100
    Write-Host $skipMessageTracks
  }
  else {
    $getMessageTracks = $false
    break
  }
}
Write-Host "Exporting private certificates done."