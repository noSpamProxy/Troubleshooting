<#
.SYNOPSIS
  Name: CheckNspCertificates.ps1
  Check a certificate including its chain by thumbprint.

.DESCRIPTION
  This script can be used to check a certificate and its chain by its thumbprint.
  Only into the NoSpamProxy imported certificates can be checked.
  The validation is done for each connected gateway role.

.PARAMETER TenantPrimaryDomain
  Connect to a specific tenant using the given primary domain name.

.PARAMETER Thumbprint
  Set the thumbprint of the certificate which should be checked.

.OUTPUTS
  The result is printed direclty into the console.

.NOTES
  Version:        1.0.1
  Author:         Finn Schulte
  Creation Date:  2024-11-18
  Purpose/Change: added v14 compatibility
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.de 
  https://www.github.com/noSpamProxy

.EXAMPLE
  .\CheckNspCertificates.ps1 -Thumbprint 0F4F9209E172B6D81022C0219CF253EFD29689F6

.EXAMPLE
  .\CheckNspCertificates.ps1 -Thumbprint 0F4F9209E172B6D81022C0219CF253EFD29689F6 -TenantPrimaryDomain "example.com"
#>
param (
	[Parameter(Mandatory = $true)][string] $Thumbprint,
	# only needed for v14 with enabled provider mode
	[Parameter(Mandatory = $false)][string] $TenantPrimaryDomain
)

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

$Roles = Get-NspGatewayRole
Foreach ($Role in $Roles) {
 #Test on each gateway
	$Result = Test-NspCertificate -Thumbprint $Thumbprint -GatewayRole $Role.Name 
	#Check the certificate on each gateway
	Foreach ($Partialresult in $Result) {
		#Split the result in root, intermediate and public or private certificate
		$Subject = ((($Partialresult.Subject).Split(","))[0]).Split("=")
		#Cut the address or common name out of the subject
	
		If ($Partialresult.ValidationResult -notlike "NoError") {
			Write-Host -ForegroundColor Yellow $Role.Name": " -NoNewline
			Write-Host -ForegroundColor White $Subject[1]"  " -NoNewline
			Write-Host -ForegroundColor Red $Partialresult.ValidationResult 
		}
		#Outputs gateway name, common name and error in red
		Else {
			Write-Host -ForegroundColor Yellow $Role.Name": " -NoNewline
			Write-Host -ForegroundColor White $Subject[1]"  " -NoNewline
			Write-Host -ForegroundColor Green $Partialresult.ValidationResult 
		}
		#Outputs gateway name, common name and "no error" in green
	}
	Write-Host ""
	#Adds a line for better separation between gateways
}