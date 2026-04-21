#Requires -Version 5.1
<#
.SYNOPSIS
    Requests new SMIME certificates for NoSpamProxy users via certificate enrollment providers.

.DESCRIPTION
    Loads all local NoSpamProxy users and selects mail addresses that have an existing
    SMIME certificate matching the given criteria (-ValidFromDate and optionally -Issuer).
    A new certificate is then requested for each matching address via the named enrollment
    provider using the WCF IIdentityController.RequestCertificates API.

    Addresses with no certificate, or no certificate matching -Issuer, are skipped.

.PARAMETER ValidFromDate
    Required. Lower bound (inclusive) on the certificate ValidFrom date. Only addresses
    whose most recent matching certificate has ValidFrom >= ValidFromDate are selected.

.PARAMETER ValidTo
    Optional. Upper bound (inclusive) on the certificate ValidFrom date. When combined
    with -ValidFromDate, only certificates whose ValidFrom falls within the range
    [ValidFromDate, ValidTo] are considered. Addresses whose most recent matching
    certificate falls outside this range are skipped.

.PARAMETER ProviderName
    Required. The exact name of the certificate enrollment provider configured in
    NoSpamProxy (as shown in the management console under Users and Keys ->
    Certificate Enrollment Providers). If the name is not found, available provider
    names are listed and the script aborts.

.PARAMETER Issuer
    Optional. Case-insensitive substring match against the IssuerDistinguishedName of
    each certificate. When supplied, only certificates whose issuer contains this string
    are considered. Addresses with no certificate from that issuer are skipped.

.PARAMETER NspUrl
    URI of the NoSpamProxy Intranet Role endpoint. Default: https://localhost:6060.

.PARAMETER Credential
    PSCredential for authenticating to NoSpamProxy.
    If omitted, Windows default credentials are used.

.PARAMETER ApiKey
    A SecureString containing a NoSpamProxy Personal Access Token / API key.
    Mutually exclusive with -Credential.

.PARAMETER IgnoreServerCertificateErrors
    When set, TLS certificate validation errors on the connection to the
    NoSpamProxy Intranet Role are suppressed.

.EXAMPLE
    .\Request-SmimeCertificate.ps1 -ValidFromDate 2026-01-01 -ProviderName "SwissSign Gold ID"
    Requests new certificates for all addresses that have any cert with ValidFrom >= 2026-01-01.

.EXAMPLE
    .\Request-SmimeCertificate.ps1 -ValidFromDate 2025-01-01 -ValidTo 2026-01-01 -ProviderName "SwissSign Gold ID"
    Only addresses whose most recent matching cert has ValidFrom in [2025-01-01, 2026-01-01] are selected.

.EXAMPLE
    .\Request-SmimeCertificate.ps1 -ValidFromDate 2026-01-01 -ProviderName "SwissSign Gold ID" -Issuer "SwissSign"
    Same as the first example, but only considers certificates whose issuer DN contains 'SwissSign'.

.EXAMPLE
    .\Request-SmimeCertificate.ps1 -ValidFromDate 2026-01-01 -ProviderName "SwissSign Gold ID" `
        -NspUrl https://nsp-mgmt.contoso.local:6060 `
        -ApiKey (Read-Host -AsSecureString 'API Key')
    Connects to a remote NSP instance with API key authentication.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true)]
    [DateTime]$ValidFromDate,

    [Parameter(Mandatory = $false)]
    [DateTime]$ValidTo,

    [Parameter(Mandatory = $true)]
    [string]$ProviderName,

    [Parameter(Mandatory = $false)]
    [string]$Issuer,

    [Parameter(Mandatory = $false)]
    [Uri]$NspUrl,

    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential,

    [Parameter(Mandatory = $false)]
    [SecureString]$ApiKey,

    [Parameter(Mandatory = $false)]
    [switch]$IgnoreServerCertificateErrors
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 1. Import NoSpamProxy PowerShell module
# ---------------------------------------------------------------------------
$nspModule = Get-Module -ListAvailable -Name 'NoSpamProxy*' |
    Sort-Object Version -Descending |
    Select-Object -First 1

if (-not $nspModule) {
    throw 'NoSpamProxy PowerShell module not found. Install it or adjust $env:PSModulePath.'
}

Import-Module $nspModule.Name -Force -ErrorAction Stop
Write-Verbose "Imported module: $($nspModule.Name) v$($nspModule.Version)"

# ---------------------------------------------------------------------------
# 2. Connect to NoSpamProxy
# ---------------------------------------------------------------------------
$connectParams = @{}
if ($NspUrl)                        { $connectParams['Url']                        = $NspUrl }
if ($Credential)                    { $connectParams['Credential']                 = $Credential }
elseif ($ApiKey)                    { $connectParams['ApiKey']                     = $ApiKey }
if ($IgnoreServerCertificateErrors) { $connectParams['IgnoreServerCertificateErrors'] = $true }

$urlDisplay = if ($NspUrl) { $NspUrl.ToString() } else { 'https://localhost:6060 (default)' }
Write-Host "Connecting to NoSpamProxy at $urlDisplay ..."
Connect-Nsp @connectParams
Write-Host 'Connected.'

# ---------------------------------------------------------------------------
# 3. Retrieve JWT from Connect-Nsp's in-process token cache and compile
#    a thin C# helper that calls IIdentityController via WCF.
#    This mirrors the approach used by Update-NspTlsCertificate.ps1.
# ---------------------------------------------------------------------------
$jwt = [System.Runtime.Caching.MemoryCache]::Default['token']
if ([string]::IsNullOrEmpty($jwt)) {
    throw 'No JWT token found in MemoryCache. Ensure Connect-Nsp completed successfully.'
}

$wcfHost = if ($NspUrl) { $NspUrl.Host } else { 'localhost' }
$wcfPort = if ($NspUrl -and $NspUrl.Port -gt 0) { $NspUrl.Port } else { 6060 }

if (-not ('NspSmimeHelper' -as [type])) {
    # Collect assembly paths from types already loaded by the NSP module
    $refs = @(
        [Netatwork.NoSpamProxy.Wcf.Client.Channel].Assembly.Location
        [Netatwork.NoSpamProxy.Wcf.Services.IIdentityController].Assembly.Location
        [Netatwork.NoSpamProxy.WebServices.CertificateEnrollmentProviderConfiguration].Assembly.Location
        [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.CertificateEnrollmentInfo].Assembly.Location
    ) | Select-Object -Unique

    # netstandard.dll facade required on PS 5.1 (.NET Framework)
    $netstdPath = @(
        Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'Facades\netstandard.dll'
        Join-Path ([Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'netstandard.dll'
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($netstdPath) { $refs += $netstdPath }

    Add-Type -ReferencedAssemblies $refs -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Net;
using System.Threading.Tasks;
using Netatwork.NoSpamProxy.Wcf.Client;
using Netatwork.NoSpamProxy.Wcf.Services;
using Netatwork.NoSpamProxy.WebServices;
using Netatwork.NoSpamProxy.WebServices.CertificateEnrollment;

public static class NspSmimeHelper
{
    public static Channel CreateChannel(string host, int port, string jwt)
    {
        return Channel.CreateHttpChannelWitUserNameAuthentication(
            host, port,
            (svc, invalid, token) => Task.FromResult(
                new TokenizedCredential(
                    new NetworkCredential("Bearer", jwt),
                    token.GetValueOrDefault())));
    }

    public static IList<CertificateEnrollmentProviderConfiguration> GetProviders(Channel channel)
    {
        return channel.Execute((IIdentityController c) => c.GetCertificateEnrollmentProviders());
    }

    public static void RequestCertificates(
        Channel channel,
        IList<CertificateEnrollmentInfo> infos,
        Guid providerId)
    {
        channel.Execute((IIdentityController c) => c.RequestCertificates(infos, providerId));
    }
}
'@
    Write-Verbose 'Compiled NspSmimeHelper C# helper.'
}

$channel = [NspSmimeHelper]::CreateChannel($wcfHost, $wcfPort, $jwt)
Write-Verbose "WCF Channel created."

# ---------------------------------------------------------------------------
# 4. Resolve certificate enrollment provider by exact name
# ---------------------------------------------------------------------------
$allProviders = [NspSmimeHelper]::GetProviders($channel)

$selectedProvider = $allProviders | Where-Object { $_.Name -eq $ProviderName } | Select-Object -First 1

if ($null -eq $selectedProvider) {
    Write-Host "Available certificate enrollment providers:"
    $allProviders | ForEach-Object { Write-Host "  $($_.Name)" }
    throw "No provider named '$ProviderName' found. Check the name and try again."
}

$selectedProviderType = $selectedProvider.GetType().Name
Write-Host "Using provider: $($selectedProvider.Name) [$($selectedProvider.Id)] (type: $selectedProviderType)"


# ---------------------------------------------------------------------------
# 5. Load users and find those needing certificate renewal
# ---------------------------------------------------------------------------
Write-Host "Loading users from NoSpamProxy ..."
$allUsers = @(Get-NspUser)
Write-Host "Loaded $($allUsers.Count) user(s)."

# Collect (user, address) pairs that have a matching certificate:
#   - address must have at least one certificate matching -Issuer (if given)
#   - the most recent such certificate must have ValidFrom >= ValidFromDate
#   - if -ValidTo is given, the most recent matching cert must also have ValidFrom <= ValidTo
$candidates = [System.Collections.ArrayList]::new()

foreach ($user in $allUsers) {
    if ($null -eq $user.MailAddresses) { continue }

    foreach ($address in $user.MailAddresses) {
        if (-not $address.IsEnabled) { continue }

        $certs = $address.AssociatedCertificates
        if ($null -eq $certs -or $certs.Count -eq 0) { continue }

        # When -Issuer is given, restrict to certs from that issuer
        $relevantCerts = if ($Issuer) {
            $certs | Where-Object { $_.IssuerDistinguishedName -like "*$Issuer*" }
        } else {
            $certs
        }

        if ($null -eq $relevantCerts -or @($relevantCerts).Count -eq 0) { continue }

        # Include if the most recent matching certificate falls within the requested date range
        $newestValidFrom = ($relevantCerts | Measure-Object -Property ValidFrom -Maximum).Maximum
        $inRange = $newestValidFrom -ge $ValidFromDate
        if ($inRange -and $ValidTo) { $inRange = $newestValidFrom -le $ValidTo }
        if ($inRange) {
            $newestCert = $relevantCerts | Sort-Object ValidFrom -Descending | Select-Object -First 1
            $null = $candidates.Add([PSCustomObject]@{
                User            = $user
                Address         = $address
                NewestValidFrom = $newestValidFrom
                IssuerDN        = $newestCert.IssuerDistinguishedName
            })
        }
    }
}

Write-Host "Found $($candidates.Count) address(es) requiring a new certificate."

if ($candidates.Count -eq 0) {
    Write-Host 'Nothing to do.'
    return
}

# Print a dry-run summary so the user can verify before any request is submitted
Write-Host ""
Write-Host "Affected addresses:"
$candidates | ForEach-Object {
    [PSCustomObject]@{
        Address         = $_.Address.Address
        DisplayName     = $_.User.Displayname
        'ValidFrom'     = $_.NewestValidFrom.ToString('yyyy-MM-dd')
        Issuer          = $_.IssuerDN
    }
} | Format-Table -AutoSize | Out-String | Write-Host
Write-Host "Total: $($candidates.Count) certificate request(s) would be submitted to provider '$($selectedProvider.Name)'."
Write-Host ""

# ---------------------------------------------------------------------------
# 6. Build CertificateEnrollmentInfo list (type depends on provider)
# ---------------------------------------------------------------------------
$enrollmentInfoType = 'Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.CertificateEnrollmentInfo'
$enrollmentInfos    = New-Object "System.Collections.Generic.List[$enrollmentInfoType]"

foreach ($candidate in $candidates) {
    $u       = $candidate.User
    $address = $candidate.Address

    # Instantiate the right subclass based on the provider's concrete type
    $info = switch -Wildcard ($selectedProviderType) {
        '*SwissSign*' {
            $ei              = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.SwissSignCertificateEnrollmentInfo]::new()
            $ei.DisplayName  = $u.Displayname
            $ei.Organization = $u.Organization
            $ei
        }
        '*WindowsCA*' {
            $ei                    = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.WindowsCACertificateEnrollmentInfo]::new()
            $ei.DisplayName        = $u.Displayname
            $ei.Organization       = $u.Organization
            $ei.OrganizationalUnit = $u.Department
            $ei.City               = $u.Locality
            $ei
        }
        '*GlobalSignAtlas*' {
            $ei                    = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.GlobalSignAtlasCertificateEnrollmentInfo]::new()
            $ei.GivenName          = $u.Givenname
            $ei.Surname            = $u.Surname
            $ei.Organization       = $u.Organization
            $ei.OrganizationalUnit = $u.Department
            $ei.Locality           = $u.Locality
            $ei.State              = $u.StateOrProvince
            $ei
        }
        '*GlobalSign*' {
            $ei                    = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.GlobalSignCertificateEnrollmentInfo]::new()
            $ei.DisplayName        = $u.Displayname
            $ei.OrganizationalUnit = $u.Department
            $ei
        }
        '*DTrust*' {
            $ei                    = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.DTrustCertificateEnrollmentInfo]::new()
            $ei.DisplayName        = $u.Displayname
            $ei.GivenName          = $u.Givenname
            $ei.Surname            = $u.Surname
            $ei.Organization       = $u.Organization
            $ei.OrganizationalUnit = $u.Department
            $ei.City               = $u.Locality
            $ei.StateOrProvince    = $u.StateOrProvince
            $ei
        }
        '*DigiCert*' {
            $ei             = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.DigiCertCertificateEnrollmentInfo]::new()
            $ei.DisplayName = $u.Displayname
            $ei.GivenName   = $u.Givenname
            $ei.Surname     = $u.Surname
            $ei
        }
        '*Sectigo*' {
            $ei           = [Netatwork.NoSpamProxy.WebServices.CertificateEnrollment.SectigoCertificateEnrollmentInfo]::new()
            $ei.GivenName = $u.Givenname
            $ei.Surname   = $u.Surname
            $ei
        }
        default {
            Write-Warning "Unrecognised provider type '$selectedProviderType' - skipping $($address.Address)."
            $null
        }
    }

    if ($null -eq $info) { continue }

    $info.MailAddress = $address.Address
    $enrollmentInfos.Add($info)
    Write-Verbose "Queued: $($address.Address)"
}

if ($enrollmentInfos.Count -eq 0) {
    Write-Warning 'No enrollment infos could be built (unsupported provider type?).'
    return
}

# ---------------------------------------------------------------------------
# 7. Request certificates via WCF
# ---------------------------------------------------------------------------
if ($PSCmdlet.ShouldProcess(
        "Submit $($enrollmentInfos.Count) certificate request(s) to provider '$($selectedProvider.Name)'",
        'Request SMIME certificate')) {

    Write-Host "Requesting $($enrollmentInfos.Count) certificate(s) ..."
    [NspSmimeHelper]::RequestCertificates($channel, $enrollmentInfos, $selectedProvider.Id)
    Write-Host "Certificate request submitted for $($enrollmentInfos.Count) address(es)."
    Write-Host "  Provider : $($selectedProvider.Name)"
    Write-Host "  Threshold: certificates issued on or after $($ValidFromDate.ToString('yyyy-MM-dd'))"
    Write-Host "Check the NoSpamProxy management console for request status."
}
