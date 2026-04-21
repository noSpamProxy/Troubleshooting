# Request-SmimeCertificate.ps1

Requests new SMIME certificates for NoSpamProxy users via the WCF `IIdentityController.RequestCertificates`
API — the same endpoint used by the NoSpamProxy management console.

## How it works

1. Connects to NoSpamProxy using `Connect-Nsp` (Windows auth, credentials, or API key).
2. Resolves the certificate enrollment provider by **exact name** (`-ProviderName`).
3. Loads all local users via `Get-NspUser` and inspects each enabled mail address.
4. **Skips** addresses with no certificate, or whose certificates do not match `-Issuer` (if given).
5. **Selects** addresses where the most recent matching certificate has `ValidFrom >= ValidFromDate`.
   If `-ValidTo` is also given, the cert's `ValidFrom` must additionally satisfy `ValidFrom <= ValidTo`.
6. Prints a table of all selected addresses so you can verify before anything is submitted.
7. Submits a batch `RequestCertificates` WCF call for all qualifying addresses.

Enrollment info type (SwissSign, WindowsCA, GlobalSign, DTrust, DigiCert, Sectigo) is selected
automatically based on the provider's concrete .NET type.

## Prerequisites

- PowerShell 5.1 or later
- NoSpamProxy PowerShell module installed (`NoSpamProxy*` on `$env:PSModulePath`)
- At least one SMIME certificate enrollment provider configured in NoSpamProxy

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `-ValidFromDate` | Yes | Lower bound (inclusive) on the cert's `ValidFrom` date — selects addresses whose most recent matching cert has `ValidFrom >= ValidFromDate` |
| `-ValidTo` | No | Upper bound (inclusive) on the cert's `ValidFrom` date — when combined with `-ValidFromDate`, only certs with `ValidFrom` in `[ValidFromDate, ValidTo]` are considered |
| `-ProviderName` | Yes | Exact name of the certificate enrollment provider as shown in the management console |
| `-Issuer` | No | Substring filter on `IssuerDistinguishedName`. Addresses with no cert from this issuer are skipped |
| `-NspUrl` | No | Intranet Role URL. Default: `https://localhost:6060` |
| `-Credential` | No | PSCredential for authentication. Uses Windows credentials if omitted |
| `-ApiKey` | No | SecureString Personal Access Token. Mutually exclusive with `-Credential` |
| `-IgnoreServerCertificateErrors` | No | Suppresses TLS certificate validation errors |
| `-WhatIf` | No | Dry run: shows who would be affected and how many certificates would be requested, without submitting |

## Usage

### Basic — local NSP, all issuers

```powershell
.\Request-SmimeCertificate.ps1 -ValidFromDate 2026-01-01 -ProviderName "SwissSign Gold ID"
```

Selects all enabled addresses that have a certificate with `ValidFrom >= 2026-01-01`.

### Date range — only certs issued within a window

```powershell
.\Request-SmimeCertificate.ps1 -ValidFromDate 2025-01-01 -ValidTo 2026-01-01 -ProviderName "SwissSign Gold ID"
```

Selects addresses whose most recent matching certificate has `ValidFrom` between 2025-01-01 and 2026-01-01 (both inclusive).

### Filter by certificate issuer

```powershell
.\Request-SmimeCertificate.ps1 -ValidFromDate 2026-01-01 -ProviderName "SwissSign Gold ID" -Issuer "SwissSign"
```

Only considers certificates whose issuer DN contains `SwissSign`. Addresses with no SwissSign cert are skipped.

### Remote NSP with API key

```powershell
.\Request-SmimeCertificate.ps1 `
    -ValidFromDate 2026-01-01 `
    -ProviderName "SwissSign Gold ID" `
    -NspUrl https://nsp-mgmt.contoso.local:6060 `
    -ApiKey (Read-Host -AsSecureString 'API Key')
```

### Dry run

```powershell
.\Request-SmimeCertificate.ps1 -ValidFromDate 2026-01-01 -ProviderName "SwissSign Gold ID" -WhatIf
```

Connects to NSP, validates the provider name, applies all filters, then prints the full list of affected
addresses (with display name, newest cert `ValidFrom`, and issuer DN) and the total count.
No certificate request is submitted.

## Error handling

| Message | Resolution |
|---|---|
| `NoSpamProxy PowerShell module not found` | Install the NSP PowerShell module |
| `No JWT token found in MemoryCache` | Verify `Connect-Nsp` succeeded |
| `No provider named '...' found` | Check the name against the list printed by the script; spelling is case-sensitive |
| `Unrecognised provider type` | Open a support request; the provider type may be new |

## After execution

Open the NoSpamProxy management console → **Users and Keys** → **Pending Certificate Requests**
to monitor status and download the issued certificates once the CA has approved them.
