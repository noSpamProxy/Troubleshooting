# Check-NspCertificates.ps1

Outputs the state of a certificate and its chain to the console.
It is only possible to check certificates which are imported into the NoSpamProxy.
The validation is done for each connected gateway role.

## Usage 

```ps
Check-NspCertificates.ps1 -Thumbprint
```

## Parameters
### TenantPrimaryDomain
    Connect to a specific tenant using the given primary domain name.

### Thumbprint
    Set the thumbprint of the certificate which should be checked.
	
## Examples
### Examples 1

```ps
.\Check-NspCertificates.ps1 -Thumbprint 0F4F9209E172B6D81022C0219CF253EFD29689F6
```
### Example 2
```ps
.\CheckNspCertificates.ps1 -Thumbprint 0F4F9209E172B6D81022C0219CF253EFD29689F6 -TenantPrimaryDomain "example.com"
```