# Check-NspPrivateCertificates.ps1

Outputs the validation state of all private certificates stored in NoSpamProxy to the console.
The validation is done for each connected gateway role.

## Usage 

```ps
.\Check-NspPrivateCertificates.ps1
```
```
.\Check-NspPrivateCertificates.ps1 -SqlServer -SqlInstance
```
```
.\Check-NspPrivateCertificates.ps1 -SqlServer -SqlInstance -SqlDatabase -SqlCredential
```

## Parameters
### SqlCredential
	Receives a PowerShell PowerShell-Credential-Object for logging in to the database.
    By default, the script will use the credentials of the current user.
    
### SqlDatabase
	Specifies the name of the Database of the Intranet Role the script will use to gather private key thumbprints (default value: "NoSpamProxyAddressSynchronization").

### SqlInstance
    Specifies the SQL-instance name the script will connect to (default value: "NoSpamProxy").

### SqlServer
    Specifies the SQL-server the script will connect to (default value: "(local)").

## Example
```ps
  .\Check-NspPrivateCertificates.ps1
```
```ps
.EXAMPLE
  .\Check-NspPrivateCertificates.ps1 -SqlServer 10.0.0.2 -SqlInstance Mail
```
```ps
.EXAMPLE
	.\Check-NspPrivateCertificates.ps1 -SqlServer 10.0.0.2 -SqlInstance Mail -SqlDatabase NSPIntranet -SqlCredential $Credentials
```
