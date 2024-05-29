# Export-MyNspCertificates.ps1
This script will export all own valid certificates including the private key.
A random password is generated for each certificate and will be safed in a seperate text file.

## Usage

```ps
Export-MyNspCertificates.ps1 [-CertificateFolder] <String> [-PasswordFolder] <String> [[-TenantPrimaryDomain] <String>]
```

## Parameters
### CertificateFolder
Mandatory parameter  
Input the full path to a folder where the certificates should be saved.  

### PasswordFolder
Mandatory parameter  
Input the full path to a folder where the passwords for the certificates should be saved.  

### TenantPrimaryDomain
Used to login into the desired NoSpamProxy tenant to run this script on.  
Only required if NoSpamProxy v14 is used in provider mode.  

## Outputs
Certificates are saved under the given path in <CertificateFolder> parameter.   
Passwords are saved under the given path in <PasswordFolder> parameter.  
A file called "ListOfThumbprints.txt" will be saved into the <CertificateFolder> path.   
It contains all exported thumbprints and is used to prevent to export all certificates at every time.  

## Examples
### Example 1
```ps 
.\Export-MyNspCertificates.ps1 -CertificateFolder "C:\certs" -PasswordFolder "C:\pw"
```