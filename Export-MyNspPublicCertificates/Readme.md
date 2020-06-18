# Export-MyNspPublicCertificates.ps1

This script will export all own valid public certificates.  
Certificates are saved under the given path in <CertificateFolder> parameter.  
A file called "ListOfThumbprints.txt" will be saved into the <CertificateFolder> path.   
It contains all exported thumbprints and is used to prevent to export all certificates at every time.  
  
## Usage

`.\Export-MyNspPublicCertificates.ps1 -CertificateFolder "PATH_TO_FOLDER"`

## Example

`.\Export-MyNspPublicCertificates.ps1 -CertificateFolder "C:\certs"`
