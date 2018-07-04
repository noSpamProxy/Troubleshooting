# Publish-CertificateToOpenKeys

Publish one or more certificates to [OpenKeys](https://openkeys.de)

## Syntax

```powershell
    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1 -CertificateFileInfo <FileInfo> [<CommonParameters>]

    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1 -FileName <String> [<CommonParameters>]

    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1 -Thumbprint <String> [<CommonParameters>]
```

## Parameters

- **CertificateFileInfo** \<FileInfo>: A FileInfo pointing to the file containing a public certificate. Commonly used in conjection with Get-ChildItem.
- **FileName** \<String>: The filename of a public certificate.
- **Thumbprint** \<String>: The thumbprint of a certificate stored in NoSpamProxy. The thumbprint can be obtained by calling Get-NspCertificate

## Examples

```powershell
C:\PS>Get-NspCertificate -KeyValidity Valid -StoreIds My | .\Publish-CertificateToOpenKeys.ps1
```

Publishes all private certificates to OpenKeys.
Note: The Private key is NOT published. Only the public part of the certificate is.

```powershell
C:\PS>Get-ChildItem c:\certs -Filter *.cer | .\Publish-CertificateToOpenKeys.ps1
```

Publishes the certificate stored in the directory c:\certs to OpenKeys.

```powershell
C:\PS>.\Publish-CertificateToOpenKeys.ps1 -FileName c:\certs\mycert.cer
```

Publishes the certificate c:\certs\mycert.cer to OpenKeys.

```powershell
C:\PS>.\Publish-CertificateToOpenKeys.ps1 -Thumbprint A2782D3679FB89677089247AC2B1FD81F561688E
```

Publishes the certificate with the thumbprint A2782D3679FB89677089247AC2B1FD81F561688E from NoSpamProxy to OpenKeys.