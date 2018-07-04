# Publish-CertificateToOpenKeys

```
Get-Help .\Publish-CertificateToOpenKeys.ps1 -Detailed

NAME
    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1

SYNOPSIS
    Publish one or more certificates to OpenKeys (https://openkeys.de)


SYNTAX
    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1 -CertificateFileInfo <FileInfo> [<CommonParameters>]

    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1 -FileName <String> [<CommonParameters>]

    C:\src\nsp-public-troubleshooting\Publish-CertificateToOpenKeys\Publish-CertificateToOpenKeys.ps1 -Thumbprint <String> [<CommonParameters>]


DESCRIPTION


PARAMETERS
    -CertificateFileInfo <FileInfo>
        A FileInfo pointing to the file containing a public certificate. Commonly used in conjection with Get-ChildItem.

    -FileName <String>
        The filename of a public certificate.

    -Thumbprint <String>
        The thumbprint of a certifacate stored in NoSpamProxy. The thumbprint can be obtained by calling Get-NspCertificate

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216).

    -------------------------- EXAMPLE 1 --------------------------

    C:\PS>Get-NspCertificate -KeyValidity Valid -StoreIds My | .\Publish-CertificateToOpenKeys.ps1

    Publishes all private certificates to OpenKeys.
    Note: The Private key is NOT published. Only the public part of the certificate is.

    -------------------------- EXAMPLE 2 --------------------------

    C:\PS>Get-ChildItem c:\certs -Filter *.cer | .\Publish-CertificateToOpenKeys.ps1

    Publishes the certificate stored in the directory c:\certs to OpenKeys.

    -------------------------- EXAMPLE 3 --------------------------

    C:\PS>.\Publish-CertificateToOpenKeys.ps1 -FileName c:\certs\mycert.cer

    Publishes the certificate c:\certs\mycert.cer to OpenKeys.

    -------------------------- EXAMPLE 4 --------------------------

    C:\PS>.\Publish-CertificateToOpenKeys.ps1 -Thumbprint A2782D3679FB89677089247AC2B1FD81F561688E

    Publishes the certificate with the thumbprint A2782D3679FB89677089247AC2B1FD81F561688E from NoSpamProxy to OpenKeys.

```