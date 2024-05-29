<#
.SYNOPSIS
  Name: Export-MyNspPublicCertificates.ps1
  Exports all own valid certifcifates as PFX with a random password.

.DESCRIPTION
  This script will export all own valid public certificates.

.PARAMETER CertificateFolder
  Mandatory parameter
  Input the full path to a folder where the certificates should be saved.

.OUTPUTS
  Certificates are saved under the given path in <CertificateFolder> parameter.
  A file called "ListOfThumbprints.txt" will be saved into the <CertificateFolder> path. 
  It contains all exported thumbprints and is used to prevent to export all certificates at every time.

.NOTES
  Version:        1.0.1
  Author:         Jan Jaeschke
  Creation Date:  2024-05-29
  Purpose/Change: added v14 compatibility
  
.LINK
  https://www.nospamproxy.de
  https://forum.nospamproxy.com
  https://github.com/noSpamProxy

.EXAMPLE
  .\Export-MyNspPublicCertificates.ps1 -CertificateFolder "C:\certs" 
#>
param (
  [Parameter(Mandatory = $true)][string] $CertificateFolder,
  # only needed for v14 with enabled provider mode
	[Parameter(Mandatory=$false)][string] $TenantPrimaryDomain
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
      $cerFile = "$CertificateFolder\$certMail-$certThumbprint.cer"
      # generates a random 16 char password, 4 chars have to be not alphanumeric
      $bytetCertificate = (Export-NspCertificate -Thumbprint $certThumbprint).Certificate
      [io.file]::WriteAllBytes("$cerFile", $bytetCertificate)
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
$getCertificates = $true
$skipCertificates = 0

while ($getCertificates -eq $true) {	
  if ($skipCertificates -eq 0) {
    # get all own certificate thubmbprints
    $myCertificates = (Get-NspCertificate -StoreIds My -KeyValidity Valid -First 100)
  }
  else {
    $myCertificates = (Get-NspCertificate -StoreIds My -KeyValidity Valid -First 100 -Skip $skipCertificates)
  }

  processCertificates $myCertificates

  # exit condition
  if ($certificates) {
    $skipCertificates = $skipCertificates + 100
    Write-Host $skipCertificates
  }
  else {
    $getCertificates = $false
    break
  }
}
# SIG # Begin signature block
# MIIoEQYJKoZIhvcNAQcCoIIoAjCCJ/4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBDJArEYTosOrSC
# A/b6DbAly/Bo4sTSmsUVG7oWiH0bU6CCISgwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMy
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXH
# JQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMf
# UBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w
# 1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRk
# tFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYb
# qMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUm
# cJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP6
# 5x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzK
# QtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo
# 80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjB
# Jgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXche
# MBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd
# 4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiC
# qBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl
# /Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeC
# RK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYT
# gAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/
# a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37
# xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmL
# NriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0
# YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJ
# RyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIG
# wjCCBKqgAwIBAgIQBUSv85SdCDmmv9s/X+VhFjANBgkqhkiG9w0BAQsFADBjMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRp
# Z2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENB
# MB4XDTIzMDcxNDAwMDAwMFoXDTM0MTAxMzIzNTk1OVowSDELMAkGA1UEBhMCVVMx
# FzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdEaWdpQ2VydCBUaW1l
# c3RhbXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKNTRYcd
# g45brD5UsyPgz5/X5dLnXaEOCdwvSKOXejsqnGfcYhVYwamTEafNqrJq3RApih5i
# Y2nTWJw1cb86l+uUUI8cIOrHmjsvlmbjaedp/lvD1isgHMGXlLSlUIHyz8sHpjBo
# yoNC2vx/CSSUpIIa2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaaPPDFLnkPG2ZS48jW
# Pl/aQ9OE9dDH9kgtXkV1lnX+3RChG4PBuOZSlbVH13gpOWvgeFmX40QrStWVzu8I
# F+qCZE3/I+PKhu60pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyyt1/f8O52fTGZZUdV
# nUokL6wrl76f5P17cz4y7lI0+9S769SgLDSb495uZBkHNwGRDxy1Uc2qTGaDiGhi
# u7xBG3gZbeTZD+BYQfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2dD9/O0V6MqqtQFcmz
# yrzXxDtoRKOlO0L9c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TWdc2PEnZWpST618Rr
# IbroHzSYLzrqawGw9/sqhux7UjipmAmhcbJsca8+uG+W1eEQE/5hRwqM/vC2x9XH
# 3mwk8L9CgsqgcT2ckpMEtGlwJw1Pt7U20clfCKRwo+wK8REuZODLIivK8SgTIUlR
# fgZm0zu++uuRONhRB8qUt+JQofM604qDy0B7AgMBAAGjggGLMIIBhzAOBgNVHQ8B
# Af8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAg
# BgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZ
# bU2FL3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFKW27xPn783QZKHVVqllMaPe1eNJ
# MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAG
# CCsGAQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2Vy
# dC5jb20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9E
# aWdpQ2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQw
# DQYJKoZIhvcNAQELBQADggIBAIEa1t6gqbWYF7xwjU+KPGic2CX/yyzkzepdIpLs
# jCICqbjPgKjZ5+PF7SaCinEvGN1Ott5s1+FgnCvt7T1IjrhrunxdvcJhN2hJd6Pr
# kKoS1yeF844ektrCQDifXcigLiV4JZ0qBXqEKZi2V3mP2yZWK7Dzp703DNiYdk9W
# uVLCtp04qYHnbUFcjGnRuSvExnvPnPp44pMadqJpddNQ5EQSviANnqlE0PjlSXcI
# WiHFtM+YlRpUurm8wWkZus8W8oM3NG6wQSbd3lqXTzON1I13fXVFoaVYJmoDRd7Z
# ULVQjK9WvUzF4UbFKNOt50MAcN7MmJ4ZiQPq1JE3701S88lgIcRWR+3aEUuMMsOI
# 5ljitts++V+wQtaP4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKWfpd5OEhfysLcPTLf
# ddY2Z1qJ+Panx+VPNTwAvb6cKmx5AdzaROY63jg7B145WPR8czFVoIARyxQMfq68
# /qTreWWqaNYiyjvrmoI1VygWy2nyMpqy0tg6uLFGhmu6F/3Ed2wVbK6rr3M66ElG
# t9V/zLY4wNjsHPW2obhDLN9OTH0eaHDAdwrUAuBcYLso/zjlUlrWrBciI0707NMX
# +1Br/wd3H3GXREHJuEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee8PBf4NvZrZCARK+A
# EEGKMIIG5jCCBM6gAwIBAgIQd70OA6G3CPhUqwZyENkERzANBgkqhkiG9w0BAQsF
# ADBTMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEpMCcG
# A1UEAxMgR2xvYmFsU2lnbiBDb2RlIFNpZ25pbmcgUm9vdCBSNDUwHhcNMjAwNzI4
# MDAwMDAwWhcNMzAwNzI4MDAwMDAwWjBZMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQ
# R2xvYmFsU2lnbiBudi1zYTEvMC0GA1UEAxMmR2xvYmFsU2lnbiBHQ0MgUjQ1IENv
# ZGVTaWduaW5nIENBIDIwMjAwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoIC
# AQDWQk3540/GI/RsHYGmMPdIPc/Q5Y3lICKWB0Q1XQbPDx1wYOYmVPpTI2ACqF8C
# AveOyW49qXgFvY71TxkkmXzPERabH3tr0qN7aGV3q9ixLD/TcgYyXFusUGcsJU1W
# Bjb8wWJMfX2GFpWaXVS6UNCwf6JEGenWbmw+E8KfEdRfNFtRaDFjCvhb0N66WV8x
# r4loOEA+COhTZ05jtiGO792NhUFVnhy8N9yVoMRxpx8bpUluCiBZfomjWBWXACVp
# 397CalBlTlP7a6GfGB6KDl9UXr3gW8/yDATS3gihECb3svN6LsKOlsE/zqXa9Fko
# jDdloTGWC46kdncVSYRmgiXnQwp3UrGZUUL/obLdnNLcGNnBhqlAHUGXYoa8qP+i
# x2MXBv1mejaUASCJeB+Q9HupUk5qT1QGKoCvnsdQQvplCuMB9LFurA6o44EZqDjI
# ngMohqR0p0eVfnJaKnsVahzEaeawvkAZmcvSfVVOIpwQ4KFbw7MueovE3vFLH4wo
# eTBFf2wTtj0s/y1KiirsKA8tytScmIpKbVo2LC/fusviQUoIdxiIrTVhlBLzpHLr
# 7jaep1EnkTz3ohrM/Ifll+FRh2npIsyDwLcPRWwH4UNP1IxKzs9jsbWkEHr5DQwo
# sGs0/iFoJ2/s+PomhFt1Qs2JJnlZnWurY3FikCUNCCDx/wIDAQABo4IBrjCCAaow
# DgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBIGA1UdEwEB/wQI
# MAYBAf8CAQAwHQYDVR0OBBYEFNqzjcAkkKNrd9MMoFndIWdkdgt4MB8GA1UdIwQY
# MBaAFB8Av0aACvx4ObeltEPZVlC7zpY7MIGTBggrBgEFBQcBAQSBhjCBgzA5Bggr
# BgEFBQcwAYYtaHR0cDovL29jc3AuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25pbmdy
# b290cjQ1MEYGCCsGAQUFBzAChjpodHRwOi8vc2VjdXJlLmdsb2JhbHNpZ24uY29t
# L2NhY2VydC9jb2Rlc2lnbmluZ3Jvb3RyNDUuY3J0MEEGA1UdHwQ6MDgwNqA0oDKG
# MGh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vY29kZXNpZ25pbmdyb290cjQ1LmNy
# bDBWBgNVHSAETzBNMEEGCSsGAQQBoDIBMjA0MDIGCCsGAQUFBwIBFiZodHRwczov
# L3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAIBgZngQwBBAEwDQYJKoZI
# hvcNAQELBQADggIBAAiIcibGr/qsXwbAqoyQ2tCywKKX/24TMhZU/T70MBGfj5j5
# m1Ld8qIW7tl4laaafGG4BLX468v0YREz9mUltxFCi9hpbsf/lbSBQ6l+rr+C1k3M
# EaODcWoQXhkFp+dsf1b0qFzDTgmtWWu4+X6lLrj83g7CoPuwBNQTG8cnqbmqLTE7
# z0ZMnetM7LwunPGHo384aV9BQGf2U33qQe+OPfup1BE4Rt886/bNIr0TzfDh5uUz
# oL485HjVG8wg8jBzsCIc9oTWm1wAAuEoUkv/EktA6u6wGgYGnoTm5/DbhEb7c9kr
# QrbJVzTHFsCm6yG5qg73/tvK67wXy7hn6+M+T9uplIZkVckJCsDZBHFKEUtaZMO8
# eHitTEcmZQeZ1c02YKEzU7P2eyrViUA8caWr+JlZ/eObkkvdBb0LDHgGK89T2L0S
# mlsnhoU/kb7geIBzVN+nHWcrarauTYmAJAhScFDzAf9Eri+a4OFJCOHhW9c40Z4K
# ip2UJ5vKo7nb4jZq42+5WGLgNng2AfrBp4l6JlOjXLvSsuuKy2MIL/4e81Yp4jWb
# 2P/ppb1tS1ksiSwvUru1KZDaQ0e8ct282b+Awdywq7RLHVg2N2Trm+GFF5opov3m
# CNKS/6D4fOHpp9Ewjl8mUCvHouKXd4rv2E0+JuuZQGDzPGcMtghyKTVTgTTcMIIH
# MTCCBRmgAwIBAgIMMFidGzNEBznAfRFWMA0GCSqGSIb3DQEBCwUAMFkxCzAJBgNV
# BAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMS8wLQYDVQQDEyZHbG9i
# YWxTaWduIEdDQyBSNDUgQ29kZVNpZ25pbmcgQ0EgMjAyMDAeFw0yMjAzMjgxMzQz
# MzhaFw0yNTAzMjgxMzQzMzhaMIGeMQswCQYDVQQGEwJERTEcMBoGA1UECBMTTm9y
# ZHJoZWluLVdlc3RmYWxlbjESMBAGA1UEBxMJUGFkZXJib3JuMRkwFwYDVQQKExBO
# ZXQgYXQgV29yayBHbWJIMRkwFwYDVQQDExBOZXQgYXQgV29yayBHbWJIMScwJQYJ
# KoZIhvcNAQkBFhhzYWxlc29mZmljZUBuZXRhdHdvcmsuZGUwggIiMA0GCSqGSIb3
# DQEBAQUAA4ICDwAwggIKAoICAQC6eJO4sCjJF/nHirVF+xe/0Dl0pRbFbzzKlb+O
# fK92xzWOClez2XkIDi1VIG333L0+OpCpcGoBuJKhWRgRIqZ2GBbyWrfnyVO2k/sf
# IqbfYufo8FbK4NvDozQpyMIvyncT6fNDSRRGv5LJMY0Z4Rah6SlAHQ2J0qNTrfEW
# JOLdRc+a4nNtYW0Qr2tgw6Ays4wzowA40cUQOQ8BR3txF4xisBh9asmyUwzSdY1E
# TE+8GNfDxDtv5xv5EyHUF5nJcQbQhZ4/vuyxoI3iK+Gq/f84NNT0qNevIK9D+k3w
# H07qc+lIJdHou6/ies9oRNUNQvG03zVbYxLp/cuVaoxt6itI0YbX6M1VnEQEooSL
# mQ2q+4j9LyTc/2ZjkGBV6NHTcxv1iqvLCCW4jYrZH+mpdMSxdOEgRCr8fxo8yprw
# H2Aj4kCm4vlYcZAyL3I8UBoucRgc4l3DwfOdjv/orfdPEgez+uokEgQpuxrfUrfX
# XQStHCdypQBL4pu2mKfkmbdasRiBq3VBmy4fUBohYRZxFi/sLuRgKY3Dv8T2Polm
# AtADbdMQzvWuSq5y1jhmeSxzWMqkk34Vl2HdYx6tGOTAulKsZ/e3wtW6A2sGKawz
# YNywwyqtOYPMqqPgWwehSwyXiCLUWedlhjapHrE6IprFHYcYCTRppdPlUdXRaWjh
# mcAknQIDAQABo4IBsTCCAa0wDgYDVR0PAQH/BAQDAgeAMIGbBggrBgEFBQcBAQSB
# jjCBizBKBggrBgEFBQcwAoY+aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNvbS9j
# YWNlcnQvZ3NnY2NyNDVjb2Rlc2lnbmNhMjAyMC5jcnQwPQYIKwYBBQUHMAGGMWh0
# dHA6Ly9vY3NwLmdsb2JhbHNpZ24uY29tL2dzZ2NjcjQ1Y29kZXNpZ25jYTIwMjAw
# VgYDVR0gBE8wTTBBBgkrBgEEAaAyATIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93
# d3cuZ2xvYmFsc2lnbi5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMAkGA1UdEwQC
# MAAwRQYDVR0fBD4wPDA6oDigNoY0aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9n
# c2djY3I0NWNvZGVzaWduY2EyMDIwLmNybDATBgNVHSUEDDAKBggrBgEFBQcDAzAf
# BgNVHSMEGDAWgBTas43AJJCja3fTDKBZ3SFnZHYLeDAdBgNVHQ4EFgQUD1XYzs8P
# 9KbTTxB+NQjZBCU2FpswDQYJKoZIhvcNAQELBQADggIBAL8f73n5oalZk1ParbVZ
# 8gI8RBqhitBHb5pwwbtmvEBI+Sm+0u3TY7aWGeyUfyf2CC8QH77bhPuvn8hRjNWU
# N1FPMYSbG2gXxMgHaq/R+Bl44m/TTc84L+nMt99JHPyLAJrtlYZgmJW5NYRotPYB
# VM0XGmC/md1eg6OOesHTHej6aXZLW/6vhcwOHwAVmRk4467qLg5/g5wBipQ/ge3Q
# /tphpNv/LQF/VhGPIt2bb3tVsaOlAi97+RRD4XeduqtBLPBVZapgUq9Gg0Irk7fm
# nTt/7rzplK0jTpVO+6e6JqOzxHtv9lo2+PYP4IDvvj6FQgt3sBktR7RVwJ1/X3RU
# 4eTs95V8nFCGNWm21K62TECV2idFm98cFxbE1VG1VynKsYgWIyd9uDFwuC2lxWAh
# Rr3kkKUZ2p8+pDQR0k2TDJLxv596ddc9CYeZij8bfkpIF62vxwQUFCo4ce6CFEDm
# MczWExs+LlP3XLiX7H4akThb/CBDjR8Uqoc9Z+ZxMYwi9T6JVUvlNvJ82h+rHawn
# b2A2QFobm6lkVUolaUaImZX6Ca7MP+eP3TuShipyqdU931rRuy2ltPtN034D3qjD
# +jfiwJM/uhgT2HREtAz9w8lK0bRrf7vUUn40oeN+Qpv+i2OHO6pD14ue2hrWgBFO
# h9M5ws6vPbqMIIWnZeQEmHUMMYIGPzCCBjsCAQEwaTBZMQswCQYDVQQGEwJCRTEZ
# MBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEvMC0GA1UEAxMmR2xvYmFsU2lnbiBH
# Q0MgUjQ1IENvZGVTaWduaW5nIENBIDIwMjACDDBYnRszRAc5wH0RVjANBglghkgB
# ZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJ
# AzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8G
# CSqGSIb3DQEJBDEiBCBQqxTmFMbgEP51rPKuBo/uxGbWEVn0QPoUtpyKzCNFUTAN
# BgkqhkiG9w0BAQEFAASCAgCkNo9cGhGUdco2tNEnrijr4yL5oMT1xqjeZNOwOOlg
# 4zEUFEBpemXkvMWNohft+j0aDi8q87inLCSJ4YYdhC4npvMcmMvbtvujQxQL4qJ8
# lHK8CSH1FfG8cgz2NiSdS9QzBWTdIVDHqPXrNkN2v6pSKETtW10CuPhgQACAsnzR
# XlIrvA+KljsaOHRpsImj0HLGXXiTbVTfJSRf5wphfqYlteljah1e/eBbvdXxeXCg
# 8Nily56C8c2wkbbTYnYd+Cu4TCJoC+qqZ+l6iQuFi53dTxWVqHrLgcgRuBJ0NVQ9
# qCjgBUE5UjVjZ/BgdjS7McS5a5Y3M2cJkrTK2Lp/KT2kasp4yL65BleLY6nev+Cv
# 4cy1zTlsawnCc1mYkKI64/GtviVXzAvwFHx3JF3h09kjhodgZgqlajbTHQrmLyKx
# iF3L/ok1MT5inyOWI/OR0Dd9VNc0z59sz7tRCwcI8x+k7kSutdWilfwr9fkYTGPk
# q8HXcTHQCunf6dbuY1ZCxCXU3O60LRvPXaA6N+CCfBd51+loXSza7/owM9J1oAe6
# Md/kRfYdpYHY4dC4DG/6DnJFsoYIXnyioSPBMKFcTw7g05YkWy3sMZMTb6LV8U53
# Xr8ZWsgoC85RQNVQiMsb6W5YvxZ+wnUv5kAu9hw4npHqtAQNX7Z5Lkf8RYhSgP/e
# BaGCAyAwggMcBgkqhkiG9w0BCQYxggMNMIIDCQIBATB3MGMxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1
# c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0ECEAVEr/OUnQg5
# pr/bP1/lYRYwDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0yNDA1MjkwNjI1MDRaMC8GCSqGSIb3DQEJBDEi
# BCB4x1xNzHPjeygY+zXb4xrByLcY4RyfspqDsWeyPjjwZDANBgkqhkiG9w0BAQEF
# AASCAgBCza/ii2nHK4OAulRuVNerReabdJX9CJegGL+DkJlnUbCmNjBHGUTFrCvp
# jlbk5iGhHuH/5v3WJV70NfMIJSZdtOpjBkZirWUcJVqdM/iQ6vBtasGynh5rFd78
# DfAPPdy7XdnfJvG/L6//7Sfx6qau9AFxIQmWyxJcmn3rlrOj9Vyt4xke2uwg2gba
# nHiR5//i8/pBEKRMhRVeEg6gAkpPhRA3ZVY4Uo1NYfrT/QWK4yVTY1duyPFPSKud
# 8ViiRcmImwcrIS7TiB+PXRQPOHg1jmC3lHVev+zMbRhh5mHlFtOw36EvfLsahxKA
# WnxMTaSkGiebJQeJzMWQjKxfWkkErE4bGAVLopSPDPNTqR/7ZLKO/gIRrwF59os0
# wmRKUdDhDYwfGCx6kNQYS8DNGdHBDeL+rs1SwrS0prbU3LxfOmxe0YsIiEulxbnp
# H5zuey1Rrmhb0CDEDQN9emF845ydPAqPfskpeBBv6VB1ku+w6ujlJjcMAIamj854
# jsxvQmdaBF+GQKk3gyiBRxKdKpzEBvENQdTdwhLagImzhlBtstlOIoX3G9KS9U4v
# /h4v5DmuJyLjco0gSd41/aRukBVvj4tpmE9evzy+cz2eyPN8MDg380cYb7ppIUch
# HV7rUziK87nBVoJEeSzKdAJDVTF0OHj3+dvNRa67rmkZScacmw==
# SIG # End signature block
