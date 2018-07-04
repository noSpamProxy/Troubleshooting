<#
.SYNOPSIS
    Publish one or more certificates to OpenKeys (https://openkeys.de)
.PARAMETER CertificateFileInfo
    A FileInfo pointing to the file containing a public certificate. Commonly used in conjection with Get-ChildItem.
.PARAMETER FileName
    The filename of a public certificate.
.PARAMETER Thumbprint
    The thumbprint of a certifacate stored in NoSpamProxy. The thumbprint can be obtained by calling Get-NspCertificate
.EXAMPLE
    C:\PS> Get-NspCertificate -KeyValidity Valid -StoreIds My | .\Publish-CertificateToOpenKeys.ps1
    Publishes all private certificates to OpenKeys.
    Note: The Private key is NOT published. Only the public part of the certificate is.
.EXAMPLE
    C:\PS> Get-ChildItem c:\certs -Filter *.cer | .\Publish-CertificateToOpenKeys.ps1
    Publishes the certificate stored in the directory c:\certs to OpenKeys.
.EXAMPLE
    C:\PS> .\Publish-CertificateToOpenKeys.ps1 -FileName c:\certs\mycert.cer
    Publishes the certificate c:\certs\mycert.cer to OpenKeys.
.EXAMPLE
    C:\PS> .\Publish-CertificateToOpenKeys.ps1 -Thumbprint A2782D3679FB89677089247AC2B1FD81F561688E
    Publishes the certificate with the thumbprint A2782D3679FB89677089247AC2B1FD81F561688E from NoSpamProxy to OpenKeys.
.NOTES
    Author: Henning Krause
    Date:   July 4th, 2018   
#>
param(
    [parameter(
        Mandatory         = $true,
        ValueFromPipeline = $true,
        ParameterSetName = "FileInfo"
        )]
    [System.IO.FileInfo] $CertificateFileInfo,
    [parameter(
        Mandatory         = $true,
        ValueFromPipeline = $true,
        ParameterSetName = "FileName"
        )]
    [string] $FileName,
    [parameter(
        Mandatory         = $true,
        ValueFromPipelineByPropertyName = $true,
        ParameterSetName = "Thumbprint"
        
        )]
    [string] $Thumbprint
)

Begin {
    $license = Get-NspLicense
    if ($license -eq $null) {
        throw "Failed to get your current license from NoSpamProxy."
    }
    $apiKey = $license.ApiKey
    if ($apiKey -eq $null) {
        throw "You'll need an updated license. Please contact Net at Work."
    }

    $publishUrl = "https://api.openkeys.de/api/certificate/publish"
}

Process {

    function Get-ResponseBody($response) {
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd()
        if ($ResponseBody.StartsWith('{')) {
            $responseBody = ConvertFrom-Json $ResponseBody 
        }
        return $responseBody
    }

    if ($Thumbprint) {
        $cert = (Export-Nspcertificate $thumbprint)
        $certData = [Convert]::ToBase64String($cert.Certificate)
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ($cert.Certificate, "cer")
    }
    elseif ($CertificateFileInfo) {
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ($CertificateFileInfo.FullName)
        $certData = [Convert]::ToBase64String($certificate.GetRawCertData())
    }
    elseif ($FileName) {
        $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList ($FileName)
        $certData = [Convert]::ToBase64String($certificate.GetRawCertData())
    }

    $requestObject = @{}
    $requestObject.FullName = $certificate.Subject
    $requestObject.Overwrite = $false
    $requestObject.SerializedCertificate = $certData

    $request = ConvertTo-Json $requestObject

    try {
        $result = Invoke-RestMethod -Uri $publishUrl -Headers @{"Authorization"="BEARER $apikey"} -Method POST -ContentType "application/json; charset=utf-8" -Body ([System.Text.Encoding]::UTF8.GetBytes($request))
        if ($result.Response.StatusCode -eq "Created") {
            $result = new-object -TypeName PSObject -Property  @{"SubjectName"=$certificate.Subject; "Thumbprint"=$certificate.Thumbprint; "Status"="Ok"}
            Write-Output $result        
        }
    }
    catch {
        $response = $_.Exception.Response
        if ($response) {
            switch ($response.StatusCode) {
                "BadRequest" { 
                    $errorCode ="InvalidCertificate" 
                    $responseBody = Get-ResponseBody $response
                }
                "Conflict" {
                    $responseBody = $null;
                    $errorCode ="AlreadyPresent"}
                Default { 
                    $errorCode = "UnknownError"
                    $responseBody = Get-ResponseBody $response
                }
            }
            if ($response.StatusCode -eq "BadRequest") {
                $errorCode = "InvalidCertificate"
            }

        }
        else {
            $errorCode = "UnknownError"
            $responseBody = $_.Exception.Message
        }  
        $result = new-object -TypeName PSObject -Property  @{"SubjectName"=$certificate.Subject; "Thumbprint"=$certificate.Thumbprint; "Status"=$errorCode; "Response"=$responseBody}
        Write-Output $result        
    }
}

End {

}

# SIG # Begin signature block
# MIIMSwYJKoZIhvcNAQcCoIIMPDCCDDgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUTivsk7hLzH9oUuED9lLRZvTB
# EV2gggmqMIIElDCCA3ygAwIBAgIOSBtqBybS6D8mAtSCWs0wDQYJKoZIhvcNAQEL
# BQAwTDEgMB4GA1UECxMXR2xvYmFsU2lnbiBSb290IENBIC0gUjMxEzARBgNVBAoT
# Ckdsb2JhbFNpZ24xEzARBgNVBAMTCkdsb2JhbFNpZ24wHhcNMTYwNjE1MDAwMDAw
# WhcNMjQwNjE1MDAwMDAwWjBaMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFs
# U2lnbiBudi1zYTEwMC4GA1UEAxMnR2xvYmFsU2lnbiBDb2RlU2lnbmluZyBDQSAt
# IFNIQTI1NiAtIEczMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjYVV
# I6kfU6/J7TbCKbVu2PlC9SGLh/BDoS/AP5fjGEfUlk6Iq8Zj6bZJFYXx2Zt7G/3Y
# SsxtToZAF817ukcotdYUQAyG7h5LM/MsVe4hjNq2wf6wTjquUZ+lFOMQ5pPK+vld
# sZCH7/g1LfyiXCbuexWLH9nDoZc1QbMw/XITrZGXOs5ynQYKdTwfmOPLGC+MnwhK
# kQrZ2TXZg5J2Yl7fg67k1gFOzPM8cGFYNx8U42qgr2v02dJsLBkwXaBvUt/RnMng
# Ddl1EWWW2UO0p5A5rkccVMuxlW4l3o7xEhzw127nFE2zGmXWhEpX7gSvYjjFEJtD
# jlK4PrauniyX/4507wIDAQABo4IBZDCCAWAwDgYDVR0PAQH/BAQDAgEGMB0GA1Ud
# JQQWMBQGCCsGAQUFBwMDBggrBgEFBQcDCTASBgNVHRMBAf8ECDAGAQH/AgEAMB0G
# A1UdDgQWBBQPOueslJF0LZYCc4OtnC5JPxmqVDAfBgNVHSMEGDAWgBSP8Et/qC5F
# JK5NUPpjmove4t0bvDA+BggrBgEFBQcBAQQyMDAwLgYIKwYBBQUHMAGGImh0dHA6
# Ly9vY3NwMi5nbG9iYWxzaWduLmNvbS9yb290cjMwNgYDVR0fBC8wLTAroCmgJ4Yl
# aHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9yb290LXIzLmNybDBjBgNVHSAEXDBa
# MAsGCSsGAQQBoDIBMjAIBgZngQwBBAEwQQYJKwYBBAGgMgFfMDQwMgYIKwYBBQUH
# AgEWJmh0dHBzOi8vd3d3Lmdsb2JhbHNpZ24uY29tL3JlcG9zaXRvcnkvMA0GCSqG
# SIb3DQEBCwUAA4IBAQAVhCgM7aHDGYLbYydB18xjfda8zzabz9JdTAKLWBoWCHqx
# mJl/2DOKXJ5iCprqkMLFYwQL6IdYBgAHglnDqJQy2eAUTaDVI+DH3brwaeJKRWUt
# TUmQeGYyDrBowLCIsI7tXAb4XBBIPyNzujtThFKAzfCzFcgRCosFeEZZCNS+t/9L
# 9ZxqTJx2ohGFRYzUN+5Q3eEzNKmhHzoL8VZEim+zM9CxjtEMYAfuMsLwJG+/r/uB
# AXZnxKPo4KvcM1Uo42dHPOtqpN+U6fSmwIHRUphRptYCtzzqSu/QumXSN4NTS35n
# fIxA9gccsK8EBtz4bEaIcpzrTp3DsLlUo7lOl8oUMIIFDjCCA/agAwIBAgIMUfr8
# J+jCyr4Ay7YNMA0GCSqGSIb3DQEBCwUAMFoxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMTAwLgYDVQQDEydHbG9iYWxTaWduIENvZGVTaWdu
# aW5nIENBIC0gU0hBMjU2IC0gRzMwHhcNMTYwNzI4MTA1NjE3WhcNMTkwNzI5MTA1
# NjE3WjCBhzELMAkGA1UEBhMCREUxDDAKBgNVBAgTA05SVzESMBAGA1UEBxMJUGFk
# ZXJib3JuMRkwFwYDVQQKExBOZXQgYXQgV29yayBHbWJIMRkwFwYDVQQDExBOZXQg
# YXQgV29yayBHbWJIMSAwHgYJKoZIhvcNAQkBFhFpbmZvQG5ldGF0d29yay5kZTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJWtx+QDzgovn6AmkJ8UCTNr
# xtFJbRCHKNkfev6k35mMkNlibsVnFxooABDKSvaB21nXojMz63g+KLUEN5S4JiX3
# FKq5h2XahwWHvar/r2HMK2uJZ76360ePhuSZTnkifsxvwNxByQ9ot2S1O40AyVU5
# xfEUsBh7vVADMbjqBVlXuNAfsfpfvgjoR0CsOfgKk0CEDZ1wP0bXIkrk021a7lAO
# Yq9kqVDFv8K8O5WYvNcvbtAg3QW5JEaFnM3TMaOOSaWZMmIo7lw3e+B8rqknwmcS
# 66W2E0uayJXKqh/SXfS/xCwO2EzBT9Q1x0XiFR1LlEHQ0T/tfenBUlefIxfDZnEC
# AwEAAaOCAaQwggGgMA4GA1UdDwEB/wQEAwIHgDCBlAYIKwYBBQUHAQEEgYcwgYQw
# SAYIKwYBBQUHMAKGPGh0dHA6Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0
# L2dzY29kZXNpZ25zaGEyZzNvY3NwLmNydDA4BggrBgEFBQcwAYYsaHR0cDovL29j
# c3AyLmdsb2JhbHNpZ24uY29tL2dzY29kZXNpZ25zaGEyZzMwVgYDVR0gBE8wTTBB
# BgkrBgEEAaAyATIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFsc2ln
# bi5jb20vcmVwb3NpdG9yeS8wCAYGZ4EMAQQBMAkGA1UdEwQCMAAwPwYDVR0fBDgw
# NjA0oDKgMIYuaHR0cDovL2NybC5nbG9iYWxzaWduLmNvbS9nc2NvZGVzaWduc2hh
# MmczLmNybDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUZLedJVdZSZd5
# lwNJFEgIc8KbEFEwHwYDVR0jBBgwFoAUDzrnrJSRdC2WAnODrZwuST8ZqlQwDQYJ
# KoZIhvcNAQELBQADggEBADYcz/+SCP59icPJK5w50yiTcoxnOtoA21GZDpt4GGVf
# RQJDWCDJMkU62xwu5HzqwimbwmBykrAf5Log1fLbggI83zIE4sMjkUe/BnnHpHgK
# LYv+3eLEwglMw/6Gmlq9IqNSD8YmTncGZFoFhrCrgAZUkA6RiVxuZrx2wiluueBI
# vfGs+tRA+7Tgx6Ed9kBybnc+xbAiTCNIcSo9OkPZfc3Q9saMgjIehBMXHLgMdrhv
# N5HXv/r4+aZ6asgv3ggArHrS1Pxp0f60hooVK4bA4Ph1td6YZ5lf8HA4uMmHvOjQ
# iNS0UjXqu5Vs6leIRM3pBjuX45xL6ydUsMlLhZQfansxggILMIICBwIBATBqMFox
# CzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9iYWxTaWduIG52LXNhMTAwLgYDVQQD
# EydHbG9iYWxTaWduIENvZGVTaWduaW5nIENBIC0gU0hBMjU2IC0gRzMCDFH6/Cfo
# wsq+AMu2DTAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZ
# BgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYB
# BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUIsIMh9DFoSza/ZlagW3eYlwx48owDQYJ
# KoZIhvcNAQEBBQAEggEAZoHqXpbGppRUEmx4B30hc5CLx3ejqASl2i+Kmgfy65HW
# cfYjfHC27GeraqQJurt8hLoirVTYVHGmchdGXEsw+W79fXTG8bxGT/6KBvT6geJ+
# gNKKqrvCBd40ycNDC9hAQyYvM8nZDcCRav02SxSr1UbPrtVe3WCHajQ2ZsIa7iPl
# bAlu2inO04am0tt5VAL/5+bAjqLuwm0ZbTsnOGYk+IsO1+KfJx5R8yMJtF/mfQB8
# a62MOSSMTbtF95Hy98eOTUmQNw8ZqMiID/D8T6HMEw7P0TJVlKT3cy7mnVpOzFHQ
# vb+FuNK42Xw79+plTAVv3X6ITWiAkbahbJZn8tSIvA==
# SIG # End signature block
