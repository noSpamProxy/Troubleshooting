$path =$0
$fullpath =$0
$pathstate =$0

function processQueue($path) 
{
	$queue = $path.Name
	$fullpath = $path.FullName
	$pathstate = test-path ($fullpath + "\*")
	if ($pathstate -like "False")
	{
	Remove-Item $fullpath
	}
}
Write-Host "Getting Queue-Path..."

[xml]$gwrconfig = get-content 'C:\ProgramData\Net at Work Mail Gateway\Configuration\Gateway Role.config'
$configpath = $gwrconfig.configuration.'netatwork.nospamproxy.storageconfiguration'.storageLocation.path
$queuepath = ($configpath+"\MailQueues\")
Write-Host "Queue Path is " $queuepath
Write-Host "Looking for empty folders and delete them..."

get-childitem $queuepath | % { processQueue $_ }
Write-Host "Done."
Write-Host "Restarting Gateway Role Service.."
Restart-Service NetatworkMailGatewayGatewayRole
Write-Host "Done."
# SIG # Begin signature block
# MIIMNAYJKoZIhvcNAQcCoIIMJTCCDCECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+n62CuSdulP/yyW7mCZQD0y4
# hHKgggmdMIIEmTCCA4GgAwIBAgIQcaC3NpXdsa/COyuaGO5UyzANBgkqhkiG9w0B
# AQsFADCBqTELMAkGA1UEBhMCVVMxFTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEoMCYG
# A1UECxMfQ2VydGlmaWNhdGlvbiBTZXJ2aWNlcyBEaXZpc2lvbjE4MDYGA1UECxMv
# KGMpIDIwMDYgdGhhd3RlLCBJbmMuIC0gRm9yIGF1dGhvcml6ZWQgdXNlIG9ubHkx
# HzAdBgNVBAMTFnRoYXd0ZSBQcmltYXJ5IFJvb3QgQ0EwHhcNMTMxMjEwMDAwMDAw
# WhcNMjMxMjA5MjM1OTU5WjBMMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMdGhhd3Rl
# LCBJbmMuMSYwJAYDVQQDEx10aGF3dGUgU0hBMjU2IENvZGUgU2lnbmluZyBDQTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJtVAkwXBenQZsP8KK3TwP7v
# 4Ol+1B72qhuRRv31Fu2YB1P6uocbfZ4fASerudJnyrcQJVP0476bkLjtI1xC72Ql
# WOWIIhq+9ceu9b6KsRERkxoiqXRpwXS2aIengzD5ZPGx4zg+9NbB/BL+c1cXNVeK
# 3VCNA/hmzcp2gxPI1w5xHeRjyboX+NG55IjSLCjIISANQbcL4i/CgOaIe1Nsw0Rj
# gX9oR4wrKs9b9IxJYbpphf1rAHgFJmkTMIA4TvFaVcnFUNaqOIlHQ1z+TXOlScWT
# af53lpqv84wOV7oz2Q7GQtMDd8S7Oa2R+fP3llw6ZKbtJ1fB6EDzU/K+KTT+X/kC
# AwEAAaOCARcwggETMC8GCCsGAQUFBwEBBCMwITAfBggrBgEFBQcwAYYTaHR0cDov
# L3QyLnN5bWNiLmNvbTASBgNVHRMBAf8ECDAGAQH/AgEAMDIGA1UdHwQrMCkwJ6Al
# oCOGIWh0dHA6Ly90MS5zeW1jYi5jb20vVGhhd3RlUENBLmNybDAdBgNVHSUEFjAU
# BggrBgEFBQcDAgYIKwYBBQUHAwMwDgYDVR0PAQH/BAQDAgEGMCkGA1UdEQQiMCCk
# HjAcMRowGAYDVQQDExFTeW1hbnRlY1BLSS0xLTU2ODAdBgNVHQ4EFgQUV4abVLi+
# pimK5PbC4hMYiYXN3LcwHwYDVR0jBBgwFoAUe1tFz6/Oy3r9MZIaarbzRutXSFAw
# DQYJKoZIhvcNAQELBQADggEBACQ79degNhPHQ/7wCYdo0ZgxbhLkPx4flntrTB6H
# novFbKOxDHtQktWBnLGPLCm37vmRBbmOQfEs9tBZLZjgueqAAUdAlbg9nQO9ebs1
# tq2cTCf2Z0UQycW8h05Ve9KHu93cMO/G1GzMmTVtHOBg081ojylZS4mWCEbJjvx1
# T8XcCcxOJ4tEzQe8rATgtTOlh5/03XMMkeoSgW/jdfAetZNsRBfVPpfJvQcsVncf
# hd1G6L/eLIGUo/flt6fBN591ylV3TV42KcqF2EVBcld1wHlb+jQQBm1kIEK3Osgf
# HUZkAl/GR77wxDooVNr2Hk+aohlDpG9J+PxeQiAohItHIG4wggT8MIID5KADAgEC
# AhAh36cYPt9rQMtVY5K+Zf5LMA0GCSqGSIb3DQEBCwUAMEwxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwx0aGF3dGUsIEluYy4xJjAkBgNVBAMTHXRoYXd0ZSBTSEEyNTYg
# Q29kZSBTaWduaW5nIENBMB4XDTE1MDkyMTAwMDAwMFoXDTE2MDkyMDIzNTk1OVow
# gZUxCzAJBgNVBAYTAkRFMRwwGgYDVQQIExNOb3JkcmhlaW4gV2VzdGZhbGVuMRIw
# EAYDVQQHFAlQYWRlcmJvcm4xKTAnBgNVBAoUIE5ldCBhdCBXb3JrIE5ldHp3ZXJr
# c3lzdGVtZSBHbWJIMSkwJwYDVQQDFCBOZXQgYXQgV29yayBOZXR6d2Vya3N5c3Rl
# bWUgR21iSDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAMIsja4vgLIG
# rpdvUkdCsS8HCjLwaFXt8TIXG8NYIed1aaG+tV0cmScVlsVRUSRfdKVlaTrg7ZDa
# v17t5rFle0fI8XlaMTt86mp8ujdo+svKpHSXiWL51LiADwRETzqIQfUXkdZqgXGg
# wBrTu0zzIH6NvRm7o7o43sSw5rHTHyKPJUDNEE+gAfPsH/69xDmMuH/2r6iMe5GZ
# dRyAmEtB+sEOdhCIX45gXCEGtc3lPeUDCi4I0P6+oqwHzmgfh3IIBF/PCda4V8yP
# lk65x3+6X1eNox3hWQxNQX2cOx1Yd8yaH9ZYdY8y+RwYauaiGOhzf5XvQtfuka6P
# GR270YqN7/ECAwEAAaOCAY4wggGKMAkGA1UdEwQCMAAwHwYDVR0jBBgwFoAUV4ab
# VLi+pimK5PbC4hMYiYXN3LcwHQYDVR0OBBYEFH3SkQBtD02UoOt/MNcKFgvst2+/
# MCsGA1UdHwQkMCIwIKAeoByGGmh0dHA6Ly90bC5zeW1jYi5jb20vdGwuY3JsMA4G
# A1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzBzBgNVHSAEbDBqMGgG
# C2CGSAGG+EUBBzACMFkwJgYIKwYBBQUHAgEWGmh0dHBzOi8vd3d3LnRoYXd0ZS5j
# b20vY3BzMC8GCCsGAQUFBwICMCMMIWh0dHBzOi8vd3d3LnRoYXd0ZS5jb20vcmVw
# b3NpdG9yeTAdBgNVHQQEFjAUMA4wDAYKKwYBBAGCNwIBFgMCB4AwVwYIKwYBBQUH
# AQEESzBJMB8GCCsGAQUFBzABhhNodHRwOi8vdGwuc3ltY2QuY29tMCYGCCsGAQUF
# BzAChhpodHRwOi8vdGwuc3ltY2IuY29tL3RsLmNydDANBgkqhkiG9w0BAQsFAAOC
# AQEAjQCSIdnnJcXUpByMElfYuBh0o66Z9D0teIP7tstExgFpUEdV2i1QgftYTod9
# kflbJWL+kreYq0v3Ibi70X2+o46cbKMncZpkuPNgUN91mn5V0B3DONgrE7FYZ2Ts
# JP5PR+wOunVtIaKn3SbOqTocbDx3SLBaGly+bPnh5FqsudhRWqiMKzQHxy3Lh03c
# PYYRkGUjjZekS6s3cYFZremd8TZyZgiU6ifCI8e3wNK1GFv8M7DFYHa0ta27jofc
# DtJW6f0U+8GY99R3HP3B99Lw96Gf3RMjH4ItbpT0vImZLPoA5FyigphBdYnAiZ9N
# Pd0LwA/vo00NG6ZHXUliXjH4UjGCAgEwggH9AgEBMGAwTDELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDHRoYXd0ZSwgSW5jLjEmMCQGA1UEAxMddGhhd3RlIFNIQTI1NiBD
# b2RlIFNpZ25pbmcgQ0ECECHfpxg+32tAy1Vjkr5l/kswCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARYwIwYJKoZIhvcNAQkEMRYE
# FEMXL5bO3LWVsb7TkEbaDpQOSI0UMA0GCSqGSIb3DQEBAQUABIIBAAjHeUb1Bhye
# Vp4CjFbKGuJwIiy1Nv/TgHiKvHTmqwUCzcna+81qZNdv0/mWGdF03M8YfSw5Us1Z
# Wh/6Ydk5gXzpuGw4puU/vV/zkI9thVI19iBw+ySF0/6YLgdmCK9509FDU8VleHiH
# /syZvSalCa3gxitDsSxYQDtwy5LVtBFTk/2lvgVx7lw0wKO+XGzE1far66iqvejf
# 4zssbahIdNRFEFDhasmJCmFB2tNSM8Zu+8wfFfWL+pauAlGKSG9UxIgroJyWBVov
# cyXRQgFvq1kwhpcZB3iunF5ktmeAKOb9x2lWQOuh7dyUfXhIg7nNe1jNBtyN5jYh
# /gjhwhrACGQ=
# SIG # End signature block
