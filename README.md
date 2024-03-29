# Troubleshooting

This repository contains PowerShell Scripts to troubleshoot or fix problems in NoSpamProxy installations.

## [CleanupDebugLogs](CleanupDebugLogs/Readme.md)

This script can be used to delete all files and folders which are older than a specified number of days. Please pay attention before usage. Every file or folder in the specified path will be removed.

## [Check-NspCertificates](https://github.com/noSpamProxy/Troubleshooting/tree/master/Check-NspCertificates)

Check a certificate and its chain by thumbprint.

## [Check-NspPrivateCertificates](https://github.com/noSpamProxy/Troubleshooting/tree/master/Check-NspPrivateCertificates)

Validate all private certificates found in the NoSpamProxy certificate store.

## [Delete-EmptyQueueFolder](Delete-EmptyQueueFolder/readme.md)

This tiny PowerShell Script determins the Queuefolder path, deletes every empty Queuefolder and restarts the Gateway Role afterwards.

## [Publish-CertificateToOpenKeys](Publish-CertificateToOpenKeys/readme.md)

This script publishes certificate

## [Export-MyNspPublicCertificates](Export-MyNspPublicCertificates/Readme.md)

This script exports your own valid public certificates.