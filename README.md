# Troubleshooting

This repository contains PowerShell Scripts to troubleshoot or fix problems in NoSpamProxy installations.

## [CleanupDebugLogs](CleanupDebugLogs/Readme.md)

This script can be used to delete all files and folders which are older than a specified number of days. Please pay attention before usage. Every file or folder in the specified path will be removed.

## [Check-NspCertificates](https://github.com/noSpamProxy/Reports/tree/master/Check-NspCertificates)

Check a certificate and its chain by thumbprint.

## [Delete-EmptyQueueFolder](Delete-EmptyQueueFolder/readme.md)

This tiny PowerShell Script determins the Queuefolder path, deletes every empty Queuefolder and restarts the Gateway Role afterwards.

## [Publish-CertificateToOpenKeys](Publish-CertificateToOpenKeys/readme.md)

This script publishes certificate