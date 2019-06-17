# CleanupDebugLogs.ps1

This script can be used to delete all files and folders which are older than a specified number of days.
Please pay attention before usage. Every file or folder in the specified path will be removed.

## Usage 
```ps
CleanupDebugLogs.ps1 [[-NumberOfDaysToDelete] <Int32>] [[-LogFolder] <FileInfo>]
```

## Parameters
### NumberOfDaysToDelete
  Default: 10  
  Define the day age of files which should be removed. Use positive values only.

### LogFolder
  Default: C:\Windows\ServiceProfiles\LocalService\AppData\Local\Temp  
  Define the folder for cleanup.

## Outputs
  Eventlog entries for script start and stop.
  
## Examples
### Example 1
```ps
CleanupDebugLogs.ps1 -NumberOfDaysToDelete 10 -LogFolder "C:\Windows\ServiceProfiles\LocalService\AppData\Local\Temp"
```