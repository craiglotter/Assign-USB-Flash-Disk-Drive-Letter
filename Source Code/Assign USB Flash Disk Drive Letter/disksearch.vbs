ComputerName = "."
Set wmiServices  = GetObject ( _
    "winmgmts:{impersonationLevel=Impersonate}!//" _
    & ComputerName)
' Get physical disk drive
Set wmiDiskDrives =  wmiServices.ExecQuery ( _
    "SELECT Caption, DeviceID FROM Win32_DiskDrive")

For Each wmiDiskDrive In wmiDiskDrives
    WScript.Echo "Disk drive Caption: " _
        & wmiDiskDrive.Caption

'Backslash in disk drive deviceid
' must be escaped by "\"    
    strEscapedDeviceID = Replace( _
        wmiDiskDrive.DeviceID, "\", "\\",_
        1, -1, vbTextCompare)
'Use the disk drive device id to
' find associated partition 
    Set wmiDiskPartitions = wmiServices.ExecQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & strEscapedDeviceID & """} WHERE AssocClass = Win32_DiskDriveToDiskPartition")

    For Each wmiDiskPartition In wmiDiskPartitions
'Use partition device id to find logical disk
        Set wmiLogicalDisks = wmiServices.ExecQuery("ASSOCIATORS OF " & "{Win32_DiskPartition.DeviceID=""" &   wmiDiskPartition.DeviceID & """} WHERE AssocClass = Win32_LogicalDiskToPartition")

        For Each wmiLogicalDisk In wmiLogicalDisks
            WScript.Echo "Drive Type: " _
		& wmiLogicalDisk.DriveType
            WScript.Echo "Associated Drive letter: " _
                & wmiLogicalDisk.DeviceID
        Next
    Next
Next
