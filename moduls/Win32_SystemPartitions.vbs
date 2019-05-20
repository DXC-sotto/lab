
Public Function Win32_SystemPartitions(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next
Dim vol_size, vol_free,vol_DeviceID,fp
Const MBCONVERSION= 1048576
Const GBCONVERSION= 1073741824 '1074738110


if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_SystemPartitions  = "SystemName;DiskName;PartitionID;LogicalDisk.DeviceID;LogicalDisk_Size(GB);LogicalDisk_Freespace(GB)"
        Exit Function
End if


  ' create new file
Set fp = objFSO.OpenTextFile(CurrDir & results_path &"\LogicalDisk"&".csv", ForAppending,True)

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_SystemPartitions:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_SystemPartitions  = "N/A;N/A;N/A"
        Exit Function
End if

Set colDiskDrives = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")

Wscript.Echo "Number of Disk Drives:" & colDiskDrives.Count
i = 0

For Each objDrive In colDiskDrives
    Wscript.Echo "Physical Disk: " & objDrive.Caption & " -- " & objDrive.DeviceID
    Wscript.Echo " Size= " & objDrive.Size

        strDeviceID = Replace(objDrive.DeviceID, "\", "\\")
        Set colPartitions = objWMIService.ExecQuery _
            ("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & _
                strDeviceID & """} WHERE AssocClass = " & _
                    "Win32_DiskDriveToDiskPartition")

        For Each objPartition In colPartitions
            Wscript.Echo "Disk Partition: " & objPartition.DeviceID
            Set colLogicalDisks = objWMIService.ExecQuery _
                ("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & _
                    objPartition.DeviceID & """} WHERE AssocClass = " & _
                        "Win32_LogicalDiskToPartition")

            For Each objLogicalDisk In colLogicalDisks
                Wscript.Echo "Logical Disk      : " & objLogicalDisk.DeviceID
                Wscript.Echo "Logical DriveType : " & objLogicalDisk.DriveType
                Wscript.Echo "Logical Size      : " & objLogicalDisk.Size
                Wscript.Echo "Logical freespace : " & objLogicalDisk.freespace

                vol_DeviceID = objLogicalDisk.DeviceID
                vol_size = Round( (objLogicalDisk.Size / GBCONVERSION) ,3 )
                vol_free = Round( (objLogicalDisk.freespace / GBCONVERSION),3 )

                fp.Write ("" & strHost & ";" & objDrive.Caption & ";" & objPartition.DeviceID & ";" & vol_DeviceID &";"& vol_size  &";"& vol_free & vbcrlf)

            Next
            Wscript.Echo
        Next
    Wscript.Echo
Next

temp_v = ""

Win32_SystemPartitions = ""
fp.Close
Wscript.Echo "-----------------------------------"
End Function
