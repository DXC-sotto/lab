

Public Function Win32_BIOS(strComputer,strDomain,strUser,strPassword)
dim ReleaseDate, BIOSVersion, SMBIOSBIOSVersion, SerialNumber

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_BIOS  = "Bios_Date;BIOS_Version;SerialNumber"
        Exit Function
End if

On Error Resume Next
Err.Clear               ' Clear the error.
ReleaseDate = "01/01/1900"
BIOSVersion ="0.0"
SerialNumber =0

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_BIOS:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_BIOS  = ReleaseDate & ";"& BIOSVersion &";"& SerialNumber
        Exit Function
End if

Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_BIOS",,48)
For Each objItem in colItems

    Wscript.Echo "Name: " & objItem.Name
    if len(objItem.SerialNumber) > 0 then
        Wscript.Echo "SerialNumber: " & objItem.SerialNumber
        SerialNumber = objItem.SerialNumber
    else
        Wscript.Echo "SerialNumber: " &alternative_Serial_number(strComputer,strDomain,strUser,strPassword)
        SerialNumber = alternative_Serial_number(strComputer,strDomain,strUser,strPassword)
    End If
    Wscript.Echo "ReleaseDate: " & left(objItem.ReleaseDate,8)
    ReleaseDate =  left(objItem.ReleaseDate,8)
    on error resume next
    If isNull(objItem.BIOSVersion) Then
            Wscript.Echo "BIOSVersion: "
            BIOSVersion = ""
    Else
            Wscript.Echo "BIOSVersion: " & Join(objItem.BIOSVersion, ",")
            BIOSVersion = Join(objItem.BIOSVersion, ",")
    End If

    Wscript.Echo "SMBIOSBIOSVersion: " & objItem.SMBIOSBIOSVersion
    Wscript.Echo "SMBIOSMajorVersion: " & objItem.SMBIOSMajorVersion
    Wscript.Echo "SMBIOSMinorVersion: " & objItem.SMBIOSMinorVersion
    SMBIOSBIOSVersion = objItem.SMBIOSBIOSVersion
Next

Wscript.Echo "-----------------------------------"

if len(BIOSVersion ) = 0 then
        BIOSVersion = SMBIOSBIOSVersion
end if

ReleaseDate = right(ReleaseDate,2) &"."& Mid(ReleaseDate,5,2) &"."& left(ReleaseDate,4)

if len(BIOSVersion) = 0 then
		BIOSVersion = 0
End if

Win32_BIOS = ReleaseDate & ";"& BIOSVersion &";"& SerialNumber
End Function


Private Function alternative_Serial_number(strComputer,strDomain,strUser,strPassword)

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)


Set colSMBIOS = objSWbemServices.ExecQuery ("Select * from Win32_SystemEnclosure")
For Each objSMBIOS in colSMBIOS
    'Wscript.Echo "Part Number: " & objSMBIOS.PartNumber
    'Wscript.Echo "Serial Number: " & objSMBIOS.SerialNumber
    'Wscript.Echo "Asset Tag: " & objSMBIOS.SMBIOSAssetTag
    alternative_Serial_number = objSMBIOS.SerialNumber
Next

End function
