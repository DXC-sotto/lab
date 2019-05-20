Public Function Win32_TimeZone(strComputer,strDomain,strUser,strPassword,strHost)

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_TimeZone  = "TimeZone.Description"
        Exit Function
End if

On Error Resume Next
Err.Clear               ' Clear the error.
Description = "N/A"

dim TotalPhysicalMemory,NumberOfProcessors,Model,Manufacturer
Const MBCONVERSION= 1048576
Wscript.Echo "-----------------------------------"
wscript.echo "Win32_TimeZone:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_TimeZone  = "N/A"
        Exit Function
End if


Set colItems = objSWbemServices.ExecQuery( _
    "SELECT * FROM Win32_TimeZone",,48)
For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description

	
    Description = objItem.Description
	

Next
Wscript.Echo "-----------------------------------"

If len(Name) = 0 then
        Name = strHost
End If

if len(Description) = 0 then
	Description = "N/A"
End if
	
Win32_TimeZone = Description 
End Function
