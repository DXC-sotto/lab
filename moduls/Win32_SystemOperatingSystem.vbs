Public Function Win32_SystemOperatingSystem(strComputer,strDomain,strUser,strPassword)
On Error Resume Next
dim os

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_SystemOperatingSystem  = "OperatingSystem.Name;OperatingSystem.Version;OperatingSystem.BuildNumber;OperatingSystem.CSDVersion;LastBootUpTime;InstallDate"
        Exit Function
End if

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_SystemOperatingSystem:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_SystemOperatingSystem  = "0" &";"&"0"&";"& "0" & ";" & "0" & ";" & "0.0.0000"
        Exit Function
End if

Set colItems = objSWbemServices.ExecQuery("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colItems
    Wscript.Echo objOperatingSystem.Caption & "  " & objOperatingSystem.Version & " " & objOperatingSystem.CSDVersion & vbCRLF & "BootTime:" & WMIDateStringToDate(objOperatingSystem.LastBootUpTime) & " InstallDate:" & WMIDateStringToDate(objOperatingSystem.InstallDate) 
    'os = objOperatingSystem.Caption & "  " & objOperatingSystem.Version & " " & objOperatingSystem.CSDVersion
    os = objOperatingSystem.Caption &";"& objOperatingSystem.Version &";"& objOperatingSystem.BuildNumber &";"& objOperatingSystem.CSDVersion &";"& WMIDateStringToDate(objOperatingSystem.LastBootUpTime) &";"& WMIDateStringToDate(objOperatingSystem.InstallDate)
Next

Wscript.Echo "-----------------------------------"
if len(os) = 0 then  
	os = "0" &";"&"0"&";"& "0" & ";" & "0" & ";" & "0.0.0000"& "0.0.0000"
End if	
	
Win32_SystemOperatingSystem = os
End function
