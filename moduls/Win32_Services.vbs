
Public Function Win32_Services(strComputer,strDomain,strUser,strPassword,strHost)

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_Services  = "SystemName;Name;DisplayName;StartMode;State;StartName;PathName;InstallDate;Status"
        Exit Function
End if

  ' create new file
Set fs = objFSO.OpenTextFile(CurrDir & results_path &"\Service"&".csv", ForAppending,True)

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_Services:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_SystemPartitions  = "N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A"
        Exit Function
End if

Set colListOfServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service")

Wscript.Echo "Number of Services:" & colListOfServices.Count
i = 0

For Each objService In colListOfServices

	
    fs.Write ("" & strHost & ";" & objService.Name & ";" & objService.DisplayName & ";" & objService.StartMode & ";" & objService.State & ";" & objService.StartName & ";" & objService.PathName & ";" & objService.InstallDate & ";" & objService.Status & ";"   & vbcrlf)

Next

temp_v = ""

Win32_Services = ""
fs.Close
Wscript.Echo "-----------------------------------"

End Function
