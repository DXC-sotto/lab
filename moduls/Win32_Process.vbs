
Public Function Win32_Process(strComputer,strDomain,strUser,strPassword,strHost)

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_Process  = "SystemName;objService.Caption;objService.SessionId;objService.Description;objService.KernelModeTime;objService.WorkingSetSize;objService.CreationDate;objService.ExecutablePath;objService.HandleCount"
        Exit Function
End if

  ' create new file
Set fpp = objFSO.OpenTextFile(CurrDir & results_path &"\Process"&".csv", ForAppending,True)

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_Process:'"&strComputer &"'"

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

Set colListOfServices = objWMIService.ExecQuery("SELECT * FROM Win32_Process")

Wscript.Echo "Number of Process:" & colListOfServices.Count
i = 0

For Each objService In colListOfServices

	
    fpp.Write ("" & strHost & ";" & objService.Caption & ";" & objService.SessionId & ";" & objService.Description & ";" & objService.KernelModeTime & ";" & objService.WorkingSetSize & ";" & objService.CreationDate & ";" & objService.ExecutablePath & ";" & objService.HandleCount & ";"   & vbcrlf)

Next

temp_v = ""

Win32_Process = ""
fpp.Close
Wscript.Echo "-----------------------------------"

End Function
