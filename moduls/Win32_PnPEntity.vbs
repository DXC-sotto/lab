Public Function Win32_PnPEntity(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next
Dim MaxClockSpeed, Name

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_PnPEntity  = "SystemName;PnPEntity.Service;PnPEntity.Name;PnPEntity.Manufacturer"
        Exit Function
End if

On Error Resume Next
Err.Clear               ' Clear the error.

  ' create new file
Set fp = objFSO.OpenTextFile(CurrDir & results_path &"\PnPEntity"&".csv", ForAppending,True)

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_PnPEntity:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_PnPEntity  = "N/A;N/A;N/A"
        fx.Write ("N/A;N/A;N/A" & vbcrlf)
        Exit Function
End if

Set colItems = objSWbemServices.ExecQuery("Select * from Win32_PnPEntity")
Wscript.Echo "Number of Win32_PnPEntity:" & colItems.Count

For Each objItem in colItems
    if objItem.CreationClassName = "Win32_PnPEntity" then
                if objItem.Status = "OK" then
                        fp.Write ("" & strHost & ";" &objItem.Service &";"& objItem.Name  &";"& objItem.Manufacturer & vbcrlf)
                        WScript.Echo strHost & ";" & objItem.Service &";"&  objItem.Name &";"& objItem.Manufacturer &""
                End if
    End if
Next
fp.Close
Wscript.Echo "-----------------------------------"
Win32_PnPEntity = ""
End Function
