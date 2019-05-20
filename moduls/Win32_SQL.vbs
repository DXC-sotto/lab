Public Function Win32_Products(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next


if len(strDomain) = 0 then    ' calling function without a computername will return the header
        Win32_Products  = "SystemName;refProduct.Name;refProduct.Version;refProduct.InstallDate"

        Exit Function
End if

Set fx = objFSO.OpenTextFile(CurrDir & results_path &"\Products"&".csv",ForAppending,True)

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_Products:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\default", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

Set objReg = objSWbemServices.Get("StdRegProv")



if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
            fx.Write ("N/A;N/A" & vbcrlf)
        Win32_Products  = "N/A;N/A"
        Exit Function
End if

Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData 
' HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\00002109150000000000000000F01FEC\InstallProperties
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry2 = "InstallDate"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"
strEntry5 = "EstimatedSize"

objReg.EnumKey HKLM, strKey, arrSubkeys

For Each strSubkey In arrSubkeys
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, strEntry1b, strValue1
  End If

  if strValue1 <> Product_Name then

        If strValue1 <> "" Then
              Product_Name = strValue1
          'WScript.Echo VbCrLf & "Display Name: " & strValue1
        End If

        objReg.GetStringValue HKLM, strKey & strSubkey, strEntry2, strValue2
        If strValue2 <> "" Then
              Product_Install_Date = strValue2
             ' WScript.Echo "Install Date: " & strValue2
        End If
        objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry3, intValue3
        objReg.GetDWORDValue HKLM, strKey & strSubkey, strEntry4, intValue4
        If intValue3 <> "" Then
              Product_Version =  intValue3 & "." & intValue4
             ' WScript.Echo "Version: " & intValue3 & "." & intValue4
        End If

         fx.Write ("" & strHost & ";" & Product_Name &";'"& Product_Version &";'"&  Product_Install_Date & vbcrlf)
         Wscript.Echo "Product Name:'" & Product_Name & "'  Version: '" & Product_Version& "'"
  End If

Next


fx.Close
Wscript.Echo "-----------------------------------"
Win32_Products = ""
End Function
