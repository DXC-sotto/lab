Public Function Win32_LocalSecurity(strComputer,strDomain,strUser,strPassword,strHost)

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        
		Win32_LocalSecurity = "SystemName;" _
			& "Caption" & ";" _
			& "Description"  & ";" _ 
			& "Domain" & ";" _ 
			& "Local Account" & ";" _ 
			& "Name" & ";" _ 
			& "SID" & ";" _ 
			& "SID Type" & ";" _ 
			& "Status" & ";" _ 		
			& "members" 
        Exit Function
		
End if

  ' create new file
Set fls = objFSO.OpenTextFile(CurrDir & results_path &"\LocalAccountGroups"&".csv", ForAppending,True)

On Error Resume Next
Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_LocalSecurity module on:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_LocalSecurity  = "N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A"
        Exit Function
End if

Set colItems  = objWMIService.ExecQuery("Select * from Win32_Group Where LocalAccount = True",,48)
i = 0
For Each objItem in colItems 

	i = i +1
	Wscript.Echo "-----------------------------------"
	Wscript.Echo "Win32_Group instance: " & i
	Wscript.Echo "-----------------------------------"
	 
	Wscript.Echo "Caption: " & objItem.Caption 
	Wscript.Echo "Description: " & objItem.Description 
	Wscript.Echo "Domain: " & objItem.Domain 
	Wscript.Echo "Local Account: " & objItem.LocalAccount 
	Wscript.Echo "Name: " & objItem.Name 
	Wscript.Echo "SID: " & objItem.SID 
	Wscript.Echo "SID Type: " & objItem.SIDType 
	Wscript.Echo "Status: " & objItem.Status 
	Wscript.Echo 
	str_temp = MemmberOF(objItem.Name, objItem.Domain)  
	Wscript.Echo "Memebers: " & str_temp
	
    fls.Write ("" & strHost & ";" _
		& objItem.Caption & ";" _
		& objItem.Description & ";" _
		& objItem.Domain & ";" _
		& objItem.LocalAccount & ";" _
		& objItem.Name & ";" _
		& objItem.SID & ";" _
		& objItem.SIDType & ";" _
		& objItem.Status & ";" _
		& str_temp & ";" _
		& vbcrlf)
Next

Win32_LocalSecurity = ""
fls.Close
Wscript.Echo "-----------------------------------"

End Function

Function MemmberOF(strGroup,strDomain )
Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

	qry = "SELECT * FROM Win32_GroupUser WHERE GroupComponent=""Win32_Group.Domain='"&strDomain&"',Name='"&strGroup&"'"""
	strRet = ""
	on error goto 0
	Set mcolUsers = objWMIService.ExecQuery(qry,"WQL",wbemFlagReturnImmediately + wbemFlagForwardOnly)
	
	For Each mobjItem In mcolUsers
		  NameArray = Split(mobjItem.PartComponent, """", -1, 1)
		  strRet = strRet & NameArray(1) & "\" &  NameArray(3) 
		  'if colUsers.cout = 
			strRet = strRet & ","
		  'Endif
	Next
	MemmberOF = strRet
End function 
