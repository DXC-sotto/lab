Public Function Win32_IIS(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next


if len(strDomain) = 0 then    ' calling function without a computername will return the header
        Win32_IIS  = "SystemName;ID;Name;Root;Install_Date;Package_Name;Type"
        Exit Function
End if

Set fx = objFSO.OpenTextFile(CurrDir & results_path &"\IIS"&".csv",ForAppending,True)

On Error Resume Next

Err.Clear               ' Clear the error.

Wscript.Echo "-----------------------------------"
wscript.echo "Win32_IIS:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Exit Function
End if

Set colListOfServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service WHERE Name=""IISADMIN""")

if colListOfServices.Count >= 1 then

	REM root\microsoftiisv2 requires encryption
	objSWbemLocator.Security_.AuthenticationLevel = 6
	Set objWMIService = objSWbemLocator.ConnectServer(strComputer, _
    	"root\microsoftiisv2", _
     	strUser, _
     	strPassword, _
     	"MS_409", _
     	"ntlmdomain:" + strDomain)

	if Err.Number <> 0 then
        	wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        	Exit Function
	End if

	Wscript.Echo "IIS is installed"

	Set ColWebVirtualDirs = objWMIService.ExecQuery("Select * from IIsWebVirtualDir")
	
	For Each ColWebVirtualDir In ColWebVirtualDirs
		fx.Write ("" & strHost & ";" & ColWebVirtualDir.AppPackageID  & ";" & ColWebVirtualDir.Name & ";" & ColWebVirtualDir.AppRoot & ";" & ColWebVirtualDir.InstallDate & ";" & ColWebVirtualDir.AppPackageName & ";Virtual;"   & vbcrlf)
		Wscript.Echo("" & strHost & ";" & ColWebVirtualDir.AppPackageID  & ";" & ColWebVirtualDir.Name & ";" & ColWebVirtualDir.AppRoot & ";" & ColWebVirtualDir.InstallDate & ";" & ColWebVirtualDir.AppPackageName & ";Virtual;"   & vbcrlf)
	next

	Set ColWebDirs = objWMIService.ExecQuery("Select * from IIsWebDirectory")
	
	For Each ColWebDir In ColWebDirs
		fx.Write ("" & strHost & ";" & ColWebDir.AppPackageID  & ";" & ColWebDir.Name & ";" & ColWebDir.AppRoot & ";" & ColWebDir.InstallDate & ";" & ColWebDir.AppPackageName & ";not-Virtual;"   & vbcrlf)
		Wscript.Echo("" & strHost & ";" & ColWebDir.AppPackageID  & ";" & ColWebDir.Name & ";" & ColWebDir.AppRoot & ";" & ColWebDir.InstallDate & ";" & ColWebDir.AppPackageName & ";not-Virtual;"   & vbcrlf)
	next
	
end if



End Function
