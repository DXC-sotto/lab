Public Function Win32_ComputerSystem(strComputer,strDomain,strUser,strPassword,strHost)

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_ComputerSystem  = "Key;HostName;IP_Adress;Manufacturer;Model;PhysicalMemory(MB)"
        Exit Function
End if

On Error Resume Next
Err.Clear               ' Clear the error.
HostName = ""
ID = CalcKeyString (strComputer,strHost)
Mufacturer = "N/A"
Model = "N/A"
TotalPhysicalMemory = 0
NumberOfProcessors = 0

dim TotalPhysicalMemory,NumberOfProcessors,Model,Manufacturer
Const MBCONVERSION= 1048576
Wscript.Echo "-----------------------------------"
wscript.echo "Win32_ComputerSystem:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_ComputerSystem  = ID &";"& HostName &";"& strComputer &";"& Manufacturer &";"& Model &";"&  TotalPhysicalMemory 
        Exit Function
End if


Set colItems = objSWbemServices.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystem",,48)
For Each objItem in colItems
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.HostName
    Wscript.Echo "NumberOfProcessors: " & objItem.NumberOfProcessors
    Wscript.Echo "TotalPhysicalMemory: " & objItem.TotalPhysicalMemory
    TotalPhysicalMemory = Round( (objItem.TotalPhysicalMemory / MBCONVERSION) + 0.5,0 )
    NumberOfProcessors = objItem.NumberOfProcessors
    HostName = objItem.Name
    Model =objItem.Model
    Manufacturer = objItem.Manufacturer

Next
Wscript.Echo "-----------------------------------"

If len(Name) = 0 then
        HostName = strHost
End If

Win32_ComputerSystem = ID &";"& HostName &";"& strComputer &";"& Manufacturer &";"& Model &";"&  TotalPhysicalMemory 
End Function
