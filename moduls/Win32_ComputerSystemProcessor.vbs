Public Function Win32_ComputerSystemProcessor(strComputer,strDomain,strUser,strPassword)
On Error Resume Next
Dim MaxClockSpeed, Name

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Win32_ComputerSystemProcessor  = "CPU_Clock;AddressWidth;CPU_Type;NumberOfProcessors;NumberOfCores;NumberOfLogicalProcessors"
        Exit Function
End if

On Error Resume Next
Err.Clear               ' Clear the error.

MaxClockSpeed = 0
Name = "N/A"
NumberOfCores = 0
NumberOfLogicalProcessors = 0
AddressWidth = 0
NumberOfProcessors = 0


Wscript.Echo "-----------------------------------"
wscript.echo "Win32_ComputerSystemProcessor:'"&strComputer &"'"

Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
    "root\cimv2", _
     strUser, _
     strPassword, _
     "MS_409", _
     "ntlmdomain:" + strDomain)

if Err.Number <> 0 then
        wscript.echo ("***Error  #0x" & Hex(Err.Number) & " " & Err.Description)
        Win32_ComputerSystemProcessor  = MaxClockSpeed &";"&AddressWidth &";"& Name &";"& NumberOfProcessors &";"& NumberOfCores&";"&NumberOfLogicalProcessors
        Exit Function
End if

Set colItems = objSWbemServices.ExecQuery("Select * from Win32_Processor")
Wscript.Echo "Number of Processors:" & colItems.Count
NumberOfProcessors = colItems.Count

For Each objItem in colItems
    'Wscript.Echo "Processor Id: " & objItem.ProcessorId
	MaxClockSpeed = objItem.MaxClockSpeed
	Name = objItem.Name
	NumberOfCores = objItem.NumberOfCores
	NumberOfLogicalProcessors = objItem.NumberOfLogicalProcessors
	if NumberOfLogicalProcessors = 0 then
		' Windows 200, W2003
		 NumberOfLogicalProcessors = NumberOfProcessors
	end if
	AddressWidth = objItem.AddressWidth
	
	Wscript.Echo "Number Of Cores			 :"	& objItem.NumberOfCores
	Wscript.Echo "Number Of LogicalProcessors:" & objItem.NumberOfLogicalProcessors
	Wscript.Echo "AddressWidth				 :"	& objItem.AddressWidth
    Wscript.Echo "Maximum Clock Speed		 :" & objItem.MaxClockSpeed
    Wscript.Echo "Processor Manufacturer	 :" & objItem.Manufacturer
    Wscript.Echo "Processor Name			 :" & objItem.Name
    Wscript.Echo "Processor Caption			 :"	& objItem.Caption
	
    Exit For
Next
Wscript.Echo "-----------------------------------"
Win32_ComputerSystemProcessor = MaxClockSpeed &";"&AddressWidth &";"& Name &";"& NumberOfProcessors &";"& NumberOfCores&";"&NumberOfLogicalProcessors
End Function
