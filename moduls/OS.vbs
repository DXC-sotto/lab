Public Function Gues_OS(strComputer,strDomain,strUser,strPassword,strHost)
'On Error Resume Next
dim result
'On Error Goto 0

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Gues_OS  = "Computing Platform"
        Exit Function
End if

'Wscript.Echo "U:'"&strUser&"' P:'" & strPassword &"' H:"&strHost&"'"
'On Error Resume Next
'Err.Clear               ' Clear the error.

result = MyGues_OS(strComputer)

Wscript.Echo "computing platform:      " & result
Wscript.Echo "-----------------------------------"


Gues_OS  = result

End function

' try to find open network ports
Function MyGues_OS(Server)
MyGues_OS = "disconnectet"
	' Try SSH
	if MyCheckPort(Server,22) then
		MyGues_OS = "Linux/UNIX"
		'exit function
	end if
	if MyCheckPort(Server,21) then
		MyGues_OS = "Linux/UNIX"

	end if 
	if MyCheckPort(Server,80) then
		MyGues_OS = "Linux/UNIX"
	end if 
	if MyCheckPort(Server,445) then
		MyGues_OS = "Wintel"
		exit function
	elseif (MyCheckPort(Server,138)) then
		MyGues_OS = "Wintel"
		exit function
	elseif (MyCheckPort(Server,3389)) then ' RDP 
		MyGues_OS = "Wintel"	
		exit function
	End If
End Function

Function MyCheckPort(Server,Port)
	on error resume next
	set SockObject=CreateObject("MSWinsock.Winsock.1")
	SockObject.Protocol=0 ' TCP
	SockObject.Close
	SockObject.Connect Server,port
	Wscript.Sleep 50
	MyCheckPort = False
	
	if(SockObject.State=6) then ' if sockect is attempting to connect i.e 
	
		'WScript.echo port & " Invalid Port"
	
	elseif (SockObject.State=7) then ' if socket connected
	
		'WScript.echo port & " Port Open"
		MyCheckPort = True
	
	elseif(SockObject.State=9) then ' If Error
	
		'WScript.echo port & " error"
	
	elseif(SockObject.State=0) then 'Closed
	
		'WScript.echo port & " connection refused"
	
	end if
SockObject.Close

set SockObject=nothing
End Function 
