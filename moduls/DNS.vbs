Public Function DNS(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next
dim os, ltemp, DNS_str
On Error Goto 0


if len(strComputer) = 0 then    ' calling function without a computername will return the header
        DNS  = "DNSName"
        Exit Function
End if

'On Error Resume Next
'Err.Clear               ' Clear the error.

DNS_str = DNSName(strComputer)
wscript.echo "System DNS Name:'"&strHost&"'"
Wscript.Echo "-----------------------------------"
Wscript.Echo "DNSName:      " & DNS_str


DNS  = DNS_str

End function

' *****************************************************************************
Public Function DNSName(strComputer)
Dim max_loop,objExec,i,temp,strPingResults

DNSName = "N/A"
On Error Goto 0

Set objExec = objShell.Exec("nslookup " & strComputer)
max_loop = 100

Do While (True)
    max_loop = max_loop - 1

    If Not objExec.StdOut.AtEndOfStream Then
        strPingResults = LCase(objExec.StdOut.Readline)
                'wscript.echo strPingResults
        If InStr(strPingResults, "name:") Then
            For i = Len(strPingResults) To 1 Step -1
                temp = Mid(strPingResults, i, 1)
                If temp = " " Then

                    DNSName = Right(strPingResults, Len(strPingResults) - i)
                    Exit Do
                End If
            Next
        Else
            DNSName = "N/A"
        End If
    End If

    If max_loop = 0 Then
        Exit Do
    End If
Loop

End function




