Public Function RemoteMgt(strComputer,strDomain,strUser,strPassword,strHost)
'On Error Resume Next
dim os, ltemp, DNS_str, lineArray,i
'On Error Goto 0

if len(strComputer) = 0 then    ' calling function without a computername will return the header
        RemoteMgt  = "RemoteMgt"
        Exit Function
End if

'On Error Resume Next
'Err.Clear               ' Clear the error.

RM_str = DNSName(strHost)
if StrComp (RM_str,"N/A") >0 then

        lineArray = Split(RM_str, ".", -1, 1)

        if UBound(lineArray) > 0 then
                 RM_str = lineArray(0) & "r"

                 for i = 1 to UBound(lineArray)
                        if Len( lineArray(i) ) > 0 then
                                RM_str = RM_str & "." & lineArray(i)
                        end if
                 Next
                 'Wscript.Echo "RRM_str:      " & RM_str
                 RM_str = DNSName(RM_str)
                'Wscript.Echo "RRM_str:      " & RM_str
        End if

        if StrComp (RM_str,"N/A") = 0 then
        if UBound(lineArray) > 0 then
                RM_str = lineArray(0) & "i4"

                for i = 1 to UBound(lineArray)
                        if Len( lineArray(i) ) > 0 then
                              RM_str = RM_str & "." & lineArray(i)
                        end if
                Next
                'Wscript.Echo "RM_str:      " & RM_str
                RM_str = DNSName(RM_str)

        End if
        End if

End If

wscript.echo "System RemoteMgt:'"&strHost&"'"
Wscript.Echo "-----------------------------------"
Wscript.Echo "RemoteMgt Name:      " & RM_str


RemoteMgt  = RM_str

End function

