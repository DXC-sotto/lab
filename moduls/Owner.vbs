Public Function Owner(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next
dim os, ltemp
On Error Goto 0


if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Owner  = "OwnerName;OwnerPhone;OwnerEMail"
        Exit Function
End if

'On Error Resume Next
'Err.Clear               ' Clear the error.

if find_Owner_strComputer(strHost) then
        'Wscript.echo strHost & " gefunden"

Else
        'Wscript.echo strHost & " NICHT gefunden"
        Add_Owner_strComputer(strHost)

End If




wscript.echo "System Owner:'"&strHost&"'"
Wscript.Echo "-----------------------------------"

Wscript.Echo "Name:      " & OwnerName(strHost)
Wscript.Echo "OwnerPhone:" & OwnerPhone(strHost)
Wscript.Echo "OwnerEMail:" & OwnerEMail(strHost)


Owner  = OwnerName(strHost)&";'"&OwnerPhone(strHost)&";"&OwnerEMail(strHost)

End function

' *****************************************************************************
Public Function OwnerName(strComputer)
Dim InPutFile, lline, f1, linea

OwnerName = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& Owner_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        OwnerName = linea(1)
                End If
        end if
Wend

End function

' *****************************************************************************
Public Function OwnerPhone(strComputer)
Dim InPutFile, lline, f1, linea

OwnerPhone = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& Owner_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        OwnerPhone = "'" & linea(2)
                End If
        end if
Wend

End Function

' *****************************************************************************
Public Function OwnerEMail(strComputer)
Dim InPutFile, lline, f1, linea

OwnerEMail = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& Owner_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        OwnerEMail = linea(3)
                End If
        end if
Wend

End Function


' *****************************************************************************
Public Function find_Owner_strComputer(strComputer)
Dim InPutFile, lline, f1, linea

find_Owner_strComputer = FALSE

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& Owner_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea= left(lline,len(strComputer))
                        'wscript.echo linea & " - "&strComputer
                if UCASE(linea) = UCASE(strComputer) then
                        find_Owner_strComputer = TRUE
                End If
        end if
Wend

End Function

' *****************************************************************************
Public Function Add_Owner_strComputer(strComputer)
Dim InPutFile, lline, f1, linea

Add_Owner_strComputer = FALSE

Set f1 = objFSO.OpenTextFile(CurrDir & data_path &"\"& Owner_file_name,ForAppending,TristateUseDefault)

 f1.Write (strComputer&";?;?;?"& VbCrLf )
 f1.Close

' Close the file to writing.
End Function

