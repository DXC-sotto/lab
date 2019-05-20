Public Function Asset(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next
dim os, ltemp
On Error Goto 0


if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Asset  = "AssetNr"
        Exit Function
End if

'On Error Resume Next
'Err.Clear               ' Clear the error.

if find_Asset_strComputer(strHost) = FALSE then
        Add_Asset_strComputer(strHost)
End If

wscript.echo "System Asset:'"&strHost&"'"
Wscript.Echo "-----------------------------------"

Wscript.Echo "AssetNr:      " & AssetNr(strHost)


Asset  = AssetNr(strHost)

End function

' *****************************************************************************
Public Function AssetNr(strComputer)
Dim InPutFile, lline, f1, linea

AssetNr = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& Asset_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        AssetNr = linea(1)
                End If
        end if
Wend

End function


' *****************************************************************************
Public Function find_Asset_strComputer(strComputer)
Dim InPutFile, lline, f1, linea

find_Asset_strComputer = FALSE

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& Asset_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea= left(lline,len(strComputer))
                        'wscript.echo linea & " - "&strComputer
                if UCASE(linea) = UCASE(strComputer) then
                        find_Asset_strComputer = TRUE
                End If
        end if
Wend

End Function

' *****************************************************************************
Public Function Add_Asset_strComputer(strComputer)
Dim InPutFile, lline, f1, linea

Add_Asset_strComputer = FALSE

Set f1 = objFSO.OpenTextFile(CurrDir & data_path &"\"& Asset_file_name,ForAppending,TristateUseDefault)

 f1.Write (strComputer&";?"& VbCrLf )
 f1.Close

' Close the file to writing.
End Function

