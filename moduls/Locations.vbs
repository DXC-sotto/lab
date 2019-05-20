Public Function Location(strComputer,strDomain,strUser,strPassword,strHost)
On Error Resume Next
dim os, ltemp
On Error Goto 0


if len(strComputer) = 0 then    ' calling function without a computername will return the header
        Location  = "Country;City;Building;Room;Shelf;Shelf_Slot"
        Exit Function
End if

'On Error Resume Next
'Err.Clear               ' Clear the error.

if find_Location_strComputer(strHost) then
        'Wscript.echo strHost & " gefunden"

Else
        'Wscript.echo strHost & " NICHT gefunden"
        Add_Location_strComputer(strHost)

End If




wscript.echo "Location:'"&strHost&"'"
Wscript.Echo "-----------------------------------"

Wscript.Echo "Country:   " & Country(strHost)
Wscript.Echo "City:      " & City(strHost)
Wscript.Echo "Building:  " & Building(strHost)
Wscript.Echo "Room:      " & Room(strHost)
Wscript.Echo "Shelf:      " & Shelf(strHost)
Wscript.Echo "Shelf_Slot: " & Shelf_Slot(strHost)

Location  = Country(strHost)&";"&City(strHost)&";"&Building(strHost)&";'"&Room(strHost)&";'"&Shelf(strHost)&";'"&Shelf_Slot(strHost)

End function

' *****************************************************************************
Public Function Country(strComputer)
Dim InPutFile, lline, f1, linea

Country = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        Country = linea(1)
                End If
        end if
Wend

End function

' *****************************************************************************
Public Function City(strComputer)
Dim InPutFile, lline, f1, linea

City = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        City = linea(2)
                End If
        end if
Wend

End Function

' *****************************************************************************
Public Function Building(strComputer)
Dim InPutFile, lline, f1, linea

Building = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        Building = linea(3)
                End If
        end if
Wend

End Function

' *****************************************************************************
Public Function Room(strComputer)
Dim InPutFile, lline, f1, linea

Room = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        Room = linea(4)
                End If
        end if
Wend


End function

' *****************************************************************************
Public Function Shelf(strComputer)
Dim InPutFile, lline, f1, linea

Shelf = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        Shelf = linea(5)
                End If
        end if
Wend

End function

' *****************************************************************************
Public Function Shelf_Slot(strComputer)
Dim InPutFile, lline, f1, linea

Shelf_Slot = ""

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea = Split(lline, tab, -1, 1)

                if UCASE(linea(0)) = UCASE(strComputer) then
                        Shelf_Slot = linea(6)
                End If
        end if
Wend


End function

' *****************************************************************************
Public Function find_Location_strComputer(strComputer)
Dim InPutFile, lline, f1, linea

find_Location_strComputer = FALSE

Set f1 = objFSO.GetFile(CurrDir & data_path &"\"& location_file_name)
Set InPutFile = f1.OpenAsTextStream(ForReading, TristateUseDefault)
lline = InPutFile.ReadLine ' skip first header line

While InPutFile.AtEndOfStream = False
        lline = InPutFile.ReadLine
        if len(lline) > 0 then    ' skip empty lines
                linea= left(lline,len(strComputer))
                        'wscript.echo linea & " - "&strComputer
                if UCASE(linea) = UCASE(strComputer) then
                        find_Location_strComputer = TRUE
                End If
        end if
Wend

End Function

' *****************************************************************************
Public Function Add_Location_strComputer(strComputer)
Dim InPutFile, lline, f1, linea

add_Location_strComputer = FALSE

Set f1 = objFSO.OpenTextFile(CurrDir & data_path &"\"& location_file_name,ForAppending,TristateUseDefault)

 f1.Write (strComputer&";?;?;?;?;?;?"& VbCrLf )
 f1.Close

' Close the file to writing.
End Function

