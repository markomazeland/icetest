Attribute VB_Name = "modIceImport"
' Copyright (C) Marko Mazeland 2003
'
' This program is free software; you can redistribute it and/or modify it under the terms of the
' GNU General Public License as published by the Free Software Foundation; either version 2 of the License,
' or (at your option) any later version.
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the
' implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
' for more details (http://www.opensource.org/licenses/gpl-license.php).
'
' You should have received a copy of the GNU General Public License along with this program; if not, write to the
' Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Option Compare Text

Sub ImportExcel()
    Dim cExcelFile As String
    
    On Local Error Resume Next
    
    ReadIniFile gcIniFile, "Import", "Excel", cExcelFile
    With frmMain.CommonDialog1
        .DefaultExt = "Xls"
        .DialogTitle = Translate("Select an Excel-sheet", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Excel (*.Xls)|*.Xls"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    On Local Error GoTo 0
    
    SetMouseHourGlass
    
    If cExcelFile <> "" And cExcelFile <> Chr$(27) Then
        WriteIniFile gcIniFile, "Import", "Excel", cExcelFile
        frmMain.Enabled = False
        MsgBox ImportXls(cExcelFile) & " " & Translate("participants processed.", mcLanguage)
        frmMain.Enabled = True
    Else
        MsgBox Translate("No proper Excel-sheet selected.", mcLanguage)
    End If
    
    StatusMessage
    
    SetMouseNormal
    
End Sub
Public Function ImportRecord(cImport As String, cColumns() As String, Optional cDelimChar As String = vbTab) As String
    Dim cSta As String
    Dim cRider As String
    Dim cFirst As String
    Dim cLast As String
    Dim cHorse As String
    Dim cClub As String
    Dim cTeam As String
    Dim cClass As String
    Dim cPersonId As String
    Dim cHorseId As String
    Dim cField() As String
    Dim cTest() As String
    Dim cTemp As String
    Dim cPosition As String
    Dim cRein As String
    Dim cColor As String
    Dim iTemp As Integer
    Dim iLow As Integer
    
    Dim cLeft As String
    Dim cRight As String
    
    Dim rstParticipant As DAO.Recordset
    Dim rstHorse As DAO.Recordset
    Dim rstRider As DAO.Recordset
    Dim rstEntry As DAO.Recordset
    
    iLow = 3
    
    cImport = Replace(cImport, Chr$(34), "'")
    cField = Split(cImport, cDelimChar)
    If UBound(cField) >= iLow - 1 Then
        cLeft = Left$(Translate("Left", mcLanguage), 1)
        cRight = Left$(Translate("Right", mcLanguage), 1)
        
        cSta = Format$(Val(cField(0)), "000")
        cRider = cField(1)
        cHorse = cField(2)
        cClub = ""
        cTeam = ""
        cClass = ""
        Select Case cColumns(3)
        Case "Club", """Club"""
            cClub = cField(3)
            iLow = iLow + 1
        Case "Class", """Class"""
            cClass = cField(3)
            iLow = iLow + 1
        Case "Team", """Team"""
            cTeam = cField(3)
            iLow = iLow + 1
        End Select
        
        Select Case cColumns(4)
        Case "Club"
            cClub = cField(4)
            iLow = iLow + 1
        Case "Class", """Class"""
            cClass = cField(4)
            iLow = iLow + 1
        Case "Team"
            cTeam = cField(4)
            iLow = iLow + 1
        End Select
        Select Case cColumns(5)
        Case "Club"
            cClub = cField(5)
            iLow = iLow + 1
        Case "Class", """Class"""
            cClass = cField(5)
            iLow = iLow + 1
        Case "Team"
            cTeam = cField(5)
            iLow = iLow + 1
        End Select
                 
        If Val(cSta) > 0 Then
            'remove old participant
            mdbMain.Execute ("DELETE * FROM Participants WHERE Sta LIKE '" & cSta & "'")
            
            Set rstRider = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE Name_First & ' ' & Name_Last LIKE " & Chr$(34) & cRider & Chr$(34))
            With rstRider
                If .RecordCount = 0 Then
                   .AddNew
                   cPersonId = CreatePersonId
                   .Fields("PersonId") = cPersonId
                   .Fields("Name_First").Value = Left$(MakeFirst(cRider), .Fields("Name_First").Size)
                   .Fields("Name_Last").Value = Left$(MakeLast(cRider), .Fields("Name_Last").Size)
                   .Update
                Else
                    cPersonId = .Fields("PersonId")
                End If
            End With
            rstRider.Close
            
            Set rstHorse = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE Name_Horse LIKE " & Chr$(34) & cHorse & Chr$(34))
            With rstHorse
                If .RecordCount = 0 Then
                    .AddNew
                    cHorseId = CreateHorseId
                    .Fields("HorseId") = cHorseId
                    .Fields("Name_Horse") = cHorse
                    .Update
                Else
                    cHorseId = .Fields("HorseId")
                End If
            End With
            rstHorse.Close
            
            Set rstParticipant = mdbMain.OpenRecordset("SELECT * FROM Participants")
            With rstParticipant
                .AddNew
                .Fields("Sta") = Left$(cSta, .Fields("STA").Size)
                .Fields("HorseId") = cHorseId
                .Fields("PersonId") = cPersonId
                .Fields("Club") = cClub
                .Fields("Class") = cClass
                .Fields("Team") = cTeam
                .Update
            End With
            rstParticipant.Close
            
            If UBound(cField) >= iLow And UBound(cColumns) >= iLow Then
                For iTemp = iLow To UBound(cField)
                    'check for empty field, avoid spaces
                    If Trim$(cField(iTemp)) <> "" Then
                        cTemp = cColumns(iTemp)
                        mdbMain.Execute ("DELETE * FROM Entries WHERE Code='" & cTemp & "' AND Sta='" & cSta & "' AND Status=0")
                        Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries")
                        With rstEntry
                            .AddNew
                            .Fields("Sta") = Left$(cSta, .Fields("STA").Size)
                            .Fields("Code") = cTemp
                            .Fields("Group") = 0
                            .Fields("Status") = 0
                            .Fields("Timestamp") = Now
                            cTemp = Trim$(Replace(cField(iTemp), "  ", " "))
                            If Right$(cTemp, 1) = cRight Then
                                cTemp = Trim$(Left$(cTemp, Len(cTemp) - 1))
                                .Fields("RR") = True
                            ElseIf Right$(cTemp, 1) = cLeft Then
                                cTemp = Trim$(Left$(cTemp, Len(cTemp) - 1))
                                .Fields("RR") = False
                            ElseIf InStr(" " & cTemp & " ", " " & cRight & " ") > 0 Then
                                cTemp = Replace(" " & cTemp, " " & cRight & " ", " ")
                                .Fields("RR") = True
                            ElseIf InStr(" " & cTemp & " ", " " & cLeft & " ") > 0 Then
                                cTemp = Replace(" " & cTemp & " ", cLeft & " ", " ")
                                .Fields("RR") = False
                            Else
                                .Fields("RR") = False
                            End If
                            cTemp = Replace(Trim$(cTemp), "  ", " ")
                            Parse cPosition, cTemp, " "
                            .Fields("Position") = Val(cPosition)
                            Parse cColor, cTemp, " "
                            .Fields("Color") = cColor
                            .Update
                        End With
                        rstEntry.Close
                    End If
                Next iTemp
            End If
        
        End If
    End If
    Set rstRider = Nothing
    Set rstHorse = Nothing
    Set rstParticipant = Nothing
    Set rstEntry = Nothing
    
End Function
Public Function ImportXls(ImportFile As String) As Integer
    Dim xlObj As Object
    
    Dim cSta As String
    Dim cTemp As String
    Dim cPersonId As String
    Dim cHorseId As String
    Dim cName_First As String
    Dim cName_Last As String
    Dim cName_Horse As String
    Dim cHorses_FEIFId As String
    Dim cPersons_FEIFId As String
    
    Dim vValue As Variant
    Dim cValue As String
    Dim cPosition As String
    Dim cColor As String
    
    Dim cFieldList(4) As String
    Dim cLabel() As String
    
    Dim iColumn  As Integer
    Dim iColumnCount As Integer
    Dim iRow As Integer
    Dim iStartRow As Integer
    
    Dim cLeft As String
    Dim cRight As String
    
    Dim rstParticipant As DAO.Recordset
    Dim rstHorse As DAO.Recordset
    Dim rstRider As DAO.Recordset
    Dim rstEntry As DAO.Recordset
    Dim fld As DAO.Field
    
    On Local Error Resume Next
        
    SetMouseHourGlass
    
    cLeft = Left$(Translate("Left", mcLanguage), 1)
    cRight = Left$(Translate("Right", mcLanguage), 1)
    
    Set xlObj = GetObject(ImportFile)
    'need to specify the sheet on excel versions >= 8
    'the format of my excel import file specifies only
    'one worksheet
    If Val(xlObj.Application.Version) >= 8 Then
        Set xlObj = xlObj.ActiveSheet
    End If
    DoEvents
    
    cFieldList(4) = "|"
    Set rstEntry = mdbMain.OpenRecordset("Select CODE FROM Tests")
    If rstEntry.RecordCount > 0 Then
        Do While Not rstEntry.EOF
            cFieldList(4) = cFieldList(4) & rstEntry.Fields(0) & "|"
            rstEntry.MoveNext
        Loop
    End If
        
    ReDim cLabel(xlObj.Columns.Count)
        
    iStartRow = 0
    Do While iStartRow < 255
        iStartRow = iStartRow + 1
        iColumn = 0
        If xlObj.Cells(1, 1).Value & "" <> "" Then
            Do While iColumn < 255
                iColumn = iColumn + 1
                If xlObj.Cells(iStartRow, iColumn).Value & "" = "STA" Then
                    Exit Do
                ElseIf xlObj.Cells(iStartRow, iColumn).Value & "" = "" Then
                    Exit Do
                End If
            Loop
        End If
        If xlObj.Cells(iStartRow, iColumn).Value & "" = "STA" Then
            Exit Do
        End If
    Loop
    
    For iColumn = 1 To xlObj.Columns.Count
        If xlObj.Cells(iStartRow, iColumn).Value & "" = "" Then
            Exit For
        End If
        cLabel(iColumn) = UnDotSpace(xlObj.Cells(iStartRow, iColumn).Value)
        If cLabel(iColumn) = "Horse" Then
            cLabel(iColumn) = "Name_horse"
        End If
    Next iColumn
    
    iColumnCount = iColumn - 1
    ReDim Preserve cLabel(iColumnCount)
    
    iRow = iStartRow
    
    Do While xlObj.Cells(iRow, 1).Value <> ""
        iRow = iRow + 1
        cSta = ""
        cPersonId = ""
        cHorseId = ""
        cName_First = ""
        cName_Last = ""
        cName_Horse = ""
        cHorses_FEIFId = ""
        cPersons_FEIFId = ""
        
        For iColumn = 1 To iColumnCount
            Select Case xlObj.Cells(iStartRow, iColumn).Value & ""
            Case "STA"
                cSta = Trim$(xlObj.Cells(iRow, iColumn).Value)
            Case "PersonId"
                cPersonId = Trim$(xlObj.Cells(iRow, iColumn).Value)
            Case "HorseId"
                cHorseId = Trim$(xlObj.Cells(iRow, iColumn).Value)
            Case "Name_First"
                cName_First = xlObj.Cells(iRow, iColumn).Value
            Case "Name_Last"
                cName_Last = xlObj.Cells(iRow, iColumn).Value
            Case "Name_Horse", "Horse"
                cName_Horse = xlObj.Cells(iRow, iColumn).Value
            Case "Rider", "Name"
                cName_First = MakeFirst(xlObj.Cells(iRow, iColumn).Value)
                cName_Last = MakeLast(xlObj.Cells(iRow, iColumn).Value)
            Case "Persons.FEIFId", "Person_FEIFId", "Rider_FEIFId"
                cPersons_FEIFId = Trim$(xlObj.Cells(iRow, iColumn).Value)
            Case "Horses.FEIFId", "Horse_FEIFId"
                cHorses_FEIFId = Trim$(xlObj.Cells(iRow, iColumn).Value)
            Case "FEIFId"
                If Left$(xlObj.Cells(iRow, iColumn).Value, 2) = "FF" Then
                    cPersons_FEIFId = Trim$(xlObj.Cells(iRow, iColumn).Value)
                Else
                    cHorses_FEIFId = Trim$(xlObj.Cells(iRow, iColumn).Value)
                End If
            End Select
        Next iColumn
        
        
        If Val(cSta) > 0 Then
            'remove old participant
            cSta = Format$(Val(cSta), "000")
            If cPersonId = "" Then
                Set rstParticipant = mdbMain.OpenRecordset("SELECT PersonId FROM Participants WHERE Sta='" & cSta & "'")
                If rstParticipant.RecordCount > 0 Then
                    cPersonId = rstParticipant.Fields(0)
                End If
            End If
            If cPersonId = "" And cName_First <> "" And cName_Last <> "" Then
                Set rstRider = mdbMain.OpenRecordset("SELECT PersonId FROM Persons WHERE Name_First LIKE'" & cName_First & "' AND Name_Last LIKE '" & cName_Last & "'")
                If rstRider.RecordCount > 0 Then
                    cPersonId = rstRider.Fields(0)
                End If
            End If
            If cPersonId = "" Then
                cPersonId = CreatePersonId
            End If
            
            If cHorseId = "" Then
                Set rstParticipant = mdbMain.OpenRecordset("SELECT HorseId FROM Participants WHERE Sta='" & cSta & "'")
                If rstParticipant.RecordCount > 0 Then
                    cHorseId = rstParticipant.Fields(0)
                End If
            End If
            If cHorseId = "" And cName_Horse <> "" Then
                Set rstHorse = mdbMain.OpenRecordset("SELECT HorseId FROM Horses WHERE Name_Horse LIKE '" & cName_Horse & "'")
                If rstHorse.RecordCount > 0 Then
                    cHorseId = rstHorse.Fields(0)
                End If
            End If
            If cHorseId = "" Then
                cHorseId = CreateHorseId
            End If
            
            Set rstParticipant = mdbMain.OpenRecordset("SELECT * FROM Participants WHERE Sta LIKE '" & cSta & "'")
            With rstParticipant
                If .RecordCount = 0 Then
                    .AddNew
                    .Fields("STA") = cSta
                Else
                    .Edit
                End If
                .Fields("PersonId") = Left$(cPersonId, .Fields("PersonId").Size)
                .Fields("HorseId") = Left$(cHorseId, .Fields("HorseId").Size)
                If cFieldList(1) = "" Then
                    cFieldList(1) = "|"
                    For Each fld In .Fields
                        cFieldList(1) = cFieldList(1) & fld.Name & "|"
                    Next
                End If
            End With
            
            Set rstRider = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE PersonId LIKE '" & cPersonId & "'")
            With rstRider
                If .RecordCount = 0 Then
                    .AddNew
                    .Fields("PersonId") = Left$(cPersonId, .Fields("PersonId").Size)
                Else
                    .Edit
                End If
                .Fields("FEIFId") = cPersons_FEIFId
                If cFieldList(2) = "" Then
                    cFieldList(2) = "|"
                    For Each fld In .Fields
                        cFieldList(2) = cFieldList(2) & fld.Name & "|"
                    Next
                End If
            End With
            
            Set rstHorse = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE HorseId LIKE '" & cHorseId & "'")
            With rstHorse
                If .RecordCount = 0 Then
                    .AddNew
                    .Fields("HorseId") = Left$(cHorseId, .Fields("HorseId").Size)
                Else
                    .Edit
                End If
                .Fields("FEIFId") = cHorses_FEIFId
                If cFieldList(3) = "" Then
                    cFieldList(3) = "|"
                    For Each fld In .Fields
                        cFieldList(3) = cFieldList(3) & fld.Name & "|"
                    Next
                End If
            End With
            
            For iColumn = 1 To iColumnCount
                vValue = Trim$(xlObj.Cells(iRow, iColumn).Value)
                
                If cLabel(iColumn) = "Rider" Or cLabel(iColumn) = "Name" Then
                    rstRider.Fields("Name_first") = MakeFirst(CStr(vValue))
                    rstRider.Fields("Name_last") = MakeLast(CStr(vValue))
                ElseIf cLabel(iColumn) = "FEIFId" Then
                    'ignore, has been added already
                ElseIf cLabel(iColumn) = "PersonId" Then
                    'ignore, has been added already
                ElseIf cLabel(iColumn) = "HorseId" Then
                    'ignore, has been added already
                ElseIf InStr(cFieldList(1), "|" & cLabel(iColumn) & "|") > 0 Then
                    If cLabel(iColumn) = "STA" Then
                        'ignore, has been added already2
                    Else
                        rstParticipant.Fields(cLabel(iColumn)) = vValue
                    End If
                ElseIf InStr(cFieldList(2), "|" & cLabel(iColumn) & "|") > 0 Then
                    If rstRider.Fields(cLabel(iColumn)).Type = dbDate Then
                        rstRider.Fields(cLabel(iColumn)) = CDate(vValue)
                        If Err > 0 Then
                            Err = 0
                            rstRider.Fields(cLabel(iColumn)) = vValue
                        End If
                    Else
                        rstRider.Fields(cLabel(iColumn)) = vValue
                    End If
                ElseIf InStr(cFieldList(3), "|" & cLabel(iColumn) & "|") > 0 Then
                    If rstHorse.Fields(cLabel(iColumn)).Type = dbDate Then
                        rstHorse.Fields(cLabel(iColumn)) = CDate(vValue)
                        If Err > 0 Then
                            Err = 0
                            rstHorse.Fields(cLabel(iColumn)) = vValue
                        End If
                    Else
                        rstHorse.Fields(cLabel(iColumn)) = vValue
                    End If
                ElseIf InStr(cFieldList(4), "|" & cLabel(iColumn) & "|") > 0 Then
                    cValue = CStr(vValue)
                    Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE STA='" & cSta & "' AND Code='" & cLabel(iColumn) & "' AND Status=0")
                    With rstEntry
                        If .RecordCount = 0 And cValue = "" Then
                            'ignore
                        ElseIf .RecordCount = 0 And cValue <> "" Then
                            .AddNew
                            .Fields("Sta") = Left$(cSta, .Fields("STA").Size)
                            .Fields("Code") = Left$(cLabel(iColumn), .Fields("Code").Size)
                            .Fields("Group") = 0
                            .Fields("Status") = 0
                            .Fields("Deleted") = 0
                            .Fields("Timestamp") = Now
                        ElseIf .RecordCount > 0 And cValue = "" Then
                            .Delete
                        ElseIf cValue <> "" Then
                            .Edit
                        End If
                        If cValue <> "" Then
                            If Right$(cValue, 1) = cRight Then
                                cValue = Trim$(Left$(cValue, Len(cValue) - 1))
                                .Fields("RR") = True
                            ElseIf InStr(" " & cValue & " ", " " & cRight & " ") > 0 Then
                                cValue = Replace(" " & cValue, " " & cRight & " ", " ")
                                .Fields("RR") = True
                            Else
                                .Fields("RR") = False
                            End If
                            cValue = Replace(Trim$(cValue), "  ", " ")
                            Parse cPosition, cValue, " "
                            .Fields("Position") = Val(cPosition)
                            Parse cColor, cValue, " "
                            .Fields("Color") = cColor
                            .Update
                        End If
                    End With
                End If
            Next iColumn
            
            rstParticipant.Update
            rstRider.Update
            rstHorse.Update
            
            StatusMessage Translate("Importing Participants", mcLanguage) & " [" & cSta & " - " & cName_First & " " & cName_Last & " - " & cName_Horse & "]"
            
            ImportXls = iRow - iStartRow
            
        End If
    Loop
    
    rstParticipant.Close
    rstRider.Close
    rstHorse.Close
    
    Set rstRider = Nothing
    Set rstHorse = Nothing
    Set rstParticipant = Nothing
    Set rstEntry = Nothing
    
    On Local Error GoTo 0
    
End Function
Public Sub ExportCsv()
    Dim rstXls As DAO.Recordset
    Dim rstTest As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    Dim cExcelFile As String
    Dim iExcelFile As Integer
    Dim iKey As Integer
    
    On Local Error Resume Next
    
    ReadIniFile gcIniFile, "Export", "Csv", cExcelFile
    If cExcelFile = "" Then
        cExcelFile = NameOfFile(Dir$(mcDatabaseName)) & ".Csv"
    End If
    With frmMain.CommonDialog1
        .DefaultExt = "Csv"
        .DialogTitle = Translate("Enter a file name", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Comma-delimited (*.Csv)|*.Csv"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    On Local Error GoTo 0
    
    SetMouseHourGlass
    If Dir$(cExcelFile) <> "" Then
        iKey = MsgBox(Translate("Overwrite existing file", mcLanguage) & " '" & cExcelFile & "'?", vbQuestion + vbYesNo)
    Else
        iKey = MsgBox(Translate("Create file", mcLanguage) & " '" & cExcelFile & "'?", vbQuestion + vbYesNo)
    End If
    If iKey = vbYes Then
        StatusMessage Translate("Exporting Participants", mcLanguage) & "..."
        
        frmMain.Enabled = False
        
        iExcelFile = FreeFile
        Open cExcelFile For Output Access Write As #iExcelFile
        
        Set rstXls = mdbMain.OpenRecordset("SELECT DISTINCT Participants.STA, Persons.Name_First & ' ' & Persons.Name_Last AS Rider, Horses.Name_Horse AS Horse, Participants.Club as Club, Participants.Team as Team , Participants.Class as Class FROM Entries, (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID ORDER BY Participants.STA;")
        
        Print #iExcelFile, "STA";
        Print #iExcelFile, mcExcelSeparator & "Rider";
        Print #iExcelFile, mcExcelSeparator & "Horse";
        Print #iExcelFile, mcExcelSeparator & "Club";
        Print #iExcelFile, mcExcelSeparator & "Team";
        Print #iExcelFile, mcExcelSeparator & "Class";
        
        Set rstTest = mdbMain.OpenRecordset("SELECT DISTINCT CODE FROM Entries WHERE Code IN (SELECT Code FROM Tests) ORDER BY Code;")
        If rstTest.RecordCount > 0 Then
            Do While Not rstTest.EOF
                Print #iExcelFile, mcExcelSeparator & Chr$(34) & rstTest.Fields("Code") & Chr$(34);
                rstTest.MoveNext
            Loop
        End If
        Print #iExcelFile, ""
        
        If rstXls.RecordCount > 0 Then
            Do While Not rstXls.EOF
                If rstXls.AbsolutePosition Mod 10 = 0 Then
                    StatusMessage Translate("Exporting Participants", mcLanguage) & " [" & rstXls.AbsolutePosition & "]"
                End If
                Print #iExcelFile, rstXls.Fields("STA");
                Print #iExcelFile, mcExcelSeparator & Replace(rstXls.Fields("Rider"), mcExcelSeparator, ",");
                Print #iExcelFile, mcExcelSeparator & Replace(rstXls.Fields("Horse"), mcExcelSeparator, ",");
                Print #iExcelFile, mcExcelSeparator & Replace(rstXls.Fields("Club") & "", mcExcelSeparator, ",");
                Print #iExcelFile, mcExcelSeparator & Replace(rstXls.Fields("Team") & "", mcExcelSeparator, ",");
                Print #iExcelFile, mcExcelSeparator & Replace(rstXls.Fields("Class") & "", mcExcelSeparator, ",");
                If rstTest.RecordCount > 0 Then
                    rstTest.MoveFirst
                    Do While Not rstTest.EOF
                        Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code = '" & rstTest.Fields("Code") & "' AND STA='" & rstXls.Fields("STA") & "' AND Status=0")
                        If rstEntries.RecordCount > 0 Then
                            Print #iExcelFile, mcExcelSeparator & Format$(rstEntries.Fields("Position") + 0) & IIf(rstEntries.Fields("RR") + 0 <> 0, "R", "L");
                        Else
                            Print #iExcelFile, mcExcelSeparator;
                        End If
                        rstEntries.Close
                        rstTest.MoveNext
                    Loop
                    Print #iExcelFile, ""
                End If
                rstXls.MoveNext
            Loop
        End If
        
        Close #iExcelFile
        StatusMessage
        
        WriteIniFile gcIniFile, "Export", "Csv", cExcelFile
        MsgBox rstXls.RecordCount & " " & Translate("Participants have been exported to", mcLanguage) & " '" & cExcelFile & "'."
        rstXls.Close
        rstTest.Close
    Else
        MsgBox Translate("No file created.", mcLanguage)
    End If
    Set rstXls = Nothing
    Set rstTest = Nothing
    Set rstEntries = Nothing
    
    StatusMessage
    
    SetMouseNormal

End Sub
Public Sub ExportExcel()
    Dim rstXls As DAO.Recordset
    Dim rstTest As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    Dim rstFld As DAO.Recordset
    Dim fld As DAO.Field
    Dim cExcelFile As String
    Dim cQry As String
    
    Dim iExcelFile As Integer
    Dim iKey As Integer
    
    Dim xlObj As Object
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iColumnCount As Integer
    Dim cTemp As String
        
    On Local Error Resume Next
        
    ReadIniFile gcIniFile, "Export", "Excel", cExcelFile
    If cExcelFile = "" Then
        cExcelFile = NameOfFile(Dir$(mcDatabaseName)) & ".Xls"
    End If
    With frmMain.CommonDialog1
        .DefaultExt = "Xls"
        .DialogTitle = Translate("Enter a file name", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Excel (*.Xls)|*.Xls"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    On Local Error GoTo 0
    
    
    SetMouseHourGlass
    If Dir$(cExcelFile) <> "" Then
        KillFile cExcelFile
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Create file", mcLanguage) & " '" & cExcelFile & "'?", vbQuestion + vbYesNo)
    End If
    If iKey = vbYes Then
        
        frmMain.Enabled = False
        
        StatusMessage Translate("Exporting Participants", mcLanguage) & "..."
        
        iColumnCount = 0
        ReDim cColumn(50)
        
        Set xlObj = CreateObject("Excel.Sheet")

        xlObj.Application.Visible = False
        
        cQry = "SELECT DISTINCT "
        cQry = cQry & " "
        Set rstFld = mdbMain.OpenRecordset("SELECT [Field] FROM [Fields] WHERE [Table]='Participants' AND Type<>'Memo' AND Status>1 ORDER BY Seq")
        If rstFld.RecordCount = 0 Then
            Set rstFld = mdbMain.OpenRecordset("SELECT [Field] FROM [Fields] WHERE [Table]='Participants' AND Type<>'Memo' AND NOT Comment &'' LIKE '*IPZV*' ORDER BY Seq")
        End If
        If rstFld.RecordCount > 0 Then
            Do While Not rstFld.EOF
                iColumnCount = iColumnCount + 1
                cQry = cQry & "[Participants].[" & rstFld.Fields(0) & "], "
                rstFld.MoveNext
            Loop
        End If
        
        Set rstFld = mdbMain.OpenRecordset("SELECT [Field] FROM [Fields] WHERE [Table]='Persons' AND Type<>'Memo' AND Status>1 ORDER BY Seq")
        If rstFld.RecordCount > 0 Then
            Do While Not rstFld.EOF
                iColumnCount = iColumnCount + 1
                If rstFld.Fields(0) = "FEIFId" Then
                    cQry = cQry & "[Persons].[" & rstFld.Fields(0) & "] AS Person_FEIFId, "
                Else
                    cQry = cQry & "[Persons].[" & rstFld.Fields(0) & "], "
                End If
                rstFld.MoveNext
            Loop
        End If
        
        Set rstFld = mdbMain.OpenRecordset("SELECT [Field] FROM [Fields] WHERE [Table]='Horses' AND Type<>'Memo' AND Status>1 ORDER BY Seq")
        If rstFld.RecordCount > 0 Then
            Do While Not rstFld.EOF
                iColumnCount = iColumnCount + 1
                If rstFld.Fields(0) = "FEIFId" Then
                    cQry = cQry & "[Horses].[" & rstFld.Fields(0) & "] AS Horse_FEIFId, "
                Else
                    cQry = cQry & "[Horses].[" & rstFld.Fields(0) & "], "
                End If
                rstFld.MoveNext
            Loop
        End If
        cQry = Left$(cQry, Len(cQry) - 2)
        
        cQry = cQry & " FROM (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID ORDER BY Participants.STA;"
        Set rstXls = mdbMain.OpenRecordset(cQry)
        
        If frmMain.mnuTestAll.Checked = True Then
            Set rstTest = mdbMain.OpenRecordset("SELECT DISTINCT Code FROM Tests ORDER BY Code;")
        Else
            Set rstTest = mdbMain.OpenRecordset("SELECT DISTINCT Tests.Code FROM Tests INNER JOIN TestInfo ON Tests.Code=TestInfo.Code Where Testinfo.Nr>0 ORDER BY Tests.Code;")
        End If
        
        iColumn = 0
        iRow = 1
        For iColumn = 1 To iColumnCount
            xlObj.Application.Cells(iRow, iColumn).Value = rstXls.Fields(iColumn - 1).Name
        Next
        
        If rstTest.RecordCount > 0 Then
            rstTest.MoveLast
            rstTest.MoveFirst
            iColumnCount = rstTest.RecordCount + iColumnCount
            For iTemp = rstXls.Fields.Count + 1 To iColumnCount
                xlObj.Application.Cells(iRow, iTemp).Value = rstTest.Fields("Code")
                rstTest.MoveNext
            Next iTemp
        End If
        
        If rstXls.RecordCount > 0 Then
            Do While Not rstXls.EOF
                iRow = iRow + 1
                For iTemp = 1 To rstXls.Fields.Count
                    If rstXls.Fields(iTemp - 1).Type = dbDate Then
                        xlObj.Application.Cells(iRow, iTemp).Value = Format$(rstXls.Fields(iTemp - 1).Value, "DD/MM/YYYY")
                    Else
                        xlObj.Application.Cells(iRow, iTemp).Value = rstXls.Fields(iTemp - 1).Value
                    End If
                Next iTemp
                If rstTest.RecordCount > 0 Then
                    rstTest.MoveFirst
                    For iTemp = rstXls.Fields.Count + 1 To iColumnCount
                        Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code = '" & rstTest.Fields("Code") & "' AND STA='" & rstXls.Fields("STA") & "' AND Status=0")
                        If rstEntries.RecordCount > 0 Then
                            xlObj.Application.Cells(iRow, iTemp).Value = Format$(rstEntries.Fields("Position") + 0) & IIf(rstEntries.Fields("RR") + 0 <> 0, "R", "L")
                        End If
                        rstEntries.Close
                        rstTest.MoveNext
                    Next
                End If
                StatusMessage Translate("Exporting Participants", mcLanguage) & " [" & rstXls.Fields("STA") & " - " & rstXls.Fields("Name_first") & " " & rstXls.Fields("Name_Last") & " - " & rstXls.Fields("Name_horse") & "]"
                rstXls.MoveNext
            Loop
        End If
        
        xlObj.SaveAs cExcelFile
        ' Close Excel with the Quit method on the Application object.
        xlObj.Application.Quit
        ' Release the object variable.
        Set xlObj = Nothing

        StatusMessage
        
        WriteIniFile gcIniFile, "Export", "Excel", cExcelFile
        iKey = MsgBox(rstXls.RecordCount & " " & Translate("Participants have been exported to", mcLanguage) & " '" & cExcelFile & "'." & vbCrLf & Translate("Open", mcLanguage) & " '" & Dir$(cExcelFile) & "'?", vbDefaultButton1 + vbYesNo)
        rstXls.Close
        rstTest.Close
        If iKey = vbYes Then
            ShowDocument cExcelFile, frmMain
        End If
    Else
        MsgBox Translate("No file created.", mcLanguage)
    End If
    Set rstXls = Nothing
    Set rstTest = Nothing
    Set rstEntries = Nothing
    
    StatusMessage
    
    SetMouseNormal

End Sub

Public Function MakeFirst(cRider As String) As String
    Dim iTemp As Integer
    iTemp = InStr(cRider, "&")
    If iTemp = 0 Then
        iTemp = InStr(cRider, " ")
    End If
    If iTemp > 1 Then
        MakeFirst = RTrim$(Left$(cRider, iTemp - 1))
    Else
        MakeFirst = cRider
    End If
End Function
Public Sub ExportTestToExcel()
    Dim rstXls As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    Dim cExcelFile As String
    Dim cQry As String
    
    Dim iExcelFile As Integer
    Dim iKey As Integer
    
    Dim xlObj As Object
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iRow As Integer
    Dim iSection As Integer
    Dim iColumn As Integer
    Dim iColumnCount As Integer
    Dim cTemp As String
        
    On Local Error Resume Next
        
    cExcelFile = mcExcelDir & frmMain.TestCode & "_" & Left$(frmMain.tbsSelFin.SelectedItem.Caption, 5) & ".Xls"
    With frmMain.CommonDialog1
        .DefaultExt = "Xls"
        .DialogTitle = Translate("Enter a file name", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Excel (*.Xls)|*.Xls"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    On Local Error GoTo 0
    
    SetMouseHourGlass
    
    If Dir$(cExcelFile) <> "" Then
        KillFile cExcelFile
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Create file", mcLanguage) & " '" & cExcelFile & "'?", vbQuestion + vbYesNo + vbDefaultButton1)
    End If
    
    If iKey = vbYes Then
        
        frmMain.Enabled = False
        
        StatusMessage Translate("Exporting", mcLanguage) & frmMain.TestCode & "..."
        
        iKey = MsgBox(Translate("Order by starting order (Yes) or by start number (No)", mcLanguage) & "?", vbQuestion + vbDefaultButton1 + vbYesNo)
        
        iColumnCount = 0
        ReDim cColumn(50)
        
        Set xlObj = CreateObject("Excel.Sheet")
        If Err > 0 Then
            Err = 0
            Set xlObj = GetObject("", "Excel.Sheet")
        End If
        
        xlObj.Application.Visible = False
        
        cQry = "SELECT Entries.STA, Persons.Name_First & ' ' & Persons.Name_Last AS Rider, Horses.Name_Horse"
        cQry = cQry & " FROM ((Entries INNER JOIN Participants ON Entries.STA = Participants.STA) "
        cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) "
        cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
        cQry = cQry & " WHERE Entries.Code='" & frmMain.TestCode & "' "
        cQry = cQry & " AND Deleted<>1 "
        cQry = cQry & " AND Entries.Status=" & frmMain.TestStatus
        If iKey = vbYes Then
            cQry = cQry & " ORDER BY Entries.Position;"
        Else
            cQry = cQry & " ORDER BY Entries.STA;"
        End If
        
        Set rstEntries = mdbMain.OpenRecordset(cQry)
                
        iRow = 1
        xlObj.Application.Cells(iRow, 1) = "STA"
        xlObj.Application.Cells(iRow, 2) = "Rider"
        xlObj.Application.Cells(iRow, 3) = "Horse"
        If frmMain.txtTime.Visible = False Then
            xlObj.Application.Cells(iRow, 4) = "Section"
            xlObj.Application.Cells(iRow, 5) = "Section Name"
            xlObj.Application.Cells(iRow, 6) = "Mark1"
            xlObj.Application.Cells(iRow, 7) = "Mark2"
            xlObj.Application.Cells(iRow, 8) = "Mark3"
            xlObj.Application.Cells(iRow, 9) = "Mark4"
            xlObj.Application.Cells(iRow, 10) = "Mark5"
            iColumnCount = 10
        Else
            xlObj.Application.Cells(iRow, 4) = "Run"
            xlObj.Application.Cells(iRow, 5) = "Section Name"
            xlObj.Application.Cells(iRow, 6) = "Time"
            iColumnCount = 5
        End If
        
        If rstEntries.RecordCount > 0 Then
            Do While Not rstEntries.EOF
                For iSection = 1 To frmMain.tbsSection(frmMain.TestStatus).Tabs.Count
                    iRow = iRow + 1
                    For iTemp = 1 To rstEntries.Fields.Count
                        xlObj.Application.Cells(iRow, iTemp).Value = rstEntries.Fields(iTemp - 1).Value
                    Next iTemp
                    xlObj.Application.Cells(iRow, rstEntries.Fields.Count + 1).Value = iSection
                    xlObj.Application.Cells(iRow, rstEntries.Fields.Count + 2).Value = ClipAmp(frmMain.tbsSection(frmMain.TestStatus).Tabs(iSection).Caption)
                    Set rstXls = mdbMain.OpenRecordset("SELECT STA,Mark1,Mark2,Mark3,Mark4,Mark5 FROM Marks WHERE STA='" & rstEntries.Fields(0) & "' AND Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus & " AND Section=" & iSection)
                    If rstXls.RecordCount > 0 Then
                        For iTemp = 1 To IIf(frmMain.txtTime.Visible = True, 1, 5)
                            xlObj.Application.Cells(iRow, rstEntries.Fields.Count + iTemp + 2).Value = Format$(rstXls.Fields(iTemp).Value)
                        Next iTemp
                    End If
                    rstXls.Close
            
                Next iSection
                StatusMessage Translate("Exporting", mcLanguage) & " [" & rstEntries.Fields("STA") & "]"
                rstEntries.MoveNext
            Loop
        End If
        
        xlObj.SaveAs cExcelFile
        ' Close Excel with the Quit method on the Application object.
        xlObj.Application.Quit
        ' Release the object variable.
        
        Set xlObj = Nothing

        StatusMessage
        
        DoEvents
        
        iKey = MsgBox(iRow - 1 & " " & Translate("Rows have been exported to", mcLanguage) & " '" & cExcelFile & "'." & vbCrLf & Translate("Open", mcLanguage) & " '" & Dir$(cExcelFile) & "'?", vbYesNo + vbQuestion)
        If iKey = vbYes Then
            ShowDocument cExcelFile, frmMain
        End If
    
    Else
        MsgBox Translate("No file created.", mcLanguage)
    End If
    Set rstXls = Nothing
    
    StatusMessage
    
    SetMouseNormal

End Sub
Public Sub ExportTestToExcelForJudges()
    Dim rstXls As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    Dim cExcelFile As String
    Dim cQry As String
    
    Dim iExcelFile As Integer
    Dim iKey As Integer
    
    Dim xlObj As Object
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iRow As Integer
    Dim iSection As Integer
    Dim iColumn As Integer
    Dim iColumnCount As Integer
    Dim cTemp As String
        
    Dim sAVG() As Single
    ReDim sAVG(5)
    Dim sMin As Single
    Dim sMax As Single
    
    On Local Error Resume Next
        
    cExcelFile = mcExcelDir & frmMain.TestCode & "_" & Left$(frmMain.tbsSelFin.SelectedItem.Caption, 5) & "_" & Translate("Judges", mcLanguage) & ".Xls"
    With frmMain.CommonDialog1
        .DefaultExt = "Xls"
        .DialogTitle = Translate("Enter a file name", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Excel (*.Xls)|*.Xls"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    On Local Error GoTo 0
    
    SetMouseHourGlass
    
    If Dir$(cExcelFile) <> "" Then
        KillFile cExcelFile
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Create file", mcLanguage) & " '" & cExcelFile & "'?", vbQuestion + vbYesNo + vbDefaultButton1)
    End If
    
    If iKey = vbYes Then
        
        frmMain.Enabled = False
        
        StatusMessage Translate("Exporting", mcLanguage) & frmMain.TestCode & "..."
        
        iKey = MsgBox(Translate("Order by starting order (Yes) or by start number (No)", mcLanguage) & "?", vbQuestion + vbDefaultButton1 + vbYesNo)
        
        iColumnCount = 0
        ReDim cColumn(50)
        
        Set xlObj = CreateObject("Excel.Sheet")
        If Err > 0 Then
            Err = 0
            Set xlObj = GetObject("", "Excel.Sheet")
        End If
        
        xlObj.Application.Visible = False
        
        cQry = "SELECT Entries.STA, Persons.Name_First & ' ' & Persons.Name_Last AS Rider, Horses.Name_Horse"
        cQry = cQry & " FROM ((Entries INNER JOIN Participants ON Entries.STA = Participants.STA) "
        cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) "
        cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
        cQry = cQry & " WHERE Entries.Code='" & frmMain.TestCode & "' "
        cQry = cQry & " AND Deleted<>1 "
        cQry = cQry & " AND Entries.Status=" & frmMain.TestStatus
        If iKey = vbYes Then
            cQry = cQry & " ORDER BY Entries.Position;"
        Else
            cQry = cQry & " ORDER BY Entries.STA;"
        End If
        
        Set rstEntries = mdbMain.OpenRecordset(cQry)
                
        iRow = 1
        xlObj.Application.Cells(iRow, 1) = "STA"
        xlObj.Application.Cells(iRow, 2) = Translate("Rider", mcLanguage)
        xlObj.Application.Cells(iRow, 3) = Translate("Horse", mcLanguage)
        If frmMain.txtTime.Visible = False Then
            xlObj.Application.Cells(iRow, 4) = Translate("Section", mcLanguage)
            xlObj.Application.Cells(iRow, 5) = Translate("Section Name", mcLanguage)
            xlObj.Application.Cells(iRow, 6) = Translate("Judge", mcLanguage) & " A"
            xlObj.Application.Cells(iRow, 7) = Translate("Judge", mcLanguage) & " B"
            xlObj.Application.Cells(iRow, 8) = Translate("Judge", mcLanguage) & " C"
            xlObj.Application.Cells(iRow, 9) = Translate("Judge", mcLanguage) & " D"
            xlObj.Application.Cells(iRow, 10) = Translate("Judge", mcLanguage) & " E"
            xlObj.Application.Cells(iRow, 11) = Translate("Range", mcLanguage)
            xlObj.Application.Cells(iRow, 12) = Translate("Average", mcLanguage)
            xlObj.Application.Cells(iRow, 13) = Translate("Deviation", mcLanguage) & " A"
            xlObj.Application.Cells(iRow, 14) = Translate("Deviation", mcLanguage) & " B"
            xlObj.Application.Cells(iRow, 15) = Translate("Deviation", mcLanguage) & " C"
            xlObj.Application.Cells(iRow, 16) = Translate("Deviation", mcLanguage) & " D"
            xlObj.Application.Cells(iRow, 17) = Translate("Deviation", mcLanguage) & " E"
            iColumnCount = 10
        End If
        
        If rstEntries.RecordCount > 0 Then
            Do While Not rstEntries.EOF
                For iSection = 1 To frmMain.tbsSection(frmMain.TestStatus).Tabs.Count
                    iRow = iRow + 1
                    For iTemp = 1 To rstEntries.Fields.Count
                        xlObj.Application.Cells(iRow, iTemp).Value = rstEntries.Fields(iTemp - 1).Value
                    Next iTemp
                    xlObj.Application.Cells(iRow, rstEntries.Fields.Count + 1).Value = iSection
                    xlObj.Application.Cells(iRow, rstEntries.Fields.Count + 2).Value = ClipAmp(frmMain.tbsSection(frmMain.TestStatus).Tabs(iSection).Caption)
                    Set rstXls = mdbMain.OpenRecordset("SELECT STA,Mark1,Mark2,Mark3,Mark4,Mark5 FROM Marks WHERE STA='" & rstEntries.Fields(0) & "' AND Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus & " AND Section=" & iSection)
                    If rstXls.RecordCount > 0 Then
                        sAVG(0) = 0
                        sMin = 10
                        sMax = 0
                        For iTemp = 1 To frmMain.TestJudges
                            xlObj.Application.Cells(iRow, rstEntries.Fields.Count + iTemp + 2).Value = CSng(rstXls.Fields(iTemp).Value)
                            sAVG(0) = sAVG(0) + CSng(rstXls.Fields(iTemp).Value)
                            If CSng(rstXls.Fields(iTemp).Value) < sMin Then
                                sMin = CSng(rstXls.Fields(iTemp).Value)
                            End If
                            If CSng(rstXls.Fields(iTemp).Value) > sMax Then
                                sMax = CSng(rstXls.Fields(iTemp).Value)
                            End If
                        Next iTemp
                        xlObj.Application.Cells(iRow, 11).Value = sMax - sMin
                        xlObj.Application.Cells(iRow, 12).Value = CSng(Format(sAVG(0) / frmMain.TestJudges, "0.000"))
                        For iTemp = 1 To frmMain.TestJudges
                            xlObj.Application.Cells(iRow, iTemp + 12).Value = CSng(Format(CSng(rstXls.Fields(iTemp).Value) - (sAVG(0) / frmMain.TestJudges), "0.00"))
                            sAVG(iTemp) = sAVG(iTemp) + CSng(Format(CSng(rstXls.Fields(iTemp).Value) - (sAVG(0) / frmMain.TestJudges), "0.00"))
                        Next iTemp
                    End If
                    rstXls.Close
                Next iSection
                
                StatusMessage Translate("Exporting", mcLanguage) & " [" & rstEntries.Fields("STA") & "]"
                rstEntries.MoveNext
            Loop
        End If
        'iRow = iRow + 1
        'For iTemp = 1 To 5
        '    xlObj.Application.Cells(iRow, 12 + iTemp).Value = CSng(Format(sAVG(iTemp), "0.00"))
        'Next iTemp
        
        xlObj.SaveAs cExcelFile
        ' Close Excel with the Quit method on the Application object.
        xlObj.Application.Quit
        ' Release the object variable.
        
        Set xlObj = Nothing

        StatusMessage
        
        DoEvents
        
        iKey = MsgBox(iRow - 1 & " " & Translate("Rows have been exported to", mcLanguage) & " '" & cExcelFile & "'." & vbCrLf & Translate("Open", mcLanguage) & " '" & Dir$(cExcelFile) & "'?", vbYesNo + vbQuestion)
        If iKey = vbYes Then
            ShowDocument cExcelFile, frmMain
        End If
    
    Else
        MsgBox Translate("No file created.", mcLanguage)
    End If
    Set rstXls = Nothing
    
    StatusMessage
    
    SetMouseNormal


End Sub
Public Sub ImportTestFromExcel()
    Dim xlObj As Object
    Dim cExcelFile As String
    Dim iKey As Integer
    Dim cSta As String
    Dim curMark1 As Variant
    Dim curMark2 As Variant
    Dim curMark3 As Variant
    Dim curMark4 As Variant
    Dim curMark5 As Variant
    Dim iSection As Integer
    Dim iStartRow As Integer
    Dim iColumn As Integer
    Dim iColumnCount As Integer
    Dim iRow As Integer
    Dim iCurrentSection As Integer
    Dim iTemp As Integer
    
    On Local Error Resume Next
        
    SetMouseHourGlass
    
    With frmMain.CommonDialog1
        .DefaultExt = "Xls"
        .DialogTitle = Translate("Select an Excel-sheet", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Excel (*.Xls)|*.Xls"
        .InitDir = mcExcelDir
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    
    If cExcelFile <> "" And cExcelFile <> Chr$(27) Then
        iKey = MsgBox(cExcelFile & ": " & Translate("Import selected file into current test", mcLanguage) & " " & frmMain.TestCode & " - " & frmMain.TestName & " - " & frmMain.tbsSelFin.SelectedItem.Caption & "?", vbYesNo + vbExclamation)
    Else
        MsgBox Translate("No proper Excel-sheet selected.", mcLanguage)
        iKey = vbNo
    End If
    
    If iKey = vbYes Then
        
        iKey = MsgBox(Translate("Please be aware that the import of marks from MS Excel might overwrite existing data!", mcLanguage) & vbCrLf & Translate("Consider to make a backup first!", mcLanguage), vbExclamation + vbOKCancel)
        If iKey = vbOK Then
            SetMouseHourGlass
            
            Set xlObj = GetObject(cExcelFile)
            
            If Val(xlObj.Application.Version) >= 8 Then
                Set xlObj = xlObj.ActiveSheet
            End If
            
            ReDim cLabel(xlObj.Columns.Count)
                
            iStartRow = 0
            Do While iStartRow < 255
                iStartRow = iStartRow + 1
                iColumn = 0
                If xlObj.Cells(1, iColumn).Value & "" <> "" Then
                    Do While iColumn < 255
                        iColumn = iColumn + 1
                        If xlObj.Cells(iStartRow, iColumn).Value & "" = "STA" Then
                            Exit Do
                        ElseIf xlObj.Cells(iStartRow, iColumn).Value & "" = "" Then
                            Exit Do
                        End If
                    Loop
                End If
                If xlObj.Cells(iStartRow, iColumn).Value & "" = "STA" Then
                    Exit Do
                End If
            Loop
            
            For iColumn = 1 To xlObj.Columns.Count
                If xlObj.Cells(iStartRow, iColumn).Value & "" = "" Then
                    Exit For
                End If
                cLabel(iColumn) = UnDotSpace(xlObj.Cells(iStartRow, iColumn).Value)
            Next iColumn
            
            iColumnCount = iColumn - 1
            ReDim Preserve cLabel(iColumnCount)
            
            iRow = iStartRow
            
            Do While xlObj.Cells(iRow, 1).Value <> ""
                iRow = iRow + 1
                cSta = ""
                
                curMark1 = Null
                curMark2 = Null
                curMark3 = Null
                curMark4 = Null
                curMark5 = Null
                iSection = 1
                
                For iColumn = 1 To iColumnCount
                    If xlObj.Cells(iRow, iColumn).Value <> "" Then
                        Select Case Trim$(xlObj.Cells(iStartRow, iColumn).Value)
                        Case "STA"
                            cSta = xlObj.Cells(iRow, iColumn).Value
                        Case "Mark1", "Time"
                            curMark1 = xlObj.Cells(iRow, iColumn).Value
                        Case "Mark2"
                            curMark2 = xlObj.Cells(iRow, iColumn).Value
                        Case "Mark3"
                            curMark3 = xlObj.Cells(iRow, iColumn).Value
                        Case "Mark4"
                            curMark4 = xlObj.Cells(iRow, iColumn).Value
                        Case "Mark5"
                            curMark5 = xlObj.Cells(iRow, iColumn).Value
                        Case "Section", "Run"
                            iSection = xlObj.Cells(iRow, iColumn).Value
                        Case "Section Name"
                            For iTemp = 1 To frmMain.tbsSection(frmMain.TestStatus).Tabs.Count
                                If xlObj.Cells(iRow, iColumn).Value = frmMain.tbsSection(frmMain.TestStatus).Tabs(iTemp).Caption Then
                                    iSection = iTemp
                                    Exit For
                                End If
                            Next iTemp
                        End Select
                    End If
                Next iColumn
                
                If iSection < 1 Then
                    iSection = 1
                ElseIf iSection > frmMain.tbsSection.Count Then
                    iSection = frmMain.tbsSection.Count
                End If
                
                If Val(cSta) > 0 And IsNull(curMark1) = False Then
                    If iSection <> iCurrentSection Then
                        frmMain.tbsSection_Click iSection - 1
                        iCurrentSection = iSection
                        DoEvents
                    End If
                    'remove old participant
                    cSta = Format$(Val(cSta), "000")
                    With frmMain
                        .txtParticipant.Text = cSta
                        
                        frmMain.LookUpParticipant
                        If .txtTime.Visible = True Then
                            .txtTime.Text = curMark1
                        Else
                            If IsNull(curMark1) = False And curMark1 >= frmMain.dtaTestSection.Recordset.Fields("Mark_low") And curMark1 <= frmMain.dtaTestSection.Recordset.Fields("Mark_hi") Then
                                .txtMarks(0) = curMark1
                            End If
                            If IsNull(curMark2) = False And curMark2 >= frmMain.dtaTestSection.Recordset.Fields("Mark_low") And curMark2 <= frmMain.dtaTestSection.Recordset.Fields("Mark_hi") Then
                                .txtMarks(1) = curMark2
                            End If
                            If IsNull(curMark3) = False And curMark3 >= frmMain.dtaTestSection.Recordset.Fields("Mark_low") And curMark3 <= frmMain.dtaTestSection.Recordset.Fields("Mark_hi") Then
                                .txtMarks(2) = curMark3
                            End If
                            If IsNull(curMark4) = False And curMark4 >= frmMain.dtaTestSection.Recordset.Fields("Mark_low") And curMark4 <= frmMain.dtaTestSection.Recordset.Fields("Mark_hi") Then
                                .txtMarks(3) = curMark4
                            End If
                            If IsNull(curMark5) = False And (curMark5 >= frmMain.dtaTestSection.Recordset.Fields("Mark_low") And curMark5 <= frmMain.dtaTestSection.Recordset.Fields("Mark_hi")) Or (frmMain.dtaTest.Recordset.Fields("Type_Time") = 3) Then
                                .txtMarks(4) = curMark5
                            End If
                        End If
                        .cmdOkClick
                    End With
                    DoEvents
                End If
            Loop
        End If
        MsgBox iRow - iStartRow - 1 & " " & Translate("Rows have been imported into", mcLanguage) & " " & frmMain.TestCode & " - " & frmMain.TestName & " - " & frmMain.tbsSelFin.SelectedItem.Caption
    End If
    
    SetMouseNormal
    
    StatusMessage
        
    On Local Error GoTo 0

End Sub

Public Function MakeLast(cRider As String) As String
    Dim iTemp As Integer
    
    iTemp = InStr(cRider, "&")
    If iTemp = 0 Then
        iTemp = InStr(cRider, " ")
    End If
    If iTemp > 0 Then
        MakeLast = LTrim$(Mid$(Replace(cRider, "&", " "), iTemp))
    Else
        MakeLast = ""
    End If
End Function

Sub ImportCsv()
    Dim cCsvFile As String
    Dim iKey As Integer
    
    On Local Error Resume Next
    
    ReadIniFile gcIniFile, "Import", "Csv", cCsvFile
    With frmMain.CommonDialog1
        .DefaultExt = "Csv"
        .DialogTitle = Translate("Select a text file (comma delimited)", mcLanguage)
        .FileName = cCsvFile
        .Filter = "Comma delimited (*.Csv)|*.Csv|Tab delimited (*.Txt)|*.Txt|All files (*.*)|*.*"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cCsvFile = .FileName
    End With
    
    On Local Error GoTo 0
    
    SetMouseHourGlass
    
    If cCsvFile <> "" And cCsvFile <> Chr$(27) Then
        WriteIniFile gcIniFile, "Import", "Csv", cCsvFile
        mdbMain.Execute ("DELETE * FROM Participants")
        ProcessText cCsvFile
    Else
        MsgBox Translate("No proper text file (comma delimited) selected.", mcLanguage)
    End If
    
    StatusMessage
    
    SetMouseNormal

End Sub
Sub ImportFeif(Optional FipoOnly As Integer = False, Optional NoDialog As Integer = False)
    Dim cCsvFile As String
    Dim iKey As Integer
    
    On Local Error Resume Next
    
    If FipoOnly = True Then
        If Dir$(App.Path & "\Fipo.Mdb") <> "" Then
            cCsvFile = App.Path & "\Fipo.Mdb"
        Else
            ReadIniFile gcIniFile, "Import", "FIPO", cCsvFile
        End If
    Else
        ReadIniFile gcIniFile, "Import", "FEIF", cCsvFile
    End If
    
    If NoDialog = False Then
        With frmMain.CommonDialog1
            .DefaultExt = "Mdb"
            .DialogTitle = Translate("Select a file (MS Access; FEIF Compatible)", mcLanguage)
            .FileName = cCsvFile
            .Filter = "MS Access (*.Mdb)|*.Mdb"
            .InitDir = mcDatabaseName
            .CancelError = True
            .ShowOpen
            If Err = cdlCancel Then
                Exit Sub
            End If
            cCsvFile = .FileName
        End With
    End If
    
    On Local Error GoTo 0
    
    SetMouseHourGlass
    
    If cCsvFile <> "" And cCsvFile <> Chr$(27) Then
        If FipoOnly = True Then
            'Allow FIPO.mdb and FIPO-XX.mdb for import
            If (InStr(cCsvFile, "Fipo.Mdb") = 0 And InStr(cCsvFile, "Fipo-") = 0) Then
                MsgBox Translate("No proper MS Access file (FEIF Compatible) selected.", mcLanguage)
            Else
                WriteIniFile gcIniFile, "Import", "FIPO", cCsvFile
                ProcessFipo cCsvFile
            End If
        Else
            WriteIniFile gcIniFile, "Import", "FEIF", cCsvFile
            ProcessMdb cCsvFile
        End If
    Else
        MsgBox Translate("No proper MS Access file (FEIF Compatible) selected.", mcLanguage)
    End If
    
    StatusMessage
    
    SetMouseNormal
End Sub
Sub ImportTab(Optional cTabFile As String)
    Dim iKey As Integer
    
    On Local Error Resume Next
    
    ReadIniFile gcIniFile, "Import", "Tab", cTabFile
    With frmMain.CommonDialog1
        .DefaultExt = "Txt"
        .DialogTitle = Translate("Select a text file (tab delimited)", mcLanguage)
        .FileName = cTabFile
        .Filter = "Tab delimited (*.Txt)|*.Txt|Comma delimited (*Csv)|*.Csv|All files (*.*)|*.*"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        cTabFile = .FileName
    End With
    
    On Local Error GoTo 0
    SetMouseHourGlass
    
    If cTabFile <> "" And cTabFile <> Chr$(27) Then
        WriteIniFile gcIniFile, "Import", "Tab", cTabFile
        mdbMain.Execute ("DELETE * FROM Participants")
        ProcessText cTabFile
    Else
        MsgBox Translate("No proper text file (tab delimited) selected.", mcLanguage)
    End If
    
    StatusMessage
    
    SetMouseNormal

End Sub
Sub ProcessText(ImportFile As String, Optional cDelimChar As String = "")
    Dim iTemp As Integer
    Dim iTextFileNum As Integer
    Dim iTabFile As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iColumnCount As Integer
    
    Dim cTemp As String
    Dim cColumn() As String
    
    On Local Error Resume Next
    
    StatusMessage Translate("Importing Participants", mcLanguage)
    
    iTextFileNum = FreeFile
    Open ImportFile For Input Access Read Shared As #iTextFileNum
    Line Input #iTextFileNum, cTemp
    If cDelimChar = "" Then
        If InStr(cTemp, vbTab) > 0 Then
            cDelimChar = vbTab
        ElseIf InStr(cTemp, "|") > 0 Then
            cDelimChar = "|"
        ElseIf InStr(cTemp, ";") > 0 Then
            cDelimChar = ";"
        Else
            cDelimChar = ","
        End If
    End If
    If InStr(cTemp, cDelimChar) = 0 Then
        MsgBox ImportFile & " " & Translate("doesn't have the right format (no proper field delimiter found)", mcLanguage)
        Exit Sub
    End If
    cColumn = Split(cTemp, cDelimChar)
    For iTemp = LBound(cColumn) To LBound(cColumn)
        cColumn(iTemp) = UnDotSpace(cColumn(iTemp))
    Next iTemp
    
    iRow = 0
    
    Do While Not EOF(iTextFileNum)
        If iRow Mod 10 = 0 Then
            StatusMessage Translate("Importing Participants", mcLanguage) & " [" & iRow & "]"
        End If
        Line Input #iTextFileNum, cTemp
        cTemp = Replace(cTemp, Chr$(34) & cDelimChar & Chr$(34), cDelimChar)
        If Left$(cTemp, 1) = Chr$(34) Then
            cTemp = Trim$(Mid$(cTemp & " ", 2))
        End If
        If Right$(cTemp, 1) = Chr$(34) Then
            cTemp = Left$(cTemp, Len(cTemp) - 1)
        End If
        ImportRecord cTemp, cColumn(), cDelimChar
        iRow = iRow + 1
    Loop
    Close #iTextFileNum
    
    MsgBox iRow & " " & Translate("participants processed.", mcLanguage)
    
End Sub
Sub ProcessMdb(FEIFFile As String, Optional iNoDialog As Integer = False)
    Dim iKey As Integer
    
    If iNoDialog = True Then
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Do you want to overwrite all existing data (click 'Yes') or to replace existing data and add new data only (click 'No')?", mcLanguage), vbYesNoCancel + vbQuestion + vbDefaultButton2)
    End If
    
    If iKey = vbYes Then
        
        SetVariable "ProgramVersion", ""
        
        mdbMain.Close
        DoEvents
        
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Fields", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Values", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Persons", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Horses", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Participants", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Tests", "Code", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Testsections", "Code", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "TestTimeTables", "Code", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Entries", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Results", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Marks", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Penalties", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Finance", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Staff", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Timetable", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "CountryNames", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Forms", "Title", True
        CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "_"
        CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "@"
        CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "#"

        
        OpenDatabase mcDatabaseName
        SetVariable "Programversion", ""
        mdbMain.Close
        frmMain.RestartApp
    ElseIf iKey = vbNo Then
        SetVariable "ProgramVersion", ""
        
        mdbMain.Close
        DoEvents
        
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Fields", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Values", "", True
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Persons", "PersonId", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Horses", "HorseId", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Participants", "Sta", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Tests", "Code", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Testsections", "Code", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "TestTimeTables", "Code", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Entries", "Sta", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Results", "Sta", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Marks", "Sta", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Penalties", "Sta", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Finance", "PersonId", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Staff", "PersonId", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Timetable", "Code", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "CountryNames", "CountryId", False
        CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Forms", "Title", False
        CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "_"
        CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "@"
        CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "#"
        
        OpenDatabase mcDatabaseName
        SetVariable "Programversion", ""
        mdbMain.Close
        frmMain.RestartApp
    End If
    
End Sub
Sub ProcessFipo(FEIFFile As String, Optional NoRestart As Integer = False)
    
    SetVariable "ProgramVersion", ""
        
    mdbMain.Close
    DoEvents
    
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Fields", "", True
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Values", "", True
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Tests", "Code", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Testsections", "Code", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "TestTimeTables", "Code", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Combinations", "Code", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Combinationsections", "Code", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Variables", "Item", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "CountryNames", "CountryId", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Penalties", "", False
    CopyTableBetweenDatabases FEIFFile, mcDatabaseName, "Forms", "Title", False
    CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "_"
    CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "@"
    CopyQueriesBetweenDatabases FEIFFile, mcDatabaseName, "#"
    
    OpenDatabase mcDatabaseName
    SetVariable "Programversion", ""
    SetVariable "FIPO", FileDateTime(FEIFFile)
    If NoRestart = False Then
        mdbMain.Close
        frmMain.RestartApp
    End If
End Sub
