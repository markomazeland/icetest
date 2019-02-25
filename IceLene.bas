Attribute VB_Name = "modIceLene"
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

Public Function ImportLene() As Integer
    Dim mdbLene As DAO.Database
    Dim rstLene As DAO.Recordset
    Dim rstMain As DAO.Recordset
    
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim iFileNum As Integer
    Dim iValid As Integer
    Dim iPosition As Integer
    Dim cIncludeFields As String
    Dim cTests As String
    Dim cTabel1 As String
    
    Dim cLene As String
    Dim cRiderId As String
    Dim cHorseID As String
    Dim iKey As Integer
    
    On Local Error Resume Next
    ReadIniFile gcIniFile, "Import", "Lene", cLene
    With frmMain.CommonDialog1
        .DefaultExt = "*.Mdb"
        .DialogTitle = Translate("Select a database", mcLanguage)
        .FileName = cLene
        .Filter = "MS Access (*.Mdb)|*.Mdb"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Function
        End If
        cLene = .FileName
    End With
    
    SetMouseHourGlass
    
    On Local Error GoTo ImportLeneError
    
    ImportLene = True
    
    If cLene <> "" And cLene <> Chr$(27) Then
        If Dir$(cLene) <> "" Then
            WriteIniFile gcIniFile, "Import", "Lene", cLene
            cTabel1 = "Tabel1"
            Set mdbLene = DBEngine(0).OpenDatabase(cLene)
            If TableExist(mdbLene, cTabel1) = False Then
                ImportLene = False
                MsgBox "Tabel1 " & Translate("not found!", mcLanguage)
                For Each tdf In mdbLene.TableDefs
                    If tdf.Attributes = 0 Then
                        iKey = MsgBox(Translate("Use", mcLanguage) & ": " & tdf.Name, vbYesNoCancel + vbQuestion)
                    End If
                    If iKey = vbYes Or iKey = vbCancel Then Exit For
                Next
                If iKey <> vbYes Then
                    ImportLene = False
                    Set mdbLene = Nothing
                    Exit Function
                End If
            Else
                iKey = vbYes
            End If
            If iKey = vbYes Then
                Set rstLene = mdbMain.OpenRecordset("SELECT* FROM Tests")
                If rstLene.RecordCount > 0 Then
                    Do While Not rstLene.EOF
                        cTests = cTests & "|" & rstLene.Fields("Code")
                        rstLene.MoveNext
                    Loop
                End If
    
                Set rstLene = mdbLene.OpenRecordset("SELECT * FROM Tabel1")
                If rstLene.RecordCount > 0 Then
                   rstLene.MoveLast
                   rstLene.MoveFirst
                   ShowProgressbar frmMain, 2, rstLene.RecordCount
                    Do While Not rstLene.EOF
                   
                        With rstLene
                            IncreaseProgressbarValue frmMain.ProgressBar1
                            If Val(.Fields("Nr")) > 0 Then
                                Set rstMain = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE Name_First LIKE '" & .Fields("Fornavn") & "" & "' AND Name_Last LIKE '" & .Fields("Efternavn") & "" & "'")
                                If rstMain.RecordCount = 0 Then
                                    rstMain.AddNew
                                    rstMain.Fields("Name_First") = Left(.Fields("Fornavn") & "", rstMain.Fields("Name_First").Size)
                                    rstMain.Fields("Name_Last") = Left(.Fields("Efternavn") & "", rstMain.Fields("Name_Last").Size)
                                    rstMain.Fields("PersonId") = CreatePersonId
                                    cRiderId = rstMain.Fields("PersonId")
                                    rstMain.Update
                                Else
                                    cRiderId = rstMain.Fields("PersonId")
                                End If
                                
                                Set rstMain = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE Name_Horse LIKE '" & .Fields("Hest") & "" & "'")
                                If rstMain.RecordCount = 0 Then
                                    rstMain.AddNew
                                    rstMain.Fields("Name_Horse") = Left(.Fields("Hest") & "", rstMain.Fields("Name_Horse").Size)
                                    rstMain.Fields("F") = Left(.Fields("Hingst") & "", rstMain.Fields("F").Size)
                                    rstMain.Fields("HorseId") = CreateHorseId
                                    cHorseID = rstMain.Fields("HorseId")
                                    rstMain.Update
                                Else
                                    cHorseID = rstMain.Fields("HorseId")
                                End If
                                
                                mdbMain.Execute ("DELETE * FROM Participants WHERE STA LIKE '" & Format$(.Fields("Nr"), "000") & "'")
                                Set rstMain = mdbMain.OpenRecordset("SELECT * FROM Participants")
                                rstMain.AddNew
                                rstMain.Fields("STA") = Format$(Val(.Fields("Nr")), "000")
                                rstMain.Fields("HorseId") = cHorseID
                                rstMain.Fields("PersonId") = cRiderId
                                rstMain.Update
                                Set rstMain = mdbMain.OpenRecordset("SELECT * FROM Entries")
                                For Each fld In rstLene.Fields
                                   If InStr(cTests, "|" & fld.Name) > 0 Then
                                       If (fld.Value & "") <> "" Then
                                            mdbMain.Execute "DELETE * FROM Entries WHERE Code LIKE '" & UnDotSpace(fld.Name) & "' AND Sta='" & Format$(Val(.Fields("Nr")), "000") & "' AND Status=0"
                                            rstMain.AddNew
                                            rstMain.Fields("Sta") = Format$(Val(.Fields("Nr")), "000")
                                            rstMain.Fields("Code") = UnDotSpace(fld.Name)
                                            rstMain.Fields("Status") = 0
                                            If InStr(fld.Value, Left$(Translate("Right", mcLanguage), 1)) > 0 Then
                                                rstMain.Fields("RR") = True
                                            Else
                                                rstMain.Fields("RR") = False
                                            End If
                                            If Val(fld.Value) > 0 Then
                                                rstMain.Fields("Position") = Val(fld.Value)
                                            Else
                                                rstMain.Fields("Position") = 0
                                            End If
                                            rstMain.Update
                                       End If
                                   End If
                                Next
                            End If
                            .MoveNext
                        End With
                    Loop

                End If
            End If
            rstLene.Close
        End If
    Else
        MsgBox Translate("No proper database selected.", mcLanguage)
    End If
    
ImportLeneError:
    If Err > 0 Then
        ImportLene = False
        MsgBox cLene & ": " & Err.Description
    End If
    
    Set rstLene = Nothing
    Set mdbLene = Nothing
    
    ShowProgressbar frmMain, 2, 0
    
    StatusMessage
    
    SetMouseNormal
End Function


