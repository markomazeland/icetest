Attribute VB_Name = "modIceNSIJP"
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

Public Function ImportNSIJP() As Integer
    Dim mdbNSIJP As DAO.Database
    Dim rstTest As DAO.Recordset
    Dim rstNSIJP As DAO.Recordset
    Dim rstPersons As DAO.Recordset
    Dim rstHorses As DAO.Recordset
    Dim rstParticipants As DAO.Recordset
    Dim rstParent As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    
    Dim rstPersonId As DAO.Recordset
    Dim rstHorseId As DAO.Recordset
    
    Dim cNotAvailable As String
    Dim cAvailable As String
    Dim cNSIJP As String
    Dim cNSIJP2 As String
    Dim cPrevCode As String
    
    Dim iStartvolgorde As Integer
    
    Dim vF As Variant
    Dim vM As Variant
    Dim vFF As Variant
    Dim vFM As Variant
    Dim vMF As Variant
    Dim vMM As Variant
    
    cNotAvailable = "|"
    cAvailable = "|"
    
    On Local Error Resume Next
    
    ReadIniFile gcIniFile, "Import", "NSIJP", cNSIJP
    With frmMain.CommonDialog1
        .DefaultExt = "*.Mdb"
        .DialogTitle = Translate("Select a database other than", mcLanguage) & " Paarden&Ruiters"
        .FileName = cNSIJP
        .Filter = "MS Access (*.Mdb)|*.Mdb"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Function
        End If
        cNSIJP = .FileName
    End With
    
    SetMouseHourGlass
    
    On Local Error GoTo ImportNSIJPError
    
    ImportNSIJP = True
    
    If cNSIJP <> "" And cNSIJP <> Chr$(27) And InStr(cNSIJP, "Paarden&Ruiters") = 0 Then
        cNSIJP2 = Left$(cNSIJP, InStrRev(cNSIJP, "\")) & "Paarden&Ruiters.Mdb"
        If Dir$(cNSIJP) <> "" And Dir$(cNSIJP2) <> "" Then
            WriteIniFile gcIniFile, "Import", "NSIJP", cNSIJP
            
            Set rstTest = mdbMain.OpenRecordset("SELECT Code FROM Tests")
            If rstTest.RecordCount > 0 Then
                Do While Not rstTest.EOF
                    cAvailable = cAvailable & rstTest.Fields("Code") & "|"
                    rstTest.MoveNext
                Loop
            End If
            rstTest.Close
         
            Set mdbNSIJP = DBEngine.OpenDatabase(cNSIJP2)
            Set rstNSIJP = mdbNSIJP.OpenRecordset("SELECT * FROM Deelnemers")
            If rstNSIJP.RecordCount > 0 Then
                rstNSIJP.MoveLast
                rstNSIJP.MoveFirst
                ShowProgressbar frmMain, 2, rstNSIJP.RecordCount
                Do While Not rstNSIJP.EOF
                    
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    
                    If rstNSIJP.AbsolutePosition Mod 10 = 0 Then
                        StatusMessage Translate("Importing Persons", mcLanguage) & " [" & rstNSIJP.AbsolutePosition & "]"
                    End If
                    Set rstPersons = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE PersonId LIKE " & Chr$(34) & rstNSIJP.Fields("Id") & Chr$(34))
                    With rstPersons
                        If .RecordCount = 0 Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        CopyField rstNSIJP.Fields("Id"), .Fields("PersonId")
                        If InStr(rstNSIJP.Fields("Naam"), " ") > 0 Then
                            .Fields("Name_First") = Trim$(Left$(rstNSIJP.Fields("Naam"), InStr(rstNSIJP.Fields("Naam"), " ")))
                            .Fields("Name_Last") = Trim$(Mid$(rstNSIJP.Fields("Naam"), InStr(rstNSIJP.Fields("Naam"), " ")))
                        Else
                            CopyField rstNSIJP.Fields("Naam"), .Fields("Name_Last")
                        End If
                        CopyField rstNSIJP.Fields("Adres"), .Fields("Address_1")
                        CopyField rstNSIJP.Fields("Postcode"), .Fields("ZIP")
                        CopyField rstNSIJP.Fields("Woonplaats"), .Fields("City")
                        CopyField rstNSIJP.Fields("Telefoonnummer"), .Fields("Phone")
                        CopyField rstNSIJP.Fields("E-mail"), .Fields("Email")
                        .Update
                    End With
                    rstNSIJP.MoveNext
                Loop
                rstPersons.Close
            End If
            
            Set rstNSIJP = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden")
            If rstNSIJP.RecordCount > 0 Then
                rstNSIJP.MoveLast
                rstNSIJP.MoveFirst
                ShowProgressbar frmMain, 2, rstNSIJP.RecordCount
                Do While Not rstNSIJP.EOF
                    
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    
                    If rstNSIJP.AbsolutePosition Mod 10 = 0 Then
                        StatusMessage Translate("Importing Horses", mcLanguage) & " [" & rstNSIJP.AbsolutePosition & "]"
                    End If
                    Set rstHorses = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE HorseId LIKE " & Chr$(34) & rstNSIJP.Fields("Nr") & Chr$(34))
                    With rstHorses
                        If .RecordCount = 0 Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        CopyField rstNSIJP.Fields("Nr"), .Fields("HorseId")
                        CopyField rstNSIJP.Fields("Paard"), .Fields("Name_Horse")
                        CopyField rstNSIJP.Fields("Stamb nr"), .Fields("FEIFID")
                        CopyField rstNSIJP.Fields("Geboorteland"), .Fields("Country_Horse")
                        CopyField rstNSIJP.Fields("Geboortejaar"), .Fields("Birthday_Horse")
                        CopyField rstNSIJP.Fields("Kleur"), .Fields("Color")
                        .Fields("Sex_Horse") = IIf(Left$(rstNSIJP.Fields("Geslacht") & "", 1) = "H", 1, (IIf(Left$(rstNSIJP.Fields("Geslacht") & "", 1) = "M", 2, 3)))
                        CopyField rstNSIJP.Fields("Eigenaar"), .Fields("Owner")
                        CopyField rstNSIJP.Fields("Naam fokker"), .Fields("Breeder")
                        vF = rstNSIJP.Fields("Vader")
                        vM = rstNSIJP.Fields("Moeder")
                        If vF > 0 Then
                            Set rstParent = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden WHERE Nr=" & vF)
                            If rstParent.RecordCount > 0 Then
                                CopyField rstParent.Fields("Paard"), .Fields("F")
                                vFF = rstParent.Fields("Vader")
                                vFM = rstParent.Fields("Moeder")
                                If vFF > 0 Then
                                    Set rstParent = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden WHERE Nr=" & vFF)
                                    If rstParent.RecordCount > 0 Then
                                        CopyField rstParent.Fields("Paard"), .Fields("FF")
                                    Else
                                        .Fields("FF") = "-"
                                    End If
                                Else
                                    .Fields("FF") = "-"
                                End If
                                
                                If vFM > 0 Then
                                    Set rstParent = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden WHERE Nr=" & vFM)
                                    If rstParent.RecordCount > 0 Then
                                        CopyField rstParent.Fields("Paard"), .Fields("FM")
                                    Else
                                        .Fields("FM") = "-"
                                    End If
                                Else
                                    .Fields("FM") = "-"
                                End If
                            End If
                        Else
                            .Fields("F") = "-"
                        End If
                        
                        If vM > 0 Then
                            Set rstParent = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden WHERE Nr=" & vM)
                            If rstParent.RecordCount > 0 Then
                                CopyField rstParent.Fields("Paard"), .Fields("M")
                                vMF = rstParent.Fields("Vader")
                                vMM = rstParent.Fields("Moeder")
                                If vMF > 0 Then
                                    Set rstParent = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden WHERE Nr=" & vMF)
                                    If rstParent.RecordCount > 0 Then
                                        CopyField rstParent.Fields("Paard"), .Fields("MF")
                                    Else
                                        .Fields("MF") = "-"
                                    End If
                                Else
                                    .Fields("MF") = "-"
                                End If
                                If vMM > 0 Then
                                    Set rstParent = mdbNSIJP.OpenRecordset("SELECT * FROM Paarden WHERE Nr=" & vMM)
                                    If rstParent.RecordCount > 0 Then
                                        CopyField rstParent.Fields("Paard"), .Fields("MM")
                                    Else
                                        .Fields("MM") = "-"
                                    End If
                                Else
                                    .Fields("MM") = "-"
                                End If
                            End If
                        Else
                            .Fields("M") = "-"
                        End If
                        .Update
                    End With
                    rstNSIJP.MoveNext
                Loop
                rstHorses.Close
            End If
            
            Set mdbNSIJP = DBEngine.OpenDatabase(cNSIJP)
            Set rstNSIJP = mdbNSIJP.OpenRecordset("SELECT * FROM Combinaties")
            If rstNSIJP.RecordCount > 0 Then
                mdbMain.Execute "DELETE * FROM Participants"
                Set rstParticipants = mdbMain.OpenRecordset("SELECT * FROM Participants")
                StatusMessage Translate("Importing Participants", mcLanguage)
                rstNSIJP.MoveLast
                rstNSIJP.MoveFirst
                ShowProgressbar frmMain, 2, rstNSIJP.RecordCount
                Do While Not rstNSIJP.EOF
                    
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstParticipants
                        .AddNew
                        .Fields("STA") = Format$(rstNSIJP.Fields("Startnummer"), "000")
                        .Fields("PersonId") = rstNSIJP.Fields("Deelnemer") & ""
                        .Fields("HorseId") = IIf((rstNSIJP.Fields("Paard") & "") <> "", rstNSIJP.Fields("Paard") & "", "1")
                        .Update
                    End With
                    rstNSIJP.MoveNext
                Loop
                rstParticipants.Close
            End If
            
            Set rstNSIJP = mdbNSIJP.OpenRecordset("SELECT * FROM [Wedstrijd Gegevens]")
            If rstNSIJP.RecordCount > 0 Then
                SetVariable "Event_name", rstNSIJP.Fields("Naam wedstrijd")
                SetVariable "Event_date", rstNSIJP.Fields("Datum wedstrijd")
            End If
            
            Set rstNSIJP = mdbNSIJP.OpenRecordset("SELECT * FROM Starts ORDER BY [Id Onderdelen],Startvolgorde,Hand DESC,Startnummer;")
            If rstNSIJP.RecordCount > 0 Then
                StatusMessage Translate("Importing Entries", mcLanguage)
                mdbMain.Execute "DELETE * FROM Entries WHERE Status=0"
                Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Entries")
                rstNSIJP.MoveLast
                ShowProgressbar frmMain, 2, rstNSIJP.RecordCount
                Do While Not rstNSIJP.BOF
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstEntries
                        Set rstParent = mdbNSIJP.OpenRecordset("SELECT [Code Proef] FROM Onderdelen WHERE [Id Onderdelen]=" & rstNSIJP.Fields("Id Onderdelen"))
                        If rstParent.RecordCount > 0 Then
                            If InStr(cAvailable, "|" & UnDotSpace(rstParent.Fields("Code Proef")) & "|") = 0 And InStr(cNotAvailable, "|" & UnDotSpace(rstParent.Fields("Code Proef")) & "|") = 0 Then
                                MsgBox Translate("Unknown test", mcLanguage) & ": " & UnDotSpace(rstParent.Fields("Code Proef"))
                                cNotAvailable = cNotAvailable & UnDotSpace(rstParent.Fields("Code Proef")) & "|"
                            End If
                            
                            .AddNew
                            .Fields("STA") = Format$(rstNSIJP.Fields("Startnummer"), "000")
                            .Fields("Code") = UnDotSpace(rstParent.Fields("Code Proef"))
                            If .Fields("Code") <> cPrevCode Then
                                iStartvolgorde = 1
                            End If
                            If rstNSIJP.Fields("Hand") = 4 Then
                                .Fields("RR") = True
                            Else
                                .Fields("RR") = False
                            End If
                            .Fields("Group") = 0
                            .Fields("Status") = 0
                            If rstNSIJP.Fields("Startvolgorde") > 0 Then
                                .Fields("Position") = rstNSIJP.Fields("Startvolgorde")
                            Else
                                .Fields("Position") = iStartvolgorde
                            End If
                            iStartvolgorde = .Fields("Position") + 1
                            .Fields("Color") = rstNSIJP.Fields("Kleur")
                            cPrevCode = .Fields("Code")
                            .Update
                        End If
                        rstParent.Close
                    End With
                    rstNSIJP.MovePrevious
                Loop
                rstEntries.Close
            End If
            
            rstNSIJP.Close
            Set rstTest = Nothing
            Set rstNSIJP = Nothing
            Set rstPersons = Nothing
            Set rstHorses = Nothing
            Set rstEntries = Nothing
            Set rstParticipants = Nothing
            Set rstParent = Nothing
            mdbNSIJP.Close
            Set mdbNSIJP = Nothing
        End If
    Else
        MsgBox Translate("No proper database selected.", mcLanguage)
    End If
    
ImportNSIJPError:
    If Err > 0 Then
        ImportNSIJP = False
        MsgBox cNSIJP & ": " & Err.Description
    End If
    
    ShowProgressbar frmMain, 2, 0
    
    StatusMessage
    
    SetMouseNormal
End Function

