Attribute VB_Name = "modIceIPZV"
' IceIPZV.bas
' modIceIPZV: Import competition database in IPZV's format into IceTools.
' Copyright (C) Lutz Lesener 2003, 2004
' Based on Marko Mazeland's import filter modIceNSIJP.
'
' This file is part of the FEIF software project.
' See http://www.feif.org/software or https://sourceforge.net/projects/icehorsetools/ for details.
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

Public Function ImportIPZV() As Integer
    Dim mdbIPZV As DAO.Database
    Dim rstTest As DAO.Recordset
    Dim rstIPZV As DAO.Recordset
    Dim rstPersons As DAO.Recordset
    Dim rstHorses As DAO.Recordset
    Dim rstParticipants As DAO.Recordset
    Dim rstParent As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    
    Dim rstPersonId As DAO.Recordset
    Dim rstHorseId As DAO.Recordset
    
    Dim cNotAvailable As String
    Dim cAvailable As String
    Dim cIPZV As String
    
    Dim cPrevCode As String
    
    Dim iStartvolgorde As Integer
    Dim maxInvoiceID As Long
    Dim totalcosts As Currency
    Dim totalpaid As Currency

    cNotAvailable = "|"
    cAvailable = "|"
    
    On Local Error Resume Next
    
    ReadIniFile gcIniFile, "Import", "IPZV", cIPZV
    With frmMain.CommonDialog1
        .DefaultExt = "*.Mdb"
        .DialogTitle = Translate("Select a database", mcLanguage) & " (turnier.mdb)"
        .FileName = cIPZV
        .Filter = "MS Access (*.Mdb)|*.Mdb"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowOpen
        If Err = cdlCancel Then
            Exit Function
        End If
        cIPZV = .FileName
    End With
    
    SetMouseHourGlass
    
    On Local Error GoTo ImportIPZVError
    
    ImportIPZV = True
    
    If cIPZV <> "" And cIPZV <> Chr$(27) Then
        If Dir$(cIPZV) <> "" Then
            WriteIniFile gcIniFile, "Import", "IPZV", cIPZV
            
            Set rstTest = mdbMain.OpenRecordset("SELECT Code FROM Tests")
            If rstTest.RecordCount > 0 Then
                Do While Not rstTest.EOF
                    cAvailable = cAvailable & rstTest.Fields("Code") & "|"
                    rstTest.MoveNext
                Loop
            End If
            rstTest.Close
            
            Set mdbIPZV = DBEngine.OpenDatabase(cIPZV)
            
            'import riders
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Teilnehmer")
            If rstIPZV.RecordCount > 0 Then
                rstIPZV.MoveLast
                rstIPZV.MoveFirst
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.EOF
                    
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    
                    If rstIPZV.AbsolutePosition Mod 10 = 0 Then
                        StatusMessage Translate("Importing Persons", mcLanguage) & " [" & rstIPZV.AbsolutePosition & "]"
                    End If
                    Set rstPersons = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE PersonId LIKE " & Chr$(34) & rstIPZV.Fields("Reiterbarcode") & Chr$(34))
                    With rstPersons
                        If .RecordCount = 0 Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        CopyField rstIPZV.Fields("ReiterBarcode"), .Fields("PersonId")
                        CopyField rstIPZV.Fields("Vorname"), .Fields("Name_First")
                        CopyField rstIPZV.Fields("Nachname"), .Fields("Name_Last")
                        CopyField rstIPZV.Fields("Titel"), .Fields("Title")
                        CopyField rstIPZV.Fields("Anschrift1"), .Fields("Address_1")
                        CopyField rstIPZV.Fields("Anschrift2"), .Fields("Address_2")
                        CopyField rstIPZV.Fields("PLZ"), .Fields("ZIP")
                        CopyField rstIPZV.Fields("Ort"), .Fields("City")
                        CopyField rstIPZV.Fields("Bundesland"), .Fields("Region")
                        CopyField rstIPZV.Fields("Staat"), .Fields("Country")
                        CopyField rstIPZV.Fields("Telefon1"), .Fields("Phone")
                        CopyField rstIPZV.Fields("Mobil"), .Fields("Mobile")
                        CopyField rstIPZV.Fields("Telefax"), .Fields("Fax")
                        CopyField rstIPZV.Fields("eMail"), .Fields("Email")
                        CopyField rstIPZV.Fields("Geburtsdatum"), .Fields("Birthday")
                        
                        Select Case rstIPZV.Fields("Anrede")
                            Case "F"
                                .Fields("Sex") = 2
                            Case "H"
                                .Fields("Sex") = 1
                        End Select
                        
                        .Update
                    End With
                    rstIPZV.MoveNext
                Loop
                rstPersons.Close
            End If
            
            
            'import horses
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Teilnehmer")
            If rstIPZV.RecordCount > 0 Then
                rstIPZV.MoveLast
                rstIPZV.MoveFirst
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.EOF
                    
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    
                    If rstIPZV.AbsolutePosition Mod 10 = 0 Then
                        StatusMessage Translate("Importing Horses", mcLanguage) & " [" & rstIPZV.AbsolutePosition & "]"
                    End If
                    Set rstHorses = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE HorseId LIKE " & Chr$(34) & rstIPZV.Fields("PferdeBarcode") & Chr$(34))
                    With rstHorses
                        If .RecordCount = 0 Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        CopyField rstIPZV.Fields("PferdeBarcode"), .Fields("HorseId")
                        CopyField rstIPZV.Fields("Pferdename"), .Fields("Name_Horse")
                        CopyField rstIPZV.Fields("geb"), .Fields("Birthday_Horse")
                        CopyField rstIPZV.Fields("Farbe"), .Fields("Color")
                        CopyField rstIPZV.Fields("Abzeichen"), .Fields("Marking")
                        .Fields("Sex_Horse") = IIf(Left$(rstIPZV.Fields("Geschlecht") & "", 1) = "H", 1, (IIf(Left$(rstIPZV.Fields("Geschlecht") & "", 1) = "S", 2, 3)))
                        CopyField rstIPZV.Fields("Zuchtland"), .Fields("Country_Horse")
                        CopyField rstIPZV.Fields("V"), .Fields("F")
                        CopyField rstIPZV.Fields("M"), .Fields("M")
                        CopyField rstIPZV.Fields("VV"), .Fields("FF")
                        CopyField rstIPZV.Fields("VM"), .Fields("FM")
                        CopyField rstIPZV.Fields("MV"), .Fields("MF")
                        CopyField rstIPZV.Fields("MM"), .Fields("MM")
                        CopyField rstIPZV.Fields("Z"), .Fields("Breeder")
                        CopyField rstIPZV.Fields("B"), .Fields("Owner")
                        .Update
                    End With
                    rstIPZV.MoveNext
                Loop
                rstHorses.Close
            End If
            
            'import participants
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Teilnehmer")
            If rstIPZV.RecordCount > 0 Then
                mdbMain.Execute "DELETE * FROM Participants"
                Set rstParticipants = mdbMain.OpenRecordset("SELECT * FROM Participants")
                StatusMessage Translate("Importing Participants", mcLanguage)
                rstIPZV.MoveLast
                rstIPZV.MoveFirst
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.EOF
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstParticipants
                        .AddNew
                        .Fields("STA") = Format$(rstIPZV.Fields("STA"), "000")
                        .Fields("PersonId") = rstIPZV.Fields("ReiterBarcode") & ""
                        .Fields("HorseId") = rstIPZV.Fields("PferdeBarcode") & ""
                        .Fields("Club") = rstIPZV.Fields("Verein") & ""
                        .Fields("Team") = rstIPZV.Fields("Team") & ""
                        .Fields("Stable") = rstIPZV.Fields("Stall") & ""
                        
                        Select Case LCase$(rstIPZV.Fields("Status") & "")
                            Case "genannt"
                                .Fields("Status") = 0
                            Case "anwesend"
                                .Fields("Status") = 1
                            Case "gestrichen"
                                .Fields("Status") = 2
                        End Select
                        
                        'use temporary fields:
                        .Fields("nenngeld") = rstIPZV.Fields("nenngeld")
                        .Fields("startgeld") = rstIPZV.Fields("startgeld")
                        .Fields("stallgeld") = rstIPZV.Fields("stallgeld")
                        .Fields("helferfonds") = rstIPZV.Fields("helferfonds")
                        .Fields("programmheft") = rstIPZV.Fields("programmheft")
                        .Fields("sonstiges") = rstIPZV.Fields("sonstiges")
                        .Fields("extra") = rstIPZV.Fields("extra")
                        .Fields("summe") = rstIPZV.Fields("summe")
                        .Fields("perscheck") = rstIPZV.Fields("perscheck")
                        .Fields("perüberweisung") = rstIPZV.Fields("perüberweisung")
                        .Fields("perbar") = rstIPZV.Fields("perbar")
                        .Fields("rückerstattet") = rstIPZV.Fields("rückerstattet")
                        If Not IsNull(rstIPZV.Fields("scheck-nr")) Then
                            .Fields("scheck-nr") = rstIPZV.Fields("scheck-nr") & ""
                        Else
                            .Fields("scheck-nr") = "-"
                        End If
                        If Not IsNull(rstIPZV.Fields("extras")) Then
                            .Fields("extras") = rstIPZV.Fields("extras") & ""
                        Else
                            .Fields("extras") = "-"
                        End If
                        
                        .Update
                    End With
                    rstIPZV.MoveNext
                Loop
                rstParticipants.Close
            End If
            
            
            'import variables
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Turnier")
            If rstIPZV.RecordCount > 0 Then
                SetVariable "Event_name", rstIPZV.Fields("Name")
                SetVariable "Event_date_start", rstIPZV.Fields("Anfangsdatum")
                SetVariable "Event_date_end", rstIPZV.Fields("Enddatum")
            End If
            rstIPZV.Close
            Set rstIPZV = Nothing
            
            
            'import entries
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT *, [IPO-Code] AS ipoc FROM Nennungen")
            If rstIPZV.RecordCount > 0 Then
                StatusMessage Translate("Importing Entries", mcLanguage)
                mdbMain.Execute "DELETE * FROM Entries WHERE Status=0"
                Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Entries")
                rstIPZV.MoveLast
                rstIPZV.MoveFirst
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.EOF
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstEntries
                        .AddNew
                        .Fields("STA") = rstIPZV.Fields("STA")
                        Select Case LCase$(rstIPZV.Fields("Status"))
                            Case "ve"
                                .Fields("Status") = 0
                            Case "af"
                                .Fields("Status") = 1
                            Case "bf"
                                .Fields("Status") = 2
                            Case Else
                                .Fields("Status") = 3
                        End Select
                        .Fields("Group") = 0
                        .Fields("Code") = Left$(rstIPZV.Fields("ipoc"), 8)
                        .Fields("Position") = rstIPZV.Fields("Position")
                        .Fields("Late_Entry") = rstIPZV.Fields("Nachnennung")
                        .Fields("Timestamp") = rstIPZV.Fields("Nennzeit")
                        .Fields("RR") = rstIPZV.Fields("rH")
                        .Fields("Qualification") = rstIPZV.Fields("QPunkte")
                        .Update
                    End With
                    rstIPZV.MoveNext
                Loop
                rstEntries.Close
            End If
            
            
            'import financial data
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Teilnehmer")
            If rstIPZV.RecordCount > 0 Then
                'delete financial data to avoid duplicates:
                mdbMain.Execute "DELETE * FROM finance_costs;"
                mdbMain.Execute "DELETE * FROM finance_payments;"
                mdbMain.Execute "DELETE * FROM finance_invoices;"
                mdbMain.Execute "DELETE * FROM finance_transactions;"
                
                StatusMessage Translate("Importing financial data", mcLanguage)
                rstIPZV.MoveLast
                rstIPZV.MoveFirst
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.EOF
                    Set rstParticipants = mdbMain.OpenRecordset("SELECT * FROM finance_costs")
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    
                    'reset counters for new participant:
                    totalcosts = 0
                    totalpaid = 0
                    
                    With rstParticipants
                        If rstIPZV.Fields("Nenngeld") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Nenngeld")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Nenngeld")
                            .Fields("description") = Translate("Nenngeld", mcLanguage)
                            .Fields("debtor_type") = 2       'Participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Startgeld") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Startgeld")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Startgeld")
                            .Fields("description") = Translate("Startgeld", mcLanguage)
                            .Fields("debtor_type") = 2       'Participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Stallgeld") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Stallgeld")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Stallgeld")
                            .Fields("description") = Translate("Stallgeld", mcLanguage)
                            .Fields("debtor_type") = 2       'Participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Helferfonds") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Helferfonds")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Helferfonds")
                            .Fields("description") = Translate("Helferfonds", mcLanguage)
                            .Fields("debtor_type") = 2       'Participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Programmheft") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Programmheft")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Programmheft")
                            .Fields("description") = Translate("Programmheft", mcLanguage)
                            .Fields("debtor_type") = 2       'Participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Sonstiges") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Sonstiges")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Sonstiges")
                            .Fields("description") = Translate("Sonstiges", mcLanguage)
                            .Fields("debtor_type") = 2       'Participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Extra") <> 0 Then
                            totalcosts = totalcosts + rstIPZV.Fields("Extra")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Extra")
                            .Fields("description") = rstIPZV.Fields("Extras")
                            .Fields("debtor_type") = 2       'participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                        If rstIPZV.Fields("Rückerstattet") <> 0 Then
                            totalcosts = totalcosts - rstIPZV.Fields("Rückerstattet")
                            .AddNew
                            .Fields("amount") = rstIPZV.Fields("Rückerstattet") * (-1)
                            .Fields("description") = Translate("Rückerstattung", mcLanguage)
                            .Fields("debtor_type") = 2       'participant
                            .Fields("debtor_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("comment") = "Imported via IceIPZV"
                            .Fields("created") = Now()
                            .Update
                        End If
                        
                    End With
                    rstParticipants.Close
                    
                    
                        
                    
                    'Payments
                    Set rstParticipants = mdbMain.OpenRecordset("SELECT * FROM finance_payments")
                    With rstParticipants
                    
                        If rstIPZV.Fields("perScheck") <> 0 Then
                            totalpaid = totalpaid + rstIPZV.Fields("perScheck")
                            .AddNew
                            .Fields("Amount") = rstIPZV.Fields("perScheck")
                            .Fields("created") = Now()
                            .Fields("paidby_type") = 1      'person
                            .Fields("paidby_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("type") = 1             'cheque
                            .Fields("type_description") = rstIPZV.Fields("Scheck-Nr")
                            .Update
                        End If
                        
                        If rstIPZV.Fields("perBar") <> 0 Then
                            totalpaid = totalpaid + rstIPZV.Fields("perBar")
                            .AddNew
                            .Fields("Amount") = rstIPZV.Fields("perBar")
                            .Fields("created") = Now()
                            .Fields("paidby_type") = 1      'person
                            .Fields("paidby_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("type") = 3             'cash
                            .Update
                        End If
                        
                        If rstIPZV.Fields("perÜberweisung") <> 0 Then
                            totalpaid = totalpaid + rstIPZV.Fields("perÜberweisung")
                            .AddNew
                            .Fields("Amount") = rstIPZV.Fields("perÜberweisung")
                            .Fields("created") = Now()
                            .Fields("paidby_type") = 1      'person
                            .Fields("paidby_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("sta") = rstIPZV.Fields("sta")
                            .Fields("type") = 2             'bank transaction
                            .Update
                        End If
                    End With
                    rstParticipants.Close
                    
                    'Create a new invoice for the total costs of this participant:
                    If totalcosts > 0 Then
                        Set rstParticipants = mdbMain.OpenRecordset("SELECT * FROM finance_invoices")
                        With rstParticipants
                            .AddNew
                            .Fields("addressee_id") = rstIPZV.Fields("ReiterBarcode")
                            .Fields("created") = Now()
                            'the field "sender" is currently left blank:
                            '.Fields("sender") = id_of_person_issueing_the_invoice
                            .Update
                        End With
                        rstParticipants.Close
                        
                        'Retrieve the id of the invoice just created:
                        Set rstTest = mdbMain.OpenRecordset("SELECT id FROM finance_invoices ORDER BY id DESC")
                            If Not IsNull(rstTest.Fields("id")) Then
                                maxInvoiceID = CLng(rstTest.Fields("id"))
                            Else
                                maxInvoiceID = 0
                            End If
                        rstTest.Close
                        
                        'Update all matching records in finance_costs with this invoice id:
                        mdbMain.Execute "UPDATE finance_costs SET invoice_id=" & maxInvoiceID & " WHERE sta='" & rstIPZV.Fields("sta") & "';"
                    End If
                    
                    'Create a new transaction record if the total costs match the total payments:
                    '(normally a transaction should *always* be created if any payment is made, but this would require
                    ' much additional code to split payments between individual cost records... since this import is meant
                    ' to be run only once and is only a temporary solution anyway, I guess we can live with this simplification.)
                    If totalcosts = totalpaid And totalcosts > 0 Then
                        Set rstParticipants = mdbMain.OpenRecordset("SELECT * FROM finance_transactions")
                        With rstParticipants
                            .AddNew
                            .Fields("created") = Now()
                            .Update
                        End With
                        rstParticipants.Close
                        
                        'Retrieve the id of the transaction just created:
                        Set rstTest = mdbMain.OpenRecordset("SELECT id FROM finance_transactions ORDER BY id DESC")
                            If Not IsNull(rstTest.Fields("id")) Then
                                maxInvoiceID = CLng(rstTest.Fields("id"))
                            Else
                                maxInvoiceID = 0
                            End If
                        rstTest.Close
                        
                        'Update all matching records in finance_costs and finance_payments with this transaction id:
                        mdbMain.Execute "UPDATE finance_costs SET transaction_id=" & maxInvoiceID & " WHERE sta='" & rstIPZV.Fields("sta") & "';"
                        mdbMain.Execute "UPDATE finance_payments SET transaction_id=" & maxInvoiceID & " WHERE sta='" & rstIPZV.Fields("sta") & "';"
                    End If
                    
                    
                    rstIPZV.MoveNext
                Loop
            End If
            
            
            'import results
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Ergebnisse")
            If rstIPZV.RecordCount > 0 Then
                StatusMessage Translate("Importing results", mcLanguage)
                mdbMain.Execute "DELETE * FROM Results"
                Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Results")
                rstIPZV.MoveLast
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.BOF
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstEntries
                        .AddNew
                        .Fields("STA") = rstIPZV.Fields("Startnummer")
                        
                        Select Case LCase$(rstIPZV.Fields("Status"))
                            Case "ve"
                                .Fields("Status") = 0
                            Case "af"
                                .Fields("Status") = 1
                            Case "bf"
                                .Fields("Status") = 2
                        End Select
                        
                        .Fields("Code") = rstIPZV.Fields("IPO")
                        .Fields("Disq") = rstIPZV.Fields("Disq")
                        .Fields("FR") = rstIPZV.Fields("Endresultat")
                        .Fields("Timestamp") = rstIPZV.Fields("Datum")
                        .Fields("Score") = rstIPZV.Fields("Punkte")
                        .Fields("Time") = rstIPZV.Fields("Zeit")
        
                        .Update
                    End With
                    rstIPZV.MovePrevious
                Loop
                rstEntries.Close
            End If
            
            
            'import single marks
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM Einzelnoten")
            If rstIPZV.RecordCount > 0 Then
                StatusMessage Translate("Importing marks", mcLanguage)
                mdbMain.Execute "DELETE * FROM Marks"
                Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Marks")
                rstIPZV.MoveLast
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.BOF
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstEntries
                        .AddNew
                        .Fields("STA") = rstIPZV.Fields("Startnummer")
                        
                        Select Case LCase$(rstIPZV.Fields("Status"))
                            Case "ve"
                                .Fields("Status") = 0
                            Case "af"
                                .Fields("Status") = 1
                            Case "bf"
                                .Fields("Status") = 2
                        End Select
                        
                        .Fields("Code") = rstIPZV.Fields("IPO")
                        .Fields("Mark1") = rstIPZV.Fields("W1")
                        .Fields("Mark2") = rstIPZV.Fields("W2")
                        .Fields("Mark3") = rstIPZV.Fields("W3")
                        .Fields("Mark4") = rstIPZV.Fields("W4")
                        .Fields("Mark5") = rstIPZV.Fields("W5")
                        .Fields("Score") = rstIPZV.Fields("Punkte")
                        .Fields("Section") = rstIPZV.Fields("Aufgabenteil")
                        .Fields("Timestamp") = rstIPZV.Fields("Datum")
                        .Update
                    End With
                    rstIPZV.MovePrevious
                Loop
                rstEntries.Close
            End If
            
            
            'import penalties
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT * FROM OM")
            If rstIPZV.RecordCount > 0 Then
                StatusMessage Translate("Importing penalties", mcLanguage)
                mdbMain.Execute "DELETE * FROM Penalties"
                Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Penalties")
                rstIPZV.MoveLast
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.BOF
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    With rstEntries
                        .AddNew
                        .Fields("STA") = rstIPZV.Fields("STA")
                        .Fields("Cause") = rstIPZV.Fields("Anlass")
                        .Fields("Comments") = rstIPZV.Fields("Kommentar") & " / " & rstIPZV.Fields("Verantwortlicher")
                        .Fields("Measure") = rstIPZV.Fields("Massnahme")
                        .Update
                    End With
                    rstIPZV.MovePrevious
                Loop
                rstEntries.Close
            End If
            
            
            
            'import tests
            Set rstIPZV = mdbIPZV.OpenRecordset("SELECT *, [IPO-Code] AS ipoc, [Darf_RH] as darfrh, [Q-Code] as qcode, Gruppengrösse as gg, Prüfungsnummer as pn FROM Ausschreibung")
            If rstIPZV.RecordCount > 0 Then
                StatusMessage Translate("Importing tests", mcLanguage)
                rstIPZV.MoveLast
                rstIPZV.MoveFirst
                ShowProgressbar frmMain, 2, rstIPZV.RecordCount
                Do While Not rstIPZV.EOF
                    
                    IncreaseProgressbarValue frmMain.ProgressBar1
                    
                    Set rstPersons = mdbMain.OpenRecordset("SELECT * FROM Tests WHERE Code LIKE " & Chr$(34) & rstIPZV.Fields("ipoc") & Chr$(34))
                    With rstPersons
                        If .RecordCount = 0 Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        'CopyField rstIPZV.Fields("ipoc"), .Fields("Code")
                        CopyField rstIPZV.Fields("Nenngeld"), .Fields("Fee1")
                        CopyField rstIPZV.Fields("gg"), .Fields("GroupSize")
                        CopyField rstIPZV.Fields("Gruppenzeit"), .Fields("GroupTime")
                        CopyField rstIPZV.Fields("pn"), .Fields("Nr")
                        
                        CopyField rstIPZV.Fields("darfrh"), .Fields("RR")
                        CopyField rstIPZV.Fields("Sortierzeichen"), .Fields("SortChar")
                        CopyField rstIPZV.Fields("Sortierstelle"), .Fields("SortDigit")
                        CopyField rstIPZV.Fields("Sponsor"), .Fields("Sponsor")
                        CopyField rstIPZV.Fields("Titel"), .Fields("Test")
                        
                        'If rstIPZV.Fields("qCode") & "" = "" Then
                        '    .Fields("Qualification") = "none"
                        'Else
                        '    CopyField rstIPZV.Fields("qcode"), .Fields("Qualification")
                        'End If
                            
                        .Update
                    End With
                    rstIPZV.MoveNext
                Loop
                rstPersons.Close
            End If
                        
            
            
            rstIPZV.Close
            Set rstTest = Nothing
            Set rstIPZV = Nothing
            Set rstPersons = Nothing
            Set rstHorses = Nothing
            Set rstEntries = Nothing
            Set rstParticipants = Nothing
            Set rstParent = Nothing
            mdbIPZV.Close
            Set mdbIPZV = Nothing
        End If
    Else
        MsgBox Translate("No proper database selected.", mcLanguage)
    End If
    
ImportIPZVError:
    If Err > 0 Then
        ImportIPZV = False
        MsgBox cIPZV & ": " & Err.Description
    End If
    
    ShowProgressbar frmMain, 2, 0
    
    StatusMessage
    
    SetMouseNormal
End Function


