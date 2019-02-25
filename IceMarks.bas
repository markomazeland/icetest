Attribute VB_Name = "modIceMarks"
Option Explicit
Option Compare Text
Public Sub ExtractFromTempTable()
    Dim iJudge As Integer
    Dim iSection As Integer
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iValue As Integer
    Dim iRow As Integer
    
    Dim rstSectionmarks As DAO.Recordset
    Dim rstSections As DAO.Recordset
    Dim rstTempMarks As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    Dim rstEntries As DAO.Recordset
    Dim rstResults As DAO.Recordset
    Dim rstMark As DAO.Recordset
    
    Dim cSta As String
    Dim cTemp As String
    Dim cOldSta As String
    
    Dim curMark As Currency
    Dim curHi As Currency
    Dim curLo As Currency
    Dim curScore As Currency
    Dim curTemp As Currency
    
    SetMouseHourGlass
        
    mdbMain.Execute ("DELETE * FROM SectionMarks WHERE Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
     
    LogLine mcTempTableName
    
    '***update sectionmarks table
    For iJudge = 1 To frmMain.TestJudges
        Set rstSectionmarks = mdbMain.OpenRecordset("SELECT * FROM SectionMarks")
        Set rstTempMarks = mdbMain.OpenRecordset("SELECT * FROM [" & mcTempTableName & "] WHERE Judge=" & iJudge & " ORDER BY STA")
        If rstTempMarks.RecordCount > 0 Then
            Do
                With rstSectionmarks
                    cTemp = rstTempMarks.Fields("Total") & ""
                    If Trim$(cTemp) <> "" Then
                        curTemp = MakeStringValue(cTemp)
                    Else
                        curTemp = -1
                    End If
                    LogLine CStr(curTemp)
                    .AddNew
                    .Fields("Code") = frmMain.TestCode
                    .Fields("Status") = frmMain.TestStatus
                    .Fields("Judge") = iJudge
                    .Fields("Section") = 0
                    .Fields("Mark") = curTemp
                    .Fields("STA") = Left$(rstTempMarks.Fields("STA"), .Fields("STA").Size)
                    .Update
                End With
                
                For iSection = 1 To miSectionCount
                    With rstSectionmarks
                        cTemp = rstTempMarks.Fields("SEC" & Format$(iSection)) & ""
                        If Trim$(cTemp) <> "" Then
                            curTemp = MakeStringValue(cTemp)
                        Else
                            curTemp = -1
                        End If
                        .AddNew
                        .Fields("Code") = frmMain.TestCode
                        .Fields("Status") = frmMain.TestStatus
                        .Fields("Judge") = iJudge
                        .Fields("Section") = iSection
                        .Fields("Mark") = curTemp
                        .Fields("STA") = Left$(rstTempMarks.Fields("STA"), .Fields("STA").Size)
                        .Update
                    End With
                Next iSection
                
                rstTempMarks.MoveNext
            Loop While Not rstTempMarks.EOF
        End If
        rstSectionmarks.Close
        rstTempMarks.Close
    Next iJudge
    
    '***update marks table
    Set rstEntries = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
    If rstEntries.RecordCount > 0 Then
        Do While Not rstEntries.EOF
            Set rstMark = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Sta='" & rstEntries.Fields("Sta") & "' AND Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
            With rstMark
                iValue = 0
                If .RecordCount > 0 Then
                    .Edit
                 Else
                    .AddNew
                    .Fields("Sta") = rstEntries.Fields("Sta")
                    .Fields("Status") = frmMain.TestStatus
                    .Fields("Code") = frmMain.TestCode
                    .Fields("Section") = frmMain.TestSection
                    .Fields("Flag") = 0
                 End If
                 curHi = 0
                 curLo = 0
                 curScore = 0
                 For iJudge = 1 To frmMain.TestJudges
                    Set rstSectionmarks = mdbMain.OpenRecordset("SELECT * FROM Sectionmarks WHERE Sta='" & rstEntries.Fields("Sta") & "' AND Judge=" & iJudge & " AND Mark>=0 AND Section=0 AND Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
                    If rstSectionmarks.RecordCount > 0 Then
                        If rstSectionmarks.Fields("Mark") >= 0 Then
                            iValue = iValue + 1
                            curMark = MakeStringValue(rstSectionmarks.Fields("Mark"))
                            .Fields("Mark" & Format$(iJudge)) = curMark
                            If iJudge = 1 Then
                                curLo = curMark
                            End If
                            If curMark < curLo Then
                                curLo = curMark
                            ElseIf curMark > curHi Then
                                curHi = curMark
                            End If
                            curScore = curScore + curMark
                        Else
                            .Fields("Mark" & Format$(iJudge)) = Null
                            curLo = 0
                        End If
                    Else
                        curLo = 0
                    End If
                 Next iJudge
                 If iValue = 0 Then
                    .CancelUpdate
                 Else
                    If frmMain.TestJudges = 5 Then
                        curScore = (curScore - curLo - curHi) / 3
                    Else
                        curScore = curScore / frmMain.TestJudges
                    End If
                    .Fields("Score") = curScore
                    .Fields("TimeStamp") = Now
                    .Update
                End If
            End With
            rstEntries.MoveNext
        Loop
        rstMark.Close
        rstSectionmarks.Close
    End If
    rstEntries.Close
    
    '*** update results table
    Set rstMark = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
    If rstMark.RecordCount > 0 Then
        Do While Not rstMark.EOF
            Set rstResults = mdbMain.OpenRecordset("SELECT * FROM RESULTS WHERE STA='" & rstMark.Fields("STA") & "' AND Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
            With rstResults
                If .RecordCount = 0 Then
                    .AddNew
                    .Fields("STA") = rstMark.Fields("STA")
                    .Fields("Code") = frmMain.TestCode
                    .Fields("Status") = frmMain.TestStatus
                    .Fields("Disq") = 0
                Else
                    .Edit
                End If
                .Fields("Score") = frmMain.CalculateResult(rstMark.Fields("STA"))
                .Fields("Position") = 0
                .Fields("Timestamp") = Now
                .Update
            End With
            rstMark.MoveNext
        Loop
        rstResults.Close
    End If
    rstMark.Close
        
    Set rstTempMarks = Nothing
    Set rstMark = Nothing
    Set rstEntries = Nothing
    Set rstSectionmarks = Nothing
    Set rstResults = Nothing
    
    mdbMain.Execute ("DROP TABLE [" & mcTempTableName & "]")
    
    frmMain.ClearMarks
    frmMain.LookUpRelevantParticipants
    
    SetMouseNormal

End Sub
Public Function PopulateTempTable() As Integer
    Dim rstEntries As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    Dim rstSectionmarks As DAO.Recordset
    Dim rstSections As DAO.Recordset
    Dim rstTemp As DAO.Recordset
    
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iJudge As Integer
    Dim iSection As Integer
    Dim cQry As String
    Dim tdf As DAO.TableDef
    
    '***how many sections in this test ?
    If frmMain.TestStatus > 0 Then
        iTemp2 = 4
    Else
        iTemp2 = 3
    End If
    For iTemp = iTemp2 To 0 Step -1
        Set rstSections = mdbMain.OpenRecordset("SELECT * FROM TestSections WHERE Code LIKE '" & frmMain.TestCode & "' AND Status=" & iTemp & " ORDER BY Section")
        If rstSections.RecordCount > 0 Then
            rstSections.MoveLast
            miSectionCount = rstSections.RecordCount
            Exit For
        End If
    Next iTemp
    
    If miSectionCount = 0 Then
        rstSections.Close
        Set rstSections = Nothing
        PopulateTempTable = False
        Exit Function
    Else
        '***read section info
        PopulateTempTable = True
    End If
    
    ReDim mcSectionName(1 To miSectionCount)
    ReDim mcurSectionFactor(1 To miSectionCount)
    mcurTestFactor = 0
    rstSections.MoveFirst
    iTemp = 0
    Do While Not rstSections.EOF
        iTemp = iTemp + 1
        mcSectionName(iTemp) = Translate(rstSections.Fields("Name"), mcLanguage)
        mcurSectionFactor(iTemp) = rstSections.Fields("Factor")
        mcurTestFactor = mcurTestFactor + rstSections.Fields("Factor")
        rstSections.MoveNext
    Loop
        
    '***add missing combinations
    For iJudge = 1 To frmMain.TestJudges
        Set rstEntries = mdbMain.OpenRecordset("SELECT STA FROM ENTRIES WHERE Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus & " AND NOT STA IN (SELECT STA FROM SectionMarks WHERE Code='" & frmMain.TestCode & "' AND Judge=" & iJudge & " AND Status=" & frmMain.TestStatus & " )")
        If rstEntries.RecordCount > 0 Then
            Set rstSectionmarks = mdbMain.OpenRecordset("SELECT * FROM SectionMarks")
            Do While Not rstEntries.EOF
                With rstSectionmarks
                    .AddNew
                    .Fields("STA") = rstEntries.Fields("STA")
                    .Fields("Code") = frmMain.TestCode
                    .Fields("Status") = frmMain.TestStatus
                    .Fields("Judge") = iJudge
                    .Fields("Section") = 0
                    .Fields("Mark") = -1 '*** -1 to show no mark has been entered yet
                    .Update
                End With
                For iSection = 1 To miSectionCount
                    With rstSectionmarks
                        .AddNew
                        .Fields("STA") = rstEntries.Fields("STA")
                        .Fields("Code") = frmMain.TestCode
                        .Fields("Status") = frmMain.TestStatus
                        .Fields("Judge") = iJudge
                        .Fields("Section") = iSection
                        .Fields("Mark") = -1 '*** -1 to show no mark has been entered yet
                        .Update
                    End With
                Next iSection
                rstEntries.MoveNext
            Loop
            rstSectionmarks.Close
        End If
    Next iJudge
    
    rstEntries.Close
    Set rstEntries = Nothing
    
    '***fill temp table
    mcTempTableName = Replace("_Temp-" & MachineName & "-" & frmMain.TestCode & "-" & Format$(frmMain.TestStatus), ".", "_")
    If TableExist(mdbMain, "[" & mcTempTableName & "]") Then
        mdbMain.Execute ("Drop Table [" & mcTempTableName & "]")
    End If
    
    CreateTable mdbMain, mcTempTableName, "STA", dbText, 3
    
    mdbMain.TableDefs.Refresh
    
    Set tdf = mdbMain.TableDefs("[" & mcTempTableName & "]")
    
    AppendDeleteField tdf, "APPEND", "Rider", dbText, 255
    AppendDeleteField tdf, "APPEND", "Horse", dbText, 255
    AppendDeleteField tdf, "APPEND", "Judge", dbLong
    For iSection = 1 To miSectionCount
        AppendDeleteField tdf, "APPEND", "SEC" & Format$(iSection), dbText, 10
    Next iSection
    AppendDeleteField tdf, "APPEND", "Total", dbText, 10
    AppendDeleteField tdf, "APPEND", "POS", dbLong
    
    For iJudge = 1 To frmMain.TestJudges
        cQry = " INSERT INTO [" & mcTempTableName & "] "
        cQry = cQry & " SELECT DISTINCT SectionMarks.STA,"
        cQry = cQry & " Persons.Name_First & ' ' & Persons.Name_Last & IIf(Participants.Class<>'',' [' & Participants.Class & ']','') AS Rider,"
        cQry = cQry & " Horses.Name_Horse AS Horse,"
        cQry = cQry & " SectionMarks.Judge AS Judge,"
        
        rstSections.MoveFirst
        For iSection = 1 To miSectionCount
            If iSection = 1 Then
                cQry = cQry & " IIF(SectionMarks.Mark>=0,FORMAT(SectionMarks"
            Else
                cQry = cQry & " IIF(SectionMarks_" & Format$(iSection - 1) & ".Mark>=0,FORMAT(SectionMarks_" & Format$(iSection - 1)
            End If
            cQry = cQry & ".Mark,'0.0'),'') AS Sec" & iSection & ","
            rstSections.MoveNext
        Next iSection
        cQry = cQry & " IIF(SectionMarks_" & Format$(miSectionCount) & ".Mark>=0,FORMAT(SectionMarks_" & Format$(miSectionCount) & ".Mark,'" & frmMain.TestMarkFormat & "'),'') AS Total "
        cQry = cQry & ",Entries.Position AS POS"
        
        cQry = cQry & " FROM "
        
        If miSectionCount > 1 Then
            cQry = cQry & " " & String$(miSectionCount + 1, "(")
        End If
        
        cQry = cQry & " SectionMarks "
        If miSectionCount > 1 Then
            cQry = cQry & " INNER JOIN SectionMarks AS SectionMarks_1"
            cQry = cQry & " ON (SectionMarks.Code = SectionMarks_1.Code) "
            cQry = cQry & " AND (SectionMarks.Judge = SectionMarks_1.Judge) "
            cQry = cQry & " AND (SectionMarks.Status = SectionMarks_1.Status) "
            cQry = cQry & " AND (SectionMarks.STA = SectionMarks_1.STA)) "
        End If
        
        If miSectionCount > 2 Then
            For iSection = 3 To miSectionCount
                cQry = cQry & " INNER JOIN SectionMarks AS SectionMarks_" & iSection - 1
                cQry = cQry & " ON (SectionMarks_" & iSection - 2 & ".Code = SectionMarks_" & iSection - 1 & ".Code) "
                cQry = cQry & " AND (SectionMarks_" & iSection - 2 & ".Judge = SectionMarks_" & iSection - 1 & ".Judge) "
                cQry = cQry & " AND (SectionMarks_" & iSection - 2 & ".Status = SectionMarks_" & iSection - 1 & ".Status) "
                cQry = cQry & " AND (SectionMarks_" & iSection - 2 & ".STA = SectionMarks_" & iSection - 1 & ".STA)) "
            Next iSection
        End If
                
        cQry = cQry & " INNER JOIN SectionMarks AS SectionMarks_" & miSectionCount
        cQry = cQry & " ON (SectionMarks_" & miSectionCount - 1 & ".Code = SectionMarks_" & miSectionCount & ".Code) "
        cQry = cQry & " AND (SectionMarks_" & miSectionCount - 1 & ".Judge = SectionMarks_" & miSectionCount & ".Judge) "
        cQry = cQry & " AND (SectionMarks_" & miSectionCount - 1 & ".Status = SectionMarks_" & miSectionCount & ".Status) "
        cQry = cQry & " AND (SectionMarks_" & miSectionCount - 1 & ".STA = SectionMarks_" & miSectionCount & ".STA)) "
        
        cQry = cQry & " INNER JOIN ((Participants "
        cQry = cQry & " INNER JOIN Horses "
        cQry = cQry & " ON Participants.HorseID = Horses.HorseID) "
        cQry = cQry & " INNER JOIN Persons "
        cQry = cQry & " ON Participants.PersonID = Persons.PersonID) "
        cQry = cQry & " ON SectionMarks.STA = Participants.STA) "
        cQry = cQry & " INNER JOIN Entries ON Participants.STA = Entries.STA"
        cQry = cQry & " WHERE (((SectionMarks.Code)='" & frmMain.TestCode & "')"
        cQry = cQry & " AND ((SectionMarks.Status)=" & frmMain.TestStatus & ")"
        cQry = cQry & " AND ((SectionMarks.Judge)=" & iJudge & ") "
        cQry = cQry & " AND ((SectionMarks.Section)=1) "
        If miSectionCount > 1 Then
            For iSection = 2 To miSectionCount
                cQry = cQry & " AND ((SectionMarks_" & iSection - 1 & ".Section)=" & iSection & ") "
            Next iSection
        End If
        cQry = cQry & " AND ((SectionMarks_" & miSectionCount & ".Section)=0) "
        cQry = cQry & " AND ((Entries.Status)=" & frmMain.TestStatus & ") "
        cQry = cQry & " AND ((Entries.Code)='" & frmMain.TestCode & "')"
        cQry = cQry & ")"
        
        mdbMain.Execute cQry
            
    Next iJudge
    
    Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code='" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus)
    If rstMarks.RecordCount > 0 Then
        Do While Not rstMarks.EOF
            Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM [" & mcTempTableName & "] WHERE STA='" & rstMarks.Fields("STA") & "'")
            If rstTemp.RecordCount > 0 Then
                Do While Not rstTemp.EOF
                    If rstMarks.Fields("Mark" & Format$(rstTemp.Fields("Judge"))) >= 0 Then
                        rstTemp.Edit
                        rstTemp.Fields("Total") = Format$(rstMarks.Fields("Mark" & Format$(rstTemp.Fields("Judge"))), frmMain.TestMarkFormat)
                        rstTemp.Update
                    End If
                    rstTemp.MoveNext
                Loop
            End If
            rstMarks.MoveNext
        Loop
        rstTemp.Close
        Set rstTemp = Nothing
    End If
    
    rstMarks.Close
    Set rstMarks = Nothing
    
    rstSections.Close
    Set rstSections = Nothing
    
    '***lunatic effort to overcome an EOF error
    Dim rst As DAO.Recordset
    Set rst = mdbMain.OpenRecordset("SELECT * FROM [" & mcTempTableName & "]")
    With rst
        .AddNew
        .Fields("STA") = "000"
        .Fields("Rider") = Left$(UserName, .Fields("Rider").Size)
        .Fields("Horse") = Left$(MachineName, .Fields("Horse").Size)
        .Fields("Judge") = 6
        .Fields("POS") = 999
        .Update
    End With
    rst.Close
    Set rst = Nothing
    
End Function

