Attribute VB_Name = "modIceCombination"
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

Function CalculateCombination(CombinationCode As String, Optional ShowMessage As Integer = True) As Integer

    Dim rstComb As DAO.Recordset
    Dim rstResult As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    Dim rstTemp As DAO.Recordset
    Dim tdfComb As DAO.TableDef
    Dim iGroupcount As Integer
    Dim iGroup As Integer
    Dim iTmpGroup As Integer
    Dim iHasScore As Integer
    Dim iKey As Integer
    Dim curScore As Currency
    Dim iPosition As Integer
    
    Set rstComb = mdbMain.OpenRecordset("SELECT * FROM Combinations WHERE Code LIKE '" & CombinationCode & "'")
    If rstComb.RecordCount > 0 Then
        If ShowMessage = False Then
            iKey = vbYes
        Else
            iKey = MsgBox(Translate("Calculate and print combination", mcLanguage) & " " & Translate(rstComb.Fields("Combination"), mcLanguage) & "?", vbYesNo + vbQuestion)
        End If
        If iKey = vbYes Then
            SetMouseHourGlass
            Set rstComb = mdbMain.OpenRecordset("SELECT * FROM CombinationSections WHERE Code LIKE '" & CombinationCode & "' ORDER BY Group")
            If rstComb.RecordCount > 0 Then
                rstComb.MoveLast
                iGroupcount = rstComb.Fields("Group")
                If TableExist(mdbMain, "[_Temp-" & MachineName & "]") = True Then
                    mdbMain.Execute ("DROP Table [_Temp-" & MachineName & "]")
                End If
                If iGroupcount >= 1 Then
                    CreateTable mdbMain, "_Temp-" & MachineName, "STA", dbText, 3
                    mdbMain.TableDefs.Refresh
                    Set tdfComb = mdbMain.TableDefs("[_Temp-" & MachineName & "]")
                    For iGroup = 1 To iGroupcount
                        AppendDeleteField tdfComb, "APPEND", "Score" & Format$(iGroup), dbCurrency
                        AppendDeleteField tdfComb, "APPEND", "Code" & Format$(iGroup), dbText, 8
                    Next iGroup
                    AppendDeleteField tdfComb, "APPEND", "Score", dbCurrency
                    For iGroup = 1 To iGroupcount
                        Set rstComb = mdbMain.OpenRecordset("SELECT * FROM CombinationSections WHERE [Group]=" & iGroup & " AND Code LIKE '" & CombinationCode & "'")
                        Do While Not rstComb.EOF
                            Set rstResult = mdbMain.OpenRecordset("SELECT * FROM Results INNER JOIN Tests ON Tests.Code=Results.Code WHERE Results.Code='" & rstComb.Fields("Test") & "' AND Results.Score>0 AND Results.Status=0 AND Results.Disq=0")
                            If rstResult.RecordCount > 0 Then
                                Do While Not rstResult.EOF
                                    Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM [_TEMP-" & MachineName & "] WHERE STA='" & rstResult.Fields("STA") & "'")
                                    If rstTemp.RecordCount = 0 Then
                                        rstTemp.AddNew
                                        rstTemp.Fields("STA") = rstResult.Fields("STA")
                                        If rstResult.Fields("Type_Pre") > 2 Then
                                            rstTemp.Fields("Score" & Format$(iGroup)) = Time2Mark(rstResult.Fields("Score"), rstComb.Fields("Test")) * rstComb.Fields("Factor")
                                        Else
                                            rstTemp.Fields("Score" & Format$(iGroup)) = rstResult.Fields("Score") * rstComb.Fields("Factor")
                                        End If
                                        rstTemp.Fields("Code" & Format$(iGroup)) = rstResult.Fields("Results.Code")
                                    Else
                                        rstTemp.Edit
                                        If IsNull(rstTemp.Fields("Score" & Format$(iGroup))) Then
                                            If rstResult.Fields("Type_Pre") > 2 Then
                                                rstTemp.Fields("Score" & Format$(iGroup)) = Time2Mark(rstResult.Fields("Score"), rstComb.Fields("Test")) * rstComb.Fields("Factor")
                                            Else
                                                rstTemp.Fields("Score" & Format$(iGroup)) = rstResult.Fields("Score") * rstComb.Fields("Factor")
                                            End If
                                            rstTemp.Fields("Code" & Format$(iGroup)) = rstResult.Fields("Results.Code")
                                        Else
                                            If rstResult.Fields("Type_Pre") < 2 Then
                                                If rstTemp.Fields("Score" & Format$(iGroup)) < rstResult.Fields("Score") * rstComb.Fields("Factor") Then
                                                    rstTemp.Fields("Score" & Format$(iGroup)) = rstResult.Fields("Score") * rstComb.Fields("Factor")
                                                    rstTemp.Fields("Code" & Format$(iGroup)) = rstResult.Fields("Results.Code")
                                                End If
                                            ElseIf rstResult.Fields("Type_Pre") > 2 Then
                                                If rstTemp.Fields("Score" & Format$(iGroup)) < Time2Mark(rstResult.Fields("Score"), rstComb.Fields("Test")) * rstComb.Fields("Factor") Then
                                                    rstTemp.Fields("Score" & Format$(iGroup)) = Time2Mark(rstResult.Fields("Score"), rstComb.Fields("Test")) * rstComb.Fields("Factor")
                                                    rstTemp.Fields("Code" & Format$(iGroup)) = rstResult.Fields("Results.Code")
                                                End If
                                            End If
                                        End If
                                    End If
                                    If iGroup = iGroupcount Then
                                        iHasScore = True
                                        curScore = 0
                                        For iTmpGroup = 1 To iGroupcount
                                            If IsNull(rstTemp.Fields("Score" & Format$(iTmpGroup))) Or rstTemp.Fields("Score" & Format$(iTmpGroup)) = 0 Then
                                                iHasScore = False
                                            Else
                                                curScore = curScore + Val(Replace(Format$(rstTemp.Fields("Score" & Format$(iTmpGroup)), frmMain.TestTotalFormat), ",", "."))
                                            End If
                                        Next iTmpGroup
                                        If iHasScore = True Then
                                            rstTemp.Fields("Score") = curScore / iGroupcount
                                        End If
                                    End If
                                    rstTemp.Update
                                    rstResult.MoveNext
                                Loop
                                rstTemp.Close
                            End If
                            rstComb.MoveNext
                            rstResult.Close
                        Loop
                    Next iGroup
                    
                End If
            End If
            
            mdbMain.Execute ("DELETE * FROM Results WHERE Code='" & CombinationCode & "'")
            mdbMain.Execute ("DELETE * FROM Marks WHERE Code='" & CombinationCode & "'")
            DoEvents
            
            Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM [_TEMP-" & MachineName & "] WHERE Score>0 ORDER BY Score DESC")
            If rstTemp.RecordCount > 0 Then
                Set rstResult = mdbMain.OpenRecordset("SELECT * FROM Results")
                Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks")
                iPosition = 0
                curScore = 0
                Do While Not rstTemp.EOF
                    If curScore <> rstTemp.Fields("Score") Then
                        iPosition = rstTemp.AbsolutePosition + 1
                        curScore = rstTemp.Fields("Score")
                    End If
                    With rstResult
                        .AddNew
                        .Fields("Sta") = rstTemp.Fields("STA")
                        .Fields("Code") = CombinationCode
                        .Fields("Status") = 0
                        .Fields("Disq") = 0
                        .Fields("FR") = False
                        .Fields("Score") = rstTemp.Fields("Score")
                        .Fields("Position") = iPosition
                        .Fields("TimeStamp") = Now
                        .Update
                    End With
                    For iGroup = 1 To iGroupcount
                        With rstMarks
                            .AddNew
                            .Fields("Sta") = rstTemp.Fields("STA")
                            .Fields("Code") = CombinationCode
                            .Fields("Status") = 0
                            .Fields("Flag") = 0
                            .Fields("Section") = iGroup
                            .Fields("Score") = rstTemp.Fields("Score" & Format$(iGroup))
                            .Fields("Judge1") = rstTemp.Fields("Code" & Format$(iGroup))
                            .Fields("TimeStamp") = Now
                            .Update
                        End With
                    Next iGroup
                    rstTemp.MoveNext
                Loop
                
            End If
            
            rstTemp.Close
            
            If ShowMessage = True Then
                frmMain.PrintCombination CombinationCode, "_Temp-" & MachineName
                            
            End If
            SetCombinationStatus CombinationCode, 1
        End If
        SetMouseNormal
    End If
    
    Set rstTemp = Nothing
    Set rstResult = Nothing
    Set rstMarks = Nothing
    
    rstComb.Close
    Set rstComb = Nothing
    
End Function

Public Function CreateCombinationInfoTable() As Integer
    Dim tdf As DAO.TableDef
    
    If TableExist(mdbMain, "CombinationInfo") = False Then
        CreateTable mdbMain, "CombinationInfo", "Code", dbText, 8
        mdbMain.TableDefs.Refresh
    End If
    
    Set tdf = mdbMain.TableDefs("CombinationInfo")
    AppendDeleteField tdf, "APPEND", "Status", dbInteger
    AppendDeleteField tdf, "APPEND", "Nr", dbInteger
    AppendDeleteField tdf, "APPEND", "Sponsor", dbText, 255
    
    'seems redundant but is necessary for backwards compatibility
    AppendDeleteField tdf, "APPEND", "Code", dbText, 8
    
    Set tdf = Nothing
    
End Function

Public Function SetCombinationStatus(cCombinationCode As String, iStatus As Integer) As Integer
    Dim rst As DAO.Recordset
    
    CreateCombinationInfo cCombinationCode
    Set rst = mdbMain.OpenRecordset("SELECT * FROM CombinationInfo WHERE Code='" & cCombinationCode & "'")
    With rst
        If .RecordCount > 0 Then
            If .Fields("Status") < iStatus Then
                .Edit
                .Fields("Status") = iStatus
                .Update
            End If
            SetCombinationStatus = iStatus
        Else
            SetCombinationStatus = 0
        End If
        .Close
    End With
    Set rst = Nothing

End Function
Public Function GetCombinationStatus(cCombinationCode As String) As Integer
    Dim rst As DAO.Recordset
    
    CreateCombinationInfo cCombinationCode
    Set rst = mdbMain.OpenRecordset("SELECT * FROM CombinationInfo WHERE Code='" & cCombinationCode & "'")
    With rst
        If .RecordCount > 0 Then
            GetCombinationStatus = .Fields("Status")
        Else
            GetCombinationStatus = 0
        End If
        .Close
    End With
    Set rst = Nothing

End Function
Public Function CreateCombinationInfo(cCombinationCode As String) As Integer
    Dim rst As DAO.Recordset
    
    CreateCombinationInfoTable
    Set rst = mdbMain.OpenRecordset("SELECT * FROM CombinationInfo WHERE Code='" & cCombinationCode & "'")
    With rst
        If .RecordCount = 0 Then
            .AddNew
            .Fields("Code") = Left(cCombinationCode, .Fields("Code").Size)
            .Fields("Status") = 0
            .Fields("Nr") = 0
            .Update
        End If
        .Close
    End With
    Set rst = Nothing
End Function

Public Function GetCombinationSponsor(cCombinationCode As String) As String
    Dim rst As DAO.Recordset
    
    CreateCombinationInfoTable
    Set rst = mdbMain.OpenRecordset("SELECT * FROM CombinationInfo WHERE Code='" & cCombinationCode & "'")
    With rst
        If .RecordCount > 0 Then
            GetCombinationSponsor = .Fields("Sponsor") & ""
        End If
        .Close
    End With
    Set rst = Nothing
End Function

Function CalculateTeamCombination(CombinationCode As String, Optional ShowMessage As Integer = True) As Integer

    Dim rstComb As DAO.Recordset
    Dim rstResult As DAO.Recordset
    Dim rstTemp As DAO.Recordset
    Dim rstTemp2 As DAO.Recordset
    Dim tdfComb As DAO.TableDef
    Dim iGroupcount As Integer
    Dim iGroup As Integer
    Dim iTmpGroup As Integer
    Dim iHasScore As Integer
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim cTemp As String
    Dim curScore As Currency
    Dim iPosition As Integer
    Dim cQry As String
    Dim iNumResults As Integer
    Dim iNumTests As Integer
    Dim cTestList As String
    Dim cTeam As String
    Dim lTemp As Long
    
    '* Extra check added in the kitchen of FEIF President Jens Iversen
    '*
    Set rstComb = mdbMain.OpenRecordset("SELECT DISTINCT CODE FROM TESTS WHERE Code IN (SELECT Code FROM Results)")
    If rstComb.RecordCount = 0 Then
        rstComb.Close
        Set rstComb = Nothing
        MsgBox Translate("Results of at least one test are required to calculate such a combination", mcLanguage), vbOK + vbExclamation, CombinationCode
    Else
        rstComb.Close
    
        With frmToolBox
            .intChecked = True
            .strQry = "SELECT DISTINCT CODE FROM TESTS WHERE Code IN (SELECT Code FROM Results) ORDER BY Code"
            .strQry2 = "SELECT Test FROM Combinationsections WHERE Code LIKE '" & CombinationCode & "'"
            .Caption = Translate("Combination", mcLanguage) & ": " & Translate(CombinationCode, mcLanguage)
            .Show 1
        End With
        
        If frmMain.Tempvar <> "" Then
            cTemp = GetVariable("Comb_" & CombinationCode & "_Num")
            If cTemp = "" Then cTemp = 2
            cTemp = InputBox(Translate("How many partcipants per test", mcLanguage) & "?", CombinationCode & " " & Translate("Combination", mcLanguage), cTemp)
            If Val(cTemp) > 0 And cTemp <> Chr$(27) Then
                SetVariable "Comb_" & CombinationCode & "_Num", cTemp
                iNumResults = Val(cTemp)
            Else
                frmMain.Tempvar = ""
            End If
        Else
            MsgBox Translate("No tests selected", mcLanguage)
        End If
        
        If frmMain.Tempvar <> "" Then
        
            mdbMain.Execute "DELETE * FROM Combinations WHERE Code LIKE '" & CombinationCode & "'"
            mdbMain.Execute "DELETE * FROM Combinationsections WHERE Code LIKE '" & CombinationCode & "'"
            
            Set rstComb = mdbMain.OpenRecordset("SELECT * FROM Combinations")
            With rstComb
                .AddNew
                .Fields("Code") = CombinationCode
                .Fields("Combination") = CombinationCode
                .Fields("Userlevel") = -1
                .Update
            End With
            
            Set rstComb = mdbMain.OpenRecordset("SELECT * FROM Combinationsections")
            cTestList = frmMain.Tempvar
            Do While cTestList <> ""
                Parse cTemp, cTestList, "|"
                If cTemp <> "" Then
                    With rstComb
                        .AddNew
                        .Fields("Code") = CombinationCode
                        .Fields("Factor") = 1
                        .Fields("Group") = 1
                        .Fields("Test") = cTemp
                        .Update
                    End With
                End If
            Loop
            
            Set rstComb = mdbMain.OpenRecordset("SELECT * FROM Combinations WHERE Code LIKE '" & CombinationCode & "'")
            If rstComb.RecordCount > 0 Then
                If ShowMessage = False Then
                    iKey = vbYes
                Else
                    iKey = MsgBox(Translate("Calculate and print combination", mcLanguage) & ": " & Translate(rstComb.Fields("Combination"), mcLanguage) & "?", vbYesNo + vbQuestion)
                End If
                If iKey = vbYes Then
                    SetMouseHourGlass
                    Set rstComb = mdbMain.OpenRecordset("SELECT * FROM CombinationSections WHERE Code LIKE '" & CombinationCode & "' ORDER BY Group")
                    If rstComb.RecordCount > 0 Then
                        rstComb.MoveLast
                        iGroupcount = rstComb.Fields("Group")
                        If TableExist(mdbMain, "[_Temp-" & MachineName & "]") = True Then
                            mdbMain.Execute ("DROP Table [_Temp-" & MachineName & "]")
                        End If
                        CreateTable mdbMain, "_Temp-" & MachineName, "Team", dbText, 100
                        mdbMain.TableDefs.Refresh
                        Set tdfComb = mdbMain.TableDefs("[_Temp-" & MachineName & "]")
                        AppendDeleteField tdfComb, "APPEND", "Score", dbCurrency
                        iNumTests = 0
                        cTestList = "|"
                        Set rstComb = mdbMain.OpenRecordset("SELECT * FROM CombinationSections WHERE Code LIKE '" & CombinationCode & "'")
                        Do While Not rstComb.EOF
                            AppendDeleteField tdfComb, "APPEND", Replace(rstComb.Fields("Test"), ".", ""), dbCurrency
                            iNumTests = iNumTests + 1
                            cTestList = cTestList & Replace(rstComb.Fields("Test"), ".", "") & "|"
                            rstComb.MoveNext
                        Loop
                        
                        rstComb.MoveFirst
                        
                        Set rstTemp2 = mdbMain.OpenRecordset("SELECT DISTINCT " & CombinationCode & " FROM Participants ORDER BY " & CombinationCode)
                        If rstTemp2.RecordCount > 0 Then
                            Do While Not rstTemp2.EOF
                                Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM [_TEMP-" & MachineName & "]")
                                For iTemp = 1 To iNumResults
                                    rstTemp.AddNew
                                    rstTemp.Fields("Team") = rstTemp2.Fields(CombinationCode)
                                    rstComb.MoveFirst
                                    Do While Not rstComb.EOF
                                        rstTemp.Fields(Replace(rstComb.Fields("Test"), ".", "")) = 0
                                        rstComb.MoveNext
                                    Loop
                                    rstTemp.Update
                                Next iTemp
                                rstTemp2.MoveNext
                            Loop
                        End If
                        rstTemp2.Close
                        
                        rstComb.MoveFirst
                        Do While Not rstComb.EOF
                            cQry = "SELECT Results.*,Tests.*,Participants." & CombinationCode & " FROM (Results "
                            cQry = cQry & " INNER JOIN Tests ON Tests.Code=Results.Code) "
                            cQry = cQry & " INNER JOIN Participants ON Results.Sta=Participants.Sta "
                            cQry = cQry & " WHERE Results.Code='" & rstComb.Fields("Test") & "' "
                            cQry = cQry & " AND Results.Score>0 "
                            cQry = cQry & " AND Results.Status=0 "
                            cQry = cQry & " AND Results.Disq=0 "
                            cQry = cQry & " ORDER BY Results.Score DESC "
                            
                            Set rstResult = mdbMain.OpenRecordset(cQry)
                            If rstResult.RecordCount > 0 Then
                                Do While Not rstResult.EOF
                                    If rstResult.Fields("Type_Pre") > 2 Then
                                        curScore = Time2Mark(rstResult.Fields("Score"), rstComb.Fields("Test")) * rstComb.Fields("Factor")
                                    Else
                                        curScore = rstResult.Fields("Score") * rstComb.Fields("Factor")
                                    End If
                                    cQry = "SELECT * FROM [_TEMP-" & MachineName & "] "
                                    cQry = cQry & " WHERE Team='" & rstResult.Fields(CombinationCode) & "'"
                                    cQry = cQry & " AND " & Replace(rstResult.Fields("Results.Code"), ".", "") & "<" & Replace(Format$(curScore), ",", ".")
                                    cQry = cQry & " ORDER BY " & Replace(rstResult.Fields("Results.Code"), ".", "")
                                    Set rstTemp = mdbMain.OpenRecordset(cQry)
                                    If rstTemp.RecordCount > 0 Then
                                        rstTemp.Edit
                                        rstTemp.Fields(Replace(rstResult.Fields("Results.Code"), ".", "")) = curScore
                                        rstTemp.Update
                                    End If
                                    rstResult.MoveNext
                                Loop
                            End If
                            rstComb.MoveNext
                            rstResult.Close
                        Loop
                        
                        cQry = "SELECT * FROM [_TEMP-" & MachineName & "] "
                        cQry = cQry & " ORDER BY Team"
                        Set rstTemp = mdbMain.OpenRecordset(cQry)
                        If rstTemp.RecordCount > 0 Then
                            iTemp = 0
                            curScore = 0
                            Do While Not rstTemp.EOF
                                iTemp = iTemp + 1
                                Dim fld As DAO.Field
                                For Each fld In rstTemp.Fields
                                    If InStr(cTestList, "|" & fld.Name & "|") > 0 Then
                                        curScore = curScore + fld.Value
                                    End If
                                Next
                                If iTemp = iNumResults Then
                                    rstTemp.Edit
                                    rstTemp.Fields("Score") = curScore / (iNumTests * iNumResults)
                                    rstTemp.Update
                                    curScore = 0
                                    iTemp = 0
                                End If
                                rstTemp.MoveNext
                            Loop
                        End If
                        rstTemp.Close
                    End If
                    
                    If ShowMessage = True Then
                        frmMain.PrintTeamCombination CombinationCode, "_Temp-" & MachineName, iNumTests, CombinationCode
                    End If
                    
                End If
                SetMouseNormal
            Else
                MsgBox Translate("Not yet implemented", mcLanguage)
            End If
            
            Set rstTemp = Nothing
            Set rstResult = Nothing
            
            rstComb.Close
            Set rstComb = Nothing
                
        End If
    End If
    frmMain.Tempvar = ""

End Function

