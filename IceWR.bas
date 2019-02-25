Attribute VB_Name = "modIceWR"
' Functions related to WorldRanking

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
Function CreateWRFile() As Integer
   Dim cTemp As String
   Dim iKey As Integer
   Dim iFileNum As Integer
   Dim cFileName As String
   Dim iExported As Integer
   Dim rstMarks As DAO.Recordset
   Dim rstWR As DAO.Recordset
   Dim cQry As String
   
   On Local Error Resume Next
   
   iExported = False
   
   On Local Error Resume Next
   
   cFileName = GetVariable("WR_File")
   If cFileName = "" Then
        cFileName = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & GetVariable("WR_Code") & ".Txt"
   End If
   With frmMain.CommonDialog1
        .CancelError = True
        .DefaultExt = ".Txt"
        .DialogTitle = "Select a folder"
        .Filter = "WR Files|*.Txt|"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly
        .FileName = cFileName
        .ShowSave
        If Err = cdlCancel Then
            Exit Function
        End If
        If .FileName <> "" Then
            cFileName = NameOfFile(.FileName) & ".Txt"
        End If
    End With
    
    On Local Error GoTo CreateWrFileError
    
    If InStr(cFileName, GetVariable("WR_Code")) = 0 Then
        MsgBox cFileName & " " & Translate("is not a valid file name!", mcLanguage)
        Exit Function
    End If
    
    If cFileName = "" Or cFileName = Chr$(27) Then
    Else
        If InStr(cFileName, "\") = 0 And InStr(cFileName, ":") = 0 Then
            cFileName = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & cFileName
        End If
        iKey = MsgBox(Translate("Create/update FEIF WorldRanking File", mcLanguage) & " '" & cFileName & "'?", vbQuestion + vbYesNo)
        
        If iKey = vbYes Then
            iFileNum = FreeFile
            Open cFileName For Output Access Write Shared As #iFileNum
            CreateWRFile = True
            SetMouseHourGlass
            SetVariable "WR_file", cFileName
            
            'select WR tests, oval track classes only if 5 judges
            cQry = "SELECT Tests.Code,Tests.WRTest,Tests_1.Type_Pre,Tests_1.WR "
            cQry = cQry & " FROM (Tests INNER JOIN Tests AS Tests_1 ON Tests.WRTest = Tests_1.Code) INNER JOIN TestInfo ON Tests_1.Code=TestInfo.Code "
            cQry = cQry & " WHERE (((TestInfo.Num_J_0)=5) "
            cQry = cQry & " AND ((Tests.WRTest)<>'')) "
            cQry = cQry & " OR (((Tests.WRTest)<>'') "
            cQry = cQry & " AND ((Tests.Type_Time)>0))"
            Set rstWR = mdbMain.OpenRecordset(cQry)
            If rstWR.RecordCount > 0 Then
                ShowProgressbar frmMain, 2, rstWR.RecordCount
                Do While Not rstWR.EOF
                    IncreaseProgressbarValue frmMain.ProgressBar1, rstWR.AbsolutePosition
                    cTemp = "SELECT Trim(Name_First & ' ' & Name_Middle) & ' ' & Name_Last AS Rider"
                    cTemp = cTemp & " , Results.Score"
                    cTemp = cTemp & " , Horses.FEIFID"
                    cTemp = cTemp & " , Horses.Name_Horse"
                    cTemp = cTemp & " , Participants.Class"
                    cTemp = cTemp & " , Format(Persons.BirthDay,'yyyy-mm-dd')"
                    cTemp = cTemp & " FROM ((Results INNER JOIN Participants ON Results.STA = Participants.STA)"
                    cTemp = cTemp & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID)"
                    cTemp = cTemp & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
                    cTemp = cTemp & " WHERE Results.Code ='" & rstWR.Fields("Code") & "' AND Results.Status=0 AND Results.Disq=0"
                    cTemp = cTemp & " AND Results.Score>0"
                    cTemp = cTemp & " ORDER BY Results.Code,Results.Score"
                    Set rstMarks = mdbMain.OpenRecordset(cTemp)
                    If rstMarks.RecordCount > 0 Then
                        iExported = True
                        Do While Not rstMarks.EOF
                            Print #iFileNum, Replace(Trim$(rstMarks.Fields("Rider")), "  ", " ") & vbTab & rstWR.Fields("WRTest") & vbTab & Replace(Format$(rstMarks.Fields("Score"), frmMain.TestTotalFormat), ",", ".") & vbTab & rstMarks.Fields("FEIFId") & "" & vbTab & IIf(rstMarks.Fields("FEIFId") & "" <> "", rstMarks.Fields("Name_Horse"), "")
                            rstMarks.MoveNext
                        Loop
                    End If
                    rstMarks.Close
                    rstWR.MoveNext
                Loop
            End If
            rstWR.Close
        
            'select NR tests, oval track classes only if 3 judges
            cQry = "SELECT Tests.Code,Tests.NRTest,Tests_1.Type_Pre "
            cQry = cQry & " FROM (Tests INNER JOIN Tests AS Tests_1 ON Tests.NRTest = Tests_1.Code) INNER JOIN TestInfo ON Tests_1.Code=TestInfo.Code "
            cQry = cQry & " WHERE (((TestInfo.Num_J_0)>=3) "
            cQry = cQry & " AND ((Tests.NRTest)<>'')) "
            cQry = cQry & " OR (((Tests.NRTest)<>'') "
            cQry = cQry & " AND ((Tests.Type_Time)>0))"
            Set rstWR = mdbMain.OpenRecordset(cQry)
            If rstWR.RecordCount > 0 Then
                ShowProgressbar frmMain, 2, rstWR.RecordCount
                Do While Not rstWR.EOF
                    IncreaseProgressbarValue frmMain.ProgressBar1, rstWR.AbsolutePosition
                    cTemp = "SELECT Trim(Name_First & ' ' & Name_Middle) & ' ' & Name_Last AS Rider"
                    cTemp = cTemp & " , Results.Score"
                    cTemp = cTemp & " , Horses.FEIFID"
                    cTemp = cTemp & " , Horses.Name_Horse"
                    cTemp = cTemp & " , Participants.Class"
                    cTemp = cTemp & " , Format(Persons.BirthDay,'yyyy-mm-dd')"
                    cTemp = cTemp & " FROM ((Results INNER JOIN Participants ON Results.STA = Participants.STA)"
                    cTemp = cTemp & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID)"
                    cTemp = cTemp & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
                    cTemp = cTemp & " WHERE Results.Code ='" & rstWR.Fields("Code") & "' AND Results.Status=0 AND Results.Disq=0"
                    cTemp = cTemp & " AND Results.Score>0"
                    cTemp = cTemp & " ORDER BY Results.Code,Results.Score"
                    Set rstMarks = mdbMain.OpenRecordset(cTemp)
                    If rstMarks.RecordCount > 0 Then
                        iExported = True
                        Do While Not rstMarks.EOF
                            Print #iFileNum, Replace(Trim$(rstMarks.Fields("Rider")), "  ", " ") & vbTab & rstWR.Fields("NRTest") & vbTab & Replace(Format$(rstMarks.Fields("Score"), frmMain.TestTotalFormat), ",", ".") & vbTab & rstMarks.Fields("FEIFId") & "" & vbTab & IIf(rstMarks.Fields("FEIFId") & "" <> "", rstMarks.Fields("Name_Horse"), "")
                            rstMarks.MoveNext
                        Loop
                    End If
                    rstMarks.Close
                    rstWR.MoveNext
                Loop
            End If
            rstWR.Close
            
            Print #iFileNum, "[END]"
            
            'report judges
            cQry = " SELECT Persons.Name_First & ' ' & Persons.Name_Last AS Judge, Tests.WRTest, Count(Results.STA) AS [Number]"
            cQry = cQry & " FROM ((((Results INNER JOIN TestJudges ON (Results.Code = TestJudges.Code) AND (Results.Status = TestJudges.Status))"
            cQry = cQry & " INNER JOIN Persons ON TestJudges.JudgeId = Persons.PersonID)"
            cQry = cQry & " INNER JOIN Tests ON Results.Code = Tests.Code)"
            cQry = cQry & " INNER JOIN TestInfo ON Results.Code = TestInfo.Code)"
            cQry = cQry & " Where Results.Status = 0 AND ((TestInfo.Num_J_0=5  AND Tests.WRTest<>'')  OR (Tests.WRTest<>'' AND Tests.Type_Time>0))"
            cQry = cQry & " GROUP BY Persons.Name_First & ' ' & Persons.Name_Last, Tests.WRTest"
            Set rstWR = mdbMain.OpenRecordset(cQry)
            If rstWR.RecordCount > 0 Then
                Print #iFileNum, vbCrLf & "[JUDGES]"
                Do While Not rstWR.EOF
                    Print #iFileNum, rstWR.Fields(0) & vbTab & rstWR.Fields(1) & vbTab & rstWR.Fields(2)
                    rstWR.MoveNext
                Loop
                Print #iFileNum, "[END]"
            End If
            
            cQry = "SELECT Persons.Name_First & ' ' & Persons.Name_Last AS Judge, FORMAT(Results.Timestamp,'YYYY-MM-DD') AS [Date] "
            cQry = cQry & " FROM ((Results INNER JOIN TestJudges ON (Results.Code = TestJudges.Code) AND (Results.Status = TestJudges.Status)) "
            cQry = cQry & " INNER JOIN Persons ON TestJudges.JudgeId = Persons.PersonID) "
            cQry = cQry & " GROUP BY Persons.Name_First & ' ' & Persons.Name_Last, FORMAT(Results.Timestamp,'YYYY-MM-DD') "
            cQry = cQry & " ORDER BY Persons.Name_First & ' ' & Persons.Name_Last,FORMAT(Results.Timestamp,'YYYY-MM-DD')"
            Set rstWR = mdbMain.OpenRecordset(cQry)
            If rstWR.RecordCount > 0 Then
                Print #iFileNum, vbCrLf & "[JUDGES DAYS]"
                Do While Not rstWR.EOF
                    Print #iFileNum, rstWR.Fields(0) & vbTab & rstWR.Fields(1)
                    rstWR.MoveNext
                Loop
                Print #iFileNum, "[END]"
            End If
            
            rstWR.Close
        
            If iExported = False Then
                CreateWRFile = False
                MsgBox Translate("No riders found.", mcLanguage), vbExclamation
                KillFile cFileName
            End If
        Else
            MsgBox Translate("No FEIF WorldRanking tests found (import Sport Rules again).", mcLanguage)
        End If
    End If
    
CreateWrFileError:
    If Err > 0 Then
        CreateWRFile = False
        MsgBox cFileName & ": " & Err.Description, vbCritical
    End If
    
    If iFileNum Then
        Close #iFileNum
    End If
    
    Set rstMarks = Nothing
    Set rstWR = Nothing
    
    ShowProgressbar frmMain, 2, 0
    SetMouseNormal

End Function
Public Function WrTest(Test As String) As String
    'check if current test is in WorldRanking
    Dim rstWR As DAO.Recordset
    Set rstWR = mdbMain.OpenRecordset("SELECT * FROM Tests WHERE Code='" & Test & "'")
    If rstWR.RecordCount > 0 Then
        If rstWR.Fields("WR") > 0 Then
            WrTest = rstWR.Fields("WRTest") & ""
        End If
    End If
    rstWR.Close
    Set rstWR = Nothing
End Function
Public Function MakeKeyFromWrCode() As String
    Dim cTemp As String
    Dim cTemp2 As String
    
    cTemp = GetVariable("WR_Code")
    If Len(cTemp) = 9 Then
        cTemp2 = Right$(Format$(Val(Right$(cTemp, 7)) Mod 11), 1)
        MakeKeyFromWrCode = cTemp & cTemp2 & App.FileDescription
    Else
        MakeKeyFromWrCode = cTemp & App.FileDescription
    End If
    
End Function
Public Function WRLimit(Test As String) As Currency
    Dim rstWR As DAO.Recordset
    Set rstWR = mdbMain.OpenRecordset("SELECT * FROM Tests WHERE Code='" & Test & "'")
    If rstWR.RecordCount > 0 Then
        If rstWR.Fields("WR") > 0 Then
            WRLimit = rstWR.Fields("WR")
        End If
    End If
    rstWR.Close
    Set rstWR = Nothing
End Function
