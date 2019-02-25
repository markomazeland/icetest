Attribute VB_Name = "modIceTestInfo"
' Functions related additional info for tests

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
Public Function GetHighestPosition(cTestCode As String, Optional iTestStatus As Integer = 2) As Integer
    Dim rst As DAO.Recordset
    
    On Local Error Resume Next
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestInfo WHERE Code='" & cTestCode & "'")
    With rst
        If iTestStatus = 0 Or iTestStatus = 1 Then
            GetHighestPosition = 1
        ElseIf iTestStatus = 3 Then
            If .RecordCount > 0 Then
                GetHighestPosition = IIf(IsNull(.Fields("CFinal")), 0, .Fields("CFinal"))
                If GetHighestPosition = 0 Then
                    GetHighestPosition = 11
                End If
            Else
                GetHighestPosition = 6
            End If
        Else
            If .RecordCount > 0 Then
                GetHighestPosition = IIf(IsNull(.Fields("BFinal")), 0, .Fields("BFinal"))
                If GetHighestPosition = 0 Then
                    GetHighestPosition = 6
                End If
            Else
                GetHighestPosition = 1
            End If
        End If
        .Close
    End With
    Set rst = Nothing
    On Local Error GoTo 0
    
End Function
Public Function SetHighestPosition(cTestCode As String, iPosition As Integer, Optional iStatus As Integer = 2) As Integer

    Dim rst As DAO.Recordset
    
    On Local Error Resume Next
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestInfo WHERE Code='" & cTestCode & "'")
    With rst
        If .RecordCount > 0 And iStatus = 3 Then
            .Edit
            .Fields("CFinal") = iPosition
            .Update
            SetHighestPosition = iPosition
        ElseIf .RecordCount > 0 Then
            .Edit
            .Fields("BFinal") = iPosition
            .Update
            SetHighestPosition = iPosition
        Else
            SetHighestPosition = 1
        End If
        .Close
    End With
    Set rst = Nothing
    On Local Error GoTo 0
    
End Function
Public Function SetTestStatus(cTestCode As String, iStatus As Integer, Optional iForce As Integer = False) As Integer
    Dim rst As DAO.Recordset
    
    On Local Error Resume Next
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestInfo WHERE Code='" & cTestCode & "'")
    With rst
        If .RecordCount > 0 Then
            If .Fields("Status") < iStatus Or iForce = True Then
                .Edit
                .Fields("Status") = iStatus
                .Update
            End If
            SetTestStatus = iStatus
        Else
            SetTestStatus = -1
        End If
        .Close
    End With
    Set rst = Nothing
    On Local Error GoTo 0

End Function
Public Function GetTestStatus(cTestCode As String) As Integer
    Dim rst As DAO.Recordset
    
    On Local Error Resume Next
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestInfo WHERE Code='" & cTestCode & "'")
    With rst
        If .RecordCount > 0 Then
            GetTestStatus = .Fields("Status")
        Else
            GetTestStatus = 0
        End If
        .Close
    End With
    Set rst = Nothing
    On Local Error GoTo 0

End Function
Public Function CreateTestInfoAll() As Integer
    Dim rst As DAO.Recordset
    Dim iTestStatus As Integer
    iTestStatus = frmMain.TestStatus
    
    On Local Error Resume Next
    Set rst = mdbMain.OpenRecordset("SELECT Code FROM Tests")
    With rst
        If .RecordCount > 0 Then
            Do While Not .EOF
                CreateTestInfo .Fields("Code")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rst = Nothing
    frmMain.TestStatus = iTestStatus
    
    
    On Local Error GoTo 0
End Function

Public Function CreateTestInfo(cTestCode As String, Optional iTestStatus As Integer = 0, Optional iNum_J As Integer = 5) As Integer
    Dim rst As DAO.Recordset
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestInfo WHERE Code='" & cTestCode & "'")
    With rst
        If .RecordCount = 0 Then
            .AddNew
            .Fields("Code") = cTestCode
            .Fields("Status") = 0
            .Fields("CFinal") = 0
            .Fields("BFinal") = 0
            .Fields("Handling") = 2
            .Fields("SplitFinals") = 0
            .Fields("Sponsor") = ""
            .Fields("Nr") = 0
            .Fields("SortDigit") = 0
            .Fields("SortChar") = ""
            .Fields("num_j") = iNum_J
            .Fields("num_j_0") = iNum_J
            .Fields("num_j_1") = iNum_J
            .Fields("num_j_2") = iNum_J
            .Fields("num_j_3") = iNum_J
            .Update
        End If
        .Close
    End With
    
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestJudges WHERE Code='" & cTestCode & "' AND Status=" & iTestStatus)
    With rst
        If .RecordCount = 0 Then
            For iTemp = 1 To 5
                .AddNew
                .Fields("Code") = cTestCode
                .Fields("Status") = iTestStatus
                .Fields("Position") = iTemp
                .Fields("JudgeId") = ""
                .Update
            Next iTemp
        End If
        .Close
    End With
    
      
    Set rst = Nothing
    On Local Error GoTo 0
End Function
Public Function GetSponsor(cTestCode As String) As String
    Dim rst As DAO.Recordset
    
    On Local Error Resume Next
    Set rst = mdbMain.OpenRecordset("SELECT * FROM TestInfo WHERE Code='" & cTestCode & "'")
    With rst
        If .RecordCount > 0 Then
            GetSponsor = .Fields("Sponsor") & ""
        End If
        If GetSponsor = "" Then
            Set rst = mdbMain.OpenRecordset("SELECT * FROM Tests WHERE Code='" & cTestCode & "'")
            GetSponsor = .Fields("Sponsor") & ""
        End If
        .Close
    End With
    Set rst = Nothing
    On Local Error GoTo 0
End Function

