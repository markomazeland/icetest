VERSION 5.00
Begin VB.Form frmSplitFinals 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Split Finals"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cmbSplit 
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblClass 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmSplitFinals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (C) Marko Mazeland 2006
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

Option Compare Text
Option Explicit
Public fiTakeAllClasses As Integer
Public fcTestCode As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim iTemp As Integer
    Dim rstSplit As DAO.Recordset
    Dim rstTestInfo As DAO.Recordset
    Dim rstResults As DAO.Recordset
    Dim cTemp As String
    Dim cTemp2 As String
    Dim iKey As Integer
    
    For iTemp = 0 To cmbSplit.Count - 1
        cTemp = Trim$(Left$(cmbSplit(iTemp).Text, InStr(cmbSplit(iTemp), " - ")))
        If cTemp <> fcTestCode Then
            Set rstResults = mdbMain.OpenRecordset("SELECT * FROM Results WHERE Code Like '" & cTemp & "' AND Status=0")
            If rstResults.RecordCount > 0 Then
                MsgBox Translate("Marks have been entered in the preliminary rounds of", mcLanguage) & " " & cmbSplit(iTemp).Text & "." & vbCrLf & Translate("Please select another test.", mcLanguage), vbExclamation
                rstResults.Close
                Set rstResults = Nothing
                Exit Sub
            Else
                rstResults.Close
                Set rstResults = Nothing
            End If
        End If
        
        Set rstSplit = mdbMain.OpenRecordset("SELECT * FROM TestSplits WHERE TestToSplit LIKE '" & fcTestCode & "' AND [Class] LIKE '" & lblClass(iTemp).Tag & "'")
        With rstSplit
            If .RecordCount > 0 Then
                .Edit
                If .Fields("SplitToTest") <> cTemp Then
                    cTemp2 = .Fields("SplitToTest")
                Else
                    cTemp2 = ""
                End If
            Else
                .AddNew
                .Fields("Class") = lblClass(iTemp).Tag
                .Fields("TestToSplit") = fcTestCode
            End If
            .Fields("SplitToTest") = cTemp
            .Update
            .Close
        End With
        
        If cTemp2 <> fcTestCode And cTemp2 <> "" Then
            Set rstTestInfo = mdbMain.OpenRecordset("SELECT Handling FROM TestInfo WHERE Code LIKE '" & cTemp2 & "'")
            If rstTestInfo.RecordCount > 0 Then
                With rstTestInfo
                    If .Fields(0) = 6 Then
                        .Edit
                        .Fields(0) = 5
                        .Update
                    ElseIf .Fields(0) > 2 And .Fields(0) < 5 Then
                        .Edit
                        .Fields(0) = .Fields(0) - 2
                        .Update
                    End If
                    .Close
                End With
            End If
            Set rstTestInfo = Nothing
        End If
        
        If cTemp <> fcTestCode Then
            Set rstTestInfo = mdbMain.OpenRecordset("SELECT Handling FROM TestInfo WHERE Code LIKE '" & cTemp & "'")
            If rstTestInfo.RecordCount > 0 Then
                With rstTestInfo
                    If .Fields(0) = 5 Then
                        .Edit
                        .Fields(0) = 6
                        .Update
                    ElseIf .Fields(0) <= 2 Then
                        .Edit
                        .Fields(0) = .Fields(0) + 2
                        .Update
                    ElseIf .Fields(0) < 1 Or IsNull(.Fields(0)) Then
                        .Edit
                        .Fields(0) = 4
                        .Update
                    End If
                    .Close
                End With
            End If
            Set rstTestInfo = Nothing
        End If
    Next iTemp
    
    Set rstSplit = Nothing
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rstClass As DAO.Recordset
    Dim rstTests As DAO.Recordset
    Dim rstSplit As DAO.Recordset
    Dim cQry As String
    Dim iTemp As Integer
    
    If fiTakeAllClasses = True Then
        cQry = "SELECT DISTINCT '-' AS Expr1, Participants.Class"
        cQry = cQry & " FROM Participants "
        cQry = cQry & " ORDER BY Participants.Class;"
    Else
        cQry = "SELECT Count(Participants.Sta) AS Expr1,Participants.Class"
        cQry = cQry & " FROM Participants INNER JOIN Results ON Participants.STA = Results.STA"
        cQry = cQry & " WHERE Results.Code Like '" & fcTestCode & "' AND Results.Status = 0 AND Results.Disq > -1"
        cQry = cQry & " GROUP BY Participants.Class"
        cQry = cQry & " ORDER BY Participants.Class;"
    End If
    
    Set rstClass = mdbMain.OpenRecordset(cQry)
    If rstClass.RecordCount > 0 Then
        Do While Not rstClass.EOF
            With rstClass
                If lblClass.Count < .AbsolutePosition + 1 Then
                    Load lblClass.Item(lblClass.Count)
                    Load Me.cmbSplit.Item(cmbSplit.Count)
                End If
                lblClass(.AbsolutePosition).Caption = IIf(.Fields(1) & "" = "", "-", .Fields(1)) & " (" & .Fields(0) & ")"
                lblClass(.AbsolutePosition).Tag = IIf(.Fields(1) & "" = "", "none", .Fields(1))
                .MoveNext
            End With
        Loop
    End If
    rstClass.Close
    Set rstClass = Nothing
    
    cQry = "SELECT Code,Test,Qualification FROM Tests "
    cQry = cQry & " WHERE Type_Final>0"
    cQry = cQry & " ORDER BY Code,Test,Qualification"
    Set rstTests = mdbMain.OpenRecordset(cQry)
    If rstTests.RecordCount > 0 Then
        With rstTests
            For iTemp = 0 To cmbSplit.Count - 1
                .MoveFirst
                Do While Not .EOF
                    cmbSplit(iTemp).AddItem .Fields(0) & " - " & .Fields(1) & " [" & .Fields(2) & "]"
                    .MoveNext
                Loop
            Next iTemp
        End With
    End If
    rstTests.Close
    Set rstTests = Nothing
    
    For iTemp = 0 To cmbSplit.Count - 1
        Set rstSplit = mdbMain.OpenRecordset("SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & fcTestCode & "' AND [Class] LIKE '" & lblClass(iTemp).Tag & "'")
        If rstSplit.RecordCount > 0 Then
            cmbSplit(iTemp).Text = rstSplit.Fields(0)
        Else
            cmbSplit(iTemp).Text = fcTestCode
        End If
        rstSplit.Close
    Next iTemp
    Set rstSplit = Nothing
    
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    Me.Caption = Me.Caption & ": " & fcTestCode
    
    DoEvents

End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    For iTemp = 0 To lblClass.Count - 1
        
        With lblClass(iTemp)
            .Left = 50
            .Top = (iTemp * cmbSplit(iTemp).Height) + 50
            .Width = (ScaleWidth - 200) \ 3
            .Visible = True
        End With
        
        With Me.cmbSplit(iTemp)
            .Top = lblClass(iTemp).Top
            .Width = (ScaleWidth - 200) * 2 \ 3
            .Left = lblClass(iTemp).Left + lblClass(iTemp).Width + 100
            .Visible = True
        End With
    Next iTemp
    
    With cmdCancel
        .Top = ScaleHeight - .Height - 50
        .Left = ScaleWidth - .Width - 50
    End With
    
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - .Width - 50
    End With
    
    
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFormPosition Me

End Sub
