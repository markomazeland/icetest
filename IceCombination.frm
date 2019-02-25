VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCombination 
   Caption         =   "Combinations of Tests"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   4065
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   5655
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         ToolTipText     =   "Close this window"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         ToolTipText     =   "Remove a test from this combination of tests"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         ToolTipText     =   "Add a test to this combination of tests"
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Data dtaTests 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame fraTests 
      Caption         =   "Tests"
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   5175
      Begin VB.TextBox txtTest 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   9
         ToolTipText     =   "Name of the section (preferrably in English)"
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtGroup 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtFactor 
         Height          =   285
         Index           =   0
         Left            =   4560
         TabIndex        =   10
         ToolTipText     =   "Factor (usually 1)"
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblFactor 
         Caption         =   "&Factor"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblTest 
         Caption         =   "&Test"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblGroup 
         Caption         =   "&Group"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Data dtaFinals 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame fraCombination 
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.TextBox txtName 
         DataField       =   "Combination"
         DataSource      =   "dtaCombination"
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "What is the code for this test?"
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Data dtaCombination 
      Caption         =   "Combination of Tests"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4095
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmCombination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple editor to create and edit new tests (needs more options)
'
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
Public fcCode As String
Public fcCombination As String
Dim fiIndex As Integer
Dim fiChanged As Integer

Private Sub cmdAdd_Click()
    Dim iTemp As Integer
    Dim cTemp As String
    
    For iTemp = 0 To txtTest.Count - 1
        If txtTest(iTemp).Visible = True Then
            If cTemp = "" Then
                cTemp = "'" & txtTest(iTemp).Text & "'"
            Else
                cTemp = cTemp & ",'" & txtTest(iTemp).Text & "'"
            End If
        End If
    Next iTemp
    
    With frmToolBox
        .strQry = "SELECT Code From Tests WHERE ISNULL(Removed) OR Removed=0"
        If cTemp <> "" Then
            .strQry = .strQry & " AND Code NOT IN (" & cTemp & ")"
        End If
        .strQry = .strQry & " AND Code NOT IN (SELECT Test FROM CombinationSections WHERE Code LIKE " & Chr$(34) & fcCode & Chr$(34) & ") ORDER BY Code"
        .Caption = ClipAmp(cmdAdd.Caption)
        .Show 1
    End With
    If frmMain.Tempvar <> "" Then
        For iTemp = 0 To txtTest.Count - 1
            If txtTest(iTemp).Visible = False Then
                txtGroup(iTemp).Visible = True
                txtTest(iTemp).Visible = True
                txtFactor(iTemp).Visible = True
                Exit For
            ElseIf iTemp = txtTest.Count - 1 Then
                Load txtGroup(txtGroup.Count)
                Load txtTest(txtTest.Count)
                Load txtFactor(txtFactor.Count)
                txtGroup(txtGroup.Count - 1).Visible = True
                txtTest(txtTest.Count - 1).Visible = True
                txtFactor(txtFactor.Count - 1).Visible = True
                Form_Resize
                iTemp = iTemp + 1
                Exit For
            End If
        Next iTemp
        If iTemp > 0 Then
            txtGroup(iTemp).Text = txtGroup(iTemp - 1).Text
        Else
            txtGroup(iTemp).Text = "1"
        End If
        txtFactor(iTemp).Text = "1"
        txtTest(iTemp).Text = frmMain.Tempvar
        frmMain.Tempvar = ""
        SelectRow iTemp
        fiChanged = True
    End If

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim iKey As Integer
    If fiIndex >= 0 Then
        iKey = MsgBox(txtTest(fiIndex).Text & ": " & Translate("Remove this test from this combination of tests?", mcLanguage), vbQuestion + vbYesNo)
        If iKey = vbYes Then
            txtTest(fiIndex).Visible = False
            fiChanged = True
            MoveRecord
        End If
    Else
        MsgBox Translate("Select a test to remove first!", mcLanguage)
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub dtacombination_Reposition()
    MoveRecord
End Sub

Private Sub Form_Load()
    Dim fld As DAO.Field
    Dim iTemp As Integer
    
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    DoEvents
    
    With dtaCombination
        .Caption = "Combination of Tests"
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT * FROM Combinations"
        .Refresh
    End With
    
    If fcCombination <> "" Then
        dtaCombination.Recordset.FindFirst "Combination LIKE " & Chr$(34) & fcCombination & Chr$(34)
    End If
    
    fcCode = dtaCombination.Recordset.Fields("Code")
    
    With dtaTests
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT * FROM CombinationSections WHERE Code LIKE " & Chr$(34) & fcCode & Chr$(34) & " ORDER BY Group,Test"
        .Refresh
    End With
    
    fiChanged = False
    
    MoveRecord
    Form_Resize
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    With fraCombination
        .Width = ScaleWidth
        .Height = Me.txtName.Top + txtName.Height + 50
    End With
    
    With txtName
        .Left = 50
        .Width = fraCombination.Width - 100
    End With
    
    With fraButtons
        .Height = cmdAdd.Height
        .Top = Me.StatusBar1.Top - 100 - .Height
        .Width = ScaleWidth
    End With
    
    With cmdOK
        .Top = 0
        .Left = .Container.Width - cmdOK.Width
    End With
    
    With cmdDelete
        .Top = 0
        .Left = cmdOK.Left - cmdDelete.Width - 50
    End With
    
    With cmdAdd
        .Top = 0
        .Left = cmdDelete.Left - 50 - cmdAdd.Width
    End With
    
    With fraTests
        .Width = ScaleWidth
        .Top = fraCombination.Top + fraCombination.Height + 50
        .Height = fraButtons.Top - .Top - 100
    End With
    
    With lblGroup
        .Top = 250
        .Left = 50
        .Width = (.Container.Width - 100) / 10
    End With
    With lblTest
        .Top = 250
        .Left = lblGroup.Left + lblGroup.Width + 50
        .Width = (.Container.Width - 100) * 0.8 - 200
    End With
    With lblFactor
        .Top = 250
        .Left = lblTest.Left + lblTest.Width + 50
        .Width = (.Container.Width - 100) / 10
    End With
    For iTemp = 0 To txtTest.Count - 1
        With txtGroup(iTemp)
            .Top = lblGroup.Top + lblGroup.Height + txtTest(0).Height * iTemp + 250
            .Height = txtTest(0).Height
            .Left = 50
            .Width = (.Container.Width - 100) / 10
        End With
        With txtTest(iTemp)
            .Width = (.Container.Width - 100) * 0.8 - 200
            .Top = txtGroup(iTemp).Top
            .Height = txtTest(0).Height
            .Left = txtGroup(iTemp).Left + txtGroup(iTemp).Width + 50
        End With
        With txtFactor(iTemp)
            .Width = (.Container.Width - 100) / 10
            .Top = txtGroup(iTemp).Top
            .Height = txtGroup(0).Height
            .Left = txtTest(iTemp).Left + txtTest(iTemp).Width + 50
        End With
    Next iTemp
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MoveRecord
    
    WriteFormPosition Me

End Sub

Private Sub txtFactor_Change(Index As Integer)
    fiChanged = True
End Sub

Private Sub txtFactor_Click(Index As Integer)
    SelectRow Index
End Sub

Private Sub txtGroup_Change(Index As Integer)
    fiChanged = True
End Sub

Private Sub txtGroup_Click(Index As Integer)
    SelectRow Index
End Sub

Private Sub txtTest_Change(Index As Integer)
    fiChanged = True
End Sub

Private Sub MoveRecord()
    Dim iTemp As Integer
    Dim rstCombination As DAO.Recordset
    
    On Local Error Resume Next
    
    If fiChanged = True And fraCombination.Enabled = True Then
        mdbMain.Execute "DELETE * FROM Combinationsections WHERE Code Like " & Chr$(34) & fcCode & Chr$(34)
        Set rstCombination = mdbMain.OpenRecordset("SELECT * FROM Combinationsections")
        With rstCombination
            For iTemp = 0 To txtTest.Count - 1
                If txtTest(iTemp).Visible = True Then
                    .AddNew
                    .Fields("Code") = fcCode
                    .Fields("Test") = UCase$(txtTest(iTemp).Text)
                    .Fields("Group") = Val(txtGroup(iTemp).Text)
                    .Fields("Factor") = Val(Replace(txtFactor(iTemp).Text, ",", "."))
                    .Update
                End If
            Next iTemp
        End With
        rstCombination.Close
        Set rstCombination = Nothing
    End If
    
    DoEvents
        
    Caption = dtaCombination.Recordset.Fields("Combination")
    If dtaCombination.Recordset.Fields("Userlevel") = 1 Then
        fraCombination.Enabled = True
    Else
        fraCombination.Enabled = False
        Caption = Caption & " [" & Translate("Read Only", mcLanguage) & "]"
    End If
    
    fraTests.Enabled = fraCombination.Enabled
    cmdAdd.Enabled = fraCombination.Enabled
    cmdDelete.Enabled = fraCombination.Enabled
    
    If dtaCombination.Recordset.Fields("Combination") <> "" Then
        fcCode = dtaCombination.Recordset.Fields("Code")
        With dtaTests
            .DatabaseName = mcDatabaseName
            .RecordSource = "SELECT * FROM CombinationSections WHERE Code LIKE " & Chr$(34) & fcCode & Chr$(34) & " ORDER BY Group,Test"
            .Refresh
            If .Recordset.RecordCount > 0 Then
                Do While Not .Recordset.EOF
                    If txtTest.Count < .Recordset.AbsolutePosition + 1 Then
                        Load txtGroup(txtGroup.Count)
                        Load txtTest(txtTest.Count)
                        Load txtFactor(txtFactor.Count)
                    End If
                    With txtGroup(.Recordset.AbsolutePosition)
                        .Visible = True
                        .Text = dtaTests.Recordset.Fields("Group")
                    End With
                    With txtTest(.Recordset.AbsolutePosition)
                        .Visible = True
                        .Text = dtaTests.Recordset.Fields("Test")
                    End With
                    With txtFactor(.Recordset.AbsolutePosition)
                        .Visible = True
                        .Text = dtaTests.Recordset.Fields("Factor")
                    End With
                    .Recordset.MoveNext
                Loop
            End If
            For iTemp = .Recordset.RecordCount To txtTest.Count - 1
                txtGroup(iTemp).Visible = False
                txtTest(iTemp).Visible = False
                txtFactor(iTemp).Visible = False
                DoEvents
            Next iTemp
        End With
        Form_Resize
    End If
    fiChanged = False
    fiIndex = -1
End Sub
Sub SelectRow(Index As Integer)
    Dim iTemp As Integer
    For iTemp = 0 To txtFactor.Count - 1
        If iTemp = Index Then
            fiIndex = Index
            txtFactor(iTemp).FontBold = True
            txtTest(iTemp).FontBold = True
            txtGroup(iTemp).FontBold = True
        Else
            txtFactor(iTemp).FontBold = False
            txtTest(iTemp).FontBold = False
            txtGroup(iTemp).FontBold = False
        End If
    Next iTemp
End Sub

Private Sub txtTest_Click(Index As Integer)
    SelectRow Index
End Sub
