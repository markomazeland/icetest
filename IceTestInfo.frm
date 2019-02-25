VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBList32.ocx"
Begin VB.Form frmTestInfo 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Information"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
   Begin VB.CheckBox chkJudges 
      Caption         =   "&Show judges only"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      ToolTipText     =   "Show judges previously entered only when selecting judges."
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Data dtaTestJudge 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Index           =   4
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaTestJudge 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Index           =   3
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaTestJudge 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Index           =   2
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaTestJudge 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Index           =   1
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaTestJudge 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   300
      Index           =   0
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      ToolTipText     =   "Close window"
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add Name"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      TabIndex        =   14
      ToolTipText     =   "Add a new name to the list of persons"
      Top             =   4680
      Width           =   735
   End
   Begin VB.Data dtaPersons 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo dbcJudge 
      Bindings        =   "IceTestInfo.frx":0000
      DataField       =   "JudgeId"
      DataSource      =   "dtaTestJudge(0)"
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "List of persons"
      Top             =   2280
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "PersonId"
      Text            =   ""
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5235
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSponsor 
      DataField       =   "Sponsor"
      DataSource      =   "dtaTestInfo"
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      ToolTipText     =   "Name of the sponsor(s)"
      Top             =   360
      Width           =   4575
   End
   Begin VB.Data dtaTestInfo 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBCombo dbcJudge 
      Bindings        =   "IceTestInfo.frx":0019
      DataField       =   "JudgeId"
      DataSource      =   "dtaTestJudge(1)"
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "List of persons"
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "PersonId"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo dbcJudge 
      Bindings        =   "IceTestInfo.frx":0033
      DataField       =   "JudgeId"
      DataSource      =   "dtaTestJudge(2)"
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "List of persons"
      Top             =   3240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "PersonId"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo dbcJudge 
      Bindings        =   "IceTestInfo.frx":004C
      DataField       =   "JudgeId"
      DataSource      =   "dtaTestJudge(3)"
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "List of persons"
      Top             =   3720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "PersonId"
      Text            =   ""
   End
   Begin MSDBCtls.DBCombo dbcJudge 
      Bindings        =   "IceTestInfo.frx":0065
      DataField       =   "JudgeId"
      DataSource      =   "dtaTestJudge(4)"
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "List of persons"
      Top             =   4200
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Name"
      BoundColumn     =   "PersonId"
      Text            =   ""
   End
   Begin VB.Label lblJudges 
      Caption         =   "&Judges"
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label lblJudge 
      Caption         =   "&E:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblJudge 
      Caption         =   "&D:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label lblJudge 
      Caption         =   "&C:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblJudge 
      Caption         =   "&B:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblJudge 
      Caption         =   "&A:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblSponsor 
      Caption         =   "&Sponsor:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple form to create and edit local test info
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

Private Sub chkJudges_Validate(Cancel As Boolean)
    WriteIniFile gcIniFile, Me.Name, "ShowJudgesOnly", chkJudges.Value
    SelectPersons
End Sub

Private Sub cmdNew_Click()
    Dim cId As String
    Dim rst As DAO.Recordset
    
    cId = AddNewPerson
    If cId <> "" Then
        
        Me.dtaPersons.Refresh
        
        Me.chkJudges.Value = 0
        Set rst = mdbMain.OpenRecordset("SELECT Name_First & ' ' & Name_Last as Name FROM Persons WHERE PersonId LIKE '" & cId & "'")
        If rst.RecordCount > 0 Then
            Me.dbcJudge(cmdNew.Tag) = rst.Fields("Name")
        End If
        rst.Close
        Set rst = Nothing
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub dbcJudge_GotFocus(Index As Integer)
    cmdNew.Enabled = True
    cmdNew.Tag = Index
End Sub
Private Sub Form_Load()
    Dim iTemp As Integer
    Dim cTemp As String
    Dim cQry As String
    
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    Me.Caption = frmMain.TestCode & " - " & Translate(frmMain.TestName, mcLanguage)
    
    
    If frmMain.TestSection <= 2 Then
        lblJudges.Caption = lblJudges.Caption & " - " & ClipAmp(frmMain.tbsSelFin.SelectedItem.Caption) & ": "
    Else
        lblJudges.Caption = lblJudges.Caption & ": "
    End If
    
    CreateTestInfo frmMain.TestCode
    
    With dtaTestInfo
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT * FROM TestInfo WHERE Code LIKE '" & frmMain.TestCode & "'"
        .Refresh
    End With
    
    For iTemp = 0 To 4
        With dtaTestJudge(iTemp)
            .DatabaseName = mcDatabaseName
            .RecordSource = "SELECT * FROM TestJudges WHERE Code LIKE '" & frmMain.TestCode & "' AND Status=" & frmMain.TestStatus & " AND Position=" & iTemp + 1
            .Refresh
        End With
    Next iTemp
    
    ReadIniFile gcIniFile, Me.Name, "ShowJudgesOnly", cTemp
    chkJudges.Value = Val(cTemp)
    
    SelectPersons
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.LookUpJudges
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    Dim iVisible As Integer
    
    On Local Error Resume Next
    With lblSponsor
        .Left = 50
    End With
    
    With txtSponsor
        .Top = lblSponsor.Top + lblSponsor.Height + 50
        .Left = 50
        .Width = ScaleWidth - 100
    End With
    
    With lblJudges
        .Top = txtSponsor.Top + txtSponsor.Height + 50
        .Left = 50
        .Width = (ScaleWidth - 100)
    End With
    
    For iTemp = 0 To 4
        With lblJudge(iTemp)
            .Left = 50
            If iTemp = 0 Then
                .Top = lblJudges.Top + lblJudges.Height + 50
            Else
                .Top = dbcJudge(iTemp - 1).Top + dbcJudge(iTemp - 1).Height + 50
            End If
            .Caption = Chr$(64) + 1
            If iTemp < frmMain.TestJudges Then
                .Visible = True
            Else
                .Visible = False
            End If
        End With
        With dbcJudge(iTemp)
            .Left = lblJudge(iTemp).Left + lblJudge(iTemp).Width + 50
            .Top = lblJudge(iTemp).Top
            .Width = ScaleWidth - .Left - 50
            If iTemp < frmMain.TestJudges Then
                .Visible = True
                iVisible = iTemp
            Else
                .Visible = False
            End If
        End With
    Next
        
    With cmdOK
        .Left = ScaleWidth - cmdNew.Width - 50
        .Top = dbcJudge(iVisible).Top + dbcJudge(iVisible).Height + 50
    End With
    
    With cmdNew
        .Left = cmdOK.Left - cmdNew.Width - 50
        .Top = dbcJudge(iVisible).Top + dbcJudge(iVisible).Height + 50
    End With

    With chkJudges
        .Top = cmdOK.Top
        .Left = 50
        .Width = cmdNew.Left - 100
    End With
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFormPosition Me
End Sub

Private Sub txtSponsor_Click()
    cmdNew.Enabled = False
End Sub
Private Sub SelectPersons()
    Dim cQry As String
    With dtaPersons
        .DatabaseName = mcDatabaseName
        If chkJudges.Value <> 0 Then
            '*show already selected judges
            '*
            cQry = "SELECT DISTINCT [Persons].[Name_First] & ' ' & [Persons].[Name_Last] AS Name, PersonId"
            cQry = cQry & " FROM Persons WHERE PersonId IN (SELECT JudgeId FROM TestJudges)"
            cQry = cQry & " ORDER BY [Persons].[Name_First] & ' ' & [Persons].[Name_Last]"
        Else
            'LL: Sort persons' list by field STATUS first. Setting status>0 for judges in
            'table Persons will bring them up first in list.
            cQry = "SELECT  [Persons].[Name_First] & ' ' & [Persons].[Name_Last] AS Name, PersonId"
            cQry = cQry & " FROM Persons "
            cQry = cQry & " ORDER BY [Persons].[Status] DESC, [Persons].[Name_First] & ' ' & [Persons].[Name_Last]"
        End If
        .RecordSource = cQry
        .Refresh
    End With

End Sub
