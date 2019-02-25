VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmParticipant 
   Caption         =   "Participant"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10155
   Begin VB.CommandButton cmdEnd 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      ToolTipText     =   "Last participant"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   30
      ToolTipText     =   "First participant"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   29
      ToolTipText     =   "One participant up"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      ToolTipText     =   "One particpant back"
      Top             =   7080
      Width           =   975
   End
   Begin VB.Data dtaParticipant 
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
      Top             =   7440
      Visible         =   0   'False
      Width           =   9780
   End
   Begin VB.Data dtaTests 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame fraTests 
      Caption         =   "Tests"
      Height          =   1095
      Left            =   0
      TabIndex        =   26
      Top             =   6240
      Width           =   10095
      Begin VB.ListBox lstTests 
         Height          =   255
         Left            =   360
         TabIndex        =   27
         ToolTipText     =   "Use your right mouse button to add or edit tests"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame fraParticipant 
      Height          =   1575
      Left            =   0
      TabIndex        =   13
      Top             =   5400
      Width           =   10095
      Begin VB.ComboBox cmbPart 
         DataField       =   "Class"
         DataSource      =   "dtaParticipant"
         Height          =   315
         Index           =   2
         Left            =   3360
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox cmbPart 
         DataField       =   "Team"
         DataSource      =   "dtaParticipant"
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbPart 
         DataField       =   "Club"
         DataSource      =   "dtaParticipant"
         Height          =   315
         Index           =   0
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPart 
         Caption         =   "Class:"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblPart 
         Caption         =   "Team:"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label lblPart 
         Caption         =   "Club:"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&New"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      ToolTipText     =   "Add new combination"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "Remove participant"
      Top             =   6960
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   7440
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      ToolTipText     =   "Close and apply last changes"
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame fraHorse 
      Caption         =   "Horse"
      Height          =   4935
      Left            =   5280
      TabIndex        =   11
      Top             =   360
      Width           =   4695
      Begin VB.CommandButton cmdHorseId 
         Height          =   375
         Left            =   120
         Picture         =   "IceParticipant.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Check the pedigree of this horse in WorldFengur"
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox cmbHorseTxt 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         ItemData        =   "IceParticipant.frx":03A4
         Left            =   1920
         List            =   "IceParticipant.frx":03A6
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewHorse 
         Caption         =   "&New"
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         ToolTipText     =   "Add a new horse to the list of horses"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtHorse 
         DataSource      =   "dtaHorse"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   600
         Width           =   2655
      End
      Begin MSDBCtls.DBCombo cmbHorse 
         Bindings        =   "IceParticipant.frx":03A8
         DataSource      =   "dtaHorse"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "List of horses"
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cList"
         Text            =   ""
      End
      Begin VB.Label lblHorse 
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraRider 
      Caption         =   "Rider"
      Height          =   4935
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   5175
      Begin VB.CommandButton cmdRiderId 
         Height          =   375
         Left            =   240
         Picture         =   "IceParticipant.frx":03BF
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Check the name of the rider in the FEIF WorldRanking"
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cmbRiderTxt 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   20
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewRider 
         Caption         =   "&New"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "Add a new rider to the list of riders"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtRider 
         DataSource      =   "dtaRider"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   3495
      End
      Begin MSDBCtls.DBCombo cmbRider 
         Bindings        =   "IceParticipant.frx":0763
         DataSource      =   "dtaRider"
         Height          =   315
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "List of riders"
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "cList"
         Text            =   ""
      End
      Begin VB.Label lblRider 
         Height          =   315
         Index           =   0
         Left            =   315
         TabIndex        =   18
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   1215
      End
   End
   Begin MSDBCtls.DBCombo cmbParticipants 
      Bindings        =   "IceParticipant.frx":077A
      DataSource      =   "dtaParticipant"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "clist"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin VB.Data dtaHorse 
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaRider 
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "&Remove Participant"
         Index           =   0
      End
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "&Withdraw Participant"
         Index           =   1
      End
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "&Eliminate Participant"
         Index           =   2
      End
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "&Move Participant"
         Index           =   3
      End
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "&Change Rein"
         Index           =   4
      End
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuPopupPopup 
         Caption         =   "&Add Test"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmParticipant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form to assemble riders and horses into participants

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

Public FormLoading As Integer
Public NoRevLoad As Integer
Public cNewSta As String
Public cSta As String
Public fiHorseId As Integer
Public fiPersonId As Integer
Public fiName_First As Integer
Public fiName_Last As Integer

Private Sub cmbHorseTxt_Validate(Index As Integer, Cancel As Boolean)
    cmbHorseTxt2txtHorse Index
End Sub

Private Sub cmbParticipants_Validate(Cancel As Boolean)
    SaveParticipant
End Sub

Private Sub cmbRiderTxt_Change(Index As Integer)
    If FormLoading = False Then
        If cmbRiderTxt(Index).Enabled = True Then
            cmbRiderTxt2txtRider Index
        End If
    End If
End Sub

Private Sub cmbRiderTxt_Click(Index As Integer)
    cmbRiderTxt_Change Index
End Sub

Private Sub cmbHorseTxt_Change(Index As Integer)
    If FormLoading = False Then
        If cmbHorseTxt(Index).Enabled = True Then
            cmbHorseTxt2txtHorse Index
        End If
    End If
End Sub

Private Sub cmbHorseTxt_Click(Index As Integer)
    cmbHorseTxt_Change Index
End Sub

Private Sub cmbParticipants_Change()
    Dim cQry As String
    
    On Local Error Resume Next
        
    FormLoading = True
    
    With dtaParticipant.Recordset
        .Bookmark = cmbParticipants.SelectedItem
    End With
        
    With dtaRider
        .Recordset.FindFirst "PersonID LIKE '" & dtaParticipant.Recordset.Fields("Participants.PersonId") & "'"
        cmbRider = .Recordset.Fields("cList")
        cmbRider.Refresh
    End With
    
    With dtaHorse
        .Recordset.FindFirst "HorseID LIKE '" & dtaParticipant.Recordset.Fields("Participants.HorseId") & "'"
        cmbHorse = .Recordset.Fields("cList")
        cmbHorse.Refresh
    End With
    
    lstTests.Clear
    
    cSta = Left$(cmbParticipants, 3)
    If cSta <> "" Then
        
        cQry = "SELECT Entries.Code & ' - ' & Tests.Test & ': ' & IIf(Entries.Status=3,'" & UCase$(Left$(Translate("C-Final", mcLanguage), 5)) & "',IIf(Entries.Status=2,'" & UCase$(Left$(Translate("B-Final", mcLanguage), 5)) & "',IIf(Entries.Status=1,'" & UCase$(Left$(Translate("A-Final", mcLanguage), 5)) & "','" & UCase$(Left$(Translate("Preliminary Round", mcLanguage), 5)) & "'))) & ': ' "
        cQry = cQry & " & IIf(Entries.Deleted=-2,'[" & Translate("Withdrawn", mcLanguage) & "]',IIf(Entries.Deleted=-1,'[" & UCase$(Translate("Eliminated", mcLanguage)) & "]',IIf(IsNull(Results.Score),'" & Translate("Start", mcLanguage) & ": ' & Entries.Position & IIf(Entries.RR<>0,'-R',''),'" & Translate("Result", mcLanguage) & ": ' & Format(Results.Score,'0.00') & ' (' & Results.Position & ')'))) AS cList, Entries.Sta,Entries.Deleted,Results.Score,Entries.Code,Tests.RR,Results.Disq,Entries.Status,Entries.RR"
        cQry = cQry & " FROM Tests INNER JOIN (Entries LEFT JOIN Results ON (Entries.STA = Results.STA) AND (Entries.Status = Results.Status) AND (Entries.Code = Results.Code)) ON Tests.Code = Entries.Code"
        cQry = cQry & " Where (((Entries.STA) ='" & cSta & "'))"
        cQry = cQry & " ORDER BY Entries.STA, Entries.Code, Entries.Status;"
        
        dtaTests.RecordSource = Replace(Replace(cQry, " - ", vbTab), ": ", vbTab)
        dtaTests.Refresh
        If dtaTests.Recordset.RecordCount > 0 Then
            Do While Not dtaTests.Recordset.EOF
                lstTests.AddItem dtaTests.Recordset.Fields("cList")
                lstTests.ItemData(lstTests.NewIndex) = dtaTests.Recordset.AbsolutePosition
                If dtaTests.Recordset.Fields("Deleted") = -2 Then
                    lstTests.Selected(lstTests.NewIndex) = False
                Else
                    lstTests.Selected(lstTests.NewIndex) = True
                End If
                dtaTests.Recordset.MoveNext
            Loop
            lstTests.ListIndex = -1
        End If
    End If
        
    Dim iKey As Integer
        
    If txtRider(fiPersonId).Text <> "" And ValidRiderFEIFId(txtRider(fiPersonId).Text) = False Then
        iKey = MsgBox(Translate("This is not a valid FEIFId for this rider. Request a new one (YES), remove (NO) or leave and correct it (CANCEL)?", mcLanguage), vbYesNoCancel + vbQuestion)
        If iKey = vbNo Then
            txtRider(fiPersonId).Text = ""
        ElseIf iKey = vbYes Then
            cmdRiderId_Click
        Else
            SetFocusTo txtRider(fiPersonId)
        End If
    End If
    
    DoEvents
    
    FormLoading = False
    
    On Local Error GoTo 0
    
End Sub
Private Sub cmbHorse_Change()
    
    On Local Error Resume Next
    DoEvents
    If FormLoading = False Then
        With dtaHorse.Recordset
            .Bookmark = cmbHorse.SelectedItem
        End With
    End If
    On Local Error GoTo 0
    
End Sub

Private Sub cmbRider_Change()
    On Local Error Resume Next
    DoEvents
    If FormLoading = False Then
        With dtaRider.Recordset
            .Bookmark = cmbRider.SelectedItem
        End With
    End If
    On Local Error GoTo 0
End Sub


Private Sub cmbRiderTxt_Validate(Index As Integer, Cancel As Boolean)
    cmbRiderTxt2txtRider Index
End Sub

Private Sub cmdAdd_Click()
    Dim cTemp As String
    Dim rstSta As Recordset
    
    Set rstSta = mdbMain.OpenRecordset("SELECT Sta FROM Participants ORDER BY Sta DESC")
    If rstSta.RecordCount > 0 Then
        cTemp = Format$(rstSta.Fields("Sta") + 1, "000")
    Else
        cTemp = "001"
    End If
    rstSta.Close
    Set rstSta = Nothing
    
    cTemp = InputBox$(Translate("Enter a new start number for this new Participant", mcLanguage), , cTemp)
    
    Me.FindAddParticipant cTemp
    
    cmbParticipants = dtaParticipant.Recordset.Fields("cList")
    
End Sub


Private Sub cmdStart_Click()
    If Not dtaParticipant.Recordset.BOF Then
        Me.dtaParticipant.Recordset.MoveFirst
        cmdUp.Enabled = True
        cmdEnd.Enabled = True
        cmdStart.Enabled = False
        cmdDown.Enabled = False
        If Not dtaParticipant.Recordset.BOF Then
            Me.cmbParticipants.Text = dtaParticipant.Recordset.Fields("cList")
        End If
    End If

End Sub


Private Sub cmdDown_Click()
   If Not dtaParticipant.Recordset.BOF Then
        Me.dtaParticipant.Recordset.MovePrevious
        cmdUp.Enabled = True
        cmdEnd.Enabled = True
        If Not dtaParticipant.Recordset.BOF Then
            Me.cmbParticipants.Text = dtaParticipant.Recordset.Fields("cList")
        Else
            cmdStart.Enabled = False
            cmdDown.Enabled = False
       End If
    End If
End Sub

Private Sub cmdUp_Click()
    If Not dtaParticipant.Recordset.EOF Then
        Me.dtaParticipant.Recordset.MoveNext
        cmdStart.Enabled = True
        cmdDown.Enabled = True
        If Not dtaParticipant.Recordset.EOF Then
            Me.cmbParticipants.Text = dtaParticipant.Recordset.Fields("cList")
        Else
            cmdUp.Enabled = False
            cmdEnd.Enabled = False
        End If
    End If
End Sub
Private Sub cmdEnd_Click()
    If Not dtaParticipant.Recordset.EOF Then
        Me.dtaParticipant.Recordset.MoveLast
        cmdUp.Enabled = False
        cmdEnd.Enabled = False
        cmdStart.Enabled = True
        cmdDown.Enabled = True
        If Not dtaParticipant.Recordset.EOF Then
            Me.cmbParticipants.Text = dtaParticipant.Recordset.Fields("cList")
        End If
    End If
End Sub
Private Sub cmdHorseId_Click()
    
    Dim rstHorse As DAO.Recordset
    Dim cUrl As String
    Dim cUser As String
    Dim cMsg As String
    Dim cPassword As String
    Dim cXML As String
    Dim iTemp As Integer
    Dim cTemp As String
    Dim cTemp2 As String
    Dim cTemp3 As String
    Dim cWR() As String
    Dim iKey As Integer
    Dim cSlash As String
    Dim cOrigin As String
    Dim cWR_Code As String
    Dim rstTemp As Recordset
    
    On Local Error Resume Next
    
    SetFocusTo txtHorse(0)
    
    cWR_Code = GetVariable("WR_Code")
    cUrl = GetVariable("WR_Url_Horse")
    
    If ValidWRCode(cWR_Code) = False Then
        cMsg = Translate("This button is used to check the pedigree of the horse on line in the FEIF WorldRanking database (FEIF WorldRanking events only).", mcLanguage)
        MsgBox cMsg, vbInformation
        Exit Sub
    ElseIf cUrl = "" Then
        cMsg = Translate("This button is used to check the pedigree of the horse on line in the FEIF WorldRanking database (FEIF WorldRanking events only).", mcLanguage)
        cMsg = cMsg & vbCrLf & Translate("An update of Sport Rules is required to enable this function.", mcLanguage)
        MsgBox cMsg, vbInformation
        Exit Sub
    End If
    
    cUrl = cUrl & "&wr=" & cWR_Code
    cUrl = cUrl & "&feifid="
    
    Set rstHorse = mdbMain.OpenRecordset("SELECT Name_Horse, FEIFID FROM Horses WHERE FEIFID LIKE '" & txtHorse(fiHorseId).Text & "'")
    If rstHorse.RecordCount = 0 Then
        iKey = MsgBox(Translate("Add a new horse to IceTest", mcLanguage) & "?", vbYesNo + vbQuestion + vbDefaultButton1)
        If iKey <> vbYes Then
            Exit Sub
        End If
    Else
        If cmdNewHorse.Tag <> "" Then
            iKey = vbYes
        Else
            iKey = MsgBox(Translate("Replace current data by data from FEIF WorldRanking (Internet connection required)", mcLanguage) & "?", vbYesNo + vbQuestion + vbDefaultButton1)
        End If
        If iKey <> vbYes Then
            cmdNewHorse.Tag = ""
            Me.Enabled = True
            Me.cmdHorseId.Enabled = True
            Exit Sub
        Else
            cmbHorse.Text = rstHorse.Fields("Name_Horse") & ""
            cmbHorse_Change
        End If
    End If
    rstHorse.Close

    StatusMessage Translate("Requesting data from FEIF WorldRanking", mcLanguage)
    
    Me.MousePointer = vbHourglass
    SetMouseHourGlass
    
    cTemp = UTF8_Encode(txtHorse(fiHorseId).Text)
    cXML = RequestHorseXML(cUrl, cTemp)

    If cXML = "" Then
        cmdNewHorse.Tag = ""
        Me.Enabled = True
        Me.cmdHorseId.Enabled = True
        MsgBox Translate("Service not available, check your connection to Internet.", mcLanguage), vbCritical
        StatusMessage ""
        Me.MousePointer = vbNormal
        SetMouseNormal
        Exit Sub
    ElseIf InStr(cXML, "Username or password incorrect") Then
        cmdNewHorse.Tag = ""
        Me.Enabled = True
        Me.cmdHorseId.Enabled = True
        MsgBox Translate("Username and/or password incorrect.", mcLanguage), vbCritical
        StatusMessage ""
        Me.MousePointer = vbNormal
        SetMouseNormal
        Exit Sub
    End If
    
    ReDim cWR(0 To txtHorse.Count - 1)
    cWR(0) = txtHorse(fiHorseId).Text
    For iTemp = 0 To txtHorse.Count - 1
        cTemp = ""
        Select Case txtHorse(iTemp).DataField
        Case "Source"
            cTemp = " WorldFengur " & Format$(Now, "dd-mmm-yyyy hh:mm:ss")
        Case "Name_Horse"
           cTemp = XmlParse(cXML, "Name_Horse", "", True)
        Case "Birthday_horse"
            cTemp = XmlParse(cXML, "Birthday_Horse", "")
            If cTemp <> "" Then
                cTemp = "01-01-" & cTemp
            End If
        Case "Country_Horse"
            cTemp = Left(txtHorse(fiHorseId).Text, 2)
        Case Else
           cTemp = XmlParse(cXML, txtHorse(iTemp).DataField, "")
        End Select
        If cTemp <> "" Then
            cWR(iTemp) = Trim(UTF8_Decode(cTemp))
        End If
    Next iTemp
    
    cMsg = ""
    For iTemp = 0 To txtHorse.Count - 1
        cMsg = cMsg & lblHorse(iTemp).Caption & ": " & vbTab & vbTab & cWR(iTemp) & vbCrLf
    Next iTemp
        
    StatusMessage ""
    cmdNewHorse.Tag = ""
    Me.MousePointer = vbNormal
    SetMouseNormal

    iKey = MsgBox(cMsg, vbYesNo + vbQuestion + vbDefaultButton1, "Add to IceTest?")
    If iKey = vbYes Then
        Set rstHorse = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE Horseid LIKE '" & dtaHorse.Recordset.Fields("HorseId") & "'")
        If rstHorse.RecordCount = 0 Then
            rstHorse.AddNew
            rstHorse.Fields("Horseid") = dtaHorse.Recordset.Fields("HorseId")
        Else
            rstHorse.Edit
        End If
        
        For iTemp = 0 To txtHorse.Count - 1
            If cWR(iTemp) <> "" Then
                txtHorse(iTemp).Text = cWR(iTemp)
                If cmbHorseTxt(iTemp).Visible = True Then
                    Set rstTemp = mdbMain.OpenRecordset("SELECT [Code] FROM [Values] WHERE [Field] LIKE '" & txtHorse(iTemp).DataField & "' AND [LABEL] LIKE '" & cWR(iTemp) & "'")
                    If rstTemp.RecordCount > 0 Then
                        rstHorse.Fields(txtHorse(iTemp).DataField) = rstTemp.Fields(0)
                    Else
                        rstHorse.Fields(txtHorse(iTemp).DataField) = cWR(iTemp)
                    End If
                    rstTemp.Close
                Else
                    rstHorse.Fields(txtHorse(iTemp).DataField) = cWR(iTemp)
                End If
            End If
        Next iTemp
        rstHorse.Update
        rstHorse.Close
        
        Set rstTemp = Nothing
    End If
    
    Set rstHorse = Nothing
    
    Me.Enabled = True
    Me.cmdHorseId.Enabled = True
    
    On Local Error GoTo 0
    
End Sub

Private Sub cmdNewHorse_Click()
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim cTemp As String
    Dim cId As String
    Dim iCounter As Integer
    
    Dim rstHorses As DAO.Recordset
    
    If cmdNewHorse.Tag <> "" Then
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Add a new horse to the list of horses based upon its FEIF ID", mcLanguage) & "?", vbYesNoCancel + vbQuestion + vbDefaultButton1)
    End If
    If iKey = vbYes Then
        SetFocusTo txtHorse(0)
        
        Do
            If cTemp <> "" Then
                cTemp = InputBox$(Translate("Enter a valid FEIF Id.", mcLanguage), "", cTemp)
            Else
                cTemp = InputBox$(Translate("Enter the FEIF Id of the horse.", mcLanguage), "", cmdNewHorse.Tag)
            End If
        Loop While cTemp <> "" And cTemp <> Chr$(27) And ValidHorseFEIFId(cTemp) = False
        If cTemp <> "" And cTemp <> Chr$(27) Then
            
            Set rstHorses = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE FEIFID LIKE " & Chr$(34) & cTemp & Chr$(34))
            If rstHorses.RecordCount > 0 Then
                MsgBox cTemp & ": " & Translate("FEIF ID already in use in IceTest!", mcLanguage)
            Else
                cId = CreateHorseId
                
                With rstHorses
                    .AddNew
                    .Fields("HorseId") = cId
                    .Fields("FEIFID") = cTemp
                    .Update
                End With
                With dtaHorse
                    .Refresh
                    .Recordset.FindFirst "HorseID LIKE '" & cId & "'"
                    cmbHorse = .Recordset.Fields("cList")
                End With
            End If
            rstHorses.Close
            Set rstHorses = Nothing
            If cmdHorseId.Enabled = True Then
                cmdHorseId_Click
            End If
        End If
    ElseIf iKey = vbNo Then
        cTemp = InputBox$(Translate("Enter the complete name of the horse.", mcLanguage))
        If cTemp <> "" And cTemp <> Chr$(27) Then
            Set rstHorses = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE Name_Horse LIKE " & Chr$(34) & cTemp & Chr$(34))
            If rstHorses.RecordCount > 0 Then
                MsgBox cTemp & ": " & Translate("Name already exists!", mcLanguage)
            Else
                cId = CreateHorseId
                
                With rstHorses
                    .AddNew
                    .Fields("HorseId") = cId
                    .Fields("Name_Horse") = cTemp
                    .Update
                End With
                With dtaHorse
                    .Refresh
                    .Recordset.FindFirst "HorseID LIKE '" & cId & "'"
                    cmbHorse = .Recordset.Fields("cList")
                End With
            End If
            rstHorses.Close
            Set rstHorses = Nothing
        End If
    End If
End Sub

Private Sub cmdNewRider_Click()
    Dim iKey As Integer
    Dim cId As String
    
    iKey = MsgBox(Translate("Add a new rider to the list of riders", mcLanguage) & "?", vbYesNo + vbQuestion + vbDefaultButton2)
    If iKey = vbYes Then
        cId = AddNewPerson
        If cId <> "" Then
            With dtaRider
                .Refresh
                .Recordset.FindFirst "PersonID LIKE '" & cId & "'"
                cmbRider = .Recordset.Fields("cList")
            End With
        End If
    End If
    
End Sub

Private Sub SaveParticipant()
    Dim rstParticipant As Recordset
    If dtaParticipant.Recordset.RecordCount > 0 Then
        cSta = dtaParticipant.Recordset.Fields("Sta") & ""
        If cSta <> "" Then
            Set rstParticipant = mdbMain.OpenRecordset("SELECT * FROM Participants WHERE Sta LIKE '" & cSta & "'")
            rstParticipant.MoveLast
            Do While rstParticipant.RecordCount > 1
                rstParticipant.Delete
                rstParticipant.Requery
                rstParticipant.MoveLast
            Loop
            With rstParticipant
                If .RecordCount > 0 Then
                    .Edit
                Else
                    .AddNew
                    .Fields("STA") = cSta
                End If
                .Fields("PersonId") = dtaRider.Recordset.Fields("PersonId")
                .Fields("HorseId") = dtaHorse.Recordset.Fields("HorseId")
                .Fields("Club") = cmbPart(0).Text
                .Fields("Team") = cmbPart(1).Text
                .Fields("Class") = cmbPart(2).Text
                .Fields("Flag") = False
                .Update
            End With
            rstParticipant.Close
            Set rstParticipant = Nothing
        
            dtaParticipant.Refresh
            dtaParticipant.Recordset.FindFirst "Sta LIKE '" & cSta & "'"
            cmbParticipants = dtaParticipant.Recordset.Fields("cList")
        End If
    End If
End Sub

Private Sub cmdOK_Click()

    cmdOK.Enabled = False
    
    
    Unload Me

End Sub

Private Sub cmdRemove_Click()
    Dim iKey As Integer
    Dim cList As String
    
    iKey = MsgBox(Translate("Remove", mcLanguage) & " '" & dtaParticipant.Recordset.Fields("cList") & "'?", vbYesNo + vbQuestion + vbDefaultButton2)
    If iKey = vbYes Then
        cSta = dtaParticipant.Recordset.Fields("Sta")
        iKey = MsgBox("'" & dtaParticipant.Recordset.Fields("cList") & "' " & Translate("will be removed from the whole event, including all results (no way back)!", mcLanguage), vbYesNo + vbExclamation + vbDefaultButton2)
        If iKey = vbYes Then
            mdbMain.Execute ("DELETE * FROM Entries WHERE STA = '" & cSta & "'")
            mdbMain.Execute ("DELETE * FROM Results WHERE STA = '" & cSta & "'")
            mdbMain.Execute ("DELETE * FROM Marks WHERE STA = '" & cSta & "'")
            mdbMain.Execute ("DELETE * FROM Participants WHERE STA = '" & cSta & "'")
            dtaParticipant.Refresh
            cmbParticipants = dtaParticipant.Recordset.Fields("cList")
        End If
    End If
End Sub

Private Sub cmdRiderId_Click()
    
    Dim rstRider As DAO.Recordset
    Dim cUrl As String
    Dim cMsg As String
    Dim cXML As String
    Dim cRider As String
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim cTemp As String
    Dim cTemp2 As String
    Dim cTemp3 As String
    Dim cWR() As String
    Dim iKey As Integer
    Dim cSlash As String
    Dim cOrigin As String
    Dim cQry As String
    Dim cFEIFID As String
    Dim cFirst As String
    Dim cLast As String
    Dim cField() As String
    Dim cWR_Code As String
    Dim iCheckOnRiderId As Integer
    
    On Local Error Resume Next
    
    SetFocusTo txtRider(0)
    
    cWR_Code = GetVariable("WR_Code")
    cUrl = GetVariable("WR_Url_Rider")
    If cUrl = "" Then
        cUrl = GetVariable("WR_Url")
        If cUrl <> "" Then
            SetVariable "WR_Url_Rider", cUrl
        End If
    End If
    
    If ValidWRCode(cWR_Code) = False Then
        cMsg = Translate("This button is used to check the name of rides and their FEIF RiderId on line in the FEIF WorldRanking database (FEIF WorldRanking events only).", mcLanguage)
        MsgBox cMsg, vbInformation
        Exit Sub
    ElseIf cUrl = "" Then
        cMsg = Translate("This button is used to check the name of riders and their FEIF RiderId on line (FEIF WorldRanking).", mcLanguage)
        cMsg = cMsg & vbCrLf & Translate("An update of Sport Rules is required to enable this function.", mcLanguage)
        MsgBox cMsg, vbInformation
        Exit Sub
    End If
    
    If cWR_Code <> "" Then
        cUrl = cUrl & "&wr=" & cWR_Code
    End If
    cUrl = cUrl & "&search="
    
    If ValidRiderFEIFId(txtRider(fiPersonId).Text) = True Then
        cRider = txtRider(fiPersonId).Text
        cQry = "SELECT Persons.Name_First & ' ' & Persons.Name_Last AS cList,*"
        cQry = cQry & " FROM Persons WHERE Persons.FEIFId LIKE '" & cRider & "'"
        iCheckOnRiderId = 1
    Else
        cRider = txtRider(fiName_First).Text & " " & txtRider(fiName_Last).Text
        cRider = Trim$(Replace(cRider, "  ", " "))
        cQry = "SELECT Persons.Name_First & ' ' & Persons.Name_Last AS cList,*"
        cQry = cQry & " FROM Persons WHERE Persons.Name_First & ' ' & Persons.Name_Last LIKE '" & cRider & "'"
        iCheckOnRiderId = 0
    End If
    
    If Len(cRider) <= 2 Then
        MsgBox "'" & cRider & "' " & Translate("is not enough information to perform a check.", mcLanguage), vbExclamation
        Exit Sub
    End If
    
    Set rstRider = mdbMain.OpenRecordset(cQry)
    
    If rstRider.RecordCount = 0 Then
        If iCheckOnRiderId = 1 Then
        Else
            iKey = MsgBox(Replace(Translate("Check '%s' in the FEIF WorldRanking (YES) or add '%s' to IceTest without check (NO)", mcLanguage) & "?", "%s", cRider) & vbCrLf & Translate("(Internet connection required.)", mcLanguage), vbYesNoCancel + vbQuestion + vbDefaultButton1)
        End If
        If iKey = vbYes Then
            cmbRider.Text = rstRider.Fields("cList") & ""
            cmbRider_Change
        ElseIf iKey <> vbNo Then
            Exit Sub
        End If
    Else
        If cmdNewRider.Tag <> "" Or miConnectedToInternet = True Then
            iKey = vbYes
        Else
            iKey = MsgBox(Replace(Translate("Check '%s' in the FEIF WorldRanking", mcLanguage) & "?", "%s", cRider) & vbCrLf & Translate("(Internet connection required.)", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton1)
        End If
        If iKey <> vbYes Then
            cmdNewRider.Tag = ""
            Me.Enabled = True
            Me.cmdRiderId.Enabled = True
            Exit Sub
        Else
            cmbRider.Text = rstRider.Fields("cList") & ""
            cmbRider_Change
        End If
    End If
    rstRider.Close
    Set rstRider = Nothing
    
    StatusMessage Translate("Requesting data from FEIF WorldRanking", mcLanguage)
    
    Me.MousePointer = vbHourglass
    SetMouseHourGlass
    
    cTemp = UTF8_Encode(Replace(Replace(cRider, ".", ""), " ", "+"))
    cXML = RequestRiderXML(cUrl, cTemp)
    
    If cXML = "" Then
        StatusMessage ""
        cmdNewRider.Tag = ""
        Me.MousePointer = vbNormal
        SetMouseNormal
        Exit Sub
    End If
    
    ReDim cWR(0 To txtRider.Count - 1)
    For iTemp = 0 To txtRider.Count - 1
        cTemp = ""
        Select Case txtRider(iTemp).DataField
        Case "Source"
            cTemp = "FEIF WorldRanking " & Format$(Now, "dd-mmm-yyyy hh:mm:ss")
        Case "Name_First"
           cTemp = XmlParse(cXML, "Rider", "", True)
           iTemp2 = InStr(cTemp, " ")
           If iTemp2 > 0 Then
              cTemp = Trim$(Left$(cTemp, iTemp2))
           End If
        Case "Name_Last"
           cTemp = XmlParse(cXML, "Rider", "", True)
           iTemp2 = InStr(cTemp, " ")
           If iTemp2 > 0 Then
            cTemp = Trim$(Mid$(cTemp, iTemp2))
           Else
            cTemp = ""
           End If
        Case "Nationality"
           cTemp = XmlParse(cXML, "SPORTNATIONALITY", "", True)
        Case Else
           cTemp = XmlParse(cXML, txtHorse(iTemp).DataField, "")
        End Select
        If cTemp <> "" Then
            cWR(iTemp) = UTF8_Decode(cTemp)
            Select Case iTemp
            Case fiPersonId
                cFEIFID = cWR(iTemp)
            Case fiName_First
                cFirst = cWR(iTemp)
            Case fiName_Last
                cLast = cWR(iTemp)
            End Select
        End If
    Next iTemp
    
    cMsg = ""
    For iTemp = 0 To txtRider.Count - 1
        If cWR(iTemp) <> "" Then
            cMsg = cMsg & Left$(lblRider(iTemp).Caption & ":" & Space$(20), 20) & vbTab & cWR(iTemp) & vbCrLf
        End If
    Next iTemp
        
    StatusMessage ""
    cmdNewRider.Tag = ""
    Me.MousePointer = vbNormal
    SetMouseNormal
                    
    If cMsg <> "" Then
        If cFirst & " " & cLast = cRider And cFEIFID = txtRider(fiPersonId).Text Then
            iKey = MsgBox(Translate("Name of rider and FEIFId are correct.", mcLanguage))
        Else
            iKey = MsgBox(cMsg, vbYesNo + vbQuestion + vbDefaultButton1, "Add to IceTest?")
        End If
        If iKey = vbYes Then
            If cWR(0) <> "" Then
                For iTemp = 0 To txtRider.Count - 1
                    If cWR(iTemp) <> "" Then
                        txtRider(iTemp).Text = cWR(iTemp)
                    End If
                Next iTemp
            End If
        End If
    Else
        MsgBox "'" & cRider & "' " & Translate("not found in FEIF WorldRanking.", mcLanguage), vbExclamation
    End If
    
    Me.cmbRider.SelectedItem = cRider
    
    Me.Enabled = True
    Me.cmdRiderId.Enabled = True
    
    On Local Error GoTo 0

End Sub

Private Sub Form_Load()

    Dim cQry As String
    Dim cTmpSta As String
    Dim rstTeam As DAO.Recordset
    Dim rstClub As DAO.Recordset
    Dim rstClass As DAO.Recordset
    Dim rst As DAO.Recordset
    Dim rst2 As DAO.Recordset
    Dim iTemp As Integer
    
    FormLoading = True
    
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    
    TranslateControls Me
    
    Me.Caption = StrConv(Me.Caption, vbProperCase)
       
    Me.dtaParticipant.DatabaseName = mcDatabaseName
    Me.dtaRider.DatabaseName = mcDatabaseName
    Me.dtaHorse.DatabaseName = mcDatabaseName
    Me.dtaTests.DatabaseName = mcDatabaseName
    
    cQry = "SELECT Participants.Sta "
    cQry = cQry & " & '  -  ' & Persons.Name_First"
    cQry = cQry & " & ' ' & Persons.Name_Last"
    cQry = cQry & " & ' - ' & Horses.Name_Horse as cList,*"
    cQry = cQry & " FROM (Participants"
    cQry = cQry & " INNER JOIN Persons"
    cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
    cQry = cQry & " INNER JOIN Horses"
    cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
    cQry = cQry & " ORDER BY Participants.Sta"
   
    Me.dtaParticipant.RecordSource = cQry
    Me.dtaParticipant.Refresh
    
    If cNewSta <> "" Or cSta <> "" Then
        If cNewSta <> "" Then
            Me.cmbParticipants.Visible = False
            Me.Caption = Translate("New Participant", mcLanguage) & " " & cNewSta
            cTmpSta = Format$(Val(cNewSta), "000")
        Else
            cTmpSta = Format$(Val(cSta), "000")
        End If
        
        FindAddParticipant cTmpSta
        
        Me.cmbParticipants.Locked = False
        
    End If
    
    cQry = "SELECT Persons.Name_First & ' ' & Persons.Name_Middle & ' ' & Persons.Name_Last AS cList,*"
    cQry = cQry & " FROM Persons ORDER BY Persons.Name_First,Persons.Name_Middle,Persons.Name_Last"
    
    dtaRider.RecordSource = cQry
    dtaRider.Refresh
    
    If dtaRider.Recordset.RecordCount > 0 Then
        cmbRider.Text = dtaRider.Recordset.Fields("cList")
    Else
        Unload Me
    End If
        
        
    If cTmpSta <> "" Then
        cQry = "SELECT Entries.Code & '-' & Tests.Test & ': ' & IIf(Entries.Status=3,'C-FIN',IIf(Entries.Status=2,'B-FIN',IIf(Entries.Status=1,'A-FIN','PREL'))) & ': ' & IIf(IsNull(Results.Score),'Start: ' & Entries.Position & IIf(Entries.RR<>0,'R',''),'Score: ' & Format(Results.Score,'Fixed') & ' (' & Results.Position & ')') & IIf(Entries.Deleted=-2,' [Withdrawn]',IIf(Entries.Deleted=-1,' [ELIMINATED]')) AS cList"
        cQry = cQry & " FROM Tests INNER JOIN (Entries LEFT JOIN Results ON (Entries.STA = Results.STA) AND (Entries.Status = Results.Status) AND (Entries.Code = Results.Code)) ON Tests.Code = Entries.Code"
        cQry = cQry & " Where (((Entries.STA) ='" & cTmpSta & "'))"
        cQry = cQry & " ORDER BY Entries.STA, Entries.Code, Entries.Status;"
        dtaTests.RecordSource = cQry
        dtaTests.Refresh
    End If
    
    On Local Error Resume Next
    
    Set rst = mdbMain.OpenRecordset("SELECT [Label],[Field],[Comment],[Type],[Status] FROM [Fields] WHERE Table='Persons' AND Status>0 ORDER BY Seq")
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            With rst
                If lblRider.Count < .AbsolutePosition + 1 Then
                    Load lblRider(lblRider.Count)
                    Load txtRider(txtRider.Count)
                    Load cmbRiderTxt(cmbRiderTxt.Count)
                End If
                With lblRider(rst.AbsolutePosition)
                    .Caption = rst.Fields(0) & ""
                    .ToolTipText = rst.Fields(2) & ""
                    .Visible = True
                End With
                With txtRider(rst.AbsolutePosition)
                    .DataField = rst.Fields(1)
                    If .DataField = "FEIFID" Then
                        cmdRiderId.Visible = True
                    End If
                    .Locked = False
                    .ToolTipText = rst.Fields(2) & ""
                    .Tag = rst.Fields(3) & ""
                    If rst.Fields(4) = 3 Then
                        .Visible = True
                        .Locked = True
                    ElseIf rst.Fields(4) = 5 Or rst.Fields(4) = 6 Then
                        .Visible = False
                    Else
                        .Visible = True
                    End If
                End With
                With cmbRiderTxt(rst.AbsolutePosition)
                    If rst.Fields(4) = 5 Or rst.Fields(4) = 6 Then
                        .ToolTipText = rst.Fields(3) & ""
                        .Visible = True
                        .Enabled = True
                        If TableExist(mdbMain, "Values") = True Then
                            Set rst2 = mdbMain.OpenRecordset("SELECT [Label],[ValueId] FROM [Values] WHERE [Field] LIKE '" & rst.Fields(1) & "'")
                            If rst2.RecordCount > 0 Then
                                Do While Not rst2.EOF
                                    .AddItem Translate(rst2.Fields(0) & "", mcLanguage)
                                    .ItemData(cmbRiderTxt(rst.AbsolutePosition).NewIndex) = rst2.Fields(1)
                                    rst2.MoveNext
                                Loop
                            Else
                                Set rst2 = mdbMain.OpenRecordset("SELECT DISTINCT " & rst.Fields(1) & " FROM Persons")
                                If rst2.RecordCount > 0 Then
                                    Do While Not rst2.EOF
                                        .AddItem Translate(rst2.Fields(0) & "", mcLanguage)
                                        .ItemData(cmbRiderTxt(rst.AbsolutePosition).NewIndex) = 0
                                        rst2.MoveNext
                                    Loop
                                End If
                            End If
                        Else
                            Set rst2 = mdbMain.OpenRecordset("SELECT DISTINCT " & rst.Fields(1) & " FROM Persons")
                            If rst2.RecordCount > 0 Then
                                Do While Not rst2.EOF
                                    .AddItem Translate(rst2.Fields(0) & "", mcLanguage)
                                    .ItemData(cmbRiderTxt(rst.AbsolutePosition).NewIndex) = 0
                                    rst2.MoveNext
                                Loop
                            End If
                        End If
                    Else
                        .Enabled = False
                        .Visible = False
                    End If
                End With
                    
                .MoveNext
            End With
        Loop
    End If
    
    cQry = "SELECT Name_Horse & '' as cList,*"
    cQry = cQry & " FROM Horses ORDER BY Name_Horse"
    
    dtaHorse.RecordSource = cQry
    dtaHorse.Refresh
    
    If dtaHorse.Recordset.RecordCount > 0 Then
        cmbHorse.Text = dtaHorse.Recordset.Fields("cList")
    End If
    
    Set rst = mdbMain.OpenRecordset("SELECT [Label],[Field],[Comment],[Type],[Status] FROM [Fields] WHERE Table='Horses' AND Status>0 ORDER BY Seq")
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            With rst
                If lblHorse.Count < .AbsolutePosition + 1 Then
                    Load lblHorse(lblHorse.Count)
                    Load txtHorse(txtHorse.Count)
                    Load cmbHorseTxt(cmbHorseTxt.Count)
                End If
                With lblHorse(.AbsolutePosition)
                    .Caption = rst.Fields(0)
                    .ToolTipText = rst.Fields(2) & ""
                    .Visible = True
                End With
                With txtHorse(.AbsolutePosition)
                    .DataField = rst.Fields(1)
                    If .DataField = "FEIFID" Then
                        cmdHorseId.Visible = True ' No more waiting for WF ...
                        cmdHorseId.Enabled = True
                    End If
                    .ToolTipText = rst.Fields(2) & ""
                    .Tag = rst.Fields(3)
                    If rst.Fields(4) = 5 Or rst.Fields(4) = 6 Then
                        .Visible = False
                    Else
                        .Visible = True
                    End If
                    If rst.Fields(4) = 4 Then
                        .Locked = True
                    End If
                End With
                
                With cmbHorseTxt(.AbsolutePosition)
                    If rst.Fields(4) = 5 Or rst.Fields(4) = 6 Then
                        .ToolTipText = rst.Fields(2) & ""
                        .Visible = True
                        .Enabled = True
                        DoEvents
                        If TableExist(mdbMain, "Values") = True Then
                            Set rst2 = mdbMain.OpenRecordset("SELECT [Label],[ValueId] FROM [Values] WHERE [Field] LIKE '" & rst.Fields(1) & "'")
                            If rst2.RecordCount > 0 Then
                                Do While Not rst2.EOF
                                    .AddItem Translate(rst2.Fields(0) & "", mcLanguage)
                                    .ItemData(cmbHorseTxt(rst.AbsolutePosition).NewIndex) = rst2.Fields(1)
                                    rst2.MoveNext
                                Loop
                            Else
                                Set rst2 = mdbMain.OpenRecordset("SELECT DISTINCT " & rst.Fields(1) & " FROM Horses")
                                If rst2.RecordCount > 0 Then
                                    Do While Not rst2.EOF
                                        .AddItem Translate(rst2.Fields(0) & "", mcLanguage)
                                        .ItemData(cmbHorseTxt(rst.AbsolutePosition).NewIndex) = 0
                                        rst2.MoveNext
                                    Loop
                                End If
                            End If
                        Else
                            Set rst2 = mdbMain.OpenRecordset("SELECT DISTINCT " & rst.Fields(1) & " FROM Horses")
                            If rst2.RecordCount > 0 Then
                                Do While Not rst2.EOF
                                    .AddItem Translate(rst2.Fields(0) & "", mcLanguage)
                                    .ItemData(cmbHorseTxt(rst.AbsolutePosition).NewIndex) = 0
                                    rst2.MoveNext
                                Loop
                            End If
                        End If
                    Else
                        .Enabled = False
                        .Visible = False
                    End If
                End With
                .MoveNext
            End With
        Loop
    End If
    rst2.Close
    Set rst2 = Nothing
    
    
    rst.Close
    Set rst = Nothing
    
    Set rstClub = mdbMain.OpenRecordset("SELECT DISTINCT Club FROM Participants ORDER BY Club")
    If rstClub.RecordCount > 0 Then
        cmbPart(0).Clear
        Do While Not rstClub.EOF
            cmbPart(0).AddItem rstClub.Fields(0) & ""
            rstClub.MoveNext
        Loop
    End If
    rstClub.Close
    Set rstClub = Nothing
    
    Set rstTeam = mdbMain.OpenRecordset("SELECT DISTINCT Team FROM Participants ORDER BY Team")
    If rstTeam.RecordCount > 0 Then
        cmbPart(1).Clear
        Do While Not rstTeam.EOF
            cmbPart(1).AddItem rstTeam.Fields(0) & ""
            rstTeam.MoveNext
        Loop
    End If
    rstTeam.Close
    Set rstTeam = Nothing
    
    Set rstClass = mdbMain.OpenRecordset("SELECT DISTINCT Class FROM Participants ORDER BY Class")
    If rstClass.RecordCount > 0 Then
        cmbPart(2).Clear
        Do While Not rstClass.EOF
            cmbPart(2).AddItem rstClass.Fields(0) & ""
            rstClass.MoveNext
        Loop
    End If
    rstClass.Close
    Set rstClass = Nothing
    
    DoEvents
    
    If dtaParticipant.Recordset.RecordCount > 0 Then
        Me.cmbParticipants = dtaParticipant.Recordset.Fields("cList")
    End If
    
    For iTemp = 0 To cmbRiderTxt.Count - 1
        If cmbRiderTxt(iTemp).Enabled = True Then
            txtRider2cmbRiderTxt iTemp
        End If
    Next iTemp
    
    For iTemp = 0 To cmbHorseTxt.Count - 1
        If cmbHorseTxt(iTemp).Enabled = True Then
            txtHorse2cmbHorseTxt iTemp
        End If
    Next iTemp
    
    Form_Resize
    
    cmbParticipants.Enabled = True
    
    FormLoading = False
    
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveParticipant
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    Static iBusy As Integer
    Dim iTabSize As Integer
    
    On Local Error Resume Next
    If iBusy = False Then
        iBusy = True
        
        With cmbParticipants
            .Width = ScaleWidth - 100
            .Left = 50
            .Top = 50
        End With
        
        With cmdOK
            .Top = StatusBar1.Top - .Height - 50
            .Left = ScaleWidth - .Width - 50
        End With
        
        With cmdAdd
            .Top = cmdOK.Top
            .Left = 50
        End With
        
        With cmdRemove
            .Top = cmdOK.Top
            .Left = cmdAdd.Left + cmdAdd.Width + 50
        End With
        
        With cmdStart
            .Top = cmdOK.Top
            .Left = cmdRemove.Left + cmdRemove.Width + 50
            .Width = cmdRemove.Width \ 3
        End With
        
        With cmdDown
            .Top = cmdOK.Top
            .Left = cmdStart.Left + cmdStart.Width + 50
            .Width = cmdStart.Width
        End With
        
        With cmdUp
            .Top = cmdOK.Top
            .Left = cmdDown.Left + cmdDown.Width + 50
            .Width = cmdStart.Width
        End With
        
        With cmdEnd
            .Top = cmdOK.Top
            .Left = cmdUp.Left + cmdUp.Width + 50
             .Width = cmdStart.Width
       End With
        
        With Me.fraTests
            .Left = 50
            .Width = ScaleWidth - 100
            .Height = ScaleHeight \ 10
            .Top = ScaleHeight - .Height - cmdOK.Height - Me.StatusBar1.Height - 100
        End With
        
        With fraParticipant
            .Left = 50
            .Width = ScaleWidth - 100
            .Height = (cmbPart(0).Height + 50) * cmbPart.Count + 300
            .Top = fraTests.Top - .Height - 50
        End With
        
        With fraRider
            .Left = 50
            .Width = ScaleWidth \ 2 - 100
            .Top = cmbParticipants.Top + cmbParticipants.Height + 50
            .Height = fraParticipant.Top - .Top - 50
        End With
        
        With fraHorse
            .Left = fraRider.Left + fraRider.Width + 100
            .Width = ScaleWidth \ 2 - 100
            .Height = fraRider.Height
            .Top = fraRider.Top
        End With
        
        For iTemp = 0 To lblPart.Count - 1
            With lblPart(iTemp)
                If iTemp > 0 Then
                    .Top = lblPart(iTemp - 1).Top + cmbPart(iTemp - 1).Height + 50
                End If
                .Container = fraParticipant
                .Left = 50
                .Width = .Container.Width \ 4
            End With
            With cmbPart(iTemp)
                .Container = fraParticipant
                .Top = lblPart(iTemp).Top
                .Left = lblPart(iTemp).Left + lblPart(iTemp).Width + 50
                .Width = .Container.Width - .Left - 50
                .ToolTipText = lblPart(iTemp).Caption
            End With
        Next iTemp
        
        With cmbRider
            .Top = 300
            .Left = 50
            .Width = fraRider.Width - 100
        End With
        
        For iTemp = 0 To lblRider.Count - 1
            With lblRider(iTemp)
                .Left = 50
                .Width = (fraRider.Width - 100) \ 4
                .Height = txtRider(iTemp).Height
                If iTemp > 0 Then
                    .Top = txtRider(iTemp - 1).Top + txtRider(iTemp - 1).Height + 50
                Else
                    .Top = cmbRider.Top + cmbRider.Height + 50
                End If
            End With
        Next iTemp
        
        For iTemp = 0 To txtRider.Count - 1
            With txtRider(iTemp)
                .Left = lblRider(iTemp).Width + 100
                .Width = fraRider.Width - lblRider(iTemp).Width - 150
                If iTemp > 0 Then
                    .Top = txtRider(iTemp - 1).Top + txtRider(iTemp - 1).Height + 50
                Else
                    .Top = cmbRider.Top + cmbRider.Height + 50
                End If
                If .DataField = "FEIFID" Then
                    cmdRiderId.Top = .Top
                    cmdRiderId.Left = .Left - cmdRiderId.Width - 50
                    fiPersonId = iTemp
                ElseIf .DataField = "Name_First" Then
                    fiName_First = iTemp
                ElseIf .DataField = "Name_Last" Then
                    fiName_Last = iTemp
                End If
            End With
        Next iTemp
        
        For iTemp = 0 To cmbRiderTxt.Count - 1
            With cmbRiderTxt(iTemp)
                .Left = lblRider(iTemp).Width + 100
               .Width = fraRider.Width - lblRider(iTemp).Width - 150
                If iTemp > 0 Then
                    .Top = txtRider(iTemp - 1).Top + txtRider(iTemp - 1).Height + 50
                Else
                    .Top = cmbRider.Top + cmbRider.Height + 50
                End If
            End With
        Next iTemp
        
        With cmbHorse
            .Top = 300
            .Left = 50
            .Width = fraHorse.Width - 100
        End With
        
        For iTemp = 0 To lblHorse.Count - 1
            With lblHorse(iTemp)
                .Left = 50
                .Width = (fraHorse.Width - 100) \ 4
                .Height = txtHorse(iTemp).Height
                If iTemp > 0 Then
                    .Top = txtHorse(iTemp - 1).Top + txtHorse(iTemp - 1).Height + 50
                Else
                    .Top = cmbHorse.Top + cmbHorse.Height + 50
                End If
            End With
        Next iTemp
        
        For iTemp = 0 To txtHorse.Count - 1
            With txtHorse(iTemp)
                .Left = lblHorse(iTemp).Width + 100
                .Width = fraHorse.Width - lblHorse(iTemp).Width - 150
                If iTemp > 0 Then
                    .Top = txtHorse(iTemp - 1).Top + txtHorse(iTemp - 1).Height + 50
                Else
                    .Top = cmbHorse.Top + cmbHorse.Height + 50
                End If
                If .DataField = "FEIFID" Then
                    cmdHorseId.Top = .Top
                    cmdHorseId.Left = .Left - cmdHorseId.Width - 50
                    fiHorseId = iTemp
                End If
            End With
        Next iTemp
        
        For iTemp = 0 To cmbHorseTxt.Count - 1
            With cmbHorseTxt(iTemp)
                .Left = lblHorse(iTemp).Width + 100
                .Width = fraHorse.Width - lblHorse(iTemp).Width - 150
                If iTemp > 0 Then
                    .Top = txtHorse(iTemp - 1).Top + txtHorse(iTemp - 1).Height + 50
                Else
                    .Top = cmbHorse.Top + cmbHorse.Height + 50
                End If
            End With
        Next iTemp
        
        With lstTests
            .Container = fraTests
            .Top = 250
            .Left = 50
            .Width = .Container.Width - 100
            .Height = .Container.Height - 300
        End With
        
        iTabSize = lstTests.Width \ 300
        
        SetListTabStop Me.lstTests, iTabSize, 4 * iTabSize, 5 * iTabSize, 6 * iTabSize, 7 * iTabSize
           
        With cmdNewRider
            .Container = fraRider
            .Top = .Container.Height - .Height - 50
            .Left = .Container.Width - .Width - 50
        End With
        
        With cmdNewHorse
            .Container = fraHorse
            .Top = .Container.Height - .Height - 50
            .Left = .Container.Width - .Width - 50
        End With
        iBusy = False
    End If
    
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    WriteFormPosition Me
End Sub
Public Sub FindAddParticipant(cTmpSta As String)
    Dim rstParticipant As DAO.Recordset
    Dim cTempIdH As String
    Dim cTempIdP As String
    
    
    With dtaParticipant
        .Refresh
        .Recordset.FindFirst "Sta LIKE '" & cTmpSta & "'"
        If .Recordset.NoMatch = True Then
            
            cTempIdP = CreatePersonId
            cTempIdH = CreateHorseId
            
            Set rstParticipant = mdbMain.OpenRecordset("SELECT * FROM Participants WHERE STA LIKE '" & cTmpSta & "'")
            With rstParticipant
                If rstParticipant.RecordCount = 0 Then
                    .AddNew
                    .Fields("STA") = cTmpSta
                Else
                    .Edit
                End If
                .Fields("PersonId") = cTempIdP
                .Fields("HorseId") = cTempIdH
                .Fields("Flag") = False
                .Update
                .Close
            End With
            
            Set rstParticipant = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE PersonId LIKE '" & cTempIdP & "'")
            If rstParticipant.RecordCount = 0 Then
                With rstParticipant
                    .AddNew
                    .Fields("PersonId") = cTempIdP
                    .Fields("Name_First") = "(" & Translate("Unknown", mcLanguage)
                    .Fields("Name_Last") = Translate("Rider", mcLanguage) & ")"
                    .Update
                    .Close
                End With
            End If
            
            Set rstParticipant = mdbMain.OpenRecordset("SELECT * FROM Horses WHERE HorseId LIKE '" & cTempIdH & "'")
            If rstParticipant.RecordCount = 0 Then
                With rstParticipant
                    .AddNew
                    .Fields("HorseId") = cTempIdH
                    .Fields("Name_Horse") = "(" & Translate("Unknown horse", mcLanguage) & ")"
                    .Update
                    .Close
                End With
            End If
            .Refresh
            FindAddParticipant cTmpSta
        End If
    End With
    
    Set rstParticipant = Nothing
End Sub

Private Sub lstTests_DblClick()
      StartMenuPopUp
End Sub

Private Sub lstTests_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub
Private Sub mnuPopupPopup_Click(Index As Integer)
    Dim iKey As Integer
    Dim iTest As Integer
    Dim cSta As String
    Dim cCode As String
    Dim cTemp As String
    Dim cTemp2 As String
    Dim cTemp3 As String
    Dim cQry As String
    
    Dim iDisq As Integer
    Dim iStatus As Integer
    Dim rst As DAO.Recordset
    Dim iRR As Integer
    Dim iPosition As Integer
    
    If Index = 6 Then
        cSta = Left$(dtaParticipant.Recordset.Fields("cList"), 3)
        cQry = "SELECT Tests.Code "
        cQry = cQry & " & ' - ' & Tests.Test as cList"
        cQry = cQry & " FROM Tests INNER JOIN TestInfo ON Tests.Code=TestInfo.Code"
        cQry = cQry & " WHERE Tests.Code NOT IN (SELECT Code FROM Entries WHERE Sta='" & cSta & "')"
        If frmMain.mnuTestAll.Checked = False Then
            cQry = cQry & " AND TestInfo.Nr>0"
        End If
        cQry = cQry & " ORDER BY Tests.Code"
        
        With frmToolBox
             .intChecked = True
             .strQry = cQry
             .Caption = Translate("Select a test", mcLanguage)
             .Show 1, Me
        End With
        
        cTemp3 = frmMain.Tempvar
        Do While cTemp3 <> ""
           Parse cTemp2, cTemp3, "|"
           If cTemp2 <> "" Then
               Parse cTemp, cTemp2, " - "
               Set rst = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code LIKE '" & cTemp & "' AND STATUS=0 AND Sta LIKE '" & cSta & "'")
               If rst.RecordCount = 0 Then
                    With rst
                        .AddNew
                        .Fields("Code") = cTemp
                        .Fields("Sta") = cSta
                        .Fields("Position") = 0
                        .Fields("Group") = 1
                        .Fields("RR") = False
                        .Fields("Deleted") = 0
                        .Fields("Status") = 0
                        .Fields("Timestamp") = Now
                        .Update
                    End With
                    AddOneToPosition cTemp, 0
               Else
                    MsgBox cTemp & " " & Translate("already entered for this participant.", mcLanguage)
               End If
               rst.Close
               Set rst = Nothing
           End If
        Loop
        frmMain.Tempvar = ""
        cmbParticipants_Change
    ElseIf lstTests.ListIndex >= 0 Then
        dtaTests.Recordset.AbsolutePosition = lstTests.ListIndex
        cSta = dtaTests.Recordset.Fields("Sta")
        cCode = dtaTests.Recordset.Fields("Code")
        iDisq = IIf(IsNull(dtaTests.Recordset.Fields("Deleted")), 0, dtaTests.Recordset.Fields("Deleted"))
        iStatus = IIf(IsNull(dtaTests.Recordset.Fields("Status")), 0, dtaTests.Recordset.Fields("Status"))
        iRR = IIf(IsNull(dtaTests.Recordset.Fields("Entries.RR")), 0, dtaTests.Recordset.Fields("Entries.RR"))
        Select Case Index
            Case 0 'remove
                iKey = MsgBox(Translate("Remove participant from", mcLanguage) & " " & dtaTests.Recordset.Fields("Code") & "?", vbYesNo + vbQuestion)
                If iKey = vbYes Then
                    iKey = MsgBox(Translate("Are you sure to remove participant from", mcLanguage) & " " & dtaTests.Recordset.Fields("Code") & "?", vbYesNo + vbQuestion)
                    If iKey = vbYes Then
                        mdbMain.Execute ("DELETE * FROM Entries WHERE STA='" & cSta & "' AND Code='" & cCode & "'")
                        mdbMain.Execute ("DELETE * FROM Results WHERE STA='" & cSta & "' AND Code='" & cCode & "'")
                        mdbMain.Execute ("DELETE * FROM Marks WHERE STA='" & cSta & "' AND Code='" & cCode & "'")
                    End If
                End If
            Case 1 'withdraw
                If iDisq = 0 Then
                    iKey = MsgBox(Translate("Does this participant withdraw?", mcLanguage), vbYesNo + vbQuestion)
                    If iKey = vbYes Then
                        ParticipantDisqWith cSta, cCode, iStatus, -2
                    End If
                Else
                    iKey = MsgBox(Translate("Remove withdrawal for this participant?", mcLanguage), vbYesNo + vbQuestion)
                    If iKey = vbYes Then
                        ParticipantDisqWith cSta, cCode, iStatus, 0
                    End If
                End If
            Case 2 'eliminate
                If iDisq <> -1 Then
                    iKey = MsgBox(Translate("Eliminate this participant?", mcLanguage), vbYesNo + vbQuestion)
                    If iKey = vbYes Then
                        ParticipantDisqWith cSta, cCode, iStatus, -1
                    End If
                Else
                   iKey = MsgBox(Translate("Remove elimination for this participant?", mcLanguage), vbYesNo + vbQuestion)
                    If iKey = vbYes Then
                        ParticipantDisqWith cSta, cCode, iStatus, 0
                    End If
                End If
            Case 3 'move
                iKey = MsgBox(Translate("Move participant to the beginning of the starting order (Yes) or to the end of the staring order (No).", mcLanguage), vbYesNoCancel + vbQuestion)
                If iKey = vbYes Then
                    Set rst = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & cCode & "' AND Status=0 ORDER BY Position")
                    If rst.RecordCount > 0 Then
                        iPosition = rst.Fields("Position")
                        If iPosition < 1 Then iPosition = 1
                        Set rst = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & cCode & "' AND Status=0 AND Sta='" & cSta & "'")
                        With rst
                            .Edit
                            .Fields("Position") = iPosition - 1
                            .Update
                        End With
                    End If
                ElseIf iKey = vbNo Then
                    Set rst = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & cCode & "' AND Status=0 ORDER BY Position DESC")
                    If rst.RecordCount > 0 Then
                        iPosition = rst.Fields("Position")
                        Set rst = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & cCode & "' AND Status=0 AND Sta='" & cSta & "'")
                        With rst
                            .Edit
                            .Fields("Position") = iPosition + 1
                            .Update
                        End With
                    End If
                End If
            Case 4 'change rein
                Set rst = mdbMain.OpenRecordset("SELECT RR From Entries WHERE STA='" & cSta & "' AND Code='" & cCode & "' AND Status=0")
                If rst.RecordCount > 0 Then
                    With rst
                        .Edit
                        If iRR = True Then
                            .Fields(0) = False
                        Else
                            .Fields(0) = True
                        End If
                        .Update
                    End With
                End If
                rst.Close
                Set rst = Nothing
        End Select
        cmbParticipants_Change
    End If
End Sub

Private Sub txtHorse_Change(Index As Integer)
    Dim iTemp As Integer
    Dim rstHorse As DAO.Recordset
    Dim iKey As Integer
    
    If cmbHorseTxt(Index).Enabled = True Then
        txtHorse2cmbHorseTxt Index
    ElseIf txtHorse(Index).DataField = "FEIFID" Then
        If ValidHorseFEIFId(Trim$(txtHorse(Index).Text)) = True Then
            SetMouseHourGlass
            txtHorse(Index).Enabled = False
            txtHorse(Index).Text = Trim$(UCase$(txtHorse(Index).Text))
            For iTemp = 0 To txtHorse.Count - 1
                If txtHorse(iTemp).DataField = "Country_horse" Then
                    txtHorse(iTemp).Text = Left$(Trim$(txtHorse(Index).Text), 2)
                    Exit For
                End If
            Next iTemp
            Set rstHorse = mdbMain.OpenRecordset("SELECT Name_Horse, FEIFID FROM Horses WHERE FEIFID LIKE '" & txtHorse(fiHorseId).Text & "'")
            If rstHorse.RecordCount = 0 Then
                iKey = MsgBox(Translate("Add a new horse to IceTest", mcLanguage) & "?", vbYesNo + vbQuestion + vbDefaultButton1)
                If iKey = vbYes Then
                    cmdNewHorse.Tag = txtHorse(fiHorseId).Text
                    cmdNewHorse_Click
                End If
            Else
                If rstHorse.Fields("Name_Horse") & "" <> "" Then
                    cmbHorse.Text = rstHorse.Fields("Name_Horse") & ""
                End If
                DoEvents
            End If
            rstHorse.Close
            Set rstHorse = Nothing
            txtHorse(Index).Enabled = True
            cmdHorseId.Enabled = True
            SetMouseNormal
        Else
            cmdHorseId.Enabled = False
        End If
    End If
End Sub

Private Sub txtRider_Change(Index As Integer)
   If cmbRiderTxt(Index).Enabled = True Then
        txtRider2cmbRiderTxt Index
   End If
End Sub
Private Sub txtRider2cmbRiderTxt(Index As Integer)
    Dim iTemp As Integer
    Dim rst As DAO.Recordset
    If NoRevLoad = False Then
        FormLoading = True
        If cmbRiderTxt(Index).ItemData(0) > 0 And txtRider(Index).Text <> "" Then
            Set rst = mdbMain.OpenRecordset("SELECT [Label] FROM [Values] WHERE [Field] Like '" & txtRider(Index).DataField & "' AND Code LIKE '" & txtRider(Index).Text & "'")
            If rst.RecordCount > 0 Then
                cmbRiderTxt(Index).Text = Translate(rst.Fields(0), mcLanguage)
            Else
                cmbRiderTxt(Index).Text = txtRider(Index).Text
            End If
            rst.Close
            Set rst = Nothing
        Else
            cmbRiderTxt(Index).Text = txtRider(Index).Text
        End If
        FormLoading = False
    End If
End Sub
Private Sub txtHorse2cmbHorseTxt(Index As Integer)
    Dim iTemp As Integer
    Dim rst As DAO.Recordset
    
    If NoRevLoad = False Then
        FormLoading = True
        If cmbHorseTxt(Index).ItemData(0) > 0 And txtHorse(Index).Text <> "" Then
            Set rst = mdbMain.OpenRecordset("SELECT [Label] FROM [Values] WHERE [Field] Like '" & txtHorse(Index).DataField & "' AND Code LIKE '" & txtHorse(Index).Text & "'")
            If rst.RecordCount > 0 Then
                cmbHorseTxt(Index).Text = Translate(rst.Fields(0), mcLanguage)
            Else
                cmbHorseTxt(Index).Text = txtHorse(Index).Text
            End If
            rst.Close
            Set rst = Nothing
        Else
            cmbHorseTxt(Index).Text = txtHorse(Index).Text
        End If
        FormLoading = False
    End If
End Sub
Private Sub cmbRiderTxt2txtRider(Index As Integer)
    Dim iTemp As Integer
    Dim rst As DAO.Recordset
    
    FormLoading = True
    NoRevLoad = True
    If cmbRiderTxt(Index).ListIndex >= 0 Then
        If cmbRiderTxt(Index).ItemData(cmbRiderTxt(Index).ListIndex) > 0 Then
            Set rst = mdbMain.OpenRecordset("SELECT [Code] FROM [Values] WHERE [ValueId]=" & cmbRiderTxt(Index).ItemData(cmbRiderTxt(Index).ListIndex))
            If rst.RecordCount > 0 Then
                txtRider(Index).Text = rst.Fields(0)
            Else
                txtRider(Index).Text = ""
            End If
            rst.Close
            Set rst = Nothing
        Else
            txtRider(Index).Text = cmbRiderTxt(Index).Text
        End If
    Else
        txtRider(Index).Text = cmbRiderTxt(Index).Text
    End If
    NoRevLoad = False
    FormLoading = False
End Sub

Private Sub cmbHorseTxt2txtHorse(Index As Integer)
    Dim iTemp As Integer
    Dim rst As DAO.Recordset
    
    FormLoading = True
    NoRevLoad = True
    If cmbHorseTxt(Index).ListIndex >= 0 Then
        If cmbHorseTxt(Index).ItemData(cmbHorseTxt(Index).ListIndex) > 0 Then
            Set rst = mdbMain.OpenRecordset("SELECT [Code] FROM [Values] WHERE [ValueId]=" & cmbHorseTxt(Index).ItemData(cmbHorseTxt(Index).ListIndex))
            If rst.RecordCount > 0 Then
                txtHorse(Index).Text = rst.Fields(0)
            Else
                txtHorse(Index).Text = ""
            End If
            rst.Close
            Set rst = Nothing
        Else
            txtHorse(Index).Text = cmbHorseTxt(Index).Text
        End If
    Else
        txtHorse(Index).Text = cmbHorseTxt(Index).Text
    End If
    NoRevLoad = False
    FormLoading = False
End Sub
Private Sub txtRider_LostFocus(Index As Integer)
    Dim rst As DAO.Recordset
    If InStr(txtRider(Index).Tag, "date") > 0 Then
        If IsDate(txtRider(Index).Text) = True Then
            Set rst = mdbMain.OpenRecordset("SELECT " & txtRider(Index).DataField & " FROM Persons WHERE PersonId='" & dtaRider.Recordset.Fields("PersonId") & "'")
            If rst.RecordCount > 0 Then
                With rst
                    .Edit
                    .Fields(0).Value = CDate(txtRider(Index).Text)
                    .Update
                End With
            End If
            rst.Close
            Set rst = Nothing
        Else
            txtRider(Index).Text = ""
        End If
    End If
End Sub
Private Sub txtHorse_LostFocus(Index As Integer)
    Dim rst As DAO.Recordset
    If InStr(txtHorse(Index).Tag, "date") > 0 Then
        If IsDate(txtHorse(Index).Text) = True Then
            Set rst = mdbMain.OpenRecordset("SELECT " & txtHorse(Index).DataField & " FROM Horses WHERE HorseId='" & dtaHorse.Recordset.Fields("HorseId") & "'")
            If rst.RecordCount > 0 Then
                With rst
                    .Edit
                    .Fields(0).Value = CDate(txtHorse(Index).Text)
                    .Update
                End With
            End If
            rst.Close
            Set rst = Nothing
        Else
            txtHorse(Index).Text = ""
        End If
    End If
End Sub

Public Sub StartMenuPopUp()
    Dim iTemp As Integer
    Dim cTemp As String
    If lstTests.ListIndex < 0 Then
        For iTemp = 0 To 4
            mnuPopupPopUp(iTemp).Enabled = False
            Parse cTemp, mnuPopupPopUp(iTemp).Caption, " - "
            mnuPopupPopUp(iTemp).Caption = cTemp
        Next iTemp
    Else
        dtaTests.Recordset.AbsolutePosition = lstTests.ListIndex
        For iTemp = 0 To 4
            If iTemp = 0 And IsNull(dtaTests.Recordset.Fields("Score")) Then
                mnuPopupPopUp(iTemp).Enabled = True
            ElseIf iTemp = 1 Or iTemp = 2 Then
                If dtaTests.Recordset.Fields("Deleted") = -1 Then ' eliminated
                    mnuPopupPopUp(2).Enabled = True
                    mnuPopupPopUp(2).Checked = True
                    mnuPopupPopUp(1).Enabled = False
                    mnuPopupPopUp(1).Checked = False
                ElseIf dtaTests.Recordset.Fields("Deleted") = -2 Then ' withdrawn
                    mnuPopupPopUp(1).Enabled = True
                    mnuPopupPopUp(1).Checked = True
                    mnuPopupPopUp(2).Enabled = True
                    mnuPopupPopUp(2).Checked = False
                Else
                    If IsNull(dtaTests.Recordset.Fields("Score")) = True Then
                        mnuPopupPopUp(1).Enabled = True
                    Else
                        mnuPopupPopUp(1).Enabled = False
                    End If
                    mnuPopupPopUp(2).Enabled = True
                    mnuPopupPopUp(1).Checked = False
                    mnuPopupPopUp(2).Checked = False
                End If
            ElseIf iTemp = 3 And dtaTests.Recordset.Fields("Status") = 0 And IsNull(dtaTests.Recordset.Fields("Score")) = True Then
                mnuPopupPopUp(iTemp).Enabled = True
            ElseIf iTemp = 4 And dtaTests.Recordset.Fields("Tests.RR") = True And dtaTests.Recordset.Fields("Status") = 0 And IsNull(dtaTests.Recordset.Fields("Score")) = True Then
                mnuPopupPopUp(iTemp).Enabled = True
            Else
                mnuPopupPopUp(iTemp).Enabled = False
            End If
            Parse cTemp, mnuPopupPopUp(iTemp).Caption, " - "
            mnuPopupPopUp(iTemp).Caption = cTemp & " - " & dtaTests.Recordset.Fields("Code")
        Next iTemp
        
    End If
    
    PopupMenu mnuPopup
   
End Sub

Function ValidRiderFEIFId(cFEIFID As String) As Integer
    Dim cTemp As String
    Dim iTemp As Integer
    Dim iCheck As Integer
    
    On Local Error GoTo ValidRiderFEIFIdErr
    
    ValidRiderFEIFId = False
    If Len(cFEIFID) = 12 And Left(cFEIFID, 2) = "FF" Then
        cTemp = Mid$(cFEIFID, 3)
        For iTemp = 1 To 9
            iCheck = iCheck + (10 - iTemp) * Val(Mid$(cTemp, iTemp, 1))
        Next iTemp
        If iCheck Mod 11 = Val(Right$(cTemp, 1)) Then
            ValidRiderFEIFId = True
        End If
    End If
    
ValidRiderFEIFIdErr:
        
End Function
Public Function RequestRiderXML(cUrl, cRider As String) As String
    Dim cXML As String
    Dim cXml2 As String
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iTemp3 As Integer
    Dim cTemp As String
    Dim cName As String
    Dim cAKA As String
    Dim cFEIFID As String
    Dim cCountries As String
    Dim cField() As String
    Dim cMsg As String
    Dim iMaxRequest As Integer
    Dim iKey As Integer
    Dim cNationality As String
    
    iMaxRequest = 250
    
    With frmMain.Inet1
        .Cancel
        .Protocol = icHTTP
        .RequestTimeout = 180
        cXML = .OpenURL(cUrl & cRider)
        While .StillExecuting
            DoEvents
        Wend
    End With
    
    If cXML = "" Then
        cmdNewRider.Tag = ""
        Me.Enabled = True
        Me.cmdRiderId.Enabled = True
        MsgBox Translate("Service not available, check your connection to Internet.", mcLanguage), vbCritical
        StatusMessage ""
        Me.MousePointer = vbNormal
        SetMouseNormal
        miConnectedToInternet = False
    Else
        miConnectedToInternet = True
        cXml2 = cXML
        cField = Split(cXML, "</row>", -1, vbTextCompare)
        cMsg = ""
        iTemp2 = 0
        For iTemp = 0 To UBound(cField) - 1
            If InStr(cField(iTemp), "</RIDER>") > 0 Then
                cName = XmlParse(cField(iTemp), "RIDER")
                cAKA = XmlParse(cField(iTemp), "AKA")
                cFEIFID = XmlParse(cField(iTemp), "FEIFID")
                cCountries = XmlParse(cField(iTemp), "COUNTRIES")
                cNationality = XmlParse(cField(iTemp), "SPORTNATIONALITY")
                cTemp = cName & IIf(cNationality <> "", " - " & cNationality, "") & IIf(cCountries <> "", " - " & cCountries, "") & IIf(cAKA <> "", " (" & cAKA & ")", "") & vbTab & " [" & cFEIFID & "]"
                If InStr(cMsg, "[" & cFEIFID & "]") = 0 Then
                    iTemp3 = iTemp3 + 1
                End If
                If InStr(cMsg, UTF8_Decode(cTemp) & "|") = 0 Then
                    cMsg = cMsg & UTF8_Decode(cTemp) & "|"
                    iTemp2 = iTemp2 + 1
                End If
                If iTemp2 > iMaxRequest Then Exit For
            End If
        Next iTemp
        
        If iTemp2 > iMaxRequest Then
            iKey = MsgBox(Replace(Translate("Your request '%s' should be more specific!", mcLanguage) & vbCrLf & "Do you want to search the FEIF WorldRanking first?", "%s", Replace(cRider, "+", " ")), vbCritical + vbYesNo)
            If iKey = vbYes Then
                ShowDocument "https://www.feif.org/WorldRanking/Riders", Me
            End If
            
            StatusMessage ""
            cmdNewRider.Tag = ""
            Me.MousePointer = vbNormal
            SetMouseNormal
            cXML = ""
        ElseIf iTemp2 > 1 And iTemp3 > 1 Then
            With frmToolBox
                .strList = cMsg
                .intChecked = False
                .intSingleChoice = True
                .Caption = Translate("Select a rider", mcLanguage) & " [" & iTemp2 & "]"
                .Show 1, Me
            End With
            
            If frmMain.Tempvar <> "" Then
                cTemp = Replace(frmMain.Tempvar, "|", "")
                frmMain.Tempvar = ""
                iTemp = InStr(cTemp, "[")
                If iTemp > 0 Then
                    cTemp = Trim$(Replace(Mid$(cTemp & " ", iTemp + 1), "]", ""))
                End If
                cXML = RequestRiderXML(cUrl, UTF8_Encode(Replace(cTemp, " ", "+")))
           Else
                MsgBox Translate("You have to select one rider!", mcLanguage), vbCritical
                StatusMessage ""
                cmdNewRider.Tag = ""
                Me.MousePointer = vbNormal
                SetMouseNormal
                cXML = ""
            End If
        Else
            cXML = cXml2
        End If
    End If
    
    RequestRiderXML = cXML
    
End Function
Public Function RequestHorseXML(cUrl, cHorse As String) As String
    Dim cXML As String
    Dim cXml2 As String
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iTemp3 As Integer
    Dim cTemp As String
    Dim cName As String
    Dim cSex As String
    Dim cFEIFID As String
    Dim cColor As String
    Dim cField() As String
    Dim cMsg As String
    Dim iMaxRequest As Integer
    Dim iKey As Integer
    
    iMaxRequest = 250
    
    With frmMain.Inet1
        .Cancel
        .Protocol = icHTTP
        .RequestTimeout = 180
        cXML = .OpenURL(cUrl & cHorse)
        While .StillExecuting
            DoEvents
        Wend
    End With
    
    If cXML = "" Then
        cmdNewHorse.Tag = ""
        Me.Enabled = True
        Me.cmdHorseId.Enabled = True
        MsgBox Translate("Service not available, check your connection to Internet.", mcLanguage), vbCritical
        StatusMessage ""
        Me.MousePointer = vbNormal
        SetMouseNormal
        miConnectedToInternet = False
    Else
        miConnectedToInternet = True
        cXml2 = cXML
        cField = Split(cXML, "</row>", -1, vbTextCompare)
        cMsg = ""
        iTemp2 = 0
        For iTemp = 0 To UBound(cField) - 1
            If InStr(cField(iTemp), "</Name_Horse>") > 0 Then
                cName = XmlParse(cField(iTemp), "Name_Horse")
                cSex = XmlParse(cField(iTemp), "Sex_Horse")
                cColor = XmlParse(cField(iTemp), "Color")
                cFEIFID = XmlParse(cField(iTemp), "FEIFID")
                cTemp = cName & IIf(cSex <> "", " - " & cSex, "") & IIf(cColor <> "", " (" & cColor & ")", "") & vbTab & " [" & cFEIFID & "]"
                If InStr(cMsg, "[" & cFEIFID & "]") = 0 Then
                    iTemp3 = iTemp3 + 1
                End If
                If InStr(cMsg, UTF8_Decode(cTemp) & "|") = 0 Then
                    cMsg = cMsg & UTF8_Decode(cTemp) & "|"
                    iTemp2 = iTemp2 + 1
                End If
                If iTemp2 > iMaxRequest Then Exit For
            End If
        Next iTemp
        
        If iTemp2 > iMaxRequest Then
            iKey = MsgBox(Replace(Translate("Your request '%s' should be more specific!", mcLanguage) & vbCrLf & "Do you want to search the FEIF WorldRanking first?", "%s", Replace(cHorse, "+", " ")), vbCritical + vbYesNo)
            If iKey = vbYes Then
                ShowDocument "https://www.feif.org/WorldRanking/Horses", Me
            End If
            
            StatusMessage ""
            cmdNewHorse.Tag = ""
            Me.MousePointer = vbNormal
            SetMouseNormal
            cXML = ""
        ElseIf iTemp2 > 1 And iTemp3 > 1 Then
            With frmToolBox
                .strList = cMsg
                .intChecked = False
                .intSingleChoice = True
                .Caption = Translate("Select a horse", mcLanguage) & " [" & iTemp2 & "]"
                .Show 1, Me
            End With
            
            If frmMain.Tempvar <> "" Then
                cTemp = Replace(frmMain.Tempvar, "|", "")
                frmMain.Tempvar = ""
                iTemp = InStr(cTemp, "[")
                If iTemp > 0 Then
                    cTemp = Trim$(Replace(Mid$(cTemp & " ", iTemp + 1), "]", ""))
                End If
                cXML = RequestHorseXML(cUrl, UTF8_Encode(Replace(cTemp, " ", "+")))
           Else
                MsgBox Translate("You have to select one horse!", mcLanguage), vbCritical
                StatusMessage ""
                cmdNewHorse.Tag = ""
                Me.MousePointer = vbNormal
                SetMouseNormal
                cXML = ""
            End If
        Else
            cXML = cXml2
        End If
    End If
    
    RequestHorseXML = cXML
    
End Function

Function ValidWRCode(cWRCode As String) As Integer
    
    On Local Error GoTo ValidWRCodeErr
    
    ValidWRCode = False
    If Len(cWRCode) = 10 And Mid$(cWRCode, 3, 1) = "2" Then
        ValidWRCode = True
    End If
    
ValidWRCodeErr:
    
End Function

