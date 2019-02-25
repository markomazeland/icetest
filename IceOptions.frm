VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Options"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOptions 
      Height          =   5175
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6735
      Begin VB.CheckBox chkIceSort 
         Caption         =   "Use IceSort"
         Height          =   375
         Left            =   1800
         TabIndex        =   42
         ToolTipText     =   "Keep log of program operations in separate database"
         Top             =   4680
         Width           =   4815
      End
      Begin VB.CheckBox chkLogDB 
         Caption         =   "Write log to SQLite database "
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         ToolTipText     =   "Keep log of program operations in separate database"
         Top             =   4320
         Width           =   4815
      End
      Begin VB.TextBox txtSponsor 
         Height          =   495
         Left            =   1800
         TabIndex        =   31
         Text            =   "Sponsored by"
         ToolTipText     =   "Announce sponsors of tests as ..."
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox cmbCountry 
         Height          =   315
         ItemData        =   "IceOptions.frx":0000
         Left            =   5160
         List            =   "IceOptions.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   38
         Text            =   "cmbCountry"
         ToolTipText     =   "Select the prefferred country to select country related functions (when available)"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkHighLights 
         Caption         =   "&Highlight selected tabs"
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   3240
         Width           =   4815
      End
      Begin VB.CheckBox chkHtmlFiles 
         Caption         =   "Create &Html-files for website automatically"
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         ToolTipText     =   "Create Html-files for website"
         Top             =   2880
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.CheckBox chkExcelFiles 
         Caption         =   "Create &Extra files for external functions"
         Height          =   375
         Left            =   1800
         TabIndex        =   27
         ToolTipText     =   "Create Extra files for external functions, like a videowall to inform the public"
         Top             =   2280
         Value           =   1  'Checked
         Width           =   4815
      End
      Begin VB.CheckBox chkJudgesRanking 
         Caption         =   "Show Ranking per &Judge in Result Lists"
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         ToolTipText     =   "Show Ranking Per Judge in Result Lists"
         Top             =   1920
         Width           =   4815
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "&double click"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   12
         ToolTipText     =   "How are participants selected"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         ItemData        =   "IceOptions.frx":0004
         Left            =   1800
         List            =   "IceOptions.frx":001A
         TabIndex        =   9
         ToolTipText     =   "Set the size of the text on the screen"
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmbLanguage 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select the prefferred language "
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optSelect 
         Caption         =   "&single click"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   11
         ToolTipText     =   "How are participants selected"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblSponsor 
         Caption         =   "Sponsor link"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblCountry 
         Caption         =   "&Country"
         Height          =   375
         Left            =   4080
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSelect 
         Caption         =   "Select &participants by"
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblFontSize 
         Caption         =   "&Font size"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblLanguage 
         Caption         =   "&Language"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   4095
      Index           =   3
      Left            =   5160
      TabIndex        =   32
      Top             =   1800
      Width           =   6735
      Begin VB.ComboBox cmbCFinals 
         Height          =   315
         ItemData        =   "IceOptions.frx":0034
         Left            =   1800
         List            =   "IceOptions.frx":0047
         TabIndex        =   43
         Text            =   "30"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbBFinals 
         Height          =   315
         ItemData        =   "IceOptions.frx":005F
         Left            =   1800
         List            =   "IceOptions.frx":006F
         TabIndex        =   35
         Text            =   "20"
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox chkMarkFinals 
         Caption         =   "Mark finals in result list of preliminary rounds"
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   720
         Width           =   4815
      End
      Begin VB.CheckBox chkFinalsSequence 
         Caption         =   "Show participants in finals first to last"
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         ToolTipText     =   "Show participants in finals first to last in stead of last to first"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblCFinals 
         Caption         =   "or more participants are needed for a C-Final"
         Height          =   375
         Left            =   3000
         TabIndex        =   44
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label lblBFinals 
         Caption         =   "or more participants are needed for a B-Final"
         Height          =   375
         Left            =   3000
         TabIndex        =   36
         Top             =   1200
         Width           =   3015
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   2175
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         ToolTipText     =   "Test creation of backup"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ComboBox cmbBackup 
         Height          =   315
         ItemData        =   "IceOptions.frx":0083
         Left            =   3480
         List            =   "IceOptions.frx":0096
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optBackup 
         Caption         =   "Every ___ minutes:"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   22
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton optBackup 
         Caption         =   "Upon E&xit"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton optBackup 
         Caption         =   "&Manually"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblBackup 
         Caption         =   "&Create backup"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   3735
      Index           =   1
      Left            =   4440
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CheckBox chkHorseAge 
         Caption         =   "Print &Horses' age in startlists"
         Height          =   375
         Left            =   1800
         TabIndex        =   45
         ToolTipText     =   "Print horses' age (below 7) in startlists"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox chkLK 
         Caption         =   "Print Riders'  LK in result lists"
         Height          =   375
         Left            =   1800
         TabIndex        =   41
         ToolTipText     =   "Print Riders' LK in result lists"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtNoColor 
         Height          =   285
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "Xx"
         ToolTipText     =   "How to indicate missing colors?"
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox chkUseColors 
         Caption         =   "Use &colors in groups"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         ToolTipText     =   "Use colors when dividing riders into groups"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtColors 
         Height          =   495
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "Enter the colors used, one at each line"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkTeam 
         Caption         =   "Print Riders' &Team in result lists"
         Height          =   375
         Left            =   1800
         TabIndex        =   25
         ToolTipText     =   "Print Riders' Team name in result lists"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CheckBox chkClub 
         Caption         =   "Print Riders' &Club in result lists"
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         ToolTipText     =   "Print Riders' Club name in result lists"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CheckBox chkHorseId 
         Caption         =   "Print &Horses' ID in result lists"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         ToolTipText     =   "Print horses' (FEIF) ID number in result lists"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblColor 
         Caption         =   "&Use Colors"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4683
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Participants"
            Key             =   "Participants"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Backup"
            Key             =   "Backup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Finals"
            Key             =   "Finals"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   4695
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Close and apply last changes"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         ToolTipText     =   "Cancel changes"
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub chkClub_Click()
    chkClub.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkClub_KeyPress(KeyAscii As Integer)
    chkClub.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkFinalsSequence_Click()
    chkFinalsSequence.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkHighLights_Click()
    chkHighLights.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkHorseId_Click()
    chkHorseId.Tag = "*"
    Me.Tag = "*"
End Sub


Private Sub chkHorseId_KeyPress(KeyAscii As Integer)
    chkHorseId.Tag = "*"
    Me.Tag = "*"
End Sub
Private Sub chkHorseAge_Click()
    chkHorseAge.Tag = "*"
    Me.Tag = "*"
End Sub


Private Sub chkHorseAge_KeyPress(KeyAscii As Integer)
    chkHorseAge.Tag = "*"
    Me.Tag = "*"
End Sub


Private Sub chkIceSort_Click()
    chkIceSort.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkJudgesRanking_Click()
    chkJudgesRanking.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkExcelFiles_Click()
    chkExcelFiles.Tag = "*"
    Me.Tag = "*"
End Sub
Private Sub chkHtmlFiles_Click()
    chkHtmlFiles.Tag = "*"
    Me.Tag = "*"
End Sub
Private Sub chkLK_Click()
    chkLK.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkLK_KeyDown(KeyCode As Integer, Shift As Integer)
    chkLK.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkLogDB_Click()
    chkLogDB.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkMarkFinals_Click()
    chkMarkFinals.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkMarkFinals_KeyDown(KeyCode As Integer, Shift As Integer)
    chkMarkFinals.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkTeam_click()
    chkTeam.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkTeam_KeyDown(KeyCode As Integer, Shift As Integer)
    chkTeam.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkUseColors_Click()
    If chkUseColors.Value <> 1 Then
        Me.txtColors.Enabled = False
    Else
        Me.txtColors.Enabled = True
    End If
    chkUseColors.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub chkUseColors_KeyPress(KeyAscii As Integer)
    If chkUseColors.Value <> 1 Then
        Me.txtColors.Enabled = False
    Else
        Me.txtColors.Enabled = True
    End If
    chkUseColors.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbBackup_Change()
    optBackup(0).Tag = "*"
    Me.Tag = "*"
    Me.optBackup(2).Value = True
End Sub

Private Sub cmbBackup_click()
    optBackup(0).Tag = "*"
    Me.Tag = "*"
    Me.optBackup(2).Value = True
End Sub

Private Sub cmbBackup_Validate(Cancel As Boolean)
    optBackup(0).Tag = "*"
    Me.Tag = "*"
    Me.optBackup(2).Value = True
End Sub

Private Sub cmbBFinals_Change()
    cmbBFinals.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbCFinals_Change()
    cmbCFinals.Tag = "*"
    Me.Tag = "*"
End Sub
Private Sub cmbBFinals_Click()
    cmbBFinals.Tag = "*"
    Me.Tag = "*"
End Sub
Private Sub cmbCFinals_Click()
    cmbCFinals.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbBFinals_KeyDown(KeyCode As Integer, Shift As Integer)
    cmbBFinals.Tag = "*"
    Me.Tag = "*"
End Sub
Private Sub cmbCFinals_KeyDown(KeyCode As Integer, Shift As Integer)
    cmbCFinals.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbFontSize_Change()
    cmbFontSize.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbFontSize_Click()
    cmbFontSize.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbFontSize_KeyDown(KeyCode As Integer, Shift As Integer)
    cmbFontSize.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbLanguage_Change()
    cmbLanguage.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbLanguage_Click()
    cmbLanguage.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbLanguage_KeyDown(KeyCode As Integer, Shift As Integer)
    cmbLanguage.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbCountry_Change()
    cmbCountry.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbCountry_Click()
    cmbCountry.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmbCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    cmbCountry.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = ""
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Tag = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    CreateBackup mdbMain, GetVariable("Backup")
End Sub
Private Sub Form_Load()
    Dim cTemp As String
    Dim cLang As String
    Dim ctl As Control
    Dim iTemp As Integer
    Dim rstCountry As Recordset
    
    ReadFormPosition Me
    
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    With tbsOptions
        .TabWidthStyle = tabFixed
        .TabMinWidth = frmMain.miTabMinWidth
        .TabFixedHeight = frmMain.miTabMinHeight
        .TabFixedWidth = Me.ScaleWidth \ 4
        .Tabs("General").Caption = Translate("General", mcLanguage)
        .Tabs("Participants").Caption = Translate("Participants", mcLanguage)
        .Tabs("Backup").Caption = Translate("Backup", mcLanguage)
    End With
    
    'build a list of available languages
    cTemp = GetVariable("Languages")
    If cTemp = "" Then
        cTemp = "English Icelandic Norwegian Swedish Finnish Danish German Dutch French Italian Slovenian"
        SetVariable "Languages", cTemp
    End If
    cmbLanguage.Clear
    Do While cTemp <> ""
       Parse cLang, cTemp, " "
       If cLang <> "" Then
          cmbLanguage.AddItem cLang
       End If
    Loop
    cmbLanguage.Text = mcLanguage
    cmbFontSize.Text = msFontSize
    
    Set rstCountry = mdbMain.OpenRecordset("SELECT [Code] & ' - ' & [Label] AS [Country] FROM [Values] WHERE [Field] LIKE 'Nationality' ORDER BY Code")
    cmbCountry.Clear
    With rstCountry
        If .RecordCount > 0 Then
            Do While Not .EOF
                cmbCountry.AddItem .Fields("Country")
                If Left(.Fields("Country"), 2) = mcCountry Then
                    cTemp = .Fields("country")
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rstCountry = Nothing
    cmbCountry.Text = cTemp
    cmbCountry.FontSize = msFontSize
    
    optSelect(0).Value = miSelectBySingleClick
    optSelect(1).Value = Not optSelect(0).Value
    Select Case miBackupInterval
    Case -1
        optBackup(0).Value = True
    Case 0
        optBackup(1).Value = True
    Case Else
        optBackup(2).Value = True
        Me.cmbBackup.Text = miBackupInterval
    End Select
    
    txtColors.Text = Replace(frmMain.TestColors, ",", vbCrLf)
    txtNoColor.Text = mcNoColor
    txtSponsor.Text = GetVariable("Sponsors")
    chkUseColors.Value = miUseColors
    chkClub.Value = miShowRidersClub
    chkTeam.Value = miShowRidersTeam
    chkHorseId.Value = miShowHorseId
    chkHorseAge.Value = miShowHorseAge
    chkJudgesRanking.Value = miShowJudgesRanking
    chkExcelFiles = miExcelFiles
    chkHtmlFiles = miHtmlFiles
    chkHighLights = IIf(miUseHighLights = 0, 0, 1)
    chkFinalsSequence = miFinalsSequence
    chkMarkFinals = miMarkFinalsInResultLists
    cmbBFinals.Text = Format$(miBFinalLevel)
    cmbCFinals.Text = Format$(miCFinalLevel)
    chkLogDB.Value = miWriteLogDB
    chkIceSort.Value = miUseIceSort
    chkLK.Value = miShowRidersLK
    
    'Make LK setting only visible to IPZV users:
    If mcVersionSwitch = "ipzv" Then
        chkLK.Visible = True
    Else
        chkLK.Visible = False
    End If
    
    If chkUseColors.Value <> 1 Then
        Me.txtColors.Enabled = False
    Else
        Me.txtColors.Enabled = True
    End If
    
    If mcVersionSwitch = "ipzv" Then
        chkIceSort.Enabled = False
    Else
        If (Dir$(App.Path & "\" & SortingApplication) > "") Then
            chkIceSort.Enabled = True
        Else
            chkIceSort.Enabled = False
        End If
    End If
    
    On Local Error Resume Next
    For Each ctl In Me.Controls
        ctl.Tag = ""
    Next
    Me.Tag = ""
    On Local Error GoTo 0
    
    tbsOptions_Click
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ApplyChanges
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    With tbsOptions
        .Width = ScaleWidth
        .Height = ScaleHeight - fraButtons.Height - 100
        .TabMinWidth = frmMain.miTabMinWidth
        .TabFixedHeight = frmMain.miTabMinHeight
        .TabFixedWidth = Me.ScaleWidth \ (tbsOptions.Tabs.Count + 1)
    End With
    
    For iTemp = 0 To fraOptions.Count - 1
        With fraOptions(iTemp)
            .Left = tbsOptions.ClientLeft
            .Top = tbsOptions.ClientTop
            .Width = tbsOptions.ClientWidth
            .Height = tbsOptions.ClientHeight
        End With
    Next iTemp
    
    With fraButtons
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height - 50
        .Height = cmdCancel.Height
    End With
    
    With cmdCancel
        .Container = fraButtons
        .Top = 0
        .Left = .Container.Width - .Width - 50
    End With
    
    With cmdOK
        .Container = fraButtons
        .Top = 0
        .Left = cmdCancel.Left - .Width - 50
    End With
    
    With lblLanguage
        .Width = .Container.Width \ 5
    End With
    
    With cmbLanguage
        .Left = lblLanguage.Left + lblLanguage.Width + 50
        .Width = .Container.Width \ 2 - .Left - 150
    End With
    
    With lblCountry
        .Top = lblLanguage.Top
        .Left = .Container.Width \ 2 + 50
        .Width = lblLanguage.Width
    End With
    
    With cmbCountry
        .Left = lblCountry.Left + lblCountry.Width + 50
        .Width = .Container.Width - .Left - 150
        .Top = lblCountry.Top
    End With
    
    With lblFontSize
        .Top = lblLanguage.Top + lblLanguage.Height + 50
        .Width = lblLanguage.Width
    End With
    
    With cmbFontSize
        .Top = lblFontSize.Top
        .Left = lblFontSize.Left + lblFontSize.Width + 50
        .Width = .Container.Width \ 2 - .Left - 150
    End With
    
    With optSelect(0)
        .Top = lblFontSize.Top + lblFontSize.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With optSelect(1)
        .Top = optSelect(0).Top + optSelect(0).Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With lblSelect
        .Top = optSelect(0).Top
        .Width = lblLanguage.Width
        .Height = optSelect(0).Height * 2 + 50
    End With
    
    With chkJudgesRanking
        .Top = lblSelect.Top + lblSelect.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With Me.chkExcelFiles
        .Top = chkJudgesRanking.Top + chkJudgesRanking.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With Me.chkHtmlFiles
        .Top = chkExcelFiles.Top + chkExcelFiles.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With Me.chkHighLights
        .Top = chkHtmlFiles.Top + chkHtmlFiles.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With Me.lblSponsor
        .Top = chkHighLights.Top + chkHighLights.Height + 50
        .Width = lblLanguage.Width
    End With
    
    With Me.txtSponsor
        .Top = lblSponsor.Top
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With Me.chkLogDB
        .Top = txtSponsor.Top + lblSponsor.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With Me.chkIceSort
        .Top = chkLogDB.Top + chkLogDB.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
     
    With Me.chkFinalsSequence
        .Top = lblColor.Top
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
        
    With Me.chkMarkFinals
        .Top = chkFinalsSequence.Top + chkFinalsSequence.Height + 50
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With cmbBFinals
        .Top = chkMarkFinals.Top + chkMarkFinals.Height + 50
        .Left = cmbFontSize.Left
    End With
    
    With cmbCFinals
        .Top = cmbBFinals.Top + cmbBFinals.Height + 50
        .Left = cmbFontSize.Left
    End With
    
    With lblBFinals
        .Top = cmbBFinals.Top
        .Left = cmbBFinals.Left + cmbBFinals.Width + 50
        .Width = .Container.Width - .Left - 50 - cmbBFinals.Width - 50
    End With
    
    With lblCFinals
        .Top = cmbCFinals.Top
        .Left = cmbCFinals.Left + cmbCFinals.Width + 50
        .Width = .Container.Width - .Left - 50 - cmbCFinals.Width - 50
    End With
    
    With txtColors
        .Left = cmbFontSize.Left
        .Width = .Container.Width - .Left - 50
        .Top = lblColor.Top
        .Height = .Container.Height \ 2
    End With
    
    With chkUseColors
        .Top = txtColors.Top + txtColors.Height + 50
        .Width = .Container.Width - .Left - 100 - txtNoColor.Width
        .Left = cmbFontSize.Left
    End With
    
    With txtNoColor
        .Top = chkUseColors.Top
        .Left = chkUseColors.Left + chkUseColors.Width + 50
    End With
    
    With chkHorseId
        .Top = chkUseColors.Top + chkUseColors.Height + 50
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With chkHorseAge
        .Top = chkHorseId.Top + chkHorseId.Height + 50
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With chkTeam
        .Top = chkHorseAge.Top + chkHorseAge.Height + 50
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With chkClub
        .Top = chkTeam.Top + chkTeam.Height + 50
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With chkLK
        .Top = chkClub.Top + chkClub.Height + 50
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With optBackup(0)
        .Top = lblBackup.Top
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With optBackup(1)
        .Top = optBackup(0).Top + optBackup(0).Height + 50
        .Width = .Container.Width - .Left - 50
        .Left = cmbFontSize.Left
    End With
    
    With optBackup(2)
        .Top = optBackup(1).Top + optBackup(1).Height + 50
        .Width = .Container.Width - Me.cmbBackup.Width - .Left - 100
        .Left = cmbFontSize.Left
    End With
    
    With cmbBackup
        .Top = optBackup(2).Top
        .Left = .Container.Width - Me.cmbBackup.Width - 50
    End With
    
    With cmdTest
        .Top = .Container.Height - .Height - 50
        .Left = .Container.Width - Me.cmdTest.Width - 50
    End With
    
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFormPosition Me
End Sub

Private Sub optBackup_Click(Index As Integer)
    optBackup(0).Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub optSelect_Click(Index As Integer)
    optSelect(0).Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub tbsOptions_Click()
    Dim iTemp As Integer
    For iTemp = 0 To fraOptions.Count - 1
        If iTemp = tbsOptions.SelectedItem.Index - 1 Then
            fraOptions(iTemp).Visible = True
            tbsOptions.Tabs(tbsOptions.SelectedItem.Index).HighLighted = miUseHighLights
        Else
            fraOptions(iTemp).Visible = False
            tbsOptions.Tabs(iTemp + 1).HighLighted = False
        End If
    Next iTemp
    
End Sub
Sub ApplyChanges()
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim cColors() As String
    Dim cColorList As String
    Dim ctl As Control
    Dim iRestart As Integer
    Dim iCaption As Integer
    Dim rstColor As DAO.Recordset
    
    On Local Error Resume Next
    If Me.Tag <> "" Then
        iKey = MsgBox(Translate("Do you want to save all changes?", mcLanguage), vbYesNo + vbQuestion)
        If iKey = vbYes Then
            For Each ctl In Me.Controls
                If ctl.Tag = "*" Then
                    Select Case ctl.Name
                    Case "cmbLanguage"
                        WriteIniFile gcIniFile, frmMain.Name, "Language", cmbLanguage
                        iRestart = True
                    Case "cmbCountry"
                        SetVariable "Country", Left$(cmbCountry, 2)
                        iRestart = True
                    Case "cmbFontSize"
                        msFontSize = cmbFontSize.Text
                        If msFontSize < 6 Then
                            msFontSize = 6
                        ElseIf msFontSize > 14 Then
                            msFontSize = 14
                        End If
                        WriteIniFile gcIniFile, Me.Name, "FontSize", Format$(msFontSize)
                        ChangeFontSize frmMain, msFontSize
                    Case "optSelect"
                        miSelectBySingleClick = optSelect(0).Value
                    Case "optBackup"
                        If optBackup(0).Value = True Then
                            miBackupInterval = -1
                        ElseIf optBackup(1).Value = True Then
                            miBackupInterval = 0
                        Else
                            miBackupInterval = Val(Me.cmbBackup.Text)
                        End If
                    Case "txtNoColor"
                        mcNoColor = txtNoColor.Text
                        SetVariable "NoColor", mcNoColor
                    Case "txtColors"
                        cColorList = txtColors.Text
                        cColorList = Replace(cColorList, vbCrLf & vbCrLf, vbCrLf)
                        cColors = Split(cColorList, vbCrLf)
                        cColorList = ""
                        If UBound(cColors) < 5 Then
                            MsgBox Translate("You are recommended to use at least 6 different colors.", mcLanguage) & vbCrLf & Translate("The first two letters of every color should be different from each other.", mcLanguage)
                        End If
                        mdbMain.Execute ("DELETE * FROM Colors")
                        Set rstColor = mdbMain.OpenRecordset("SELECT * FROM Colors")
                        For iTemp = LBound(cColors) To UBound(cColors)
                            iTemp2 = 0
                            If cColors(iTemp) <> "" Then
                                cColors(iTemp) = Trim$(StrConv(LCase$(Replace(cColors(iTemp), " ", "")), vbProperCase))
                                Do While InStr("," & cColorList, "," & cColors(iTemp)) > 0
                                    iTemp2 = iTemp2 + 1
                                    cColors(iTemp) = Left$(cColors(iTemp) & "X", 1) & Format$(iTemp2, "0")
                                    If Val(Mid$(cColors(iTemp), 2)) > 9 Then
                                        Exit Do
                                    End If
                                Loop
                                Mid$(cColors(iTemp), 1, 1) = UCase$(Mid$(cColors(iTemp), 1, 1))
                                If cColorList = "" Then
                                    cColorList = cColors(iTemp)
                                Else
                                    cColorList = cColorList & "," & cColors(iTemp)
                                End If
                                With rstColor
                                    .AddNew
                                    .Fields("ColorId") = iTemp
                                    .Fields("Color") = Left$(cColors(iTemp), .Fields("Color").Size)
                                    .Update
                                End With
                            End If
                        Next iTemp
                        frmMain.TestColors = cColorList
                        rstColor.Close
                        Set rstColor = Nothing
                    Case "txtSponsor"
                        SetVariable "Sponsors", txtSponsor.Text
                    Case "chkUseColors"
                        miUseColors = chkUseColors.Value
                        SetVariable "UseColors", Format$(miUseColors)
                        If miUseColors = 1 Then
                            frmMain.chkColor.Caption = TranslateCaption("&Groups / Colors", 0, False)
                        Else
                            frmMain.chkColor.Caption = TranslateCaption("&Groups", 0, False)
                        End If
                        iCaption = True
                    Case "chkClub"
                        miShowRidersClub = chkClub.Value
                        SetVariable "ShowRidersClub", Format$(miShowRidersClub)
                    Case "chkTeam"
                        miShowRidersTeam = chkTeam.Value
                        SetVariable "ShowRidersTeam", Format$(miShowRidersTeam)
                    Case "chkLK"
                        miShowRidersLK = chkLK.Value
                        SetVariable "ShowRidersLK", Format$(miShowRidersLK)
                    Case "chkHorseId"
                        miShowHorseId = chkHorseId.Value
                        SetVariable "ShowHorseId", Format$(miShowHorseId)
                    Case "chkHorseAge"
                        miShowHorseAge = chkHorseAge.Value
                        SetVariable "ShowHorseAge", Format$(miShowHorseAge)
                    Case "chkJudgesRanking"
                        miShowJudgesRanking = chkJudgesRanking.Value
                        SetVariable "ShowJudgesRanking", Format$(miShowJudgesRanking)
                    Case "chkLogDB"
                        miWriteLogDB = chkLogDB.Value
                        SetVariable "WriteLogDB", Format$(miWriteLogDB)
                    Case "chkIceSort"
                        miUseIceSort = chkIceSort.Value
                        If (Dir$(App.Path & "\" & SortingApplication) > "") And miUseIceSort = 1 Then
                            frmMain.cmdComposeGroups.Visible = False
                            frmMain.cmbGroupSize.Visible = True
                            frmMain.cmbGroupSize.Enabled = False
                            frmMain.cmdIceSort.Visible = True
                            frmMain.mnuEditStartOrder.Enabled = False
                        Else
                            miUseIceSort = 0
                            frmMain.cmdComposeGroups.Visible = True
                            frmMain.cmbGroupSize.Visible = True
                            frmMain.cmbGroupSize.Enabled = True
                            frmMain.cmdIceSort.Visible = False
                            If frmMain.dtaAlready.Recordset.RecordCount = 0 Then
                                frmMain.mnuEditStartOrder.Enabled = True
                            End If
                        End If
                        SetVariable "UseIceSort", Format$(miUseIceSort)
                    Case "chkExcelFiles"
                        miExcelFiles = chkExcelFiles.Value
                        SetVariable "CreateExcelFiles", Format$(miExcelFiles)
                    Case "chkHtmlFiles"
                        miHtmlFiles = chkHtmlFiles.Value
                        SetVariable "CreateHtmlFiles", Format$(miHtmlFiles)
                    Case "chkHighLights"
                        miUseHighLights = IIf(chkHighLights.Value = 0, False, True)
                        SetVariable "UseHighLights", Format$(miUseHighLights)
                    Case "chkFinalsSequence"
                        miFinalsSequence = chkFinalsSequence.Value
                        SetVariable "FinalsSequence", Format$(miFinalsSequence)
                    Case "chkMarkFinals"
                        miMarkFinalsInResultLists = chkMarkFinals.Value
                        SetVariable "MarkFinals", Format$(miMarkFinalsInResultLists)
                    Case "cmbBfinals"
                        miBFinalLevel = Val(cmbBFinals.Text)
                        SetVariable "BFinalLevel", Format$(miBFinalLevel)
                    Case "cmbCfinals"
                        miCFinalLevel = Val(cmbCFinals.Text)
                        SetVariable "CFinalLevel", Format$(miCFinalLevel)
                    End Select
                End If
            Next
        End If
    End If
    If iRestart = True Then
        frmMain.Tempvar = "*"
    ElseIf iCaption = True Then
        frmMain.ChangeCaption True
    End If
    On Local Error GoTo 0

End Sub

Private Sub txtColors_Change()
    txtColors.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub txtColors_GotFocus()
    cmdOK.Default = False
End Sub

Private Sub txtColors_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lTemp As Long
    If KeyCode = 13 Then
        lTemp = txtColors.SelStart
        txtColors.Text = Left$(txtColors.Text, lTemp) & vbCrLf & Mid$(txtColors.Text, lTemp + 1)
        txtColors.SelStart = lTemp + 2
    End If
End Sub

Private Sub txtColors_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNoColor_Change()
    txtNoColor.Tag = "*"
    Me.Tag = "*"
End Sub

Private Sub txtSponsor_Change()
    txtSponsor.Tag = "*"
    Me.Tag = "*"
    txtSponsor.ToolTipText = "Print sponsor in result lists like 'Tölt T1 " & txtSponsor.Text & " Datawerken'"
    lblSponsor.ToolTipText = txtSponsor.ToolTipText
End Sub
