VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Over ..."
   ClientHeight    =   3885
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   9840
   ClipControls    =   0   'False
   HelpContextID   =   200020000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2681.496
   ScaleMode       =   0  'User
   ScaleWidth      =   9240.269
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLogfile 
      Caption         =   "&Logfile"
      Height          =   255
      Left            =   8400
      TabIndex        =   3
      ToolTipText     =   "Display system log file"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdIniFile 
      Caption         =   "&Program"
      Height          =   255
      Left            =   8400
      TabIndex        =   2
      ToolTipText     =   "Display info about program settings"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   240
      Picture         =   "IceAbout.frx":0000
      ScaleHeight     =   990.29
      ScaleMode       =   0  'User
      ScaleWidth      =   1032.43
      TabIndex        =   4
      Top             =   240
      Width           =   1470
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   0
      ToolTipText     =   "Close this window"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info"
      Height          =   255
      Left            =   8400
      TabIndex        =   1
      ToolTipText     =   "Display info about your PC's settings"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblUser 
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   7935
   End
   Begin VB.Label lblCopyright 
      Caption         =   "(c) 1998 - Datawerken IT BV, Zeist"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   8055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   6535.8
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   7935
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   240
      Width           =   7935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   6535.8
      Y1              =   1490.87
      Y2              =   1490.87
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label lblDisclaimer 
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   8175
   End
End
Attribute VB_Name = "frmAbout"
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

Option Compare Text
Option Explicit
 

Private Sub cmdIniFile_Click()
    Shell "notepad " & gcIniFile, vbNormalFocus
End Sub

Private Sub cmdLogfile_Click()
    Shell "notepad " & App.Path & "\" & App.EXEName & ".Log", vbNormalFocus
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
    Dim cTemp As String
    
    ReadFormPosition Me
    
    TranslateControls Me
    
    Me.Caption = Translate("About", mcLanguage) & " " & App.ProductName
    Me.picIcon.Picture = frmMain.Icon
    lblVersion.Caption = Translate("Version", mcLanguage) & " " & App.Major & "." & App.Minor & " (build: " & Format$(App.Revision, "000") & IIf(mcVersionSwitch <> "feif", "/" & UCase$(mcVersionSwitch), "") & ")"
    lblTitle.Caption = App.ProductName
    lblDescription.Caption = App.FileDescription
    lblCopyright.Caption = App.LegalCopyright
    lblUser.Caption = MachineName & " - [" & UserName & "]"
    cTemp = GetVariable("FIPO Version") & "/" & GetVariable("FIPO")
    Me.lblDisclaimer.Caption = Translate("Current database", mcLanguage) & ": " & mcDatabaseName & vbCrLf & Translate("Sport Rules Version", mcLanguage) & ": " & cTemp
    Me.lblDisclaimer.Caption = Me.lblDisclaimer.Caption & vbCrLf & Translate("Current settings", mcLanguage) & ": " & vbCrLf & gcIniFile
    Me.lblDisclaimer.Caption = Me.lblDisclaimer.Caption & vbCrLf & gcIniHorseFile
End Sub
Public Sub StartSysInfo()
    On Local Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    On Local Error GoTo 0
    
    Exit Sub
SysInfoErr:
    MsgBox LoadResString(104), vbInformation
End Sub

Private Sub Form_Resize()

    With lblTitle
        .Left = 1000
        .Width = ScaleWidth - .Left - 50
        .Top = 50
    End With
    With lblVersion
        .Left = lblTitle.Left
        .Width = lblTitle.Width
        .Top = lblTitle.Top + lblTitle.Height + 50
    End With
    With lblDescription
        .Left = lblTitle.Left
        .Width = lblTitle.Width
        .Top = lblVersion.Top + lblVersion.Height + 50
    End With
    With lblUser
        .Left = lblTitle.Left
        .Width = lblTitle.Width
        .Top = lblDescription.Top + lblDescription.Height + 50
    End With
    With lblCopyright
        .Left = lblTitle.Left
        .Width = lblTitle.Width
        .Top = lblUser.Top + lblUser.Height + 50
    End With
    
    
End Sub
