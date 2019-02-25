VERSION 5.00
Begin VB.Form frmWR 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "FEIF WorldRanking"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWF 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox txtWR 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtWR 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtWR 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtWR 
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblWR 
      Caption         =   "Password (repeat):"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblWR 
      Caption         =   "Password:"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblWR 
      Caption         =   "User name:"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblWR 
      Caption         =   "Event code:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmWR"
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

Option Explicit
Option Compare Binary


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    SetVariable "WR_Code", UCase$(txtWR(0).Text)
    SetVariable "WF_User", txtWR(1).Text
    SetVariable "WF_Password", Encrypt(txtWR(2).Text, MakeKeyFromWrCode)
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    txtWF.Text = Translate("The following information is required to check the pedigree of horses in WorldFengur (FEIF WorldRanking events only; Internet connection required)", mcLanguage) & ":"
    txtWR(0).Text = GetVariable("WR_Code")
    txtWR(1).Text = GetVariable("WF_User")
    txtWR(2).Text = Encrypt(GetVariable("WF_Password"), MakeKeyFromWrCode)
    txtWR(3).Text = Encrypt(GetVariable("WF_Password"), MakeKeyFromWrCode)
    
    DoEvents

    
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    
    With lblWR(0)
        .Left = 50
        .Top = 50
        .Width = (ScaleWidth - 200) \ 2
    End With
    
    With txtWR(0)
        .Left = lblWR(0).Left + lblWR(0).Width + 100
        .Top = 50
        .Width = (ScaleWidth - 200) \ 2
    End With
    
    With txtWF
        .Left = 50
        .Width = (ScaleWidth - 100)
        .Top = txtWR(0).Top + txtWR(0).Height + 50
    End With
    
    For iTemp = 1 To 3
        With lblWR(iTemp)
            If iTemp = 1 Then
                .Top = txtWR(iTemp - 1).Top + txtWR(iTemp - 1).Height + txtWF.Height + 100
            Else
                .Top = txtWR(iTemp - 1).Top + txtWR(iTemp - 1).Height
            End If
            .Left = lblWR(0).Left
            .Width = (ScaleWidth - 200) \ 2
        End With
        
        With txtWR(iTemp)
            .Top = lblWR(iTemp).Top
            .Left = txtWR(0).Left
            .Width = (ScaleWidth - 200) \ 2
            
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

Private Sub txtWR_Change(Index As Integer)
    If txtWR(2).Text <> txtWR(3).Text Then
        cmdOK.Enabled = False
    Else
        cmdOK.Enabled = True
    End If
End Sub
