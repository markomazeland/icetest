VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Pre&view"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      Picture         =   "IcePrint.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin RichTextLib.RichTextBox rtfPrint 
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      RightMargin     =   15
      TextRTF         =   $"IcePrint.frx":066A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1200
      Top             =   240
   End
   Begin VB.TextBox txtCounter 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtCounter"
      BuddyDispid     =   196612
      OrigLeft        =   1200
      OrigTop         =   120
      OrigRight       =   1440
      OrigBottom      =   855
      Max             =   100
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   375
      Left            =   -120
      TabIndex        =   2
      ToolTipText     =   "Print to printer"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Cancel printing"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblCounter 
      Alignment       =   1  'Right Justify
      Caption         =   "Copies:"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple RTF Wysiwyg Pinter
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

Public Repeat As Integer
Public fcFocus As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    SetMouseHourGlass
    
    Repeat = False
    
    frmMain.Tempvar = "Preview"
    
    cmdPrint.Enabled = False
    cmdPreview.Enabled = False
    cmdCancel.Enabled = False
    txtCounter.Enabled = False
    
    WriteFormPosition Me
    
    SetMouseNormal
    
    Unload Me
    

End Sub

Private Sub cmdPrint_Click()
    Dim iCopy As Integer
    Dim cTemp As String

    SetMouseHourGlass
    
    Repeat = False
    
    frmMain.Tempvar = "Print"
    
    cmdPrint.Enabled = False
    cmdPreview.Enabled = False
    cmdCancel.Enabled = False
    txtCounter.Enabled = False
    
    For iCopy = 1 To txtCounter.Text
        PrintRtf rtfPrint, 25, 15, 15, 10
    Next iCopy
    
    DoEvents
    
    If InStr(Caption, "[") > 0 Then
        cTemp = RTrim$(Left$(Caption, InStr(Caption, "[") - 1))
    Else
        cTemp = Caption
    End If
    WriteIniFile gcIniFile, "Print", cTemp, txtCounter.Text
    WriteFormPosition Me
    
    SetMouseNormal
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    Dim cTemp As String
    
    
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    DoEvents
    
    If fcFocus = "Preview" Then
        SetFocusTo cmdPreview
    Else
        SetFocusTo cmdPrint
    End If
    
    If Repeat = True Then
        cmdPrint_Click
    End If
End Sub

Private Sub Form_Resize()
    Static iBusy As Integer
    
    On Local Error Resume Next
    
    If iBusy = False Then
        iBusy = True
        With Me
            If ScaleWidth < cmdCancel.Width + cmdPrint.Width + cmdPreview.Width + 300 Then
                .Width = cmdCancel.Width + cmdPrint.Width + cmdPreview.Width + 300
            End If
        End With
        
        With Me.UpDown1
            .Top = 100
            .Left = ScaleWidth - 50 - .Width
            .Height = txtCounter.Height
        End With
        
        With Me.txtCounter
            .Top = UpDown1.Top
            .Left = UpDown1.Left - .Width
        End With
        
        With Me.lblCounter
            .Top = txtCounter.Top
            .Left = txtCounter.Left - 50 - .Width
        End With
        
        With Me.cmdCancel
            .Top = ScaleHeight - 50 - .Height
            .Left = ScaleWidth - 50 - .Width
        End With
        
        With Me.cmdPreview
            .Top = cmdCancel.Top
            .Left = cmdCancel.Left - 50 - .Width
        End With
        
        With Me.cmdPrint
            .Top = cmdCancel.Top
            .Left = cmdPreview.Left - 50 - .Width
        End With
        
        SetFocusTo cmdPrint
        iBusy = False
    End If
End Sub

Private Sub Picture1_DblClick()
    cmdPrint_Click
End Sub

Private Sub Timer1_Timer()
    If Repeat = True Then
        Hide
        cmdPrint_Click
    End If
End Sub
