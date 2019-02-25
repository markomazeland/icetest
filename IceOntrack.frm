VERSION 5.00
Begin VB.Form frmOntrack 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "On Track Data"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear list"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send track data"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txtTrackNumber 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   615
   End
   Begin VB.ListBox lstOntrack 
      Height          =   2205
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblTrackNumber 
      Caption         =   "Track number:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmOntrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form to provide participants' data for ontrack database interface
' Copyright (C) Lutz Lesener 2007
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
 
Private Sub cmdClear_Click()
    lstOntrack.Clear
End Sub

Private Sub cmdSend_Click()
    Dim tracknumber As Integer
    Dim t As Integer

    On Local Error Resume Next

    'make sure the track number is valid:
    tracknumber = CInt(txtTrackNumber.Text & "")
    If tracknumber < 1 Or tracknumber > 10 Then
        tracknumber = 1
    End If

    'Empty the ontrack table:
    mdbMain.Execute ("DELETE * FROM ontrack WHERE track = " & tracknumber & ";")
    
    For t = 0 To lstOntrack.ListCount - 1
        mysql = "INSERT INTO ontrack (sta, code, status, section, track, position) VALUES ('" & Left(lstOntrack.List(t), 3) & "', '" & frmMain.TestCode & "', " & frmMain.TestStatus & ", " & frmMain.TestSection & ", " & tracknumber & ", " & t + 1 & ");"
        mdbMain.Execute (mysql)
    Next t

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF8
            Call frmMain.mnuToolOutputOntrackdata_Click
    End Select
End Sub

Private Sub Form_Load()

    txtTrackNumber.Text = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmMain.mnuToolOutputOntrackdata_Click
End Sub
