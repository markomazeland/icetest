VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPick 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Pick"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstAll 
      Height          =   450
      ItemData        =   "IcePick.frx":0000
      Left            =   120
      List            =   "IcePick.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstPicked 
      Height          =   300
      IntegralHeight  =   0   'False
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "List of selected items"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   375
      Begin VB.CommandButton cmdPick 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   0
         Picture         =   "IcePick.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Move item down"
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton cmdPick 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   0
         Picture         =   "IcePick.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Move item up"
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton cmdPick 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   0
         Picture         =   "IcePick.frx":0888
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Remove item from list"
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdPick 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "IcePick.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Add item to list"
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Confirm list of selected items"
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      ToolTipText     =   "Cancel all changes"
      Top             =   2400
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2820
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
End
Attribute VB_Name = "frmPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form to select a limited list of test

' Copyright (C) Marko Mazeland 2005
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

Public QryAll As String
Public QryPicked As String
Public FieldKey As String
Public FieldSeq As String
Public TableName As String

Private Sub lstAll_Change()
    If lstAll.Text <> "" Then
        cmdPick(0).Enabled = True
    Else
        cmdPick(0).Enabled = False
    End If
    SetFocusTo lstAll
End Sub

Private Sub lstAll_Click()
    If lstAll.Text <> "" Then
        cmdPick(0).Enabled = True
    Else
        cmdPick(0).Enabled = False
    End If
    SetFocusTo lstAll
End Sub

Private Sub lstAll_DblClick()
    PickItem
End Sub

Private Sub lstAll_GotFocus()
    If lstAll.Text <> "" Then
        cmdPick(0).Enabled = True
    Else
        cmdPick(0).Enabled = False
    End If
    SetFocusTo lstAll
End Sub


Private Sub lstAll_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PickItem
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rst As DAO.Recordset
    Dim iTemp As Integer
    
    mdbMain.Execute "Update [" & TableName & "] SET " & FieldSeq & "=0"
    If lstPicked.ListCount > 0 Then
        Set rst = mdbMain.OpenRecordset("SELECT " & FieldKey & "," & FieldSeq & " FROM [" & TableName & "]")
        If rst.RecordCount > 0 Then
            Do While Not rst.EOF
                For iTemp = 0 To lstPicked.ListCount - 1
                    If InStr(lstPicked.List(iTemp), " " & rst.Fields(FieldKey) & " ") > 0 Then
                        rst.Edit
                        rst.Fields(FieldSeq) = iTemp + 1
                        rst.Update
                        Exit For
                    End If
                Next iTemp
                rst.MoveNext
            Loop
        End If
    End If
    Unload Me
End Sub

Private Sub cmdPick_Click(Index As Integer)
    Select Case Index
    Case 0 ' add item
        PickItem
    Case 1 ' remove item
        RemoveItem
    Case 2 ' move one up
        MoveItem -1
    Case 3 ' move one down
        MoveItem 1
    End Select
End Sub

Private Sub cmdPick_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Select Case Index
        Case 0
            PickItem
        Case 1
            RemoveItem
        End Select
    End Select
End Sub

Private Sub Form_Load()
    Dim rstAll As DAO.Recordset
    Dim rstPicked As DAO.Recordset

    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    If QryAll <> "" Then
        Set rstAll = mdbMain.OpenRecordset(QryAll)
        If rstAll.RecordCount > 0 Then
            With rstAll
                Do While Not .EOF
                    lstAll.AddItem .Fields("cList")
                    lstAll.ItemData(lstAll.NewIndex) = rstAll.AbsolutePosition
                    .MoveNext
                Loop
            End With
        Else
            MsgBox Translate("No list available", mcLanguage)
        End If
        rstAll.Close
        Set rstAll = Nothing
    End If
    
    If QryPicked <> "" Then
        Set rstPicked = mdbMain.OpenRecordset(QryPicked)
        If rstPicked.RecordCount > 0 Then
            With rstPicked
                Do While Not .EOF
                    lstPicked.AddItem Format$(rstPicked.AbsolutePosition + 1, "00") & ": " & .Fields("cList")
                    .MoveNext
                Loop
            End With
        End If
        rstPicked.Close
        Set rstPicked = Nothing
    End If
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    With lstAll
        .Top = 0
        .Left = 0
        .Width = (ScaleWidth - fraButtons.Width) \ 2 - 50
        .Height = ScaleHeight - cmdCancel.Height - StatusBar1.Height - 100
    End With
    
    With fraButtons
        .Width = cmdPick(0).Height
        .Height = cmdPick(9).Height * 4
        .Top = lstAll.Height - .Height
        .Left = lstAll.Left + lstAll.Width + 50
    End With
    
    For iTemp = 0 To 3
        With cmdPick(iTemp)
            .Width = .Height
        End With
    Next iTemp
    
    With Me.lstPicked
        .Top = 0
        .Left = fraButtons.Left + fraButtons.Width + 50
        .Width = (ScaleWidth - fraButtons.Width) \ 2 - 50
        .Height = ScaleHeight - cmdCancel.Height - StatusBar1.Height - 100
    End With

    With cmdCancel
        .Top = ScaleHeight - cmdCancel.Height - StatusBar1.Height - 50
        .Left = ScaleWidth - .Width - 50
    End With
    
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - 50 - .Width
    End With
    
    With cmdAdd
        .Top = cmdOK.Top
        .Left = cmdOK.Left - 50 - .Width
    End With
    
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFormPosition Me
End Sub
Private Sub PickItem()
    Dim iTemp As Integer
    
    If lstAll.ListCount > 1 Then
        For iTemp = 0 To lstPicked.ListCount - 1
            If lstPicked.List(iTemp) = lstAll.Text Then
                MsgBox lstAll.Text & " " & Translate("can only be added once!", mcLanguage), vbInformation
                Exit Sub
            End If
        Next iTemp
        lstPicked.AddItem Format$(lstPicked.ListCount + 1, "00: ") & lstAll.Text
        lstAll.RemoveItem lstAll.ListIndex
        lstAll.Text = ""
        cmdPick(0).Enabled = False
    Else
        Beep
    End If
    
End Sub
Private Sub RemoveItem()
    Dim Item As String
        
    'move item back to list of available items
    Item = Mid$(lstPicked.List(lstPicked.ListIndex), 5)
    lstAll.AddItem Item
    lstPicked.RemoveItem (lstPicked.ListIndex)
    
    ReNumberItems
    
    cmdPick(1).Enabled = False
    cmdPick(2).Enabled = False
    cmdPick(3).Enabled = False
    
    
End Sub

Private Sub lstPicked_Click()
    If lstPicked.SelCount > 0 Then
        cmdPick(1).Enabled = True
        If lstPicked.ListCount > 1 Then
            ' button up
            If lstPicked.ListIndex > 0 Then
                cmdPick(2).Enabled = True
            Else
                cmdPick(2).Enabled = False
            End If
            ' button down
            If lstPicked.ListIndex < lstPicked.ListCount - 1 Then
                cmdPick(3).Enabled = True
            Else
                cmdPick(3).Enabled = False
            End If
        End If
    Else
        cmdPick(1).Enabled = False
        cmdPick(2).Enabled = False
        cmdPick(3).Enabled = False
    End If
    SetFocusTo lstPicked
End Sub

Private Sub lstPicked_DblClick()
    RemoveItem
End Sub

Private Sub ReNumberItems()
    Dim iTemp As Integer
    Dim cTemp As String
    
    'correct numbering of list
    If lstPicked.ListCount > 0 Then
        For iTemp = 0 To lstPicked.ListCount - 1
            If Left$(lstPicked.List(iTemp), 2) <> Format$(iTemp + 1, "00") Then
                cTemp = lstPicked.List(iTemp)
                Mid$(cTemp, 1, 2) = Format$(iTemp + 1, "00")
                lstPicked.List(iTemp) = cTemp
            End If
        Next iTemp
    End If
End Sub
Private Sub MoveItem(Direction As Integer)
    Dim iTemp As Integer
    Dim cTemp As String
    
    iTemp = lstPicked.ListIndex
    cTemp = lstPicked.List(iTemp)
    lstPicked.RemoveItem iTemp
    lstPicked.AddItem cTemp, iTemp + Direction
    lstPicked.ListIndex = iTemp + Direction
         
    ReNumberItems
    
    lstPicked_Click
    
End Sub
