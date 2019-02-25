VERSION 5.00
Begin VB.Form frmToolBox 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "ToolBox"
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
   Begin VB.ListBox lstChecked 
      Height          =   2310
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton cmdNieuw 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Add new item to list"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      ToolTipText     =   "Make selection"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Cancel selection"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ListBox lstStandard 
      Height          =   2595
      ItemData        =   "IceToolBox.frx":0000
      Left            =   0
      List            =   "IceToolBox.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' General tool to list lists (like partcipants)
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

Public strQry As String
Public strQry2 As String
Public strList As String
Public intSingleChoice As Integer
Public intChecked As Integer
Public intReturnLen As Integer
Public intCheckAll As Integer

Private Sub cmdAnnuleren_Click()
   frmMain.Tempvar = ""
   intChecked = False
   Unload Me
End Sub

Private Sub cmdNieuw_Click()
   frmMain.Tempvar = "**"
   intChecked = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim iTemp As Integer
   If lstChecked.Visible = True Then
        If lstChecked.SelCount > 0 Then
           For iTemp = 0 To lstChecked.ListCount - 1
                If lstChecked.Selected(iTemp) = True Then
                    If intReturnLen > 0 Then
                        frmMain.Tempvar = frmMain.Tempvar & Left$(lstChecked.List(iTemp), intReturnLen) & "|"
                    Else
                        frmMain.Tempvar = frmMain.Tempvar & lstChecked.List(iTemp) & "|"
                    End If
                End If
           Next iTemp
        End If
   Else
        If lstStandard.SelCount > 0 Then
           frmMain.Tempvar = lstStandard.List(lstStandard.ListIndex)
        End If
   End If
   intChecked = False
   Unload Me
End Sub

Private Sub Form_Load()
   Dim rstQry As Recordset
   Dim rstQry2 As Recordset
   Dim ListBoxTabs(3) As Long
   Dim result As Long
   Dim cPrevious As String
   Dim cTemp As String
   Dim cTemp2 As String
   Dim iTemp As Integer
   
   ReadFormPosition Me
   ChangeFontSize Me, msFontSize
   
   'Set the tab stop points.
   ListBoxTabs(1) = lstStandard.Width \ 40 '200
   ListBoxTabs(2) = 2 * ListBoxTabs(1) 'out of sight
   ListBoxTabs(3) = 3 * ListBoxTabs(1) 'out of sight

   'Send LB_SETTABSTOPS StatusMessage to ListBox.
   result = SendMessage(lstStandard.hwnd, LB_SETTABSTOPS, UBound(ListBoxTabs) + 1, ListBoxTabs(1))
   result = SendMessage(lstChecked.hwnd, LB_SETTABSTOPS, UBound(ListBoxTabs) + 1, ListBoxTabs(1))
   
   TranslateControls Me
   
   'Refresh the ListBox control.
   lstStandard.Clear
   lstStandard.Refresh
   lstChecked.Clear
   lstChecked.Refresh
   If intChecked = True Then
        lstChecked.Visible = True
        lstStandard.Visible = False
   End If
   
   If strList <> "" Then
      Do While strList <> ""
         Parse cTemp, strList, "|"
         If cTemp <> "" Then
            lstStandard.AddItem cTemp
            lstChecked.AddItem cTemp
         End If
      Loop
   ElseIf strQry <> "" Then
      Set rstQry = mdbMain.OpenRecordset(strQry)
      DoEvents
      If rstQry.RecordCount > 0 Then
         Do While Not rstQry.EOF
            If rstQry.Fields(0) <> cPrevious Then
                lstStandard.AddItem rstQry.Fields(0)
                lstChecked.AddItem rstQry.Fields(0)
                If intCheckAll = True Then
                    lstChecked.Selected(lstChecked.NewIndex) = True
                End If
                cPrevious = rstQry.Fields(0)
            End If
            rstQry.MoveNext
         Loop
         If rstQry.RecordCount = 1 Then
            lstStandard.Selected(0) = True
            lstChecked.Selected(0) = True
         End If
         lstStandard.Text = frmMain.TestCode
         If strQry2 <> "" Then
            Set rstQry2 = mdbMain.OpenRecordset(strQry2)
            DoEvents
            If rstQry2.RecordCount > 0 Then
               Do While Not rstQry2.EOF
                    For iTemp = 0 To lstChecked.ListCount - 1
                        If lstChecked.List(iTemp) = rstQry2.Fields(0) Then
                            lstChecked.Selected(iTemp) = True
                            Exit For
                        End If
                    Next iTemp
                    rstQry2.MoveNext
               Loop
            End If
            rstQry2.Close
        End If
      Else
         lstStandard.AddItem Translate("Nothing found", mcLanguage)
         cmdOK.Visible = False
         cmdAnnuleren.Default = True
      End If
      rstQry.Close
   End If
   Set rstQry = Nothing
End Sub

Private Sub Form_Resize()
   Dim ListBoxTabs(1) As Long
   Static iBusy As Integer
      
   On Local Error Resume Next
   
   If iBusy = False Then
        iBusy = True
        With cmdAnnuleren
           .Top = ScaleHeight - 50 - .Height
           .Left = ScaleWidth - 50 - .Width
        End With
        
        With cmdOK
           .Top = cmdAnnuleren.Top
           .Left = cmdAnnuleren.Left - 50 - .Width
        End With
        
        With cmdNieuw
           .Top = cmdOK.Top
           .Left = cmdOK.Left - 50 - .Width
        End With
        With lstStandard
           .Width = ScaleWidth
           .Height = ScaleHeight - cmdAnnuleren.Height - 100
        End With
        With lstChecked
           .Width = ScaleWidth
           .Height = ScaleHeight - cmdAnnuleren.Height - 100
        End With
        
        ListBoxTabs(1) = lstStandard.Width \ 40 '200
        Call SendMessage(lstStandard.hwnd, LB_SETTABSTOPS, UBound(ListBoxTabs) + 1, ListBoxTabs(1))
        Call SendMessage(lstChecked.hwnd, LB_SETTABSTOPS, UBound(ListBoxTabs) + 1, ListBoxTabs(1))
        iBusy = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFormPosition Me
End Sub

Private Sub lstChecked_Click()
   Dim iTemp As Integer
   Dim lTemp As Long
   
   cmdOK.Enabled = True
End Sub

Private Sub lstChecked_ItemCheck(Item As Integer)
    Dim iTemp As Integer
   If intSingleChoice = True Then
      If lstChecked.SelCount > 1 Then
         For iTemp = 0 To lstChecked.ListCount - 1
            If iTemp <> Item Then
                lstChecked.Selected(iTemp) = False
            Else
                lstChecked.Selected(iTemp) = True
            End If
         Next iTemp
      End If
   End If

End Sub

Private Sub lstStandard_Click()
   cmdOK.Enabled = True
End Sub

Private Sub lstStandard_DblClick()
   DoEvents
   cmdOK_Click
End Sub
