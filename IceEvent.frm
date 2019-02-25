VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEvent 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Event"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWR 
      Height          =   405
      Left            =   1920
      TabIndex        =   8
      ToolTipText     =   "Code provided by FEIF for WorldRanking Events"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpLast 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64618497
      CurrentDate     =   39449
      MinDate         =   36526
   End
   Begin MSComCtl2.DTPicker dtpFirst 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Format          =   64618497
      CurrentDate     =   39449
      MinDate         =   36526
   End
   Begin VB.TextBox txtEvent 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "Name of the event"
      Top             =   90
      Width           =   3495
   End
   Begin VB.Label lblWR 
      Caption         =   "WorldRanking Code:"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblLast 
      Caption         =   "Last day"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label lblFirst 
      Caption         =   "First day"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   975
   End
   Begin VB.Label lblEvent 
      Caption         =   "Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    SetVariable "Event_name", txtEvent.Text
    SetVariable "Event_first", Format$(dtpFirst.Value, "dd-mm-yyyy")
    SetVariable "Event_last", Format$(dtpLast.Value, "dd-mm-yyyy")
    SetVariable "WR_Code", UCase$(txtWR.Text)

    If dtpLast.Value <> dtpFirst.Value Then
        If Month(dtpFirst.Value) = Month(dtpLast.Value) Then
            SetVariable "Event_date", Format$(dtpFirst.Value, "d") & "-" & Format$(dtpLast.Value, "d mmm yyyy")
        Else
            SetVariable "Event_date", Format$(dtpFirst.Value, "d mmm") & "-" & Format$(dtpLast.Value, "d mmm yyyy")
        End If
    ElseIf dtpFirst.Value > 0 Then
        SetVariable "Event_date", Format$(dtpFirst.Value, "d mmm yyyy")
    End If
    EventName = GetVariable("Event_name")
    If GetVariable("Event_date") <> "" Then
        EventName = GetVariable("Event_name") & " - " & GetVariable("Event_date")
    End If
    Unload Me
    
End Sub

Private Sub dtpFirst_Change()
    If dtpFirst.Value > dtpLast.Value Then
        dtpLast.Value = dtpFirst.Value
    End If
End Sub

Private Sub dtpLast_Change()
    If dtpLast.Value < dtpFirst.Value Then
        dtpLast.Value = dtpFirst.Value
    End If
End Sub

Private Sub Form_Load()
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
    
    DoEvents
    
    txtEvent.Text = GetVariable("Event_name")
    If GetVariable("Event_first") <> "" Then
        dtpFirst.Value = GetVariable("Event_first")
    Else
        dtpFirst.Value = Now
    End If
    If GetVariable("Event_last") <> "" Then
        dtpLast.Value = GetVariable("Event_last")
    Else
        dtpLast.Value = dtpFirst.Value
    End If
    txtWR.Text = GetVariable("WR_Code")

End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    
    With lblEvent
        .Top = 50
        .Left = 50
        .Width = ScaleWidth \ 3
    End With
    
    With txtEvent
        .Top = lblEvent.Top
        .Left = lblEvent.Left + lblEvent.Width + 50
        .Width = ScaleWidth - .Left - 50
    End With
    
    With lblFirst
        .Top = lblEvent.Top + lblEvent.Height + 50
        .Left = lblEvent.Left
        .Width = lblEvent.Width
    End With
    
    With dtpFirst
        .Top = lblFirst.Top
        .Left = lblFirst.Left + lblFirst.Width + 50
        .Width = ScaleWidth - .Left - 50
    End With
    
    With lblLast
        .Top = lblFirst.Top + lblFirst.Height + 50
        .Left = 50
        .Width = lblEvent.Width
    End With
    
    With dtpLast
        .Top = lblLast.Top
        .Left = lblLast.Left + lblLast.Width + 50
        .Width = ScaleWidth - .Left - 50
    End With
    
    With lblWR
        .Top = lblLast.Top + lblLast.Height + 50
        .Left = 50
        .Width = lblEvent.Width
    End With
    
    With txtWR
        .Top = lblWR.Top
        .Left = lblWR.Left + lblWR.Width + 50
        .Width = ScaleWidth - .Left - 50
    End With
    
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
