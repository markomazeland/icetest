VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmParticipantTkr 
   Caption         =   "Participant"
   ClientHeight    =   8385
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Width           =   1140
   End
   Begin VB.CommandButton cmdSavePart 
      Caption         =   "save"
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddPersToNewParticipant 
      Caption         =   "select person for participant"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdAddHorsToNewParticipant 
      Caption         =   "select horse for participant"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancelAddHors 
      Caption         =   "cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddHorsToParticipant 
      Caption         =   "select horse for participant"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSaveHors 
      Caption         =   "save"
      Height          =   375
      Left            =   9480
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSavePers 
      Caption         =   "save"
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelAddPers 
      Caption         =   "cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddPersToParticipant 
      Caption         =   "select person for participant"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancelAddTest 
      Caption         =   "cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddTestToParticipant 
      Caption         =   "add test to participant"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame frmShow 
      Height          =   7695
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtDetail 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblDetail 
         Caption         =   "label"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParticipantTkr.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParticipantTkr.frx":0383
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParticipantTkr.frx":06F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParticipantTkr.frx":0A85
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParticipantTkr.frx":0E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParticipantTkr.frx":11AF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treParticipants 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   13361
      _Version        =   393217
      Indentation     =   882
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnuRootPart 
      Caption         =   "ParticipantRoot"
      Visible         =   0   'False
      Begin VB.Menu mnuRootPartAdd 
         Caption         =   "add Participant"
      End
      Begin VB.Menu mnuRootPartOrderBy 
         Caption         =   "order by..."
         Begin VB.Menu mnuRootPartOrderBySta 
            Caption         =   "starting number"
         End
         Begin VB.Menu mnuRootPartOrderByRiderName 
            Caption         =   "Rider Name"
         End
         Begin VB.Menu mnuRootPartOrderByHorseName 
            Caption         =   "Horse Name"
         End
      End
   End
   Begin VB.Menu mnuRootPers 
      Caption         =   "PersonRoot"
      Visible         =   0   'False
      Begin VB.Menu mnuRootPersAdd 
         Caption         =   "add Person"
      End
   End
   Begin VB.Menu mnuRootHors 
      Caption         =   "HorseRoot"
      Visible         =   0   'False
      Begin VB.Menu mnuRootHorsAdd 
         Caption         =   "add Horse"
      End
   End
   Begin VB.Menu mnuNodePart 
      Caption         =   "Participant"
      Visible         =   0   'False
      Begin VB.Menu mnuNodePartAddTest 
         Caption         =   "add Test"
      End
      Begin VB.Menu mnuNodePartSelectPers 
         Caption         =   "select Person"
      End
      Begin VB.Menu mnuNodePartSelectHors 
         Caption         =   "select Horse"
      End
      Begin VB.Menu mnuNodePartEdit 
         Caption         =   "edit Participant"
      End
      Begin VB.Menu mnuNodePartRemove 
         Caption         =   "remove Participant"
      End
   End
   Begin VB.Menu mnuNodePers 
      Caption         =   "Person"
      Visible         =   0   'False
      Begin VB.Menu mnuNodePersEdit 
         Caption         =   "edit Person"
      End
      Begin VB.Menu mnuNodePersShow 
         Caption         =   "show Person"
      End
   End
   Begin VB.Menu mnuNodeHors 
      Caption         =   "Horse"
      Visible         =   0   'False
      Begin VB.Menu mnuNodeHorsEdit 
         Caption         =   "edit Horse"
      End
      Begin VB.Menu mnuNodeHorsShow 
         Caption         =   "show Horse"
      End
   End
   Begin VB.Menu mnuNodeEntr 
      Caption         =   "Test"
      Visible         =   0   'False
      Begin VB.Menu mnuNodeEntrEdit 
         Caption         =   "edit Test"
      End
      Begin VB.Menu mnuNodeEntrRemove 
         Caption         =   "remove Test"
      End
   End
End
Attribute VB_Name = "frmParticipantTkr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim gStrNodeKey As String

Dim gStrSta As String
Dim gStrTest As String
Dim gStrPers As String
Dim gStrHors As String

Dim gStrOrderBy As String

Private Sub Form_Load()
    ' when load form
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
   
    hideFrmShow
    gStrOrderBy = "ORDER BY Persons.Name_Last"
    Call showAllParticpants(gStrNodeKey, gStrOrderBy)
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    With Me.treParticipants
        .Left = 50
        .Top = 50
        .Width = ScaleWidth \ 2 - 100
        .Height = ScaleHeight - 100
    End With
    
    With Me.frmShow
        .Left = Me.treParticipants.Left + Me.treParticipants.Width + 50
        .Top = 50
        .Width = ScaleWidth \ 2 - 100
        .Height = ScaleHeight - 100
    End With
    
    For iTemp = 0 To txtDetail.Count
        txtDetail(iTemp).Width = txtDetail(iTemp).Container.Width - txtDetail(iTemp).Left - 50
    Next iTemp
    
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteFormPosition Me
End Sub

Private Sub treParticipants_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' when click on one treeview-entry
    
    Dim strNodeType As String
    
    Call hideFrmShow
    ' get type of the selected treeview item
    strNodeType = getNodeType()
    ' if right mouse button
    If Button = vbRightButton Then
        ' select the popUpMenu
        Select Case strNodeType
        Case "rootPart"
                PopupMenu mnuRootPart, vbPopupMenuLeftAlign
        Case "rootPers"
                PopupMenu mnuRootPers, vbPopupMenuLeftAlign
        Case "rootHors"
                PopupMenu mnuRootHors, vbPopupMenuLeftAlign
        Case "nodePart"
                PopupMenu mnuNodePart, vbPopupMenuLeftAlign
        Case "nodePers"
                PopupMenu mnuNodePers, vbPopupMenuLeftAlign
        Case "nodeHors"
                PopupMenu mnuNodeHors, vbPopupMenuLeftAlign
        Case "nodeEntr"
                PopupMenu mnuNodeEntr, vbPopupMenuLeftAlign
        End Select
     End If

End Sub

Private Sub treParticipants_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strNodeType As String
    
    ' get type of the selected treeview item
    strNodeType = getNodeType()
    ' if left mouse button
    If Button = vbLeftButton Then
        ' select the popUpMenu
        Select Case strNodeType
        Case "rootPart"
        Case "rootPers"
        Case "rootHors"
        Case "nodePart"
            Call showParticipant(gStrSta, False)
        Case "nodePers"
            Call showPerson(gStrPers, False)
        Case "nodeHors"
            Call showHorse(gStrHors, False)
        Case "nodeEntr"
            Call showEntry
        End Select
     End If
End Sub

Private Sub showAllParticpants(strActiveNode As String, strOrderBy As String)

    ' recordSets
    Dim rstPart As DAO.Recordset
    Dim rstEntr As DAO.Recordset
    ' queryStrings
    Dim qryPart As String
    Dim qryEntr As String
    Dim qryEntrResult As String
    ' nodes
    Dim nodeRoot As Node
    Dim nodePart As Node
    Dim nodePers As Node
    Dim nodeHors As Node
    Dim nodeEntr As Node
    ' strings
    Dim strNodeText As String
    Dim strNodeKey As String
    
    ' treeview leeren
    treParticipants.Nodes.Clear
    
    ' set queryStrings
    qryPart = "SELECT DISTINCT Participants.PersonId, Participants.HorseId, Participants.Sta, Persons.Name_First, Persons.Name_Last, Horses.Name_Horse FROM (Participants INNER JOIN Persons ON Participants.PersonId=Persons.PersonId) INNER JOIN Horses ON Participants.HorseId=Horses.HorseId " & strOrderBy
    qryEntr = "SELECT DISTINCT Entries.Sta, Entries.Code, Entries.RR, Tests.Test FROM Entries INNER JOIN Tests ON Entries.Code=Tests.Code WHERE Entries.Sta = """
    ' nodeRootPart
    ' nodeText
    strNodeText = Translate("Participants", mcLanguage)
    ' nodeKey
    strNodeKey = "rootPart"
    ' set node
    Set nodeRoot = treParticipants.Nodes.Add(, tvwChild, strNodeKey, strNodeText, 1)
    nodeRoot.Expanded = True
    ' get recordSet participants
    Set rstPart = mdbMain.OpenRecordset(qryPart)
    ' fill treeview for all participants
    While Not rstPart.EOF
        ' nodeParticipant
        ' nodeText
        Select Case strOrderBy
        Case "ORDER BY Participants.Sta"
            strNodeText = _
            rstPart.Fields("Sta") & " - " & _
            rstPart.Fields("Name_Last") & ", " & _
            rstPart.Fields("Name_First") & " - " & _
            rstPart.Fields("Name_Horse")
        Case "ORDER BY Persons.Name_Last"
            strNodeText = _
            rstPart.Fields("Name_Last") & ", " & _
            rstPart.Fields("Name_First") & " - " & _
            rstPart.Fields("Name_Horse") & " - " & _
            rstPart.Fields("Sta")
        Case "ORDER BY Horses.Name_Horse"
            strNodeText = _
            rstPart.Fields("Name_Horse") & " - " & _
            rstPart.Fields("Name_Last") & ", " & _
            rstPart.Fields("Name_First") & " - " & _
            rstPart.Fields("Sta")
        End Select
        ' nodeKey
        strNodeKey = _
        "nodePart~" & _
        rstPart.Fields("Sta")
        ' setNode
        Set nodePart = treParticipants.Nodes.Add(nodeRoot, tvwChild, strNodeKey, strNodeText, 3)
        nodePart.Expanded = True
        ' nodePerson
        ' nodeText
        strNodeText = _
        rstPart.Fields("Name_First") & " " & _
        rstPart.Fields("Name_Last")
        ' nodeKey
        strNodeKey = _
        "nodePers~" & _
        rstPart.Fields("PersonId") & "~" & _
        rstPart.Fields("Sta")
        ' setNode
        Set nodePers = treParticipants.Nodes.Add(nodePart, tvwChild, strNodeKey, strNodeText, 4)
        nodePers.Expanded = False
        ' nodeHorse
        ' nodeText
        strNodeText = _
        rstPart.Fields("Name_Horse")
        ' nodeKey
        strNodeKey = _
        "nodeHors~" & _
        rstPart.Fields("HorseId") & "~" & _
        rstPart.Fields("Sta")
        ' setNode
        Set nodeHors = treParticipants.Nodes.Add(nodePart, tvwChild, strNodeKey, strNodeText, 5)
        nodePers.Expanded = False
        ' getEntries
        qryEntrResult = qryEntr & rstPart.Fields("Sta") & """"
        Set rstEntr = mdbMain.OpenRecordset(qryEntrResult)
        ' get all test for participant
        While Not rstEntr.EOF
            ' nodeEntry
            ' nodeText
            strNodeText = _
            rstEntr.Fields("Code") & " - " & _
            rstEntr.Fields("Test")
            ' nodeKey
            strNodeKey = _
            "nodeEntr~" & _
            rstPart.Fields("Sta") & "~" & _
            rstEntr.Fields("Code") & "~" & _
            rstEntr.Fields("RR")
            Set nodeEntr = treParticipants.Nodes.Add(nodePart, tvwChild, strNodeKey, strNodeText, 6)
            ' next test
            rstEntr.MoveNext
        Wend
        ' next participant
        rstPart.MoveNext
    Wend
    ' participant selected in treeview
    If strActiveNode <> "" Then
        Set treParticipants.SelectedItem = treParticipants.Nodes(strActiveNode)
        treParticipants.SelectedItem.Expanded = True
        treParticipants.SelectedItem.EnsureVisible
    End If

End Sub

Private Sub showAllTests()

    ' recordSets
    Dim rstTest As DAO.Recordset
    ' queryStrings
    Dim qryTest As String
    ' nodes
    Dim nodeRoot As Node
    Dim nodeTest As Node
    ' strings
    Dim strNodeText As String
    Dim strNodeKey As String
    
    ' treeview leeren
    treParticipants.Nodes.Clear
    
    ' set queryStrings
    qryTest = "SELECT * FROM Tests ORDER BY Code"
    ' nodeRootPart
    ' nodeText
    strNodeText = Translate("Tests", mcLanguage)
    ' nodeKey
    strNodeKey = "rootTest"
    ' set node
    Set nodeRoot = treParticipants.Nodes.Add(, tvwChild, strNodeKey, strNodeText, 1)
    nodeRoot.Expanded = True
    ' get recordSet tests
    Set rstTest = mdbMain.OpenRecordset(qryTest, dbOpenDynaset)
    ' fill treeview with all tests
    While Not rstTest.EOF
        ' nodeTest
        ' nodeText
        strNodeText = _
        rstTest.Fields("Code") & " - " & _
        rstTest.Fields("Test")
        ' nodeKey
        strNodeKey = _
        "nodeTest~" & _
        rstTest.Fields("Code")
        ' setNode
        Set nodeTest = treParticipants.Nodes.Add(nodeRoot, tvwChild, strNodeKey, strNodeText, 6)
        nodeTest.Expanded = False
        ' next test
        rstTest.MoveNext
    Wend
    
End Sub

Private Sub showAllPersons(strActiveNode As String)

    ' recordSets
    Dim rstPers As DAO.Recordset
    ' queryStrings
    Dim qryPers As String
    ' nodes
    Dim nodeRoot As Node
    Dim nodePers As Node
    ' strings
    Dim strNodeText As String
    Dim strNodeKey As String
    ' node
    Dim objNode As Node
    
    ' treeview leeren
    treParticipants.Nodes.Clear
    
    ' set queryStrings
    qryPers = "SELECT * FROM Persons ORDER BY Name_Last"
    ' nodeRootPart
    ' nodeText
    strNodeText = Translate("Persons", mcLanguage)
    ' nodeKey
    strNodeKey = "rootPers"
    ' set node
    Set nodeRoot = treParticipants.Nodes.Add(, tvwChild, strNodeKey, strNodeText, 1)
    nodeRoot.Expanded = True
    ' get recordSet persons
    Set rstPers = mdbMain.OpenRecordset(qryPers, dbOpenDynaset)
    ' fill treeview with all tests
    While Not rstPers.EOF
        ' nodePerson
        ' nodeText
        strNodeText = _
        rstPers.Fields("Name_Last") & ", " & _
        rstPers.Fields("Name_First")
        ' nodeKey
        strNodeKey = _
        "nodePers~" & _
        rstPers.Fields("PersonId")
        ' setNode
        Set nodePers = treParticipants.Nodes.Add(nodeRoot, tvwChild, strNodeKey, strNodeText, 4)
        nodePers.Expanded = False
        ' next person
        rstPers.MoveNext
    Wend
    ' person selected in treeview
    If strActiveNode <> "" Then
        Set treParticipants.SelectedItem = treParticipants.Nodes(strActiveNode)
        treParticipants.SelectedItem.EnsureVisible
    End If
    
End Sub

Private Sub showAllHorses()

    ' recordSets
    Dim rstHors As DAO.Recordset
    ' queryStrings
    Dim qryHors As String
    ' nodes
    Dim nodeRoot As Node
    Dim nodeHors As Node
    ' strings
    Dim strNodeText As String
    Dim strNodeKey As String
    
    ' treeview leeren
    treParticipants.Nodes.Clear
    
    ' set queryStrings
    qryHors = "SELECT * FROM Horses ORDER BY Name_Horse"
    ' nodeRootPart
    ' nodeText
    strNodeText = Translate("Horses", mcLanguage)
    ' nodeKey
    strNodeKey = "rootHors"
    ' set node
    Set nodeRoot = treParticipants.Nodes.Add(, tvwChild, strNodeKey, strNodeText, 1)
    nodeRoot.Expanded = True
    ' get recordSet persons
    Set rstHors = mdbMain.OpenRecordset(qryHors, dbOpenDynaset)
    ' fill treeview with all tests
    While Not rstHors.EOF
        ' nodeHorse
        ' nodeText
        If Not IsNull(rstHors.Fields("Name_Horse")) Then
            strNodeText = _
            rstHors.Fields("Name_Horse")
        Else
            strNodeText = _
            Translate("unknown", mcLanguage)
        End If
        ' nodeKey
        strNodeKey = _
        "nodeHorse~" & _
        rstHors.Fields("HorseId")
        ' setNode
        Set nodeHors = treParticipants.Nodes.Add(nodeRoot, tvwChild, strNodeKey, strNodeText, 5)
        nodeHors.Expanded = False
        ' next horse
        rstHors.MoveNext
    Wend
    
End Sub

Private Sub showPerson(strQueryKey As String, boolEditPerson As Integer)
    ' recordSet
    Dim rstPers As DAO.Recordset
    ' queryString
    Dim qryPers As String
    Dim qryPersResult As String
    
    Dim i As Integer
    
    ' set queryString
    qryPers = "SELECT * FROM Persons WHERE Persons.PersonId = """
    ' getEntries
    qryPersResult = qryPers & strQueryKey & """"
    Set rstPers = mdbMain.OpenRecordset(qryPersResult)
    
    frmShow.Visible = True
    frmShow.Caption = "Person"
    frmShow.Refresh
    
    showDetails rstPers
    Exit Sub
    
    For i = 0 To rstPers.Fields.Count - 1
        Me.Controls("lbl_" & CStr(i)).Caption = Translate(rstPers.Fields(i).Name, mcLanguage)
        If rstPers.RecordCount > 0 Then
            If Not IsNull(rstPers.Fields(i).Value) Then
                Me.Controls("txt_" & CStr(i)).Text = rstPers.Fields(i).Value
            End If
        Else
            If UCase(rstPers.Fields(i).Name) = "PERSONID" Then
                Me.Controls("txt_" & CStr(i)).Text = gStrPers
            End If
        End If
        Me.Controls("txt_" & CStr(i)).Visible = True
        If boolEditPerson And UCase(rstPers.Fields(i).Name) <> "PERSONID" Then
            Me.Controls("txt_" & CStr(i)).Enabled = True
        Else
            Me.Controls("txt_" & CStr(i)).Enabled = False
        End If
    Next i
    If boolEditPerson Then
        Me.Controls("cmdSavePers").Visible = True
        txt_8.SetFocus
    Else
        Me.Controls("cmdSavePers").Visible = False
    End If
    frmShow.Refresh
    
End Sub

Private Sub showHorse(strQueryKey As String, boolEditHorse As Integer)
    ' recordSet
    Dim rstHors As DAO.Recordset
    ' queryString
    Dim qryHors As String
    Dim qryHorsResult As String
    
    Dim i As Integer
    
    ' set queryString
    qryHors = "SELECT * FROM Horses WHERE Horses.HorseId = """
    ' getEntries
    qryHorsResult = qryHors & strQueryKey & """"
    Set rstHors = mdbMain.OpenRecordset(qryHorsResult)
    
    frmShow.Visible = True
    frmShow.Caption = "Horse"
    frmShow.Refresh
    
    For i = 0 To rstHors.Fields.Count - 1
        Me.Controls("lbl_" & CStr(i)).Caption = Translate(rstHors.Fields(i).Name, mcLanguage)
        If rstHors.RecordCount > 0 Then
            If Not IsNull(rstHors.Fields(i).Value) Then
                Me.Controls("txt_" & CStr(i)).Text = rstHors.Fields(i).Value
            End If
        Else
            If UCase(rstHors.Fields(i).Name) = "HORSEID" Then
                Me.Controls("txt_" & CStr(i)).Text = gStrHors
            End If
        End If
        Me.Controls("txt_" & CStr(i)).Visible = True
        If boolEditHorse And UCase(rstHors.Fields(i).Name) <> "HORSEID" Then
            Me.Controls("txt_" & CStr(i)).Enabled = True
        Else
            Me.Controls("txt_" & CStr(i)).Enabled = False
        End If
    Next i
    If boolEditHorse Then
        Me.Controls("cmdSaveHors").Visible = True
        txt_15.SetFocus
    Else
        Me.Controls("cmdSaveHors").Visible = False
    End If
    frmShow.Refresh
        
End Sub

Private Sub showParticipant(strQueryKey As String, boolEditParticipant As Integer)
    ' recordSet
    Dim rstPart As DAO.Recordset
    ' queryString
    Dim qryPart As String
    Dim qryPartResult As String
    
    Dim i As Integer
    
    ' set queryString
    qryPart = "SELECT * FROM Participants WHERE Participants.Sta = """
    ' getEntries
    qryPartResult = qryPart & strQueryKey & """"
    Set rstPart = mdbMain.OpenRecordset(qryPartResult)
    
    frmShow.Visible = True
    frmShow.Caption = "Participant"
    frmShow.Refresh
    ' show all fields
    
    showDetails rstPart
    
    Exit Sub
    
    ' build up details frame
    For i = 0 To rstPart.Fields.Count - 1
        Me.Controls("lbl_" & CStr(i)).Caption = Translate(rstPart.Fields(i).Name, mcLanguage)
        If rstPart.RecordCount > 0 Then
            If Not IsNull(rstPart.Fields(i).Value) Then
                Me.Controls("txt_" & CStr(i)).Text = rstPart.Fields(i).Value
            End If
        Else
            If UCase(rstPart.Fields(i).Name) = "STA" Then
                Me.Controls("txt_" & CStr(i)).Text = gStrSta
            End If
        End If
        Me.Controls("txt_" & CStr(i)).Visible = True
        If boolEditParticipant And UCase(rstPart.Fields(i).Name) <> "HORSEID" And UCase(rstPart.Fields(i).Name) <> "PERSONID" Then
            Me.Controls("txt_" & CStr(i)).Enabled = True
        Else
            Me.Controls("txt_" & CStr(i)).Enabled = False
        End If
    Next i
    If boolEditParticipant Then
        Me.Controls("cmdSavePart").Visible = True
        txt_1.SetFocus
    Else
        Me.Controls("cmdSavePart").Visible = False
    End If
    frmShow.Refresh
        
End Sub

Private Sub showEntry()
    ' recordSet
    Dim rstEntr As DAO.Recordset
    ' queryString
    Dim qryEntr As String
    
    Dim i As Integer
    
    ' set queryString
    'qryEntr = "SELECT DISTINCT Entries.Sta, Entries.Code, Entries.RR, Tests.Test, Marks.Mark1, Marks.Mark2, Marks.Mark3, Marks.Mark4, Marks.Mark5, Marks.Score FROM (Entries INNER JOIN Tests ON Entries.Code=Tests.Code) INNER JOIN Marks ON (Marks.Code=Entries.Code) AND (Marks.Sta=Entries.Sta)WHERE Entries.Sta=""" & gStrSta & """ AND Entries.Code = """ & gStrTest & """"
    qryEntr = "SELECT Entries.Sta,  Entries.Code & "" - "" & Tests.Test AS Test, Entries.RR, Entries.Color, Entries.Group, Entries.Position, Entries.Late_Entry  from ( Entries  INNER JOIN Tests ON Entries.Code=Tests.Code) WHERE Entries.Sta=""" & gStrSta & """ AND Entries.Code = """ & gStrTest & """"

    ' getEntries
    Set rstEntr = mdbMain.OpenRecordset(qryEntr)
    
    frmShow.Visible = True
    frmShow.Caption = "Test"
    frmShow.Refresh
    ' show all fields
    For i = 0 To rstEntr.Fields.Count - 1
        Me.Controls("lbl_" & CStr(i)).Caption = Translate(rstEntr.Fields(i).Name, mcLanguage)
        If rstEntr.RecordCount > 0 Then
            If Not IsNull(rstEntr.Fields(i).Value) Then
                Me.Controls("txt_" & CStr(i)).Text = rstEntr.Fields(i).Value
            End If
        Else
            If UCase(rstEntr.Fields(i).Name) = "STA" Then
                Me.Controls("txt_" & CStr(i)).Text = gStrSta
            End If
        End If
        Me.Controls("txt_" & CStr(i)).Visible = True
        Me.Controls("txt_" & CStr(i)).Enabled = False
    Next i
    Me.Controls("cmdSavePart").Visible = False
    frmShow.Refresh
        
End Sub

Private Sub hideFrmShow()
    Dim i As Integer

    For i = 0 To 19
        Me.Controls("txt_" & CStr(i)).Text = ""
        Me.Controls("txt_" & CStr(i)).Enabled = False
        Me.Controls("txt_" & CStr(i)).Visible = False
        Me.Controls("lbl_" & CStr(i)).Caption = ""
    Next i
    Me.Controls("cmdSavePers").Visible = False
    Me.Controls("cmdSaveHors").Visible = False
    Me.Controls("cmdSavePart").Visible = False

    frmShow.Visible = False
End Sub

Private Sub addPersToParticipant()
    Dim rstPart As DAO.Recordset
    Dim i As Integer
    
    Set rstPart = mdbMain.OpenRecordset("SELECT * FROM Participants WHERE Sta =""" & gStrSta & """")
    If rstPart.RecordCount = 0 Then
        With rstPart
            .AddNew
        End With
    Else
        With rstPart
            .Edit
        End With
    End If
    With rstPart
        .Fields("PersonID") = gStrPers
        .Fields("Sta") = gStrSta
        .Update
        .Close
    End With
End Sub

Private Sub addHorsToParticipant()
    Dim rstPart As DAO.Recordset
    Dim i As Integer
    
    Set rstPart = mdbMain.OpenRecordset("SELECT * FROM Participants WHERE Sta =""" & gStrSta & """")
    If rstPart.RecordCount = 0 Then
        With rstPart
            .AddNew
        End With
    Else
        With rstPart
            .Edit
        End With
    End If
    With rstPart
        .Fields("HorseID") = gStrHors
        .Fields("Sta") = gStrSta
        .Update
        .Close
    End With
End Sub

Private Sub mnuRootPartAdd_click()
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

    gStrSta = cTemp
    
    gStrPers = ""
    Call showAllPersons("")
    cmdAddPersToNewParticipant.Visible = True
    cmdCancelAddPers.Visible = True
End Sub

Private Sub mnuRootPartOrderBySta_click()
    gStrOrderBy = "ORDER BY Participants.Sta"
    Call showAllParticpants(gStrNodeKey, gStrOrderBy)
End Sub

Private Sub mnuRootPartOrderByRiderName_click()
    gStrOrderBy = "ORDER BY Persons.Name_Last"
    Call showAllParticpants(gStrNodeKey, gStrOrderBy)
End Sub

Private Sub mnuRootPartOrderByHorseName_click()
    gStrOrderBy = "ORDER BY Horses.Name_Horse"
    Call showAllParticpants(gStrNodeKey, gStrOrderBy)
End Sub

Private Sub mnuRootPersAdd_Click()
    gStrPers = CreatePersonId()
    Call showPerson(gStrPers, True)
End Sub

Private Sub mnuRootHorsAdd_Click()
    gStrHors = CreateHorseId()
    Call showHorse(gStrHors, True)
End Sub

Private Sub mnuNodePersShow_click()
    Call showPerson(gStrPers, False)
End Sub

Private Sub mnuNodePersEdit_click()
    Call showPerson(gStrPers, True)
End Sub

Private Sub mnuNodeHorsShow_click()
    Call showHorse(gStrHors, False)
End Sub

Private Sub mnuNodeHorsEdit_click()
    Call showHorse(gStrHors, True)
End Sub

Private Sub mnuNodePartAddTest_click()
    Call showAllTests
    cmdAddTestToParticipant.Visible = True
    cmdCancelAddTest.Visible = True
End Sub

Private Sub mnuNodePartSelectPers_click()
    Call showAllPersons("")
    cmdAddPersToParticipant.Visible = True
    cmdCancelAddPers.Visible = True
End Sub

Private Sub mnuNodePartSelectHors_click()
    Call showAllHorses
    cmdAddHorsToParticipant.Visible = True
    cmdCancelAddHors.Visible = True
End Sub

Private Sub mnuNodeEntrRemove_Click()
    Dim iKey As Integer
    Dim rstOK As Recordset
   
    If gStrSta <> "" Then
        iKey = MsgBox(Translate("You may remove data for the selected participant." & vbCrLf & "- select 'Yes' to remove the test completely " & vbCrLf & "- select 'No' to remove marks only" & vbCrLf & "- select 'Cancel' to keep this participant", mcLanguage), vbExclamation + vbYesNoCancel + vbDefaultButton3, Translate("Remove", mcLanguage) & " " & gStrTest & " " & Translate("from", mcLanguage) & " " & gStrSta)
        If iKey = vbYes Then
            GoSub RemoveEntry
            GoSub RemoveMarks
        ElseIf iKey = vbNo Then
            GoSub RemoveMarks
        End If
        Call showAllParticpants("nodePart~" & gStrSta, gStrOrderBy)
    Else
        MsgBox Translate("Select a participant first.", mcLanguage), vbExclamation
    End If
    Exit Sub

RemoveEntry:
   'is this participant already an entry?
   Set rstOK = mdbMain.OpenRecordset("SELECT * FROM ENTRIES WHERE Sta=""" & gStrSta & """ AND Code=""" & gStrTest & """")
   With rstOK
      If .RecordCount > 0 Then
         .Delete
      End If
      .Close
   End With
Return

RemoveMarks:
   'are there already marks?
   Set rstOK = mdbMain.OpenRecordset("SELECT * FROM MARKS WHERE Sta=""" & gStrSta & """ AND Code=""" & gStrTest & """")
   With rstOK
      If .RecordCount > 0 Then
        Do While Not .EOF
         .Delete
         .Requery
        Loop
      End If
      .Close
   End With
Return

End Sub

Private Sub mnuNodePartEdit_Click()
    Call showParticipant(gStrSta, True)
End Sub

Private Sub cmdCancelAddTest_Click()
    cmdAddTestToParticipant.Visible = False
    cmdCancelAddTest.Visible = False
    Call showAllParticpants("nodePart~" & gStrSta, gStrOrderBy)
End Sub

Private Sub cmdAddTestToParticipant_click()
    Dim i As Integer
    Dim strNodeType As String
    Dim strNodeKey As String
    Dim strNodeText As String
    Dim lngNodeIndex As Long
    
    ' recordSet
    Dim rstEntr As DAO.Recordset
    ' queryString
    Dim qryEntr As String
    Dim qryEntrResult As String
   
    Dim strTestCode As String
    Dim strSta As String
    Dim varTmp As Variant
    
    strNodeType = ""
    strNodeKey = ""
    strNodeText = ""
    lngNodeIndex = -1
        
    For i = 1 To treParticipants.Nodes.Count
        If treParticipants.Nodes(i).Selected Then
            ' keep nodeType
            strNodeType = Left(treParticipants.Nodes(i).Key, 8)
            ' keep nodeKey
            strNodeKey = treParticipants.Nodes(i).Key
            ' keep nodeText
            strNodeText = treParticipants.Nodes(i).Text
            ' keep nodeIndex
            lngNodeIndex = treParticipants.Nodes(i).Index
            ' keep testCode
            varTmp = arrLIST_ArrayAusString(strNodeKey, "~")
            strTestCode = varTmp(1)
            ' keep starting number local
            strSta = gStrSta
        End If
    Next i
    If lngNodeIndex = -1 Then
        ' nothing selected
        MsgBox (Translate("Please select any Test to add!", mcLanguage))
    Else
        ' try to add this test to participant
        Set rstEntr = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code =""" & strTestCode & """ AND Sta =""" & strSta & """")
        If rstEntr.RecordCount = 0 Then
            With rstEntr
                .AddNew
                .Fields("Code") = strTestCode
                .Fields("Sta") = strSta
                .Fields("Position") = 1
                .Fields("Group") = 0
                .Fields("Status") = 0
                .Fields("Timestamp") = Now
                .Update
                .Close
            End With
            cmdAddTestToParticipant.Visible = False
            cmdCancelAddTest.Visible = False
            Call showAllParticpants("nodePart~" & gStrSta, gStrOrderBy)
        Else
            MsgBox strSta & " " & Translate("already entered for this test.", mcLanguage)
        End If
    End If
    
End Sub

Private Sub cmdSavePers_Click()
    Dim rstPers As DAO.Recordset
    Dim i As Integer
    
    Set rstPers = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE PersonId LIKE '" & gStrPers & "'")
    If rstPers.RecordCount = 0 Then
        With rstPers
            .AddNew
        End With
    Else
        With rstPers
            .Edit
        End With
    End If
    With rstPers
        For i = 0 To 19
            Select Case UCase(Me.Controls("lbl_" & CStr(i)).Caption)
            Case ""
            Case "BIRTHDAY"
            Case "SEX", "STATUS"
            Case Else
                .Fields(Me.Controls("lbl_" & CStr(i)).Caption) = Me.Controls("txt_" & CStr(i)).Text
            End Select
        Next i
        .Update
        .Close
    End With
    ' hide saveButton
    Me.Controls("cmdSavePers").Visible = False

    ' highlight the new person
    showAllPersons ("nodePers~" & gStrPers)

End Sub

Private Sub cmdAddPersToParticipant_Click()
    Call getNodeType
    If gStrPers <> "" Then
        Call hideFrmShow
        Call addPersToParticipant
        cmdAddPersToParticipant.Visible = False
        cmdCancelAddPers.Visible = False
        Call showAllParticpants("nodePart~" & gStrSta, gStrOrderBy)
    Else
        MsgBox Translate("Select a person first.", mcLanguage), vbExclamation
    End If
End Sub

Private Sub cmdCancelAddPers_Click()
    Call hideFrmShow
    cmdAddPersToParticipant.Visible = False
    cmdAddPersToNewParticipant.Visible = False
    cmdCancelAddPers.Visible = False
    gStrNodeKey = "rootPart"
    Call showAllParticpants(gStrNodeKey, gStrOrderBy)
End Sub

Private Sub cmdAddHorsToParticipant_Click()
    Call getNodeType
    If gStrHors <> "" Then
        Call hideFrmShow
        Call addHorsToParticipant
        cmdAddHorsToParticipant.Visible = False
        cmdCancelAddHors.Visible = False
        Call showAllParticpants("nodePart~" & gStrSta, gStrOrderBy)
    Else
        MsgBox Translate("Select a horse first.", mcLanguage), vbExclamation
    End If
End Sub

Private Sub cmdCancelAddHors_Click()
    Call hideFrmShow
    cmdAddHorsToParticipant.Visible = False
    cmdAddHorsToNewParticipant.Visible = False
    cmdCancelAddHors.Visible = False
    gStrNodeKey = "rootPart"
    Call showAllParticpants(gStrNodeKey, gStrOrderBy)
End Sub

Private Sub cmdAddPersToNewParticipant_Click()
    Call getNodeType
    If gStrPers <> "" Then
        Call hideFrmShow
        Call addPersToParticipant
        cmdAddPersToNewParticipant.Visible = False
        cmdCancelAddPers.Visible = False
        gStrHors = ""
        Call showAllHorses
        cmdAddHorsToNewParticipant.Visible = True
        cmdCancelAddHors.Visible = True
    Else
        MsgBox Translate("Select a person first.", mcLanguage), vbExclamation
    End If
End Sub

Private Sub cmdAddHorsToNewParticipant_Click()
    Call getNodeType
    If gStrHors <> "" Then
        Call hideFrmShow
        Call addHorsToParticipant
        cmdAddHorsToNewParticipant.Visible = False
        cmdCancelAddHors.Visible = False
        Call showAllParticpants("nodePart~" & gStrSta, gStrOrderBy)
    Else
        MsgBox Translate("Select a horse first.", mcLanguage), vbExclamation
    End If
End Sub

Private Function arrLIST_ArrayAusString(strListTxt As String, strSeperator As String) As Variant
    
     ' Beschreibung:
     ' ... erzeugt aus einem übergebenen String ein array ung gibt diesen zurück
    
    
     ' Parameter:
     ' ... strListTxt   umzuwandelnder Text
     ' ... strSeperator Elementtrennzeichen
    
    
     ' Rückgabe:
     ' ... array vom typ variant
    
    
     ' Hinweis zum Einsatz:
     ' ...
    
    
     ' Fehlerbehandlung:
     ' ...
    
    
     ' Änderungen:
     ' 200301???? tkr Erstellt
    
    Const strProzName = "arrLIST_ArrayAusString"
    
    ' Alle sonstigen Variablendeklarationen
    Dim i As Integer
    Dim intPos As Integer
    Dim intStart As Integer
    Dim intLen As Integer
    Dim varRcA() As Variant
    Dim intLenSeperator As Integer
    
    ' Initialisierung
    intStart = 1
    intPos = 1
    i = 0
    intLenSeperator = Len(strSeperator)
    
    intPos = InStr(intStart, strListTxt, strSeperator)
    
    ' solange ein Seperator gefunden wird
    While intPos <> 0
        
        ReDim Preserve varRcA(i)
        
        intLen = intPos - intStart
        If intLen > 0 Then
            varRcA(i) = Mid(strListTxt, intStart, intLen)
        Else
            varRcA(i) = ""
        End If
        
        intStart = intPos + intLenSeperator
        i = i + 1
        intPos = InStr(intStart, strListTxt, strSeperator)
        If intPos = 0 Then
            ReDim Preserve varRcA(i)
            intLen = Len(strListTxt) - intStart + 1
            If intLen > 0 Then
                varRcA(i) = Mid(strListTxt, intStart, intLen)
            Else
                varRcA(i) = ""
            End If
        End If
    Wend
    
    ' wenn Seperator garnicht gefunden wurde wird der komplette String zurückgegeben
    If i = 0 Then
        ReDim Preserve varRcA(i)
        varRcA(i) = strListTxt
    End If
    
    arrLIST_ArrayAusString = varRcA
    
End Function

Private Function getNodeType() As String
    Dim i As Integer
    Dim strNodeType As String
    Dim strNodeText As String
    Dim lngNodeIndex As Long
    Dim varTmp As Variant
    
    For i = 1 To treParticipants.Nodes.Count
        If treParticipants.Nodes(i).Selected Then
            ' keep nodeType
            strNodeType = Left(treParticipants.Nodes(i).Key, 8)
            ' keep nodeKey
            gStrNodeKey = treParticipants.Nodes(i).Key
            ' keep nodeText
            strNodeText = treParticipants.Nodes(i).Text
            ' keep nodeIndex
            lngNodeIndex = treParticipants.Nodes(i).Index
            
            varTmp = arrLIST_ArrayAusString(gStrNodeKey, "~")
            Select Case strNodeType
            Case "nodePart"
                ' keep starting number global to know for which participant do the following actions
                gStrSta = varTmp(1)
                ' no test code
                gStrTest = ""
            Case "nodePers"
                ' keep personId
                gStrPers = varTmp(1)
            Case "nodeHors"
                ' keep horseId
                gStrHors = varTmp(1)
            Case "nodeEntr"
                ' keep starting number global to know for which participant do the following actions
                gStrSta = varTmp(1)
                ' keep test code to know which test should be removed
                gStrTest = varTmp(2)
            End Select
            
        End If
    Next i
    getNodeType = strNodeType
End Function
Public Function showDetails(r As DAO.Recordset)
    Dim iTemp As Integer
    
    For iTemp = 0 To lblDetail.Count - 1
        If iTemp > 0 Then
            Unload lblDetail(iTemp)
            Unload txtDetail(iTemp)
        End If
    Next iTemp
    For iTemp = 0 To r.Fields.Count - 1
        If iTemp > 0 Then
            Load lblDetail(lblDetail.Count)
            Load txtDetail(txtDetail.Count)
            lblDetail(lblDetail.Count - 1).Top = txtDetail(txtDetail.Count - 2).Top + txtDetail(txtDetail.Count - 2).Height + 50
            txtDetail(lblDetail.Count - 1).Top = txtDetail(txtDetail.Count - 2).Top + txtDetail(txtDetail.Count - 2).Height + 50
        End If
        txtDetail(txtDetail.Count - 1).Width = txtDetail(txtDetail.Count - 1).Container.Width - 50 - txtDetail(lblDetail.Count - 1).Left
        lblDetail(lblDetail.Count - 1).Caption = r.Fields(iTemp).Name
        txtDetail(txtDetail.Count - 1).Text = r.Fields(iTemp) & ""
        lblDetail(lblDetail.Count - 1).Visible = True
        txtDetail(txtDetail.Count - 1).Visible = True
    Next iTemp
    
End Function
