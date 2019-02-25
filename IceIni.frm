VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIni 
   Caption         =   "Ini-bestand"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5640
   HelpContextID   =   200030000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAnnuleren 
      Cancel          =   -1  'True
      Caption         =   "&Annuleren"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Opslaan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   3
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuBestand 
      Caption         =   "&Bestand"
      Begin VB.Menu mnuBestandOpslaan 
         Caption         =   "&Opslaan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBestandAfdrukken 
         Caption         =   "&Afdrukken"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuBestandSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBestandAfsluiten 
         Caption         =   "Af&sluiten"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "frmIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public ReadWrite As Integer

Private Sub cmdAnnuleren_Click()
    cmdOK.Enabled = False
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim cTemp As String, iTemp As Integer, iTemp2 As Integer
    Dim cLabel As String, cHeader As String
    Dim cVeld As String, cCommentaar As String
    
    iTemp = MsgBox("Wijzigingen opslaan in / Save changes in " & "" & " ?", vbYesNo + vbExclamation)
    If iTemp = vbYes Then
        If ReadWrite <> 0 Then
            For iTemp = 1 To MSFlexGrid1.Rows - 1
                cLabel = Trim$(MSFlexGrid1.TextMatrix(iTemp, 0))
                If cLabel <> "" Then
                    If Left$(cLabel, 1) <> "[" Then
                        cVeld = MSFlexGrid1.TextMatrix(iTemp, 1)
                        cCommentaar = MSFlexGrid1.TextMatrix(iTemp, 2)
                    Else
                        cHeader = cLabel
                        cLabel = ""
                        Do
                          iTemp2 = InStr(cHeader, "[")
                          If iTemp2 > 0 Then
                              Mid$(cHeader, iTemp2, 1) = " "
                          End If
                        Loop While iTemp2 > 0
                        Do
                          iTemp2 = InStr(cHeader, "]")
                          If iTemp2 > 0 Then
                              Mid$(cHeader, iTemp2, 1) = " "
                          End If
                        Loop While iTemp2 > 0
                        cHeader = Trim$(cHeader)
                        cVeld = ""
                    End If
                    cTemp = Trim$(cVeld)
                    If Trim$(cCommentaar) <> "" Then
                        cTemp = cTemp & ";" & Trim$(cCommentaar)
                    End If
                    If cLabel <> "" Then
                        Call WriteIniFile(gcIniFile, cHeader, cLabel, cTemp)
                    End If
                End If
            Next iTemp
            Call LogLine(App.EXEName & ".INI gewijzigd/changed.")
        End If
    End If
    cmdOK.Enabled = False
    Unload Me
End Sub

Private Sub Form_Activate()
    
    Dim cTemp As String, cTemp2 As String
    Dim iTemp As Integer, iTemp2 As Integer
    Dim cTempFile As String
    
    Call ReadIniFile(gcIniFile, "General", "General", cTemp)
    frmIni.Caption = UCase$("")
    MSFlexGrid1.FormatString = "Label|Waarde/Value|Commentaar/Comment"
    
    Call Form_Resize
    
    If gcIniFile <> "" Then
        
        If Dir$(gcIniFile) <> "" Then
            Dim iInifilenum As Integer
            iInifilenum = FreeFile
            Open gcIniFile For Input Access Read Shared As #iInifilenum
            Do While Not EOF(iInifilenum)
                Line Input #iInifilenum, cTemp
                cTemp = Trim$(cTemp)
                If Left$(cTemp, 1) = "[" Then
                    MSFlexGrid1.AddItem cTemp
                ElseIf InStr(cTemp, "=") > 0 Then
                    iTemp = InStr(cTemp, "=")
                    If iTemp > 0 Then Mid$(cTemp, iTemp, 1) = vbTab
                    iTemp = InStr(cTemp, ";")
                    If iTemp > 0 Then Mid$(cTemp, iTemp, 1) = vbTab
                    Do
                        iTemp = InStr(cTemp, vbTab + " ")
                        If iTemp > 0 Then
                            cTemp = Left$(cTemp, iTemp) + Trim$(Mid$(cTemp, iTemp + 1))
                        End If
                    Loop While iTemp > 0
                    If InStr(cTemp, vbTab) > 0 Then
                        MSFlexGrid1.AddItem cTemp
                        iTemp = MSFlexGrid1.Rows - 1
                        Call ReadIniFile(gcIniFile, MSFlexGrid1.TextMatrix(0, 2), MSFlexGrid1.TextMatrix(iTemp, 0), cTemp)
                        MSFlexGrid1.TextMatrix(iTemp, 2) = cTemp
                    End If
                End If
            Loop
            Close #iInifilenum
        End If
    End If
End Sub

Private Sub Form_Load()
    
    ReadFormPosition Me
    
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    Dim cTemp As String
    If WindowState = 0 Then
        MSFlexGrid1.Width = ScaleWidth
        MSFlexGrid1.Height = ScaleHeight - cmdOK.Height - 200
        Call ReadIniFile(gcIniFile, "Inifile", "LabelWidth", cTemp)
        If Val(cTemp) > 0 Then
            MSFlexGrid1.ColWidth(0) = Val(cTemp)
        Else
            MSFlexGrid1.ColWidth(0) = 1100
        End If
        Call ReadIniFile(gcIniFile, "Inifile", "ValueWidth", cTemp)
        If Val(cTemp) > 0 Then
            MSFlexGrid1.ColWidth(1) = Val(cTemp)
        Else
            MSFlexGrid1.ColWidth(1) = (ScaleWidth - MSFlexGrid1.ColWidth(0)) \ 2
        End If
        Call ReadIniFile(gcIniFile, "Inifile", "CommentWidth", cTemp)
        If Val(cTemp) > 0 Then
            MSFlexGrid1.ColWidth(2) = Val(cTemp)
        Else
            MSFlexGrid1.ColWidth(2) = (ScaleWidth - MSFlexGrid1.ColWidth(0)) \ 2
        End If
        
        cmdAnnuleren.Left = ScaleWidth - cmdAnnuleren.Width - 100
        cmdOK.Left = cmdAnnuleren.Left - cmdOK.Width - 100
        cmdOK.Top = ScaleHeight - cmdOK.Height - 100
        cmdAnnuleren.Top = cmdOK.Top
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Enabled = True Then
       Call cmdOK_Click
    End If
    WriteFormPosition Me
    Call WriteIniFile(gcIniFile, "Inifile", "LabelWidth", Format$(MSFlexGrid1.ColWidth(0)))
    Call WriteIniFile(gcIniFile, "Inifile", "ValueWidth", Format$(MSFlexGrid1.ColWidth(1)))
    Call WriteIniFile(gcIniFile, "Inifile", "CommentWidth", Format$(MSFlexGrid1.ColWidth(2)))
End Sub

Private Sub mnuBestandAfdrukken_Click()
    Me.MousePointer = 11
    Dim i, j, iLmarge
    iLmarge = 5
    Printer.Print ""
    Printer.FontSize = 11
    Printer.Print Tab(iLmarge);
    Printer.FontSize = 14
    Printer.Print App.EXEName; ".INI"
    Printer.FontSize = 11
    Printer.Print ""
        
    For i = 0 To MSFlexGrid1.Rows - 1
        If Left$(MSFlexGrid1.TextArray(i * MSFlexGrid1.Cols), 1) = "[" Then
            Printer.Print Tab(iLmarge); RTrim$(MSFlexGrid1.TextArray(i * MSFlexGrid1.Cols));
        Else
            Printer.Print Tab(iLmarge + 5); RTrim$(MSFlexGrid1.TextArray(i * MSFlexGrid1.Cols));
        End If
        Printer.Print Tab(25 + iLmarge); RTrim$(MSFlexGrid1.TextArray(i * MSFlexGrid1.Cols + 1));
        If MSFlexGrid1.TextArray(i * MSFlexGrid1.Cols + 2) <> "" Then
            Printer.Print "; " & RTrim$(MSFlexGrid1.TextArray(i * MSFlexGrid1.Cols + 2));
        End If
        Printer.Print ""
    Next i
    Printer.Print ""
    Printer.FontSize = 11
    Printer.Print Tab(iLmarge);
    Printer.FontSize = 9
    Printer.Print Format$(Now, "DD-MM-YYYY HH:MM:SS")
    Printer.FontSize = 11
    Printer.Print Tab(iLmarge);
    Printer.FontSize = 9
    Printer.Print App.EXEName & " " & Format$(App.Major) & "." & Format$(App.Minor) & " build: " & Format$(App.Revision) & " - " & App.LegalCopyright
    Printer.FontSize = 11
    Printer.EndDoc
    Me.MousePointer = 0
End Sub

Private Sub mnuBestandAfsluiten_Click()
    Unload Me
End Sub

Private Sub MSFlexGrid1_DblClick()
    Dim cTemp As String, iTemp As Integer
    Dim cLabel As String, cKolom As String
    cLabel = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    If Left$(cLabel, 1) <> "[" And cLabel <> "" Then
        cTemp = MSFlexGrid1.Text
        cKolom = MSFlexGrid1.TextMatrix(0, MSFlexGrid1.Col)
        cTemp = InputBox(cLabel & ":", cKolom, cTemp)
        If cTemp <> "" Then
            cmdOK.Enabled = True
            MSFlexGrid1.Text = Trim$(cTemp)
        End If
    End If
    mnuBestandOpslaan.Enabled = cmdOK.Enabled
    mnuBestandAfdrukken.Enabled = Not cmdOK.Enabled
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim iTemp As Integer
    Dim cTemp As String, cLabel As String
    cLabel = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    If Left$(cLabel, 1) <> "[" And cLabel <> "" Then
        Select Case KeyCode
        Case vbKeyInsert
            If ReadWrite = 2 Then
                iTemp = MsgBox("Label toevoegen ?", vbYesNo)
                If iTemp = vbYes Then
                    Do
                        cTemp = InputBox("Geef de naam van het label.")
                        If cTemp <> "" Then
                            For iTemp = 0 To MSFlexGrid1.Rows - 1
                                If UCase$(MSFlexGrid1.TextMatrix(iTemp, 0)) = UCase$(cTemp) Then
                                    iTemp = MsgBox(cTemp & " wordt al gebruikt!")
                                    cTemp = ""
                                    Exit For
                                End If
                            Next iTemp
                            If cTemp <> "" Then
                                MSFlexGrid1.AddItem cTemp, MSFlexGrid1.Row
                                cmdOK.Enabled = True
                            End If
                        Else
                            Exit Do
                        End If
                    Loop While cTemp = ""
                End If
            End If
        Case vbKeyDelete
            If ReadWrite = 2 Then
                iTemp = MsgBox("Label verwijderen ?", vbYesNo)
                If iTemp = vbYes Then
                    MSFlexGrid1.RemoveItem MSFlexGrid1.Row
                    cmdOK.Enabled = True
                End If
            End If
        Case vbKeyReturn
            Call MSFlexGrid1_DblClick
            cmdOK.Enabled = True
        End Select
    End If
    mnuBestandOpslaan.Enabled = cmdOK.Enabled
    mnuBestandAfdrukken.Enabled = Not cmdOK.Enabled
End Sub

