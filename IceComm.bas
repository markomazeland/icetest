Attribute VB_Name = "modIceComm"
' Copyright (C) Marko Mazeland and/or Datawerken Holding BV 2003
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

Public mcDatabaseDir As String
Public mcHtmlDir As String
Public mcExcelDir As String
Public mcRtfDir As String
Public mcDatabaseName As String
Public mcTempHtmlDir As String
Public gcIniFile As String
Public gcIniHorseFile As String
Public msFontSize As Single
Public mcLanguage As String
Public mcCountry As String
Public mcText() As String
Public miSelectBySingleClick As Integer
Public miBackupInterval As Integer
Public miShowJudgesRanking As Integer
Public miSponsorType As Integer
Public miExcelFiles As Integer
Public miHtmlFiles As Integer
Public miNoBackupNow As Integer
Public miBackupTicker As Integer
Public miShowRidersClub As Integer
Public miShowRidersTeam As Integer
Public miShowHorseId As Integer
Public miShowHorseAge As Integer
Public miFinalsSequence As Integer
Public miUseColors As Integer
Public miUseHighLights As Integer
Public mcAmpList As String
Public mcMnuAmpList As String
Public mcNoColor As String
Public mcExcelSeparator As String
Public miMarkFinalsInResultLists As Integer
Public miBFinalLevel As Integer
Public miCFinalLevel As Integer
Public giExternalOperation  As Integer
Public miConnectedToInternet As Integer
Public miWriteLogDB As Integer
Public miUseIceSort As Integer
Public miShowRidersLK As Integer

Public mFormsCollection As New Collection
Public mdbMain As DAO.Database
Public fSplash As frmSplash
Public mcVersionSwitch As String

Public Const mlAlertColor = vbYellow
Public Const SW_SHOWNORMAL = 1 ' Restores Window if Minimized
Public Const mcTempId = "XX9999999999"
Public Const mcNoPosition = "---"

'Constant SortingApplication
'The sorting application is expected to be in the application directory.
'If the application is missing, IceTest's internal group composing functionality is used.
Public Const SortingApplication = "icesort.exe"



Public Function FormIsThere(CollectionTag As String) As Integer
    Dim lTemp As Long
    
    FormIsThere = False
    If mFormsCollection.Count > 0 Then
        For lTemp = 1 To mFormsCollection.Count
            If mFormsCollection(lTemp).CollectionTag = CollectionTag Then
                If mFormsCollection(lTemp).WindowState = 1 Then
                    ReadFormPosition mFormsCollection(lTemp), CollectionTag
                End If
                SetFocusTo mFormsCollection(lTemp)
                FormIsThere = True
                Exit For
            End If
        Next lTemp
    End If
End Function
Public Function ParseCommand(Label As String) As String
    Dim intTemp As Integer
    Dim strTemp As String
    
    intTemp = InStr(Command$, Label)
    If intTemp > 0 Then
        strTemp = Mid$(Command$, intTemp + Len(Label))
        intTemp = InStr(strTemp, " ")
        If intTemp = 0 Then
            intTemp = InStr(strTemp, "/")
        End If
        If intTemp > 0 Then
            strTemp = Trim$(Left$(strTemp, intTemp))
        End If
        ParseCommand$ = strTemp
    Else
        ParseCommand$ = ""
    End If
End Function

Function TestIniFile(Optional strIniFile As String = "") As String
   'what is the (default) name of the ini file
   '(ini files are often more flexible than registry values and can be erased easily)
    Dim strTemp As String
    Dim intFileNum As Integer
    Dim cIniDir As String
    Dim cOldIniFile As String
    Dim lFileSize As Long
    
    On Local Error GoTo TestIniFileError
    
    cIniDir = GetSpecialFolderLocation(CSIDL_APPDATA) & "\IceHorse\"
    If Dir$(cIniDir, vbDirectory) = "" Then
        MkDir cIniDir$
    End If
    
    If strIniFile = "" Then
       strIniFile = App.EXEName & ".Ini"
    ElseIf InStr(strIniFile, "\") = 0 And InStr(strIniFile, ".") = 0 Then
       strIniFile = strIniFile & ".Ini"
    End If
    
    If InStr(strIniFile, "\") = 0 Then
        intFileNum = FreeFile
        Open cIniDir & strIniFile For Binary Access Write Shared As #intFileNum
        lFileSize = LOF(intFileNum)
        Close #intFileNum
        TestIniFile = cIniDir & strIniFile
        If lFileSize = 0 Then
            If Dir$(WinDir & strIniFile) <> "" Then
                FileCopy WinDir & strIniFile, TestIniFile
            End If
        End If
    Else
        intFileNum = FreeFile
        Open strIniFile For Binary Access Write Shared As #intFileNum
        Close #intFileNum
        TestIniFile = strIniFile
    End If

TestIniFileError:
    If Err > 0 Then
        If InStr(strIniFile, "\") = 0 Then
            If Dir$(App.Path & "\" & strIniFile) = "" Then
                LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
                MsgBox App.EXEName & ": user " & UserName & " has no proper access to ' " & cIniDir & "'." & vbCrLf & "De instellingen van " & App.EXEName & " (" & strIniFile & ") worden nu opgeslagen in '" & App.Path & "'.", vbCritical
            End If
            strIniFile = TestIniFile(App.Path & "\" & strIniFile)
            TestIniFile = strIniFile
        Else
            LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
            MsgBox App.EXEName & ": user " & UserName & " has no proper access to '" & App.Path & "'. " & strIniFile & " kan niet aangemaakt worden." & vbCrLf & "Waarschuw uw systeembeheerder." & vbCrLf & vbCrLf & "(" & Err.Description & ")", vbCritical
            Unload frmMain
            End
        End If
    End If
End Function

Public Sub ReadFormPosition(F As Form, Optional Label As String = "")
    Dim strTemp(5) As String, intTemp As Integer, jTemp As Integer
    
    On Local Error Resume Next
    
    If Label = "" Then
        Label = F.Name
    End If
    
    intTemp = InStr(Label, "[")
    If intTemp > 0 Then
        Label = RTrim$(Left$(Label, intTemp - 1))
    End If
    ReadIniFile gcIniFile, Label, "Move", strTemp(0), True
    
    If strTemp(0) <> "" Then
        jTemp = 0
        Do
            jTemp = jTemp + 1
            intTemp = InStr(strTemp(0), ",")
            If intTemp > 0 Then
                strTemp(jTemp) = Left$(strTemp(0), intTemp - 1)
                strTemp(0) = RTrim$(Mid$(strTemp(0) & " ", intTemp + 1))
            Else
                strTemp(jTemp) = strTemp(0)
                strTemp(0) = ""
            End If
        Loop While strTemp(0) <> ""
        F.Move Val(Max(0, strTemp(1))), Max(0, Val(strTemp(2))), Min(Screen.Width, Val(strTemp(3))), Min(Screen.Height, Val(strTemp(4)))
    ElseIf F.WindowState <> 1 Then
        F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2
    End If
    ReadIniFile gcIniFile, Label, "Windowstate", strTemp(0), True
    If strTemp(0) <> "" Then
        F.WindowState = Val(strTemp(0))
    End If
    DoEvents
    On Local Error GoTo 0
End Sub

Public Function AddAmp(strText As String, AmpList As String, Optional Style As Integer = 0) As String
    Dim intTemp As Integer
    Dim iTemp2 As Integer
    Dim strTemp As String
    Dim cWord() As String
    
    strTemp = strText
    If InStr(AmpList, Left$(strTemp, 1)) = 0 Then
        AmpList = AmpList & Left$(strTemp, 1)
        strTemp = "&" & strTemp
    ElseIf Style = 1 Then
        'Try to use seperate words first
        cWord = Split(strTemp, " ")
        For intTemp = LBound(cWord) To UBound(cWord)
            If InStr(AmpList, Left$(cWord(intTemp), 1)) = 0 Then
                strTemp = ""
                For iTemp2 = LBound(cWord) To UBound(cWord)
                    If intTemp = iTemp2 Then
                        AmpList = AmpList & Left$(cWord(iTemp2), 1)
                        strTemp = strTemp & "&" & cWord(iTemp2) & " "
                    Else
                        strTemp = strTemp & cWord(iTemp2) & " "
                    End If
                Next iTemp2
                Exit For
            End If
        Next intTemp
    End If
    If InStr(strTemp, "&") = 0 Then
        'If not done yet, find first unique letter
        For intTemp = 1 To Len(strTemp)
            If InStr(AmpList, Mid$(strTemp, intTemp, 1)) = 0 And Mid$(strTemp, intTemp, 1) >= "a" And Mid$(strTemp, intTemp, 1) <= "z" Then
                AmpList = AmpList & Mid$(strTemp, intTemp, 1)
                strTemp = Left$(strTemp, intTemp - 1) & "&" & Mid$(strTemp, intTemp)
            End If
        Next intTemp
    End If
    If InStr(strTemp, "&") = 0 Then
        'If not done yet, find first unique character
        For intTemp = 1 To Len(strTemp)
            If InStr(AmpList, Mid$(strTemp, intTemp, 1)) = 0 And Mid$(strTemp, intTemp, 1) Like "[A-Z]" Then
                AmpList = AmpList & Mid$(strTemp, intTemp, 1)
                strTemp = Left$(strTemp, intTemp - 1) & "&" & Mid$(strTemp, intTemp)
            End If
        Next intTemp
    End If
    AddAmp = strTemp
End Function

Public Sub LogLine(strText As String, Optional strPath As String)
   'logs a line to a log file
    Dim strTemp As String
    Dim lngMaxSize As Integer
    Dim intTemp As Integer
    Dim intLogFileNum As Integer
        
    On Local Error Resume Next
    
    If strPath = "" Then strPath = App.Path
    If InStr(strPath, ".Log") = 0 Then
        If Right$(strPath, 1) <> "\" And Right$(strPath, 1) <> ":" Then
            strPath = strPath & "\"
        End If
        strPath = strPath & App.EXEName & ".log"
    End If
    
    intLogFileNum = FreeFile
    Open strPath For Append Access Read Write Shared As #intLogFileNum
    Print #intLogFileNum, Format$(Now, "dd-mm-yyyy hh:mm:ss") & "  [" & MachineName & "/" & UserName & "] " & strText
    Close #intLogFileNum
    
    On Local Error GoTo 0

End Sub

Public Sub WriteFormPosition(F As Form, Optional Label As String = "")
    Dim strTemp As String, intTemp As Integer
    
    If Label = "" Then
        Label = F.Name
    End If
    intTemp = InStr(Label, "[")
    If intTemp > 0 Then
        Label = RTrim$(Left$(Label, intTemp - 1))
    End If
    WriteIniFile gcIniFile, Label, "Windowstate", Format$(F.WindowState), True
    If F.WindowState = 0 Then
        strTemp = Format$(F.Left) + "," + Format$(F.Top) + "," + Format$(F.Width) + "," + Format$(F.Height)
        WriteIniFile gcIniFile, Label, "Move", strTemp, True
    End If
End Sub

Public Function Max(Value1 As Variant, Value2 As Variant) As Variant
    Max = IIf(Value1 > Value2, Value1, Value2)
End Function
Public Function Min(Value1 As Variant, Value2 As Variant) As Variant
    Min = IIf(Value1 < Value2, Value1, Value2)
End Function

Public Function ClipAmp(strText As String) As String
   Dim strTemp As String
   
   strTemp = Trim$(Replace(strText, "&", ""))
   
   'LL: Remove slashes also to prevent invalid filenames:
   strTemp = Trim$(Replace(strTemp, "/", ""))
   
   Do While Right$(strTemp, 1) = "."
      strTemp = RTrim$(Left$(strTemp, Len(strTemp) - 1))
   Loop
   ClipAmp = strTemp
End Function

Public Function UnloadItems(o As Object)
   On Local Error GoTo UnloadItemsError
   Do While o.Count > 1
       o.Visible = False
        
       Unload o(o.Count - 1)
   Loop
UnloadItemsError:

End Function
Sub ChangeFontSize(F As Form, FontSize As Single, Optional FontName As String = "")
    
    Dim ctr As Control
    Dim cmd As CommandButton
    Dim strTemp As String
    Dim sOldFontSize As Single
    
    On Local Error Resume Next
    
    sOldFontSize = F.Font.Size
    
    F.Font.Size = FontSize
    
    If FontName <> "" Then
        F.Font.Name = FontName
    End If
    
    If FontSize > 0 Then
        For Each ctr In F
            If Left$(ctr.Name, 3) = "rtf" Then
                With ctr
                    .SelStart = 0
                    .SelLength = Len(ctr)
                    .SelFontSize = FontSize
                    If FontName <> "" Then
                        .SelFontName = FontName
                    Else
                        .SelFontName = F.Font.Name
                    End If
                End With
            Else
                If ctr.Font.Name <> "FixedSys" And ctr.Font.Name <> "Courier" Then
                    If FontName <> "" Then
                        ctr.Font.Name = FontName
                    Else
                        ctr.Font.Name = F.Font.Name
                    End If
                End If
                ctr.Font.Size = FontSize
                
                If Left$(LCase$(ctr.Name), 3) = "cmd" Then
                    If ctr.Caption <> "" Then
                        ctr.Width = 1450 * FontSize / 8.5
                    End If
                ElseIf Left$(LCase$(ctr.Name), 3) = "txt" Then
                    ctr.Height = FontSize * ctr.Height / sOldFontSize
                End If
            End If
        Next
    End If
    F.Refresh
    
    On Local Error GoTo 0
    DoEvents
End Sub
Public Function FitString(o As Object, TextToFit As String, WidthToFit As Integer, Optional FitType As Integer = 0, Optional FitUpto As String = ">>") As String
    Dim intTemp As Integer
    Dim iTemp2 As Integer
    Dim iTemp3 As Integer
    Dim strTemp As String
    Dim cFill As String
    Dim iAddcFill As Integer
    
    On Local Error GoTo FitStringFout
    
    strTemp = TextToFit
    
    If FitType = 0 Then 'in het midden weghalen
        cFill = "....."
        Do While o.TextWidth(strTemp) > WidthToFit
            If InStr(strTemp, cFill) = 0 Then
                Mid$(strTemp, Len(strTemp) \ 2 - 3) = cFill
            Else
                intTemp = InStr(strTemp, cFill)
                If intTemp > 3 And intTemp < Len(strTemp) - 8 Then
                    If Mid$(strTemp, intTemp - 1, 1) = "\" And Mid$(strTemp, intTemp + 5, 1) = "\" Then
                        iTemp2 = InStr(intTemp + 6, strTemp, "\")
                        If iTemp2 > 0 And iTemp2 < Len(strTemp) Then
                            strTemp = Left$(strTemp, intTemp - 1) & cFill & Mid$(strTemp, iTemp2)
                        Else
                            iTemp2 = InStrRev(strTemp, "\", intTemp - 2)
                            If iTemp2 > 1 Then
                                strTemp = Left$(strTemp, iTemp2) & cFill & Mid$(strTemp, intTemp + 5)
                            Else
                                strTemp = Left$(strTemp, intTemp - 1) & cFill & Mid$(strTemp, intTemp + 7)
                            End If
                        End If
                    ElseIf Mid$(strTemp, intTemp - 1, 1) = "\" Then
                        strTemp = Left$(strTemp, intTemp - 1) & cFill & Mid$(strTemp, intTemp + 7)
                    ElseIf Mid$(strTemp, intTemp + 5, 1) = "\" Then
                        strTemp = Left$(strTemp, intTemp - 3) & cFill & Mid$(strTemp, intTemp + 5)
                    Else
                        strTemp = Left$(strTemp, intTemp - 2) & cFill & Mid$(strTemp, intTemp + 6)
                    End If
                Else
                    Exit Do
                End If
            End If
        Loop
    ElseIf FitType = 1 Then 'aan het eind weghalen
        cFill = "..."
        iAddcFill = False
        If o.TextWidth(strTemp) > WidthToFit Then
            Do While o.TextWidth(strTemp & cFill) > WidthToFit
                strTemp = RTrim$(Left$(strTemp, Len(strTemp) - 1))
                iAddcFill = True
            Loop
        End If
        If iAddcFill = True Then
            strTemp = strTemp & cFill
        End If
    ElseIf FitType = 4 Then 'vlak voor het eind weghalen
        cFill = "..."
        iAddcFill = False
        iTemp2 = 0
        If o.TextWidth(strTemp) > WidthToFit Then
            Do While o.TextWidth(strTemp & cFill) > WidthToFit
                iTemp2 = iTemp2 + 1
                strTemp = RTrim$(Left$(strTemp, InStr(strTemp, FitUpto) - iTemp2)) & Mid$(strTemp, InStr(strTemp, FitUpto))
                iAddcFill = True
            Loop
        End If
        If iAddcFill = True Then
            strTemp = Replace(strTemp, FitUpto, cFill & FitUpto)
        End If
    Else 'alleen inkorten
        cFill = ""
        iAddcFill = False
        Do While o.TextWidth(strTemp & cFill) > WidthToFit
            strTemp = RTrim$(Left$(strTemp, Len(strTemp) - 1))
            iAddcFill = True
        Loop
        If iAddcFill = True Then
            strTemp = strTemp & cFill
        End If
    End If
    
FitStringFout:
    FitString = Trim$(strTemp)
    On Local Error GoTo 0
End Function

Public Function Encrypt(TextToEncrypt As String, PrivateKey As String) As String
      Dim iKeyLen As Integer
      Dim iChar As Integer
      Dim intTemp As Integer
      Dim cTmpEncrypt As String
      
      cTmpEncrypt = TextToEncrypt
      iKeyLen = Len(PrivateKey)
      For intTemp = 1 To Len(cTmpEncrypt)
         iChar = Asc(Mid$(PrivateKey, (intTemp Mod iKeyLen) - iKeyLen * ((intTemp Mod iKeyLen) = 0), 1))
         Mid$(cTmpEncrypt, intTemp, 1) = Chr$(Asc(Mid$(cTmpEncrypt, intTemp, 1)) Xor iChar)
      Next
      Encrypt = cTmpEncrypt
End Function

Public Sub NotYet()
    MsgBox Translate("Sorry, this function is not available at the moment!", mcLanguage), vbExclamation
End Sub

Public Function ReadTagItem(o As Object, ItemToRead As String) As String
    Dim cVeld() As String
    Dim intTemp As Integer
    Dim strTemp As String
    
    DoEvents
    ReadTagItem = ""
    cVeld = Split(o.Tag, "|")
    For intTemp = LBound(cVeld) To UBound(cVeld)
        Parse strTemp, cVeld(intTemp), "="
        If strTemp = ItemToRead Then
            ReadTagItem = cVeld(intTemp)
            Exit For
        End If
    Next intTemp
End Function
Public Sub ChangeTagItem(o As Object, ItemToRead As String, strText As String)
    Dim cVeld() As String
    Dim intTemp As Integer
    Dim strTemp As String
    Dim iGevonden As Integer
    
    cVeld = Split(o.Tag, "|")
    o.Tag = ""
    iGevonden = False
    For intTemp = LBound(cVeld) To UBound(cVeld)
        Parse strTemp, cVeld(intTemp), "="
        If strTemp <> "" Then
            If strTemp = ItemToRead Then
                iGevonden = True
                If strText <> "" Then
                    o.Tag = o.Tag & strTemp & "=" & strText & "|"
                End If
            Else
                o.Tag = o.Tag & strTemp & "=" & cVeld(intTemp) & "|"
            End If
        End If
    Next intTemp
    If iGevonden = False Then
        o.Tag = o.Tag & ItemToRead & "=" & strText & "|"
    End If
End Sub

Public Sub Parse(Word As String, strText As String, Separator As String, Optional Method As Integer = 0)
    Dim intTemp As Integer
    Dim iTemp2 As Integer
    Dim iTemp3 As Integer
    
    If Method = 1 And Len(Separator) > 1 Then
        intTemp = 0
        For iTemp2 = 1 To Len(Separator)
            iTemp3 = InStr(strText, Mid$(Separator, iTemp2, 1))
            If iTemp3 > 0 Then
                If intTemp = 0 Then
                    intTemp = iTemp3
                ElseIf iTemp3 < intTemp Then
                    intTemp = iTemp3
                End If
            End If
        Next iTemp2
    ElseIf Method = 2 Then
        intTemp = InStrRev(strText, Separator)
    Else
        intTemp = InStr(strText, Separator)
    End If
    If intTemp > 0 Then
        If Method = 2 Then
            Word = Trim$(Mid$(strText & " ", intTemp + 1))
        Else
            Word = Trim$(Left$(strText, intTemp - 1))
        End If
        If intTemp = Len(strText) - Len(Separator) + 1 Then
            strText = ""
        ElseIf Method = 1 Then
            strText = Mid$(strText, intTemp + 1)
            Do While InStr(Separator, Left$(strText, 1)) > 0 And strText <> ""
                strText = Trim$(Mid$(strText & " ", 2))
            Loop
        ElseIf Method = 2 Then
            strText = Left$(strText, intTemp - Len(Separator))
        Else
            strText = Mid$(strText, intTemp + Len(Separator))
        End If
    Else
        Word = strText
        strText = ""
    End If
        
End Sub
Public Function CopyFile(FileFrom As String, FileTo As String, KillFromFile As Integer) As Integer
    Dim dstart As Double
    Dim dStart2 As Double
        
    CopyFile = False
    
    On Local Error GoTo CopyFileFout
    
    dStart2 = Timer
    
    SetMouseHourGlass
    
    KillFile FileTo
    
    FileCopy FileFrom, FileTo
    
    If Err = 0 Then
        CopyFile = True
    End If
    
    DoEvents
    
    If KillFromFile Then KillFile FileFrom
    DoEvents
    
    On Local Error GoTo 0
    
    SetMouseNormal
Exit Function

CopyFileFout:
    LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
    Sleep 1
    If (Timer - 10) Mod 86400 < dStart2 Then
       Resume
    Else
       MsgBox "File " & FileTo & " cannot be updated." & vbCrLf & "Try again later.", vbExclamation
       Resume Next
    End If
Return

End Function

Public Sub SetMouseHourGlass()
    frmMain.MousePointer = vbHourglass
    DoEvents
End Sub
Public Sub SetMouseNormal()
    frmMain.MousePointer = vbNormal
    DoEvents
End Sub

Public Function KillFile(FileName As String) As Integer

    KillFile = False
    On Local Error GoTo KillFileEnd
    
    If Dir$(FileName) <> "" Then
        AttribNormal FileName
        Kill FileName
        If Err <> 0 Then
            KillFile = Err.Number
        Else
            KillFile = True
        End If
    End If
    
KillFileEnd:
    On Local Error GoTo 0
    
End Function

Public Sub Sleep(Seconds As Integer)
    Dim dstart As Double
    dstart = Timer
    Do While (Timer - Seconds - 0.5) Mod 86400 < dstart
        DoEvents
    Loop
End Sub

Public Function AttribNormal(FileName As String, Optional strPath As String) As Integer
    Dim cFileList As String
    Dim strTemp As String
    Dim iAttr As Integer
    
    If strPath = "" And InStr(FileName, "\") > 0 Then
        strPath = Left$(FileName, InStrRev(FileName, "\"))
    ElseIf strPath = "" And InStr(FileName, ":") > 0 Then
        strPath = Left$(FileName, InStrRev(FileName, ":"))
    End If
    If strPath <> "" Then
        If Right$(strPath, 1) <> ":" And Right$(strPath, 1) <> "\" Then
            strPath = strPath & "\"
        End If
    End If
    
    AttribNormal = True
    
    strTemp = Dir$(FileName)
    Do While strTemp <> ""
        cFileList = cFileList & strPath & strTemp & "|"
        strTemp = Dir$
    Loop
    
    On Local Error Resume Next
    
    Do While cFileList <> ""
        Parse strTemp, cFileList, "|"
        If strTemp <> "" Then
            iAttr = GetAttr(strTemp)
            If (iAttr And vbReadOnly) <> 0 Then
                SetAttr strTemp, vbNormal
                If Err > 0 Then
                    LogLine strTemp & ": file 'Read Only' (" & Err.Description & ")"
                    AttribNormal = False
                Else
                    LogLine strTemp & ": 'Read Only' removed."
                End If
            End If
        End If
    Loop
    
    On Local Error GoTo 0
End Function

Public Sub StatusMessage(Optional strText As String = "", Optional intStyle As Integer = 0)
     
   If strText = "" Then
      If frmMain.TestInfoMessage <> "" Then
        strText = frmMain.TestInfoMessage
      End If
      If frmMain.StatusBar1.Style = sbrSimple Then
          frmMain.StatusBar1.SimpleText = strText
      Else
          frmMain.StatusBar1.Panels("StatusMessage").Text = strText
      End If
      frmMain.Enabled = True

   Else
      If (intStyle And 2) = 2 Then
        frmMain.StatusBar1.Font.Bold = True
      Else
        frmMain.StatusBar1.Font.Bold = False
      End If
      
      If (intStyle And 4) = 4 Then
        frmMain.StatusBar1.Font.Italic = True
      Else
        frmMain.StatusBar1.Font.Italic = False
      End If
      If frmMain.StatusBar1.Style = sbrSimple Then
          frmMain.StatusBar1.SimpleText = strText
      Else
          frmMain.StatusBar1.Panels("StatusMessage").Text = strText
      End If
   End If
   DoEvents
End Sub
Public Function MakeStringValue(strValue As String, Optional strFormat As String = "") As Currency
   If strFormat <> "" Then
      MakeStringValue = Val(Format$(Val(Replace(strValue, ",", ".")), strFormat))
   Else
      MakeStringValue = Val(Replace(strValue, ",", "."))
   End If
End Function
Public Sub SetFocusTo(o As Object)
    On Local Error Resume Next
    If o.Visible = True And o.Enabled = True Then
        o.SetFocus
        DoEvents
    End If
    On Local Error GoTo 0
End Sub
Public Sub SetVariable(strName As String, strValue As String)
    Dim rstQry As Recordset
    On Local Error Resume Next
    Set rstQry = mdbMain.OpenRecordset("SELECT * FROM Variables WHERE Item LIKE " & Chr$(34) & strName & Chr$(34))
    If Err = 0 Then
        If rstQry.RecordCount > 0 Then
            rstQry.Edit
        Else
            rstQry.AddNew
            rstQry.Fields("Item") = Left$(strName, rstQry.Fields("Item").Size)
        End If
        rstQry.Fields("Value") = Left$(strValue, rstQry.Fields("Value").Size)
        rstQry.Update
    End If
    rstQry.Close
    
    On Local Error GoTo 0
    Set rstQry = Nothing
End Sub
Public Function GetVariable(strName As String) As String
    Dim rstQry As Recordset
    On Local Error Resume Next
    Set rstQry = mdbMain.OpenRecordset("SELECT Value FROM Variables WHERE Item LIKE " & Chr$(34) & strName & Chr$(34))
    If rstQry.RecordCount > 0 Then
        If Err = 0 Then
            GetVariable = RTrim$(rstQry.Fields("Value") & "")
        Else
            GetVariable = ""
        End If
    Else
        GetVariable = ""
    End If
    rstQry.Close
    On Local Error GoTo 0
    Set rstQry = Nothing
End Function
Public Function ShowDocument(DocumentName As String, F As Form, Optional Mode As String = "open") As Long
    
    Dim lRetVal As Long
    Dim cMsg As String
    
    SetMouseHourGlass
    
    DocumentName = Trim$(DocumentName)
    If DocumentName = "" Then
        Mode = "find"
    End If
    
    ShowDocument = 0
    
    lRetVal = ShellExecute(F.hwnd, Mode, DocumentName, vbNullString, vbNullString, SW_SHOWNORMAL)
    If lRetVal <= 32 Then ' Error
        Select Case lRetVal
        Case SE_ERR_FNF
            cMsg = "file not found"
            cMsg = "file not found"
            ShowDocument = 2
        Case SE_ERR_PNF
            cMsg = "path not found"
        Case SE_ERR_ACCESSDENIED
            cMsg = "access denied"
        Case SE_ERR_OOM
            cMsg = "insufficient memory"
            ShowDocument = 2
        Case SE_ERR_DLLNOTFOUND
            cMsg = "DLL not found"
            ShowDocument = 2
        Case SE_ERR_SHARE
            cMsg = "sharing error"
        Case SE_ERR_ASSOCINCOMPLETE
            cMsg = "invalid file link"
            ShowDocument = 2
        Case SE_ERR_DDETIMEOUT
            cMsg = "DDE time out"
            ShowDocument = 2
        Case SE_ERR_DDEFAIL
            cMsg = "DDE transaction error"
            ShowDocument = 2
        Case SE_ERR_DDEBUSY
            cMsg = "DDE busy"
            ShowDocument = 2
        Case SE_ERR_NOASSOC
            cMsg = "no association for this file type"
            ShowDocument = 3
        Case ERROR_BAD_FORMAT
            cMsg = "invalid EXE or error in EXE"
            ShowDocument = 2
        Case Else
            cMsg = "unknown error " & lRetVal
            ShowDocument = 2
        End Select
        If ShowDocument = 2 Then
            MsgBox "'" & DocumentName & "': " & Translate(cMsg, mcLanguage), vbExclamation
        End If
    End If
    
    SetMouseNormal
End Function


Public Function NameOfFile(strText) As String
    Dim i As Integer
    NameOfFile = strText
    i = InStrRev(NameOfFile, ".")
    If i > 0 Then
        NameOfFile = Left$(NameOfFile, i - 1)
    End If
End Function

Public Sub TranslateControls(F As Form)
    
    Dim ctl As Control
    Dim cOldAmpList As String
    Dim cMnuList As String
    Dim cTemp As String
    Dim cMenu As String
    
    On Local Error Resume Next
    
    'Translate form's caption first:
    F.Caption = Translate(F.Caption, mcLanguage)
    
    cMnuList = "File|Edit|Test|Comb|Tool|Help"
    mcAmpList = mcMnuAmpList
    Do While cMnuList <> ""
        Parse cMenu, cMnuList, "|"
        For Each ctl In F.Controls
            If ctl.Name = "mnu" & cMenu Then
                ctl.Caption = TranslateCaption(ctl.Caption, 0, True)
                Exit For
            End If
        Next ctl
    Loop
    
    cMnuList = "File|Edit|Test|Comb|Tool|Help"
    mcMnuAmpList = mcAmpList
    Do While cMnuList <> ""
        Parse cMenu, cMnuList, "|"
        mcAmpList = mcMnuAmpList
        For Each ctl In F.Controls
            If Left$(ctl.Name, 3 + Len(cMenu)) = "mnu" & cMenu And ctl.Name <> "mnu" & cMenu Then
                ctl.Caption = TranslateCaption(ctl.Caption, 0, True)
            End If
        Next ctl
    Loop
    
    mcAmpList = mcMnuAmpList & mcAmpList
    For Each ctl In F.Controls
        If Left$(ctl.Name, 3) <> "mnu" Then
            cTemp = ctl.Caption
            If cTemp <> "" Then
                ctl.Caption = TranslateCaption(cTemp, Max(ctl.Width, 750), True)
            End If
            If Len(ctl.ToolTipText) > 1 And mcLanguage <> "English" Then
               ctl.ToolTipText = Translate(ctl.ToolTipText, mcLanguage)
            End If
        End If
    Next ctl
    
    Set ctl = Nothing
    On Local Error GoTo 0
Exit Sub

End Sub
Public Function TranslateCaption(cCaption As String, iWidth As Integer, iProperCase As Integer) As String
    Dim cTemp As String
    Dim iTemp As Integer
    Dim iAmp As Integer
    Dim iDots As Integer
    
    On Local Error Resume Next
    
    cTemp = cCaption
    If InStr(cTemp, "&") > 0 Then
        iAmp = True
    Else
        iAmp = False
    End If
    If Right$(cTemp, 3) = "..." Then
        iDots = True
        cTemp = Left$(cTemp, Len(cTemp) - 3)
    Else
        iDots = False
    End If
    
    If mcLanguage <> "English" Then
        If InStr(cTemp, " - ") > 0 Then
            cTemp = ClipAmp(cTemp)
            
            cTemp = Left$(cTemp, InStr(cTemp, " - ") + 2) + Translate(Mid$(cTemp, InStr(cTemp, " - ") + 3), mcLanguage)
        Else
            cTemp = Translate(ClipAmp(cTemp), mcLanguage)
        End If
    Else
       cTemp = ClipAmp(cTemp)
    End If
    If iWidth > 0 Then
        cTemp = FitString(frmMain, cTemp, iWidth, 1)
    End If
    If iProperCase = True And cTemp <> "" Then
        Mid$(cTemp, 1, 1) = UCase$(Mid$(cTemp, 1, 1))
    End If
    cTemp = Trim$(cTemp)
    If iAmp = True Then
        If InStr(mcAmpList, UCase$(Left$(cTemp, 1))) = 0 And UCase$(Left$(cTemp, 1)) Like "[0-9A-Z]" Then
            mcAmpList = mcAmpList & UCase$(Left$(cTemp, 1))
            cTemp = "&" & cTemp
            iAmp = False
        Else
            iTemp = 0
            Do
                iTemp = InStr(iTemp + 1, cTemp, " ") > 0
                If iTemp > 0 Then
                    If InStr(mcAmpList, UCase$(Mid$(cTemp, iTemp + 1, 1))) = 0 And UCase$(Mid$(cTemp, iTemp + 1, 1)) Like "[0-9A-Z]" Then
                        mcAmpList = mcAmpList & UCase$(Mid$(cTemp, iTemp + 1, 1))
                        cTemp = Left$(cTemp, iTemp) & "&" & Mid$(cTemp, iTemp + 1)
                        iAmp = False
                        Exit Do
                    End If
                End If
            Loop While iTemp > 0
            If iAmp = True Then
                For iTemp = 1 To Len(cTemp)
                    If InStr(mcAmpList, UCase$(Mid$(cTemp, iTemp, 1))) = 0 And UCase$(Mid$(cTemp, iTemp + 1, 1)) Like "[0-9A-Z]" Then
                        mcAmpList = mcAmpList & UCase$(Mid$(cTemp, iTemp, 1))
                        cTemp = Left$(cTemp, iTemp - 1) & "&" & Mid$(cTemp, iTemp)
                        iAmp = False
                        Exit For
                    End If
                Next iTemp
            End If
        End If
    End If
    If iDots = True Then
        cTemp = cTemp & "..."
    End If
    TranslateCaption = cTemp
    
    On Local Error GoTo 0
    
End Function
Public Function UnDotSpace(cLine As String) As String
    Dim cDotSpace As String
    Dim cTemp As String
    Dim iTemp As Integer
    
    'Limit to letters and numbers only to avoid hazardly file names
    cTemp = ""
    For iTemp = 1 To Len(cLine)
        'If InStr(cDotSpace, Mid$(cLine, iTemp, 1)) = 0 Then
        If UCase$(Mid$(cLine, iTemp, 1)) Like "[A-Z]" Or Mid$(cLine, iTemp, 1) Like "[0-9]" Or InStr("_-.", Mid$(cLine, iTemp, 1)) > 0 Then
            cTemp = cTemp & Mid$(cLine, iTemp, 1)
        End If
    Next iTemp
    UnDotSpace = cTemp
End Function
Public Function Time2Mark(curTime As Currency, Optional cCode As String = "") As Currency
    Dim rstTestTime As DAO.Recordset
    Dim rstRank As DAO.Recordset
    Dim curFast As Currency
    Dim curSlow As Currency
    Dim curRange As Currency
    Dim curStep As Currency
    Dim curTemp As Currency
    
    If cCode = "" Then
        cCode = frmMain.TestCode
    End If
    Time2Mark = 0
    If frmMain.chkFlag.Value = 0 Then
        Set rstTestTime = mdbMain.OpenRecordset("SELECT Tests.Code, TestTimeTables.*, Tests.Type_Time, Tests.Time_Decimals FROM Tests INNER JOIN TestTimeTables ON Tests.Code = TestTimeTables.Code WHERE Tests.Code='" & cCode & "';")
        If rstTestTime.RecordCount = 0 Then
            MsgBox Translate("No proper formula found to calculate marks from times", mcLanguage) & " (" & cCode & ")." & Translate("Update Sport Rules first", mcLanguage) & "!", vbCritical
        ElseIf rstTestTime.Fields("Type_Time") = 2 Then
            If curTime = 0 Then
                Time2Mark = 0
            Else
                curRange = rstTestTime.Fields("ScaleRange")
                Set rstRank = mdbMain.OpenRecordset("SELECT STA,Score FROM Results WHERE Disq=0 AND Score>0 AND Code='" & cCode & "' ORDER BY Score")
                If rstRank.RecordCount > 0 Then
                    curStep = rstTestTime.Fields("ScaleStep")
                    curTemp = curRange
                    Do While Not rstRank.EOF
                        If rstRank.Fields("Score") = curTime Then
                            Exit Do
                        End If
                        curTemp = curTemp - curStep
                        If curTemp = 0 Then
                            Exit Do
                        End If
                        rstRank.MoveNext
                    Loop
                    Time2Mark = curTemp
                Else
                    Time2Mark = curRange
                End If
                rstRank.Close
            End If
        Else
            curFast = rstTestTime.Fields("ScaleFast")
            curSlow = rstTestTime.Fields("ScaleSlow")
            curRange = rstTestTime.Fields("ScaleRange")
            If curTime <= curFast And curTime > 0 Then
                Time2Mark = curRange
            ElseIf curTime >= curSlow Then
                Time2Mark = 0
            ElseIf curTime > 0 Then
                Time2Mark = (curSlow - curTime) * (curRange / (curSlow - curFast))
            End If
            Time2Mark = Val(Replace(Format$(Time2Mark, frmMain.TestTotalFormat), ",", "."))
        End If
        rstTestTime.Close
        Set rstTestTime = Nothing
        Set rstRank = Nothing
    End If
End Function

Public Function GetColorList() As String
    Dim rstColor As DAO.Recordset
    Dim cTemp As String
    
    cTemp = ""
    Set rstColor = mdbMain.OpenRecordset("SELECT Color FROM Colors ORDER BY ColorId")
    If rstColor.RecordCount > 0 Then
        Do While Not rstColor.EOF
            If cTemp = "" Then
                cTemp = rstColor.Fields("Color")
            Else
                cTemp = cTemp & "," & rstColor.Fields("Color")
            End If
            rstColor.MoveNext
        Loop
    End If
    If cTemp = "" Then cTemp = "Blue,Red,White,Yellow,Green,Purple"
    GetColorList = cTemp
    rstColor.Close
    Set rstColor = Nothing
End Function
Public Sub SetListTabStop(ByRef objList As VB.ListBox, ParamArray TabStops() As Variant)
    '   Variablen
    Dim lngCount As Long
    Dim lngTabStops() As Long
    
    '   Bereits vorhandene TabStops löschen
    Call SendMessage(objList.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)

    '   Sind TabStops festgelegt worden?
    If Not (IsMissing(TabStops)) Then
        '   Array für die TabStops anlegen
        ReDim lngTabStops(LBound(TabStops) To UBound(TabStops))

        '   TabStops kopieren
        For lngCount = LBound(TabStops) To UBound(TabStops)
            lngTabStops(lngCount) = TabStops(lngCount)
        Next lngCount
    
        '   Anzahl der TabStops ermitteln
        lngCount = UBound(lngTabStops) - LBound(lngTabStops) + 1

        '   Dann neue TabStops einfügen

        Call SendMessage(objList.hwnd, LB_SETTABSTOPS, lngCount, _
                                                        lngTabStops(0))
    End If
    
    '   Zuletzt Liste neu darstellen
    objList.Refresh
End Sub
Public Sub ParticipantDisqWith(cSta As String, cCode As String, iStatus As Integer, iDW As Integer)
    Dim rst As DAO.Recordset
    Dim iDoLogDB As Integer
    
    ' idw=-1 ' eliminated
    ' idw=-2 ' withdrawn
    ' idw=0  ' take part
    
    Set rst = mdbMain.OpenRecordset("SELECT Deleted FROM Entries WHERE STA='" & cSta & "' AND Code='" & cCode & "' AND Status=" & iStatus)
    If rst.RecordCount > 0 Then
        With rst
            .Edit
            .Fields(0) = iDW
            .Update
        End With
    End If
    
    Set rst = mdbMain.OpenRecordset("SELECT * FROM Results WHERE STA='" & cSta & "' AND Code='" & cCode & "' AND Status=" & iStatus)
    If rst.RecordCount > 0 Then
        With rst
            .Edit
        End With
    Else
        With rst
            .AddNew
            .Fields("Sta") = cSta
            .Fields("Status") = iStatus
            .Fields("Code") = cCode
            .Fields("Position") = 0
            .Fields("FR") = False
            .Fields("Score") = 0
        End With
    End If
    With rst
        .Fields("TimeStamp") = Now
        .Fields("Disq") = iDW
        .Update
    End With
    
  Set rst = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Sta='" & cSta & "' AND Status=" & iStatus & " AND Code='" & cCode & "' AND Section=1")
  With rst
     If .RecordCount = 0 Then
        .AddNew
       .Fields("Sta") = cSta
       .Fields("Status") = iStatus
       .Fields("Code") = cCode
       .Fields("Section") = 1
       .Fields("Flag") = 0
       .Fields("Mark1") = 0
       .Fields("Mark2") = 0
       .Fields("Mark3") = 0
       .Fields("Mark4") = 0
       .Fields("Mark5") = 0
       .Fields("Score") = 0
       .Fields("TimeStamp") = Now
       .Update
    End If
  End With
    
    If miWriteLogDB Then
        'iDoLogDB = WriteLogDBMarks(frmMain.EventName, cCode, iStatus, cSta, 5 + iDW, "", 99)
        iDoLogDB = WriteLogDBMarks2(frmMain.EventName, cCode, iStatus, cSta)
    End If
    
    rst.Close
    Set rst = Nothing
    
End Sub

Public Sub ParticipantNoStart(cSta As String, cCode As String, iStatus As Integer, iDW As Integer)
    Dim rst As DAO.Recordset
    Dim iDoLogDB As Integer
    
    ' idw=-1 ' no start
    
    Set rst = mdbMain.OpenRecordset("SELECT NoStart FROM Entries WHERE STA='" & cSta & "' AND Code='" & cCode & "' AND Status=" & iStatus)
    If rst.RecordCount > 0 Then
        With rst
            .Edit
            .Fields(0) = iDW
            .Update
        End With
    End If
    
    rst.Close
    Set rst = Nothing
    
End Sub

Sub AddOneToPosition(cCode As String, iStatus As Integer)
    Dim iOldposition As Integer
    Dim rstPosition As DAO.Recordset
    
    Set rstPosition = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & cCode & "' AND Status=" & iStatus & " ORDER BY Position")
    If rstPosition.RecordCount > 0 Then
        Do While Not rstPosition.EOF
            With rstPosition
                .Edit
                .Fields("Position") = iOldposition + 1
                iOldposition = .Fields("Position")
                .Update
                .MoveNext
            End With
        Loop
    End If
    rstPosition.Close
    Set rstPosition = Nothing
End Sub
'* From our friend Bill G
Sub ClearBit(iByte As Integer, iBit As Integer)
    Dim iMask As Integer
    ' Create a bitmask with the 2 to the nth power bit set:
    iMask = 2 ^ iBit
    ' Clear the nth Bit:
    iByte = iByte And Not iMask
End Sub
Function ExamineBit(iByte As Integer, iBit As Integer) As Integer
    Dim iMask As Integer
    ' Create a bitmask with the 2 to the nth power bit set:
    iMask = 2 ^ iBit
    ' Return the truth state of the 2 to the nth power bit:
    ExamineBit = ((iByte And iMask) > 0)
End Function
Sub SetBit(iByte As Integer, iBit As Integer)
    Dim iMask As Integer
    ' Create a bitmask with the 2 to the nth power bit set:
    iMask = 2 ^ iBit
    ' Set the nth Bit:
    iByte = iByte Or iMask
End Sub
Sub ToggleBit(iByte As Integer, iBit As Integer)
    Dim iMask As Integer
    ' Create a bitmask with the 2 to the nth power bit set:
    iMask = 2 ^ iBit
    ' Toggle the nth Bit:
    iByte = iByte Xor iMask
End Sub

Public Function OpenDefault(cExtension As String, cDescription As String) As String
    Dim cTemp As String
    Dim lTemp As Long
    Dim bTemp As Integer
    Dim cExeName As String
    
    On Local Error GoTo OpenDefaultError
    
    OpenDefault = ""
    lTemp = GetKeyValue(HKEY_CLASSES_ROOT, cExtension, "", cTemp)
    If cTemp <> "" Then
        lTemp = GetKeyValue(HKEY_CLASSES_ROOT, cTemp, "", cDescription)
        lTemp = GetKeyValue(HKEY_CLASSES_ROOT, cTemp & "\Shell\Open\Command", "", cExeName)
    End If
    
    If cExeName <> "" Then
        lTemp = InStr(cExeName, "/")
        If lTemp > 0 Then cExeName = Left$(cExeName, lTemp - 1)
        lTemp = InStr(cExeName, ",")
        If lTemp > 0 Then cExeName = Left$(cExeName, lTemp - 1)
        lTemp = InStr(cExeName, "%")
        If lTemp > 0 Then cExeName = Left$(cExeName, lTemp - 1)
        OpenDefault = Replace(cExeName, Chr$(34), "")
    End If
    
OpenDefaultError:
    If Err > 0 Then
        LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
        MsgBox App.EXEName & ". Could not find default application for '" & cExtension & "'." & vbCrLf & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
    End If
  
    On Local Error GoTo 0
        
End Function

