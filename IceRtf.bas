Attribute VB_Name = "modRtfLib"
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

Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type CharRange
   cpMin As Long     ' First character of range (0 for start of doc)
   cpMax As Long     ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
   hdc As Long       ' Actual DC to draw on
   hdcTarget As Long ' Target DC for determining text formatting
   rc As Rect        ' Region of the DC to draw to (in twips)
   rcPage As Rect    ' Region of the entire DC (page size) (in twips)
   chrg As CharRange ' Range of text to draw (see above declaration)
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal MSG As Long, ByVal wp As Long, lp As Any) As Long
Public Sub RtfSpan(r As RichTextBox, lStart As Long, cVoor As String, cAchter As String, cRetour As String)
    If lStart >= 0 Then
        r.SelStart = lStart
    End If
    r.Span cAchter, False, True
    r.Span cVoor, True, True
    cRetour = Trim$(r.SelText)
    r.SelLength = 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' PrintRTF - Prints the contents of a RichTextBox control using the
'            provided margins
'
' RTF - A RichTextBox control to print
'
' LeftMarginWidth - Width of desired left margin in twips
'
' TopMarginHeight - Height of desired top margin in twips
'
' RightMarginWidth - Width of desired right margin in twips
'
' BottomMarginHeight - Height of desired bottom margin in twips
'
' Notes - If you are also using WYSIWYG_RTF() on the provided RTF
'         parameter you should specify the same LeftMarginWidth and
'         RightMarginWidth that you used to call WYSIWYG_RTF()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintRtf(RTF As RichTextBox, Optional LeftMarginWidth As Long = 25, Optional TopMarginHeight As Long = 25, Optional RightMarginWidth As Long = 25, Optional BottomMarginHeight As Long = 25)
   Dim LeftOffset As Long, TopOffset As Long
   Dim LeftMargin As Long, TopMargin As Long
   Dim RightMargin As Long, BottomMargin As Long
   Dim fr As FormatRange
   Dim rcDrawTo As Rect
   Dim rcPage As Rect
   Dim TextLength As Long
   Dim NextCharPosition As Long
   Dim NextFFPosition As Long
   Dim r As Long

   On Local Error Resume Next
   
   ' Start a print job to get a valid Printer.hDC
   Printer.Print Space(1)
   If Err <> 0 Then
        MsgBox Translate("No printer available at the moment", mcLanguage) & "!", vbExclamation
        Exit Sub
   End If
   Printer.ScaleMode = vbTwips
   On Local Error GoTo 0

   ' Get the offsett to the printable area on the page in twips
   LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX), vbPixels, vbTwips)
   TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY), vbPixels, vbTwips)

   ' Calculate the Left, Top, Right, and Bottom margins
   LeftMargin = LeftMarginWidth * 56.7 - LeftOffset
   TopMargin = TopMarginHeight * 56.7 - TopOffset
   RightMargin = (Printer.Width - RightMarginWidth * 56.7) - LeftOffset
   BottomMargin = (Printer.Height - BottomMarginHeight * 56.7) - TopOffset

   ' Set printable area rect
   rcPage.Left = 0
   rcPage.Top = 0
   rcPage.Right = Printer.ScaleWidth
   rcPage.Bottom = Printer.ScaleHeight

   ' Set rect in which to print (relative to printable area)
   rcDrawTo.Left = LeftMargin
   rcDrawTo.Top = TopMargin
   rcDrawTo.Right = RightMargin
   rcDrawTo.Bottom = BottomMargin

   ' Set up the print instructions
   fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
   fr.hdcTarget = Printer.hdc  ' Point at printer hDC
   fr.rc = rcDrawTo            ' Indicate the area on page to draw to
   fr.rcPage = rcPage          ' Indicate entire size of page
   fr.chrg.cpMin = 0           ' Indicate start of text through
   fr.chrg.cpMax = -1          ' end of the text

   ' Get length of text in RTF
   TextLength = Len(RTF.Text)
   
   NextCharPosition = 0
   ' Loop printing each page until done
   Do While NextCharPosition < TextLength - 1
   ' Print the page by sending EM_FORMATRANGE message
        
        NextCharPosition = SendMessage(RTF.hwnd, EM_FORMATRANGE, True, fr)
        If NextCharPosition >= TextLength - 1 Then
            Exit Do                                 'If done then exit
        End If
        Do While NextCharPosition < TextLength And (Mid$(RTF.Text, NextCharPosition + 1, 1) = vbCr Or Mid$(RTF.Text, NextCharPosition + 1, 1) = vbLf Or Mid$(RTF.Text, NextCharPosition + 1, 1) = vbFormFeed)
            NextCharPosition = NextCharPosition + 1
        Loop
        If NextCharPosition < TextLength Then
            fr.chrg.cpMin = NextCharPosition - 1    ' Starting position for next page
            Printer.NewPage                         ' Move on to next page
            Printer.Print Space(1)                  ' Re-initialize hDC
            fr.hdc = Printer.hdc
            fr.hdcTarget = Printer.hdc
        End If
   Loop

   ' Commit the print job
   Printer.EndDoc

   ' Allow the RTF to free up memory
   r = SendMessage(RTF.hwnd, EM_FORMATRANGE, False, ByVal CLng(0))
   
End Sub
 
Public Sub EditRtfText(cTitle As String, cTable As String)
    Dim rstRTF As DAO.Recordset
    Dim cExeName As String
    Dim cDescription As String
    
    On Local Error GoTo EditRtfTextError
   
    Set rstRTF = mdbMain.OpenRecordset("SELECT * FROM [" & cTable & "] WHERE [Title] Like " & Chr$(34) & cTitle & Chr$(34))
    If rstRTF.RecordCount > 0 Then
        cExeName = OpenDefault(".RTF", cDescription)
        If cExeName = "" Then
            cExeName = OpenDefault(".DOC", cDescription)
        End If
        
        If cExeName = "" Then
            cExeName = OpenDefault(".Rtf", cDescription)
        End If
        If cExeName = "" Then
            cExeName = OpenDefault(".Doc", cDescription)
        End If
        If cExeName = "" Then
           MsgBox Translate("No proper editor found.", mcLanguage)
           Exit Sub
        End If
        
        frmPrint.rtfPrint.TextRTF = rstRTF.Fields("RTFText") & ""
        
        StartRTFEditor frmPrint.rtfPrint, cDescription, cExeName, True, True
         
        With rstRTF
             .Edit
             .Fields("RtfText") = frmPrint.rtfPrint.TextRTF
             .Fields("Editor") = Left(UserName, .Fields("Editor").Size)
             .Update
             .Close
        End With
        Set rstRTF = Nothing
    Else
        MsgBox Translate("Form not found.", mcLanguage)
    End If

EditRtfTextError:
    If Err > 0 Then
        LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
        MsgBox App.EXEName & ". Could not find RTF-editor" & vbCrLf & Err.Source & ": " & Err.Number & ": " & Err.Description, vbCritical
    End If

    On Local Error GoTo 0
End Sub

Public Sub StartRTFEditor(r As RichTextBox, cDescription As String, cExeName As String, Optional iMinimize As Integer = True, Optional iSaveAlways As Integer = False)
    Dim iTemp As Integer
    Dim cTemp As String
    Dim iFirstTime As Integer
    Dim iWindowstate As Integer
    
    Dim cTmpRtfFile As String
    Dim cSequence As String
    Dim vRetValue As Variant
    
    r.Enabled = False
    cSequence = "0000"
    Do
        If Dir$(TmpDir & Left$(App.EXEName, 4) & cSequence & ".RTF") = "" Then
            Exit Do
        Else
            cSequence = Format$(Val(cSequence) + 1, "0000")
        End If
    Loop While cSequence <> "9999"
    
    cTmpRtfFile = TmpDir & Left$(App.EXEName, 4) & cSequence & ".RTF"
    r.SaveFile cTmpRtfFile
    cTemp = FileDateTime(cTmpRtfFile)
    vRetValue = Shell(cExeName & " " & Chr$(34) & cTmpRtfFile & Chr$(34), vbNormalFocus)
    
    giExternalOperation = True
    iWindowstate = frmMain.WindowState
    
    If iMinimize = True Then
        frmMain.WindowState = vbMinimized
    End If
    
    Do While giExternalOperation = True
        DoEvents
        Sleep 1
        If FileDateTime(cTmpRtfFile) <> cTemp Then
            Exit Do
        ElseIf iMinimize = True And frmMain.WindowState <> vbMinimized Then
            Exit Do
        End If
    Loop
    
    If FileDateTime(cTmpRtfFile) = cTemp And iSaveAlways = False Then
       iTemp = MsgBox(Translate("Do you want to save this text in", mcLanguage) & " " & App.EXEName & " ?" & vbCrLf & Translate("Don't forget to save the text first.", mcLanguage), vbYesNo + vbQuestion)
    Else
       iTemp = vbYes
    End If
    
    If iTemp = vbYes Then
        StatusMessage Translate("Text saved in", mcLanguage) & " " & App.EXEName
        iFirstTime = True
        On Local Error Resume Next
        Do
            If iFirstTime = False Then
                DoEvents
                StatusMessage App.EXEName & " " & Translate("waits for text to be saved.", mcLanguage)
                Sleep 1
            Else
                iFirstTime = False
            End If
            r.LoadFile cTmpRtfFile
            KillFile cTmpRtfFile
        Loop While Dir$(cTmpRtfFile) <> ""
        On Local Error GoTo 0
    End If
    
    giExternalOperation = False
    frmMain.WindowState = iWindowstate
    r.Enabled = True
    
    StatusMessage ""

End Sub

Public Sub NewRTF()
    Dim cTemp As String
    Dim rstRTF As DAO.Recordset
    
    cTemp = InputBox(Translate("What is the new form called?", mcLanguage))
    If cTemp <> "" And cTemp <> Chr$(27) Then
        Set rstRTF = mdbMain.OpenRecordset("SELECT * FROM [Forms] WHERE [Title] LIKE " & Chr$(34) & cTemp & Chr$(34))
        If rstRTF.RecordCount = 0 Then
            With rstRTF
                .AddNew
                .Fields("Title") = Left$(cTemp, .Fields("Title").Size)
                .Fields("Owner") = Left$(UserName, .Fields("Owner").Size)
                .Fields("Editor") = Left$(UserName, .Fields("Editor").Size)
                .Update
            End With
        End If
        rstRTF.Close
        Set rstRTF = Nothing
        EditRtfText cTemp, "Forms"
    End If
End Sub

Public Sub DeleteRTF()
   Dim iKey As Integer
   
   SetMouseHourGlass
   
   With frmToolBox
      .intChecked = False
      .strQry = "SELECT [Title] FROM [Forms] WHERE Owner & '' <> 'FEIF' ORDER BY [Title]"
      .Caption = Translate("Select a form to delete", mcLanguage)
      SetMouseNormal
      .Show 1
   End With
   
   If frmMain.Tempvar <> "" Then
        iKey = MsgBox(Translate("Delete", mcLanguage) & " '" & frmMain.Tempvar & "' ?", vbQuestion + vbYesNo)
        If iKey = vbYes Then
            mdbMain.Execute "DELETE * FROM [Forms] WHERE [Title] LIKE " & Chr$(34) & frmMain.Tempvar & Chr$(34)
        End If
        frmMain.Tempvar = ""
   End If
End Sub

Public Sub EditRtf()
   
   SetMouseHourGlass
   With frmToolBox
      .intChecked = False
      If InStr(Command$, "/SYS") > 0 Then
          .strQry = "SELECT [Title] FROM [Forms] ORDER BY [Title]"
      Else
          .strQry = "SELECT [Title] FROM [Forms] WHERE [Owner]&''<>'FEIF' ORDER BY [Title]"
      End If
      .Caption = Translate("Select a form to edit", mcLanguage)
      SetMouseNormal
      .Show 1
   End With
   
   If frmMain.Tempvar <> "" Then
       EditRtfText frmMain.Tempvar, "Forms"
       frmMain.Tempvar = ""
   End If
   
End Sub

Public Sub PrintRtfForms()
   Dim rstRTF As DAO.Recordset
   Dim rstSta As DAO.Recordset
   Dim cFormName As String
   Dim cTemp As String
   Dim cTemp2 As String
   Dim cQry As String
   Dim iTemp As Integer
   Dim iTemp2 As Integer
   Dim lngTemp As Long
   Dim iFileNum As Integer
   Dim iKey As Integer
   Dim cStaList As String
   Dim iFormFeed As Integer
   Dim strPar As String
   Dim iSelectPerTest As Integer
   
   SetMouseHourGlass
   
   With frmToolBox
      .intChecked = False
      .strQry = "SELECT [Title] FROM [Forms] ORDER BY [Title]"
      .Caption = Translate("Select a form to compose and print", mcLanguage)
      SetMouseNormal
      .Show 1
   End With
   
   If frmMain.Tempvar <> "" Then
        iSelectPerTest = False
        cFormName = frmMain.Tempvar
        iFormFeed = True
        frmMain.Tempvar = ""
        frmPrint.rtfPrint.Text = ""
        frmMain.rtfResult.Text = ""
        Set rstRTF = mdbMain.OpenRecordset("SELECT * FROM [Forms] WHERE [Title] LIKE " & Chr$(34) & cFormName & Chr$(34))
        If rstRTF.RecordCount = 0 Then
            MsgBox Translate("Form not found.", mcLanguage)
            Exit Sub
        Else
            If rstRTF.Fields("FormType") = 1 Then
                cTemp = rstRTF.Fields("RTFText") & ""
                cTemp = Replace(cTemp, "<EVENTNAME>", frmMain.EventName & "")
                cTemp = Replace(cTemp, "<EVENTCODE>", GetVariable("WR_code"))
                cTemp = Replace(cTemp, "<TestCode>", frmMain.TestCode & "")
                cTemp = Replace(cTemp, "<TestName>", frmMain.TestName & "")
                With frmMain.rtfResult
                    .SelRTF = cTemp & vbCrLf
                End With
                
                frmMain.MakeRtfFooter
            Else
                 iKey = MsgBox(Translate("Print", mcLanguage) & " " & cFormName & " " & Translate("for all participants (Yes) or for selected participants (No).", mcLanguage), vbYesNoCancel + vbDefaultButton2 + vbQuestion)
                 If iKey = vbYes Then
                     cStaList = ""
                 ElseIf iKey = vbNo Then
                     iKey = MsgBox(Translate("Select participants by Starting number or Name (Yes) or by Test (No).", mcLanguage), vbYesNoCancel + vbDefaultButton1 + vbQuestion)
                     If iKey = vbYes Then
                        cTemp = InputBox$(Translate("Search for", mcLanguage), Translate("Participants", mcLanguage), "")
                        cQry = "SELECT Participants.Sta "
                        cQry = cQry & " & '  -  ' & Persons.Name_First"
                        cQry = cQry & " & ' ' & Persons.Name_Last"
                        cQry = cQry & " & IIF(Participants.Class<>'',' ['  & participants.Class & ']','') "
                        cQry = cQry & " & ' - ' & Horses.Name_Horse as cList"
                        cQry = cQry & " FROM (Participants"
                        cQry = cQry & " INNER JOIN Persons"
                        cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
                        cQry = cQry & " INNER JOIN Horses"
                        cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
                        cQry = cQry & " WHERE Participants.Sta & ' ' & Persons.Name_First & ' ' & Persons.Name_Last & ' - ' & Horses.Name_Horse LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
                        cQry = cQry & " ORDER BY Participants.Sta"
                     ElseIf iKey = vbNo Then
                        iSelectPerTest = True
                        cQry = "SELECT Code & ' - ' & Test AS cList"
                        cQry = cQry & " FROM Tests "
                        cQry = cQry & " WHERE Code IN (SELECT Code FROM Entries)"
                        cQry = cQry & " ORDER BY Code,Test"
                    Else
                        rstRTF.Close
                        Set rstRTF = Nothing
                        Exit Sub
                    End If
                    frmToolBox.strQry = cQry
                    frmToolBox.intChecked = True
                    frmToolBox.intReturnLen = 3
                    frmToolBox.Caption = Translate("Please select", mcLanguage)
                    frmToolBox.Show 1, frmMain
                    If frmMain.Tempvar = "" Then
                        rstRTF.Close
                        Set rstRTF = Nothing
                        Exit Sub
                    End If
                    cStaList = "'" & Replace(frmMain.Tempvar, "|", "','") & "'"
                 Else
                     rstRTF.Close
                     Set rstRTF = Nothing
                     Exit Sub
                 End If
                        
                cQry = "SELECT Participants.STA AS STA, "
                cQry = cQry & " Persons.Name_First & ' ' & Persons.Name_Last AS Rider, "
                cQry = cQry & " Participants.Club AS Club, "
                cQry = cQry & " Horses.Name_Horse AS Horse, "
                cQry = cQry & " IIF(Horses.Sex_Horse = 1,'Stallion',IIF(Horses.Sex_Horse = 2,'Mare',IIF(Horses.Sex_Horse = 3,'Gelding','-'))) AS Gender, "
                cQry = cQry & " Persons.FEIFId AS Rider_FEIFID,"
                cQry = cQry & " Persons.FEIFId AS Person_FEIFID,"
                cQry = cQry & " Horses.FEIFId AS Horse_FEIFID,"
                cQry = cQry & " *"
                cQry = cQry & " FROM (Participants "
                cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) "
                cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
                If iSelectPerTest = True Then
                    cQry = cQry & " WHERE Participants.STA IN (SELECT STA FROM Entries WHERE Code IN (" & cStaList & "))"
                ElseIf cStaList <> "" Then
                    cQry = cQry & " WHERE Participants.STA IN (" & cStaList & ")"
                End If
                cQry = cQry & " ORDER BY Participants.STA"
                
                Set rstSta = mdbMain.OpenRecordset(cQry)
                If rstSta.RecordCount > 0 Then
                    rstSta.MoveLast
                    rstSta.MoveFirst
                    iKey = MsgBox(Translate("Print", mcLanguage) & " " & rstSta.RecordCount & " " & Translate("forms", mcLanguage) & " " & Translate("with a page brake between each form (YES) or without a page break between the forms (NO)", mcLanguage) & "?", vbYesNoCancel + vbDefaultButton1)
                    
                    If iKey = vbYes Or iKey = vbNo Then
                        If iKey = vbYes Then
                            iFormFeed = True
                        Else
                            iFormFeed = False
                        End If
                        Do While Not rstSta.EOF
                            With frmMain.rtfResult
                                If rstSta.EOF = False And .SelStart > 0 Then
                                    If iFormFeed = True Then
                                        .SelText = "$#@!" & vbCrLf
                                    End If
                                End If
                            End With
                            cTemp = rstRTF.Fields("RTFText") & ""
                            cTemp = Replace(cTemp, "<EVENTNAME>", frmMain.EventName & "")
                            cTemp = Replace(cTemp, "<TestCode>", frmMain.TestCode & "")
                            cTemp = Replace(cTemp, "<TestName>", frmMain.TestName & "")
    
                            For iTemp = 0 To rstSta.Fields.Count - 1
                                cTemp2 = rstSta.Fields(iTemp).Name
                                iTemp2 = InStr(cTemp2, ".")
                                cTemp2 = RTrim$(Mid$(cTemp2 & " ", iTemp2 + 1))
                                cTemp = Replace(cTemp, "<" & cTemp2 & ">", rstSta.Fields(iTemp) & "")
                            Next iTemp
                            With frmMain.rtfResult
                                .SelRTF = cTemp & vbCrLf
                            End With
                            
                            frmMain.MakeRtfFooter
                            
                            rstSta.MoveNext
                        Loop
                    End If
                    rstSta.Close
                    Set rstSta = Nothing
                End If
                
                For iTemp = 1 To 50
                    strPar = strPar & "\par "
                Next iTemp
    
                frmPrint.rtfPrint.TextRTF = frmMain.rtfResult.TextRTF
                frmPrint.rtfPrint.TextRTF = Replace(frmPrint.rtfPrint.TextRTF, "$#@!", strPar)
                frmPrint.fcFocus = "Preview"
                frmPrint.Show 1, frmMain
                
                If frmMain.Tempvar = "Preview" Or frmMain.Tempvar = "Print" Then
                    
                    If Dir$(mcRtfDir, vbDirectory) = "" Then
                        MkDir mcRtfDir
                    End If
                    
                    cTemp = mcRtfDir & Replace(Replace(cFormName, " ", "_"), ".", "_") & ".Rtf"
                    lngTemp = KillFile(cTemp)
                    Printer.PaperSize = vbPRPSA4
                    frmMain.rtfResult.SaveFile cTemp
                    
                    iFileNum = FreeFile
                    Open cTemp For Binary Access Read Write Shared As #iFileNum
                    cTemp2 = Space$(LOF(iFileNum))
                    Get #iFileNum, 1, cTemp2
                    iTemp = InStr(cTemp2, ";}}")
                    If iTemp > 0 Then
                        cTemp2 = Left$(cTemp2, iTemp + 2) & "\paperw11907\paperh16840\margl1418\margr704\margt1418\margb1418" & Mid$(RTrim$(cTemp2 & " "), iTemp + 3)
                    End If
                    cTemp2 = Replace(cTemp2, "$#@!", "{\f1\fs20\page}")
                    Put #iFileNum, 1, cTemp2
                    Close #iFileNum
                    DoEvents
                    
                    If frmMain.Tempvar = "Preview" Then
                        If ShowDocument(cTemp, frmMain) = 3 Then
                            MsgBox Translate("Please install editor for RTF-files first (like MS Word).", mcLanguage), vbExclamation
                        End If
                    End If
                End If
                frmMain.rtfResult.Text = ""
                frmPrint.rtfPrint.Text = ""
            End If
            
            rstRTF.Close
            Set rstRTF = Nothing
            
            frmMain.Tempvar = ""
            SetMouseNormal
       End If
       frmMain.Tempvar = ""
   End If
   SetMouseNormal

End Sub
