Attribute VB_Name = "modIceHTML"
' IceHtml.bas
' modIceHTML: Provide functionality for IceTest data export in HTML.
' Copyright (C) Marko Mazeland, Lutz Lesener 2006, 2007
'
' This file is part of the FEIF software project.
' See http://www.feif.org/software or https://sourceforge.net/projects/icehorsetools/ for details.
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

' Checks the highest existing STA and
' produces accordingly detail description files (99 per file)
Public Function CreateParticipantsFile()
    Dim rstQuery As Recordset
    'Dim bigsta As Integer
    Dim sta As Integer
    Dim X As Integer
    Dim t As Integer
    
    Set rstQuery = mdbMain.OpenRecordset("SELECT sta FROM participants ORDER BY sta DESC;")
    
    If Not rstQuery.RecordCount = 0 Then
    
        rstQuery.MoveLast
        rstQuery.MoveFirst
    
         
         StatusMessage Translate("Wait", mcLanguage) & " ..."
         DoEvents
         
        'bigsta = CInt(Left$(rstQuery("sta"), 1))
        
        'For x = 1 To bigsta
        '    t = Create100HTMLDetails(CStr(x))
        'Next x
        
        'MM - 28-7-2007: back to one page per participant
        Do While Not rstQuery.EOF
            CreateHTMLDetails rstQuery.Fields("Sta"), True
            rstQuery.MoveNext
        Loop
        rstQuery.Close
        Set rstQuery = Nothing
        
        StatusMessage ""
        
    End If
    
    DoEvents
    
End Function
Function CreateHTMLHeader(Title As String) As String
    Dim cHtml As String
    
    cHtml = "<!-- Created by IceTest -->" & vbCrLf
    cHtml = cHtml & "<!DOCTYPE html>"
    cHtml = cHtml & "<html>" & vbCrLf
    cHtml = cHtml & "<!-- CreateHTMLHeader start -->" & vbCrLf
    cHtml = cHtml & "<head>" & vbCrLf
    cHtml = cHtml & "<meta charset=""iso-8859-1"">" & vbCrLf
    cHtml = cHtml & "<meta name=""generator"" content=""IceTest"">"
    cHtml = cHtml & "<meta http-equiv=""refresh"" content=""60"">"
    cHtml = cHtml & "<title>"
    cHtml = cHtml & Title
    cHtml = cHtml & "</title>" & vbCrLf
    cHtml = cHtml & "<link rel=""stylesheet"" type=""text/css"" href=""iceweb.css"">" & vbCrLf
    cHtml = cHtml & "</head>" & vbCrLf
    cHtml = cHtml & "<!-- CreateHTMLHeader end -->" & vbCrLf
    
    CreateHTMLHeader = cHtml
End Function
' Writes the details of 100 participants to a HTML file
Function Create100HTMLDetails(FirstDigit As String) As Integer
    Dim cHtml As String
    Dim cTemplate As String
    Dim cFooter As String
    Dim cHtmlFileName As String
    Dim cModel As String
    Dim iHtmlFileNum As Integer
    Dim rstHorse As Recordset
    Dim rstRider As Recordset
    Dim cWR2_Url As String
    
    cWR2_Url = GetVariable("WR2_Url")

    
    FirstDigit = Left$(FirstDigit, 1)
    
    cHtml = "<!-- Create100HTMLDetails begin -->" & vbCrLf
    cHtml = cHtml & "<h1>" & frmMain.EventName & "</h1>" & vbCrLf
    
    Set rstRider = mdbMain.OpenRecordset("SELECT Participants.*, Persons.* FROM Persons INNER JOIN Participants ON Persons.PersonID = Participants.PersonID WHERE (((Participants.STA) LIKE '" & FirstDigit & "??')) ORDER BY Participants.STA;")
    Set rstHorse = mdbMain.OpenRecordset("SELECT Participants.*, Horses.* FROM Horses INNER JOIN Participants ON Horses.HorseID = Participants.HorseID WHERE (((Participants.STA) LIKE '" & FirstDigit & "??')) ORDER BY Participants.STA;")
    
    While Not rstRider.EOF And Not rstRider.BOF
        cHtml = cHtml & "<a name=""participant" & rstRider("sta") & """>" & vbCrLf
        cHtml = cHtml & Translate("Startnumber", mcLanguage) & " " & rstRider("STA") & "<br />" & vbCrLf
        cHtml = cHtml & "<h2>" & rstRider("Name_First") & " " & rstRider("Name_Last") & "</h2>" & vbCrLf
    
        If rstRider("Class") & "" <> "" Then
            cHtml = cHtml & "Class: " & rstRider("Class") & "<br />" & vbCrLf
        End If
    
        If rstRider("Club") & "" <> "" Then
            cHtml = cHtml & "Club: " & rstRider("Club") & "<br />" & vbCrLf
        End If
        
        If rstRider("Team") & "" <> "" Then
            cHtml = cHtml & "Team: " & rstRider("team") & "<br />" & vbCrLf
        End If

        cHtml = cHtml & "<br /><br />" & vbCrLf

        cHtml = cHtml & "<h2>" & rstHorse("name_horse") & "</h2>" & vbCrLf
    
        Select Case rstHorse("Sex_Horse")
            Case 1
                cHtml = cHtml & Translate("Stallion", mcLanguage) & " " & vbCrLf
            Case 2
                cHtml = cHtml & Translate("Mare", mcLanguage) & " " & vbCrLf
            Case 3
                cHtml = cHtml & Translate("Gelding", mcLanguage) & " " & vbCrLf
        End Select
    
        cHtml = cHtml & Format$(rstHorse("Birthday_horse"), "YYYY") & vbCrLf
        cHtml = cHtml & " (" & rstHorse("Country_Horse") & ")"
    
        If rstHorse("FEIFID") & "" <> "" Then
            cHtml = cHtml & ", <a href=""" & cWR2_Url & "wr_x" & LCase$(rstHorse("FEIFID")) & ".html#top"">" & rstHorse("FEIFID") & "</a><br />" & vbCrLf
        Else
            cHtml = cHtml & "<br />" & vbCrLf
        End If
    
        cHtml = cHtml & "<br /><br /><table><tr>" & vbCrLf
        cHtml = cHtml & "<td>F: " & rstHorse("F") & "</td>" & vbCrLf
        cHtml = cHtml & "<td>&nbsp;</td>" & vbCrLf
        cHtml = cHtml & "<td>M: " & rstHorse("M") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>-FF: " & rstHorse("FF") & "</td><td >&nbsp;</td>"
        cHtml = cHtml & "<td>-MF: " & rstHorse("MF") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>-FM: " & rstHorse("FM") & "</td><td >&nbsp;</td>"
        cHtml = cHtml & "<td>-MM: " & rstHorse("MM") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "</table><br /><br />" & vbCrLf
    
        cHtml = cHtml & "<table>" & vbCrLf
        cHtml = cHtml & "<tr><td>" & Translate("Breeder", mcLanguage) & ": </td><td>" & rstHorse("Breeder") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>" & Translate("Owner", mcLanguage) & ": </td><td>" & rstHorse("Owner") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "</table>" & vbCrLf
    
        cHtml = cHtml & "<br /><a href=""javascript:history.back();"">" & Translate("back", mcLanguage) & "</a><br /><br />" & vbCrLf
        
        cHtml = cHtml & "<br /><hr /><br />" & vbCrLf
        
        
        rstRider.MoveNext
        rstHorse.MoveNext
    Wend
    
  
    cHtml = cHtml & "<!-- CreateHTMLDetails end -->" & vbCrLf
    
    cModel = GetHTMLTemplate
    cModel = Replace(cModel, "{eventname}", frmMain.EventName)
    cModel = Replace(cModel, "{title}", Translate("Participants", mcLanguage) & " " & FirstDigit & "00-" & FirstDigit & "99")
    cModel = Replace(cModel, "{body}", cHtml)
    
    cFooter = "<footer>"
    cFooter = cFooter & Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss")
    cFooter = cFooter & "<BR>" & App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
    cFooter = cFooter & "</footer>"
    cModel = Replace(cModel, "{footer}", cFooter)

    'Write details to participant_[cSTA].html:
    iHtmlFileNum = FreeFile
    cHtmlFileName = StrConv("participants" & FirstDigit & ".html", vbLowerCase)
    Open mcTempHtmlDir & cHtmlFileName For Output Access Write Shared As iHtmlFileNum
    Print #iHtmlFileNum, cModel
    Close #iHtmlFileNum
    
    'Return file name (without path) to calling code:
    'Create100HTMLDetails = cHtmlFileName

End Function
Function CreateHTMLDetails(cSta As String, Optional iForced As Integer = False) As String
    Dim cHtml As String
    Dim cModel As String
    Dim cHtmlFileName As String
    Dim iHtmlFileNum As Integer
    Dim rstHorse As Recordset
    Dim rstRider As Recordset
    Dim cQry As String
    Dim cOldCode As String
    Dim cMarkFormat As String
    Dim cTimeFormat As String
    Dim cTotalFormat As String
    Dim cFooter As String
    
    Dim iOldStatus As Integer
    Dim iOldSection As Integer
    Dim iMarkDecimals As Integer
    Dim iTimeDecimals As Integer
    Dim iTemp As Integer
    
    Dim rstMarks As Recordset
    Dim rstTest As Recordset
    
    Dim cWR2_Url As String
    
    cWR2_Url = GetVariable("WR2_Url")
    If cWR2_Url = "" Then
        cWR2_Url = "https://www.feif.org/files/wr/wr/"
    End If
    
    On Local Error Resume Next
    
    cHtmlFileName = StrConv("participant_" & cSta & ".html", vbLowerCase)
    
    'MM 28-7-2007: skip file creation if it already exists unless when it is forced (changed results)
    If Dir$(mcHtmlDir & cHtmlFileName) = "" Or iForced = True Then
        
        Set rstRider = mdbMain.OpenRecordset("SELECT Participants.*, Persons.* FROM Persons INNER JOIN Participants ON Persons.PersonID = Participants.PersonID WHERE (((Participants.STA)='" & cSta & "'));")
        Set rstHorse = mdbMain.OpenRecordset("SELECT Participants.*, Horses.* FROM Horses INNER JOIN Participants ON Horses.HorseID = Participants.HorseID WHERE (((Participants.STA)='" & cSta & "'));")
        
        cHtml = cHtml & "<!-- CreateHTMLDetails begin -->" & vbCrLf
        cHtml = cHtml & rstRider("STA") & vbCrLf
        cHtml = cHtml & " <b>" & rstRider("Name_First") & " " & rstRider("Name_Last") & " " & vbCrLf
        
        If rstRider("Class") & "" <> "" Then
            cHtml = cHtml & " [" & rstRider("Class") & "]" & vbCrLf
        End If
        
        If rstRider("Club") & "" <> "" Then
            cHtml = cHtml & " - Club: " & rstRider("Club") & vbCrLf
        End If
            
        If rstRider("Team") & "" <> "" Then
            cHtml = cHtml & " - Team: " & rstRider("team") & vbCrLf
        End If
        cHtml = cHtml & "<b><br />"
        
        cHtml = cHtml & "<b>" & rstHorse("name_horse") & "<b>" & vbCrLf
        
        If rstHorse("FEIFID") & "" <> "" Then
            If cWR2_Url <> "" Then
                cHtml = cHtml & " [<a href=""" & cWR2_Url & "wr_x" & LCase$(rstHorse("FEIFID")) & ".html#top"">" & rstHorse("FEIFID") & "</a>]" & vbCrLf
            Else
                cHtml = cHtml & " [" & rstHorse("FEIFID") & "]" & vbCrLf
            End If
            
        End If
        
        Select Case rstHorse("Sex_Horse")
            Case 1
                cHtml = cHtml & " - " & Translate("Stallion", mcLanguage) & vbCrLf
            Case 2
                cHtml = cHtml & " - " & Translate("Mare", mcLanguage) & vbCrLf
            Case 3
                cHtml = cHtml & " - " & Translate("Gelding", mcLanguage) & vbCrLf
        End Select
        
        cHtml = cHtml & " - " & Format$(rstHorse("Birthday_horse"), "YYYY") & vbCrLf
        If rstHorse("Country_Horse") & "" <> "" Then
            cHtml = cHtml & " - (" & rstHorse("Country_Horse") & ")"
        End If
        
        cHtml = cHtml & "<br /><br /><table><tr>" & vbCrLf
        cHtml = cHtml & "<td>F: " & rstHorse("F") & "</td>" & vbCrLf
        cHtml = cHtml & "<td>&nbsp;</td>" & vbCrLf
        cHtml = cHtml & "<td>M: " & rstHorse("M") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>-FF: " & rstHorse("FF") & "</td><td>&nbsp;</td>"
        cHtml = cHtml & "<td>-MF: " & rstHorse("MF") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>-FM: " & rstHorse("FM") & "</td><td>&nbsp;</td>"
        cHtml = cHtml & "<td>-MM: " & rstHorse("MM") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "</table><br />" & vbCrLf
        
        cHtml = cHtml & "<table>" & vbCrLf
        cHtml = cHtml & "<tr><td>" & Translate("Breeder", mcLanguage) & ": </td><td>" & rstHorse("Breeder") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>" & Translate("Owner", mcLanguage) & ": </td><td>" & rstHorse("Owner") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "</table>" & vbCrLf
        
        cHtml = cHtml & "<!-- CreateHTMLDetails end -->" & vbCrLf
        
        cQry = "SELECT Results.STA,Results.Code, Results.Status,Results.Disq,Results.Score,Results.Position,Marks.Section, Marks.Mark1, Marks.Mark2, Marks.Mark3, Marks.Mark4, Marks.Mark5, Marks.Score,Tests.Test,TestInfo.Num_j_0,TestInfo.Num_j_1,TestInfo.Num_j_2,TestInfo.Num_j_3"
        cQry = cQry & " FROM ((Results INNER JOIN Marks ON (Results.Status = Marks.Status) AND (Results.STA = Marks.STA) AND (Results.Code = Marks.Code)) INNER JOIN Tests ON Marks.Code = Tests.Code) INNER JOIN TestInfo ON Marks.Code = TestInfo.Code"
        cQry = cQry & " WHERE Results.STA='" & cSta & "'"
        cQry = cQry & " ORDER BY Results.STA, Results.Code, Results.Status=0, Results.Status DESC , Marks.Section;"
    
        Set rstMarks = mdbMain.OpenRecordset(cQry)
        If rstMarks.RecordCount > 0 Then
            cHtml = cHtml & "<br /><br /><table><tr>" & vbCrLf
            Do While Not rstMarks.EOF
                If rstMarks.Fields("Code") <> cOldCode Or rstMarks.Fields("Status") <> iOldStatus Or iOldSection <> rstMarks.Fields("Section") Then
                    cOldCode = rstMarks.Fields("Code")
                    iOldStatus = rstMarks.Fields("Status")
                    If rstMarks.Fields("Section") = 1 Then
                        cHtml = cHtml & "<td colspan=""3""><b>" & vbCrLf
                        cHtml = cHtml & cOldCode & " - " & Translate(rstMarks.Fields("Test"), mcLanguage) & IIf(iOldStatus = 0, "", IIf(iOldStatus = 1, " - " & Translate("A-Final", mcLanguage), IIf(iOldStatus = 2, " - " & Translate("B-Final", mcLanguage), " - " & Translate("C-Final", mcLanguage)))) & vbCrLf
                    End If
                    iOldSection = rstMarks.Fields("Section")
                    Set rstTest = mdbMain.OpenRecordset("SELECT * FROM Tests INNER JOIN Testsections ON Tests.Code=Testsections.Code WHERE Tests.Code='" & cOldCode & "' AND TestSections.Section=" & iOldSection & " AND TestSections.Status=" & IIf(iOldStatus = 0, 0, 1))
                End If
                If rstTest.RecordCount > 0 Then
                    cHtml = cHtml & "</b></td><tr><td>" & vbCrLf
                    cHtml = cHtml & "&nbsp;</td><td>" & vbCrLf
                    cHtml = cHtml & Left$(Translate(rstTest.Fields("Name"), mcLanguage), 20) & vbCrLf
                    cHtml = cHtml & "&nbsp;</td><td>" & vbCrLf
                    
                    iMarkDecimals = 1
                    iTimeDecimals = 1
                    cMarkFormat = "0.0"
                    cTimeFormat = "0.00"
                    Select Case rstTest.Fields("Type_Pre")
                    Case Is <= 2 'marks or placemarks
                        With rstTest
                            If IsNull(.Fields("Mark_Decimals")) Then
                                iMarkDecimals = 1
                            Else
                                iMarkDecimals = .Fields("Mark_Decimals")
                            End If
                        End With
                    Case Is = 3  'time
                        With rstTest
                            If IsNull(.Fields("Time_Decimals")) Then
                                iTimeDecimals = 1
                            Else
                                iTimeDecimals = .Fields("Time_Decimals")
                            End If
                        End With
                    Case Else
                    End Select
                
                    'how to format marks
                    If ((iOldStatus = 0 And rstTest.Fields("Type_pre") = 2) Or (iOldStatus <> 0 And rstTest.Fields("Type_Final") = 2)) Then
                        cMarkFormat = "0"
                    Else
                        cMarkFormat = "0." & String$(iMarkDecimals, "0")
                    End If
                    cTimeFormat = "0." & String$(iTimeDecimals, "0")
                    If rstTest.Fields("Type_Special") = 3 Then 'gaedingakeppni
                        cTotalFormat = "0.000"
                    Else
                        cTotalFormat = "0.00"
                    End If
                
                    If rstMarks.Fields("Disq") = -2 Then
                            cHtml = cHtml & mcNoPosition & vbCrLf
                    ElseIf rstTest.Fields("Type_pre") = 3 Then
                        If rstMarks.Fields("Mark1") = 0 Then
                            cHtml = cHtml & mcNoPosition & vbCrLf
                        Else
                            cHtml = cHtml & Format$(rstMarks.Fields("Mark1"), cTimeFormat) & Chr$(34) & vbCrLf
                            cHtml = cHtml & " (= " & Format$(Time2Mark(rstMarks.Fields("Mark1"), cOldCode), cTotalFormat) & ")" & vbCrLf
                        End If
                    Else
                        For iTemp = 1 To rstMarks.Fields("num_j_" & Format$(rstMarks.Fields("Status")))
                            cHtml = cHtml & Format$(rstMarks.Fields("Mark" & iTemp), cMarkFormat) & vbCrLf
                            If iTemp < rstMarks.Fields("num_j_" & Format$(rstMarks.Fields("Status"))) Then
                                cHtml = cHtml & " - " & vbCrLf
                            End If
                        Next iTemp
                        cHtml = cHtml & " = " & Format$(rstMarks.Fields("Marks.Score"), cTotalFormat) & vbCrLf
                    End If
                    cHtml = cHtml & "&nbsp;</td><td>" & vbCrLf
                    If iOldSection = 1 Then
                        If rstMarks.Fields("Disq") = -1 Then
                            cHtml = cHtml & vbTab & Translate("Eliminated", mcLanguage) & vbCrLf
                        ElseIf rstMarks.Fields("Disq") = -2 Then
                            cHtml = cHtml & vbTab & Translate("Withdrawn", mcLanguage) & vbCrLf
                        Else
                            cHtml = cHtml & vbTab & Translate("Total", mcLanguage) & ": " & Format$(rstMarks.Fields("Results.Score"), cTotalFormat) & vbCrLf
                            If rstTest.Fields("Type_pre") = 3 Then
                                cHtml = cHtml & Chr$(34) & vbCrLf
                            End If
                        End If
                    End If
                    cHtml = cHtml & "</td></tr>" & vbCrLf
                End If
                rstMarks.MoveNext
            Loop
            cHtml = cHtml & "</table><br /><br />" & vbCrLf
            cHtml = cHtml & "<!-- CreateHTMLDetails end -->" & vbCrLf
        
        End If
        rstMarks.Close
        Set rstMarks = Nothing
        
        cModel = GetHTMLTemplate
        cModel = Replace(cModel, "{eventname}", frmMain.EventName)
        cModel = Replace(cModel, "{title}", Translate("Individual marks", mcLanguage))
        cModel = Replace(cModel, "{body}", cHtml)
        cFooter = "<footer>"
        cFooter = cFooter & Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss")
        cFooter = cFooter & "<BR>" & App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
        cFooter = cFooter & "</footer>"
        cModel = Replace(cModel, "{footer}", cFooter)
        
        'Write details to participant_[cSTA].html:
        iHtmlFileNum = FreeFile
        Open mcHtmlDir & cHtmlFileName For Output Access Write Shared As iHtmlFileNum
        Print #iHtmlFileNum, cModel
        Close #iHtmlFileNum
    End If
    
    'Return file name (without path) to calling code:
    CreateHTMLDetails = cHtmlFileName
    
    On Local Error GoTo 0
    
End Function
Function CreateHTMLFooter() As String
    Dim cHtml As String
    
    
    cHtml = vbCrLf & "<!-- CreateHTMLFooter begin -->" & vbCrLf
    cHtml = cHtml & "<footer>{footer}&nbsp;<br />"
    cHtml = cHtml & App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
    cHtml = cHtml & "</footer></p></html>" & vbCrLf
    cHtml = cHtml & "<!-- CreateHTMLFooter end -->" & vbCrLf
    
    CreateHTMLFooter = cHtml

End Function
' CopyHTML: Uses API32 function to copy all files from <sourcedir> to <targetdir>
Function CopyHTML(sourcedir As String, targetdir As String) As Integer
    Dim op As SHFILEOPSTRUCT
    Dim temptargetdir As String
    Dim tempsourcedir As String
    
    On Local Error Resume Next
    
    'remove trailing slash in targetdir:
    If Right$(targetdir, 1) = "\" Then
        temptargetdir = Left$(targetdir, Len(targetdir) - 1)
    Else
        temptargetdir = targetdir
    End If
    
    'add wildcard to sourcedir:
    If Right$(sourcedir, 1) = "\" Then
        tempsourcedir = sourcedir & "*.*"
    Else
        tempsourcedir = sourcedir
    End If

    With op
        .wFunc = FO_COPY ' Set function
        .pTo = temptargetdir ' Set new path
        .pFrom = tempsourcedir ' Set current path
        .fFlags = FOF_SILENT + FOF_FILESONLY + FOF_NOCONFIRMATION
    End With
    
    ' Perform operation
    SHFileOperation op
    
    CopyHTML = 1
    
    On Local Error GoTo 0
    
End Function
' Creates CSS stylesheet file in the one and only HTML folder
Sub CreateCSS()
    Dim iCssFileNum As Integer
    Dim cCSS As String
    
    On Local Error Resume Next
        
    cCSS = ".normal, body, p, a {" & vbCrLf
    cCSS = cCSS & "color: black; background-color: white;" & vbCrLf
    cCSS = cCSS & "font-size: 0,9em;" & vbCrLf
    cCSS = cCSS & "font-family: Tahoma, Verdana, sans-serif;" & vbCrLf
    cCSS = cCSS & "margin: 0; padding: 0.1em;" & vbCrLf
    cCSS = cCSS & "    min-width: 41em; /* Minimum width avoids line break and display bug */" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    
    cCSS = cCSS & "h1 {" & vbCrLf
    cCSS = cCSS & "    font-size: 1.3em;" & vbCrLf
    cCSS = cCSS & "    margin: 0 0 0.7em; padding: 0.3em;" & vbCrLf
    cCSS = cCSS & "    text-align: center;" & vbCrLf
    cCSS = cCSS & "    background-color: #fed;" & vbCrLf
    cCSS = cCSS & "    border: 2px ridge silver;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    
    cCSS = cCSS & "h2 {" & vbCrLf
    cCSS = cCSS & "    font-size: 1.3em;" & vbCrLf
    cCSS = cCSS & "    margin: 0 0 0.7em; padding: 0.3em;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    
    cCSS = cCSS & "h3 {" & vbCrLf
    cCSS = cCSS & "    font-size: 1.3em;" & vbCrLf
    cCSS = cCSS & "    margin: 0 0 0.7em; padding: 0.3em;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    
    cCSS = cCSS & "body, h1, h2, h3 {" & vbCrLf
    cCSS = cCSS & "border-color: gray;  /* Farbangleichung an den Internet Explorer  */" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf

    cCSS = cCSS & "ul#Navigation {" & vbCrLf
    cCSS = cCSS & "    font-size: 0.83em;" & vbCrLf
    cCSS = cCSS & "    float: left; width: 18em;" & vbCrLf
    cCSS = cCSS & "    margin: 0 0 1.2em; padding: 0;" & vbCrLf
    cCSS = cCSS & "    border: 1px dashed silver;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "ul#Navigation li {" & vbCrLf
    cCSS = cCSS & "list-style: none;" & vbCrLf
    cCSS = cCSS & "margin: 0; padding: 0.5em;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "ul#Navigation a {" & vbCrLf
    cCSS = cCSS & "    display: block;" & vbCrLf
    cCSS = cCSS & "    padding: 0.2em;" & vbCrLf
    cCSS = cCSS & "    font-weight: bold;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "ul#Navigation a:link {" & vbCrLf
    cCSS = cCSS & "    color: black; background-color: #eee;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "ul#Navigation a:visited {" & vbCrLf
    cCSS = cCSS & "    color: #666; background-color: #eee;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "ul#Navigation a:hover {" & vbCrLf
    cCSS = cCSS & "    color: black; background-color: white;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "ul#Navigation a:active {" & vbCrLf
    cCSS = cCSS & "    color: white; background-color: gray;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf

    cCSS = cCSS & "div#Info {" & vbCrLf
    cCSS = cCSS & "    font-size: 0.9em;" & vbCrLf
    cCSS = cCSS & "    float: right; width: 12em;" & vbCrLf
    cCSS = cCSS & "    margin: 0 0 1.1em; padding: 0;" & vbCrLf
    cCSS = cCSS & "    background-color: #eee; border: 1px dashed silver;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "div#Info h2 {" & vbCrLf
    cCSS = cCSS & "    font-size: 1.33em;" & vbCrLf
    cCSS = cCSS & "    margin: 0.2em 0.5em;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "div#Info p {" & vbCrLf
    cCSS = cCSS & "font-size: 1em;" & vbCrLf
    cCSS = cCSS & "margin: 0.5em;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf

    cCSS = cCSS & "div#Content {" & vbCrLf
    cCSS = cCSS & "    margin: 0 12em 1em 16em;" & vbCrLf
    cCSS = cCSS & "    padding: 0 1em;" & vbCrLf
    cCSS = cCSS & "    border: 1px dashed silver;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "* html div#Content {" & vbCrLf
    cCSS = cCSS & "    height: 1em;  /* Workaround gegen den 3-Pixel-Bug des Internet Explorers */" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "div#Content h2 {" & vbCrLf
    cCSS = cCSS & "    font-size: 1.2em;" & vbCrLf
    cCSS = cCSS & "    margin: 0.2em 0;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "div#Content p {" & vbCrLf
    cCSS = cCSS & "    font-size: 1em;" & vbCrLf
    cCSS = cCSS & "    margin: 1em 0;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "div#Content ul {" & vbCrLf
    cCSS = cCSS & "    font-size: 0.8em;" & vbCrLf
    cCSS = cCSS & "    margin: 0 0 1.1em; padding: 0;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf

    cCSS = cCSS & ".footer {" & vbCrLf
    cCSS = cCSS & "    font-size: 0.8em;" & vbCrLf
    cCSS = cCSS & "    margin: 0; padding: 0.1em;" & vbCrLf
    cCSS = cCSS & "    text-align: left;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    
    cCSS = cCSS & "table {" & vbCrLf
    cCSS = cCSS & "    margin: 1em;" & vbCrLf
    cCSS = cCSS & "    border-collapse: collapse;" & vbCrLf
    cCSS = cCSS & "    font-size: 0.9em;" & vbCrLf
    cCSS = cCSS & "    text-align: left;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "td, th {" & vbCrLf
    cCSS = cCSS & "    padding: .3em;" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "thead {" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    cCSS = cCSS & "tbody {" & vbCrLf
    cCSS = cCSS & "}" & vbCrLf
    
    iCssFileNum = FreeFile
    Open mcHtmlDir & "iceweb.css" For Output Access Write Shared As iCssFileNum
    Print #iCssFileNum, cCSS
    Close #iCssFileNum

End Sub
' Creates HTML files in the one and only HTML folder
Sub CreateHTML(cBodyCurrent As String, cBodyTest As String, Optional iEmpty As Integer = False)
    Dim iHtmlFileNum As Integer
    Dim iTemp As Integer
    Dim cTemp As String
    Dim cTemp2 As String
    Dim cTempList As String
    Dim cTempHtml As String
    Dim cTempTitle As String
    Dim iTempFileNum As Integer
    Dim cIndex As String
    Dim cHtml As String
    Dim iKey As Integer
    Dim cTest As String
    Dim cModel As String
    Dim rstTests As Recordset
    Dim cQry As String
    Dim cTestTitle As String
    Dim cSponsor As String
    Dim cFooter As String
    
    Dim iErrCounter As Integer
    
    On Local Error Resume Next
    
       
    If mcHtmlDir = "" Then Exit Sub
    
    If InStr(Command$, "/NOHTML") > 0 Then Exit Sub
    
    On Local Error Resume Next
    
    On Local Error GoTo CreateHtmlError
    
    cModel = GetHTMLTemplate
    
    cTestTitle = frmMain.TestCode & " " & Translate(frmMain.TestName, mcLanguage)
    cSponsor = GetSponsor(frmMain.dtaTest.Recordset.Fields("Code"))
    If frmMain.dtaTest.Recordset.Fields("Type_pre") <= 2 And Translate(frmMain.dtaTest.Recordset.Fields("Test"), mcLanguage) <> ClipAmp(frmMain.tbsSelFin.SelectedItem.Caption) Then
        cTestTitle = cTestTitle & " - " & ClipAmp(frmMain.tbsSelFin.SelectedItem.Caption)
    End If
    If cSponsor <> "" Then
        cTestTitle = cTestTitle & "<BR>" & Trim$(GetVariable("Sponsors") & " " & cSponsor)
    End If
    
    cHtml = cModel
    cHtml = Replace(cHtml, "{eventname}", frmMain.EventName)
    cHtml = Replace(cHtml, "{title}", cTestTitle)
    cHtml = Replace(cHtml, "{body}", cBodyCurrent)
    
    cFooter = "<footer>"
    cFooter = cFooter & Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss")
    cFooter = cFooter & "<BR>" & App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
    cFooter = cFooter & "</footer>"
    cHtml = Replace(cHtml, "{footer}", cFooter)
    
    iHtmlFileNum = FreeFile
    Open mcHtmlDir & "current.html" For Output Access Write Shared As iHtmlFileNum
    Print #iHtmlFileNum, cHtml
    Close #iHtmlFileNum
        
    If iEmpty = True Then
        KillFile mcHtmlDir & LCase$(UnDotSpace(frmMain!dtaTest.Recordset.Fields("Code")) & " " & frmMain!dtaTest.Recordset.Fields("Test") & IIf(frmMain!dtaTest.Recordset.Fields("Type_pre") <= 2, " - " & ClipAmp(frmMain!tbsSelFin.SelectedItem.Caption), "")) & ".html"
    Else
        cHtml = cModel
        cHtml = Replace(cHtml, "{eventname}", frmMain.EventName)
        cHtml = Replace(cHtml, "{title}", cTestTitle)
        cHtml = Replace(cHtml, "{body}", cBodyTest)
        cFooter = "<footer>"
        cFooter = cFooter & Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss")
        cFooter = cFooter & "<BR>" & App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
        cFooter = cFooter & "</footer>"
        cHtml = Replace(cHtml, "{footer}", cFooter)
        
        iHtmlFileNum = FreeFile
        Open mcHtmlDir & LCase$(UnDotSpace(frmMain!dtaTest.Recordset.Fields("Code")) & "_" & frmMain.TestStatus) & ".html" For Output Access Write Shared As iHtmlFileNum
        Print #iHtmlFileNum, cHtml
        Close #iHtmlFileNum
    End If
    
    'Generate index.html
    
    cIndex = "<ul>" & vbCrLf
    cIndex = cIndex & "<li><b><a href=""current.html"">" & Translate("Current Test", mcLanguage) & "</a></b></li><br>" & vbCrLf
    
    cQry = "SELECT DISTINCT Tests.Code, Tests.Test, Tests.Type_pre, Tests.Type_time, Results.Status, Testinfo.Nr"
    cQry = cQry & " FROM (Tests INNER JOIN TestInfo ON Tests.Code = TestInfo.Code) INNER JOIN Results ON Tests.Code = Results.Code"
    cQry = cQry & " ORDER BY TestInfo.Nr,Results.Status;"
    
    Set rstTests = mdbMain.OpenRecordset(cQry)
    If rstTests.RecordCount > 0 Then
        Do While Not rstTests.EOF
            cTemp = LCase$(rstTests.Fields("Code") & "_" & rstTests.Fields("Status")) & ".html"
            If Dir$(mcHtmlDir & cTemp) <> "" Then
                Select Case rstTests.Fields("Status")
                Case 1
                    cTempTitle = rstTests.Fields("Code") & " - " & Translate(rstTests.Fields("Test"), mcLanguage) & " - A-" & Translate("Final", mcLanguage)
                Case 2
                    cTempTitle = rstTests.Fields("Code") & " - " & Translate(rstTests.Fields("Test"), mcLanguage) & " - B-" & Translate("Final", mcLanguage)
                Case 3
                    cTempTitle = rstTests.Fields("Code") & " - " & Translate(rstTests.Fields("Test"), mcLanguage) & " - C-" & Translate("Final", mcLanguage)
                Case Else
                    If rstTests.Fields("Type_pre") <= 2 And rstTests.Fields("Type_time") = 0 Then
                        cTempTitle = rstTests.Fields("Code") & " - " & Translate(rstTests.Fields("Test"), mcLanguage) & " - " & Translate("Preliminary round", mcLanguage)
                    Else
                        cTempTitle = rstTests.Fields("Code") & " - " & Translate(rstTests.Fields("Test"), mcLanguage)
                    End If
                End Select
                cIndex = cIndex & "<li><b><a href=" & Chr$(34) & cTemp & Chr$(34) & ">" & cTempTitle & "</a></b></li>" & vbCrLf
            End If
            rstTests.MoveNext
        Loop
    End If
    rstTests.Close
    Set rstTests = Nothing
    
    cIndex = cIndex & "</ul>" & vbCrLf
    
    cHtml = cModel
    
    cHtml = Replace(cHtml, "{eventname}", frmMain.EventName)
    cHtml = Replace(cHtml, "{title}", Translate("Index", mcLanguage))
    cHtml = Replace(cHtml, "{body}", cIndex)
    cFooter = "<footer>"
    cFooter = cFooter & Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss")
    cFooter = cFooter & "<BR>" & App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
    cFooter = cFooter & "</footer>"
    cHtml = Replace(cHtml, "{footer}", cFooter)
    
    iHtmlFileNum = FreeFile
    Open mcHtmlDir & "index.html" For Output Access Write Shared As iHtmlFileNum
    Print #iHtmlFileNum, cHtml
    Close #iHtmlFileNum
    
    'CopyHTML mcTempHtmlDir, mcHtmlDir
    
    On Local Error GoTo 0
    
Exit Sub

CreateHtmlError:
    If iErrCounter < 5 Then
        iErrCounter = iErrCounter + 1
        Sleep 1
        Resume
    Else
        Exit Sub
    End If
Return

End Sub
Public Sub CreateHTMLTemplate()
    Dim cHtml As String
    Dim iHtmlFileNum As Integer
    
    On Local Error Resume Next
    
    'check if feif logo exists in temp html dir
    If Dir$(mcHtmlDir & "feif.gif") = "" Then
        ResourceToDisk 106, "CUSTOM", mcHtmlDir & "feif.gif"
    End If
    
    If Dir$(mcHtmlDir & "template.html") = "" Then
        'check if logo file exists in html dir
        If Dir$(mcHtmlDir & "feifsoft.gif") = "" Then
            ResourceToDisk 101, "CUSTOM", mcHtmlDir & "feifsoft.gif"
        End If
        
        'check if stylesheet file exists in html dir
        If Dir$(mcHtmlDir & "iceweb.css") = "" Then
            Call CreateCSS
        End If
        
        'check if icetest icon exists in temp html dir
        If Dir$(mcHtmlDir & "icetestweb.gif") = "" Then
            ResourceToDisk 105, "CUSTOM", mcHtmlDir & "icetestweb.gif"
        End If
        
        
        cHtml = cHtml & "<!-- When you adapt this template to your own taste                   -->" & vbCrLf
        cHtml = cHtml & "<!-- Be sure to preserve the labels eventname, title, body and footer -->" & vbCrLf
        cHtml = cHtml & "<!-- Between { and }                                                  -->" & vbCrLf
        cHtml = cHtml & "<!-- As IceTest requires them                                         -->" & vbCrLf
        cHtml = cHtml & CreateHTMLHeader("{title}")
        cHtml = cHtml & "<body>" & vbCrLf
        cHtml = cHtml & "<table><tr>" & vbCrLf
        cHtml = cHtml & "<td><a href=""https://www.feif.org"" target=""_blank"">" & vbCrLf
        cHtml = cHtml & "<img border=""0"" src=""feif.gif"" alt=""FEIF"" width=""80"" height=""80"" longdesc=""FEIF""></a></td>" & vbCrLf
        cHtml = cHtml & "<td><a href=""index.html"" style=""text-decoration:none""><h1>{eventname}</h1><a></td>" & vbCrLf
        cHtml = cHtml & "</tr></table>" & vbCrLf
        cHtml = cHtml & "<p>" & vbCrLf
        cHtml = cHtml & "<p align=""right"">" & vbCrLf
        cHtml = cHtml & "<a href=""javascript:history.back();""><b>" & Translate("back", mcLanguage) & "</b></a>" & vbCrLf
        cHtml = cHtml & "</p>" & vbCrLf
        cHtml = cHtml & "<p>" & vbCrLf
        cHtml = cHtml & "<h2>{title}</h2>&nbsp;" & vbCrLf
        cHtml = cHtml & "</p>" & vbCrLf
        cHtml = cHtml & "<p>" & vbCrLf
        cHtml = cHtml & "{body}&nbsp;" & vbCrLf
        cHtml = cHtml & "</p>" & vbCrLf
        cHtml = cHtml & "<p>" & vbCrLf
        cHtml = cHtml & "<table class=""footer"">" & vbCrLf
        cHtml = cHtml & "<tr><td>" & GetVariable("Vision") & "</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>&nbsp;</td></tr>" & vbCrLf
        cHtml = cHtml & "<tr><td>{footer}</td></tr>" & vbCrLf
        cHtml = cHtml & "</table>" & vbCrLf
        cHtml = cHtml & "</body></html>" & vbCrLf
        
        iHtmlFileNum = FreeFile
        Open mcHtmlDir & "template.html" For Output Access Write Shared As iHtmlFileNum
        Print #iHtmlFileNum, cHtml
        Close #iHtmlFileNum
    End If
    
    On Local Error GoTo 0
End Sub
Public Function GetHTMLTemplate(Optional cFilename As String = "")
    Dim cTemp As String
    Dim iHtmlFileNum As Integer
    
    On Local Error Resume Next
    
    CreateHTMLTemplate
    
    iHtmlFileNum = FreeFile
    If cFilename <> "" Then
        Open StrConv(cFilename, vbLowerCase) For Binary Access Read Shared As iHtmlFileNum
    Else
        Open StrConv(mcHtmlDir & "template.html", vbLowerCase) For Binary Access Read Shared As iHtmlFileNum
    End If
    cTemp = Space$(LOF(iHtmlFileNum))
    Get #iHtmlFileNum, , cTemp
    Close iHtmlFileNum
    
    GetHTMLTemplate = cTemp
    
    On Local Error GoTo 0
    
End Function
