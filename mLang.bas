Attribute VB_Name = "mLanguages"
' VB Module mLang.bas
' Providing multi-language capabilities
' Copyright (C) Lutz Lesener 2002-2011
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
'
' Usage:
' ======
' - Add mLang.bas (this file) and cDB.cls to your project.
' - languages.mdb now resides in <CSIDL_APPDATA>\IceHorse instead of App.Path (unless you modify InitializeLanguageDB)
' - A reference to the MS ActiveX Data Objects Library is required (I used ADO 2.7)
' - Call InitializeLanguageDB once at startup
' - The TargetLanguage parameter must match the name of a column in table "Translation"
' - New SourceTexts will be added on the fly to the database
' - Example: x$ = Call Translate("Text to be translated", "German")


Option Explicit
Option Compare Text

Private LanguageDB As cDB
Private LanguageTexts() As String
'Returns the highest used StringID between 900000 and 1000000 from the Translation table.
' Used to ensure compatability with Heino Knief's translation code.
Public Function GetHighestStringID() As Long
    Dim mysql As String
    Dim zRs As ADODB.Recordset
    Dim z As Variant
    
    On Local Error GoTo GetHighestStringIDErr

    mysql = "SELECT TOP 1 StringID From [Translation] WHERE (StringID > 899999) AND (StringID < 1000000) ORDER BY StringID DESC;"

    If LanguageDB.CreateRS(zRs, mysql) Then
        'Check if recordset is empty
        If zRs.EOF And zRs.BOF Then
            GetHighestStringID = 900000
        Else
            z = zRs.Fields("StringId")
            GetHighestStringID = CLng(z)
        End If
        zRs.Close
        Set zRs = Nothing
    End If

GetHighestStringIDClose:
    On Local Error GoTo 0
    Exit Function
    
GetHighestStringIDErr:
    LogLine "Error in mLanguage: GetHighestStringID"
    On Local Error GoTo 0

End Function
'Looks up SourceText in the database and returns the equivalent in TargetLanguage
Public Function Translate(SourceText As String, TargetLanguage As String) As String
    Dim mysql As String
    Dim oRs As ADODB.Recordset
    Dim cTemp As String
    
    Dim X As Boolean
    
    Translate = SourceText
    If TargetLanguage <> "English" And Len(SourceText) > 1 Then
        'check if text is already available in array
        cTemp = FindTextFromArray(SourceText)
        'íf not, check if it is in  database
        If cTemp = "" Then
            mysql = "SELECT " & TargetLanguage & ", StringID FROM [Translation] WHERE English = " & Chr$(34) & SourceText & Chr$(34) & ";"
            If LanguageDB.CreateRS(oRs, mysql) Then
                If oRs.EOF And oRs.BOF Then
                    'SourceText was not found in database
                    AddSourceText (SourceText)
                    Translate = SourceText
                    oRs.Close
                    Set oRs = Nothing
                    Exit Function
                End If
            
                If IsNull(oRs.Fields(0)) Then
                    'No translation for the required TargetLanguage available
                    Translate = SourceText
                Else
                    'Translation is available
                    Translate = oRs.Fields(0)
                End If
                
                oRs.Close
                Set oRs = Nothing
            Else
                'no - SourceText not found in database
                Translate = SourceText
            End If
        Else
            Translate = cTemp
        End If
    End If
End Function
'Looks up SourceText in the database and returns the English equivalent
'Used to get original terms back when unified strings are needed.
Public Function TranslateBack(SourceText As String, TargetLanguage As String) As String
    Dim mysql As String
    Dim oRs As ADODB.Recordset
    Dim cTemp As String
    Dim posi As Integer, datepart As String
    Dim X As Boolean
    
    'If [] is used in SourceText, remove it and tag it on later again:
    posi = InStr(SourceText, " [")
    If posi > 0 Then
        datepart = Right(SourceText, Len(SourceText) - posi)
        SourceText = Left(SourceText, posi - 1)
    End If
    
    TranslateBack = SourceText
    If TargetLanguage <> "English" And Len(SourceText) > 1 Then
        'check if text is already available in array
        cTemp = ""
        'íf not, check if it is in  database
        If cTemp = "" Then
            mysql = "SELECT English, StringID FROM [Translation] WHERE " & TargetLanguage & " = " & Chr$(34) & SourceText & Chr$(34) & ";"
            If LanguageDB.CreateRS(oRs, mysql) Then
                If oRs.EOF And oRs.BOF Then
                    'SourceText was not found in database
                    TranslateBack = SourceText
                    oRs.Close
                    Set oRs = Nothing
                    Exit Function
                End If
            
                If IsNull(oRs.Fields(0)) Then
                    'No translation for the required TargetLanguage available
                    TranslateBack = SourceText
                Else
                    'Translation is available
                    TranslateBack = oRs.Fields(0)
                End If
                
                oRs.Close
                Set oRs = Nothing
            Else
                'no - SourceText not found in database
                TranslateBack = SourceText
            End If
        Else
            TranslateBack = cTemp
        End If
    End If
    
    If datepart <> "" Then
        TranslateBack = TranslateBack & " " & datepart
    End If
End Function
' Connect to the translation database
Public Function InitializeLanguageDB(Optional cLanguage As String, Optional cPath As String)

   If cPath = "" Then
      cPath = GetSpecialFolderLocation(CSIDL_APPDATA) & "\IceHorse\"
   End If
   If Right$(cPath, 1) <> "\" Then
      cPath = cPath & "\"
   End If
   If Dir$(cPath & "languages.mdb") = "" And Dir$(App.Path & "\languages.mdb") <> "" Then
        FileCopy App.Path & "\languages.mdb", cPath & "languages.mdb"
   End If
        
   Set LanguageDB = New cDB
   
   LanguageDB.InitDB cPath & "languages.mdb", , , , ejvJet4
    
    'Debug.Print "Using " & cPath & "languages.mdb"
    
   LoadLanguageTexts cLanguage

End Function
'Add a String that has not yet occured to the database, assuming it's in English.
Public Function AddSourceText(SourceText As String) As Boolean
    Dim mysql As String
    Dim HighID As Long
    
    HighID = GetHighestStringID + 1
    
    mysql = "INSERT INTO [Translation] (English, StringID) VALUES (" & Chr$(34) & Left$(SourceText, 255) & Chr$(34) & ", " & HighID & ");"
    LanguageDB.ExecuteSQL (mysql)
    
End Function
'Load texts in array to speed up IceTest on a slow machine/network
Public Sub LoadLanguageTexts(TargetLanguage As String)
   Dim mysql As String
   Dim myFld As ADODB.Field
   Dim oRs As ADODB.Recordset
   Dim cTemp As String
   
   mysql = "SELECT English," & TargetLanguage & " FROM [Translation] ORDER BY English"
   If LanguageDB.CreateRS(oRs, mysql) Then
      oRs.MoveLast
      oRs.MoveFirst
      ReDim LanguageTexts(0 To oRs.RecordCount, 1)
      With oRs
            Do While Not .EOF
                LanguageTexts(.AbsolutePosition, 0) = .Fields("English") & ""
                LanguageTexts(.AbsolutePosition, 1) = .Fields(TargetLanguage) & ""
                .MoveNext
            Loop
            .Close
      End With
   End If
   Set oRs = Nothing
    
End Sub
'Find text in array using binary search
Public Function FindTextFromArray(strText As String) As String
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim Lo As Integer
    Dim Hi As Integer
        
    On Local Error GoTo FindTextFromArrayError
        
    FindTextFromArray = ""
    If strText = "" Then Exit Function
    Lo = -1
    Hi = UBound(LanguageTexts) + 1
    
    Do While Hi > Lo
        iTemp2 = iTemp
        iTemp = (Hi + Lo) \ 2
        If iTemp = iTemp2 Then
            If Lo + 1 = Hi And Hi <= UBound(LanguageTexts) Then
                If strText = LanguageTexts(Hi, 0) Then
                       If LanguageTexts(iTemp, 1) <> "" Then
                           FindTextFromArray = LanguageTexts(iTemp, 1)
                       End If
                ElseIf strText = LanguageTexts(Lo, 0) Then
                       If LanguageTexts(iTemp, 1) <> "" Then
                           FindTextFromArray = LanguageTexts(iTemp, 1)
                       End If
                End If
            End If
            Exit Do
        End If
        If strText = LanguageTexts(iTemp, 0) Then
            If LanguageTexts(iTemp, 1) <> "" Then
               FindTextFromArray = LanguageTexts(iTemp, 1)
            End If
            Exit Do
        ElseIf strText < LanguageTexts(iTemp, 0) Then
            If Hi = iTemp Then
                Exit Do
            End If
            Hi = iTemp
        Else
            If Lo = iTemp Then
                Exit Do
            End If
            Lo = iTemp
        End If
    Loop
    Exit Function
    
FindTextFromArrayError:
    'an error ocurred: return an empty string to allow proper handling in calling procedure:
    LogLine "Error in mLanguages.FindTextFromArray when translating '" & strText & "': " & Err.Number & " " & Err.Description
    FindTextFromArray = ""
    
    On Local Error GoTo 0

End Function

