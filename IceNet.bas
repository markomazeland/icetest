Attribute VB_Name = "modIceNet"
' Functions related to Internet / Winsock

' Copyright (C) Lutz Lesener 2006
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

Type URL
    Scheme As String
    Host As String
    Port As Long
    URI As String
    Query As String
End Type

' returns as type URL from a string
Function ExtractUrl(ByVal strUrl As String) As URL
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    Dim retURL As URL
    
    '1 look for a scheme it ends with ://
    intPos1 = InStr(strUrl, "://")
    
    If intPos1 > 0 Then
        retURL.Scheme = Mid(strUrl, 1, intPos1 - 1)
        strUrl = Mid(strUrl, intPos1 + 3)
    End If
        
    retURL.Host = strUrl
    retURL.URI = "/"
    
    ExtractUrl = retURL
End Function
' converts all line endings to Windows CrLf line endings
Function FormatLineEndings(ByVal str As String) As String
    Dim prevChar As String
    Dim nextChar As String
    Dim curChar As String
    
    Dim strRet As String
    
    Dim X As Long
    
    prevChar = ""
    nextChar = ""
    curChar = ""
    strRet = ""
    
    For X = 1 To Len(str)
        prevChar = curChar
        curChar = Mid$(str, X, 1)
                
        If nextChar <> vbNullString And curChar <> nextChar Then
            curChar = curChar & nextChar
            nextChar = ""
        ElseIf curChar = vbLf Then
            If prevChar <> vbCr Then
                curChar = vbCrLf
            End If
            
            nextChar = ""
        ElseIf curChar = vbCr Then
            nextChar = vbLf
        End If
        
        strRet = strRet & curChar
    Next X
    
    FormatLineEndings = strRet
End Function

' url encodes a string
Function URLEncode(ByVal str As String) As String
        Dim intLen As Integer
        Dim X As Integer
        Dim curChar As Long
                Dim newStr As String
                intLen = Len(str)
        newStr = ""
                        For X = 1 To intLen
            curChar = Asc(Mid$(str, X, 1))
            
            If (curChar < 48 Or curChar > 57) And _
                (curChar < 65 Or curChar > 90) And _
                (curChar < 97 Or curChar > 122) Then
                                newStr = newStr & "%" & Hex(curChar)
            Else
                newStr = newStr & Chr(curChar)
            End If
        Next X
        
        URLEncode = newStr
End Function

' decodes a url encoded string
Function UrlDecode(ByVal str As String) As String
        Dim intLen As Integer
        Dim X As Integer
        Dim curChar As String * 1
        Dim strCode As String * 2
        
        Dim newStr As String
        
        intLen = Len(str)
        newStr = ""
        
        For X = 1 To intLen
            curChar = Mid$(str, X, 1)
            
            If curChar = "%" Then
                strCode = "&h" & Mid$(str, X + 1, 2)
                
                If IsNumeric(strCode) Then
                    curChar = Chr(Int(strCode))
                Else
                    curChar = ""
                End If
                                X = X + 2
            End If
            
            newStr = newStr & curChar
        Next X
        
        UrlDecode = newStr
End Function



