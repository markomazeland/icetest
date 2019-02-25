Attribute VB_Name = "modIceWF"
Option Explicit
Option Compare Text

Public Function ValidHorseFEIFId(cFEIFID As String) As Integer
    Dim iTemp As Integer
    
    cFEIFID = cFEIFID & ""
    ValidHorseFEIFId = True
    If Len(cFEIFID) = 12 Then
        For iTemp = 1 To 2
            If Mid$(cFEIFID, iTemp, 1) Like "[!A-Z]" Then
                ValidHorseFEIFId = False
                Exit For
            End If
        Next iTemp
        If ValidHorseFEIFId = True Then
            For iTemp = 3 To 12
                If Mid$(cFEIFID, iTemp, 1) Like "[!0-9]" Then
                    ValidHorseFEIFId = False
                    Exit For
                End If
            Next iTemp
        End If
    Else
        ValidHorseFEIFId = False
    End If
End Function

Public Function XmlParse(cXML, cLabel, Optional cSubLabel As String = "", Optional iReverse As Integer = 0) As String
    Dim iTemp As Integer
    Dim cTemp As String
    
    cTemp = cXML
    XmlParse = ""
    If iReverse <> 0 Then
        iTemp = InStrRev(cTemp, "<" & cLabel & " ")
    Else
        iTemp = InStr(cTemp, "<" & cLabel & " ")
    End If
    If iTemp = 0 Then
        If iReverse <> 0 Then
            iTemp = InStrRev(cTemp, "<" & cLabel & ">")
        Else
            iTemp = InStr(cTemp, "<" & cLabel & ">")
        End If
    End If
    If iTemp > 0 Then
        cTemp = Mid$(cTemp, iTemp)
        iTemp = InStr(cTemp, "</" & cLabel & ">")
        If iTemp > 0 Then
            cTemp = Left$(cTemp, iTemp - 1)
            iTemp = InStr(cTemp, ">")
            If iTemp > 0 Then
                cTemp = Trim$(Mid$(cTemp & " ", iTemp + 1))
                If cSubLabel <> "" Then
                    cTemp = XmlParse(cTemp, cSubLabel)
                Else
                    cTemp = cTemp
                End If
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    XmlParse = cTemp
End Function

Public Function Slash2From(cId As String, Optional cOrigin = "") As String
    Dim cTemp As String
    
    Select Case Left$(cId, 2)
    Case "AT"
        If InStr(cOrigin, "von ") > 0 Or InStr(cOrigin, "vom ") > 0 Then
            cTemp = " "
        Else
          cTemp = " von "
        End If
    Case "CH"
        If InStr(cOrigin, "von ") > 0 Or InStr(cOrigin, "vom ") > 0 Then
            cTemp = " "
        Else
          cTemp = " von "
        End If
    Case "DE"
        If InStr(cOrigin, "von ") > 0 Or InStr(cOrigin, "vom ") > 0 Then
            cTemp = " "
        Else
          cTemp = " von "
        End If
    Case "DK"
        cTemp = " fra "
    Case "FI"
        cTemp = " frá "
    Case "FR"
        cTemp = " de "
    Case "IS"
        cTemp = " frá "
    Case "NL"
        cTemp = " van "
    Case "NO"
        cTemp = " fra "
    Case "SE"
        cTemp = " från "
    Case Else
        cTemp = " / "
    End Select
    Slash2From = UTF8_Encode(cTemp)
End Function


Public Function UTF8_Decode(ByVal Text As String) As String
    ' thanks to fredlynx
    
    Dim cTemp As String
    Dim lLength As Long
    Dim sBuffer As String

    Text = StrConv(Text, vbFromUnicode)
    lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(Text), -1, 0, 0)
    sBuffer = Space$(lLength)
    lLength = MultiByteToWideChar(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer))
    cTemp = Left$(sBuffer, lLength - 1)
    cTemp = Replace(cTemp, "&amp;", "&")
    cTemp = Replace(cTemp, "&apos;", "'")
    cTemp = Replace(cTemp, "&aelig;", "æ")
    UTF8_Decode = cTemp
End Function

Public Function UTF8_Encode(ByVal Text As String) As String
    ' thanks to fredlynx

    Dim sBuffer As String
    Dim lLength As Long
    
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, 0, 0, 0, 0)
    sBuffer = Space$(lLength)
    lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(Text), -1, StrPtr(sBuffer), Len(sBuffer), 0, 0)
    sBuffer = StrConv(sBuffer, vbUnicode)
    
    UTF8_Encode = Left$(sBuffer, lLength - 1)

End Function

