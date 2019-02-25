Attribute VB_Name = "modIcePerson"
' Functions related to persons

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
Public Function AddNewPerson() As String
    Dim iTemp As Integer
    Dim cTemp As String
    Dim cFirst As String
    Dim cLast As String
    Dim cId As String
    Dim iCounter As Integer
    Dim rstPersons As DAO.Recordset
    
    cTemp = InputBox$(Translate("Enter the complete name of the person (first and last name).", mcLanguage))
    If cTemp <> "" And cTemp <> Chr$(27) Then
        Set rstPersons = mdbMain.OpenRecordset("SELECT * FROM Persons WHERE Name_First & ' ' & Name_Last LIKE " & Chr$(34) & cTemp & Chr$(34))
        If rstPersons.RecordCount > 0 Then
            MsgBox cTemp & ": " & Translate("Name already exists!", mcLanguage)
            AddNewPerson = rstPersons.Fields("PersonId")
        Else
            iTemp = InStr(cTemp, " ")
            If iTemp = 0 Then
                iTemp = InStrRev(cTemp, ".")
            End If
            If iTemp > 0 Then
                cFirst = Trim$(Left$(cTemp, iTemp))
                cLast = Trim$(Mid$(cTemp, iTemp))
            Else
                cFirst = cTemp
            End If
            cId = CreatePersonId
            
            With rstPersons
                .AddNew
                .Fields("PersonId") = cId
                .Fields("Name_First") = cFirst
                .Fields("Name_Last") = cLast
                .Update
            End With
            DoEvents
            AddNewPerson = cId
        End If
        rstPersons.Close
        Set rstPersons = Nothing
    End If

End Function
Function CreatePersonId() As String
    Dim rstId As Recordset
    Dim iCounter As Integer
    Dim cId As String
    Static iPrevCounter
    
    iCounter = iPrevCounter
    
    Do
        iCounter = iCounter + 1
        cId = "XX" & Format$(Now, "YYMMDD") & Format$(iCounter, "0000")
        Set rstId = mdbMain.OpenRecordset("SELECT PersonId from Persons WHERE PersonId LIKE '" & cId & "'")
    Loop While rstId.RecordCount > 0
    rstId.Close
    iPrevCounter = iCounter
    Set rstId = Nothing
    CreatePersonId = cId
End Function

