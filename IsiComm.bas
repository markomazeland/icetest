Attribute VB_Name = "modIsiComm"
Option Explicit
Option Compare Text

Public mcDatabaseDir As String
Public mcDatabaseName As String

Public mFormsCollection As New Collection
Public mdbMain As Database

Public Sub OpenDatabase(DatabaseName As String)
    If Dir$(mcDatabaseName) = "" Then
        CreateMdb mcDatabaseName
    End If
    
    Set mdbMain = DBEngine.OpenDatabase(DatabaseName, False, False)
End Sub
Public Function CreateTempId(Table As String, IdName As String, Optional Trunc As String = "") As String
    Dim lTeller As Long
    Dim cTemp As String
    Dim rst As Recordset
    
    If Trunc = "" Then
        Trunc = "XX" & Format$(Now, "YYMMDD")
    End If
    
    Do
        lTeller = lTeller + 1
        cTemp = Trunc & Format$(lTeller, String$(12 - Len(Trunc), "0"))
        Set rst = mdbMain.OpenRecordset("SELECT * FROM " & Table & " WHERE " & IdName & "='" & cTemp & "'")
    Loop While rst.RecordCount > 0
    rst.Close
    CreateTempId = cTemp
    
End Function

Public Function FormIsThere(CollectionTag As String) As Integer
    Dim lTemp As Long
    
    FormIsThere = False
    If mFormsCollection.Count > 0 Then
        For lTemp = 1 To mFormsCollection.Count
            If mFormsCollection(lTemp).CollectionTag = CollectionTag Then
                If mFormsCollection(lTemp).WindowState = 1 Then
                    ReadFormPosition mFormsCollection(lTemp), CollectionTag
                End If
                mFormsCollection(lTemp).SetFocus
                FormIsThere = True
                Exit For
            End If
        Next lTemp
    End If
End Function
