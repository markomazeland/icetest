Attribute VB_Name = "modIceMdb"
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

Public Sub OpenDatabase(DatabaseName As String)
    On Local Error GoTo OpenDatabaseError
    If Dir$(mcDatabaseName) = "" Then
        If Dir$(Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")), vbDirectory) = "" Then
            MkDir Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\"))
        End If
        If Dir$(Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Rtf\", vbDirectory) = "" Then
            MkDir Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Rtf"
        End If
        If Dir$(Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Html\", vbDirectory) = "" Then
            MkDir Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Html"
        End If
        
        If Dir$(mcDatabaseName) = "" Then
            ResourceToDisk 103, "CUSTOM", DatabaseName
        Else
            ResourceToDisk 103, "CUSTOM", Left$(DatabaseName, InStrRev(DatabaseName, "\")) & "Temp.Mdb"
            CopyTableBetweenDatabases Left$(DatabaseName, InStrRev(DatabaseName, "\")) & "Temp.Mdb", DatabaseName, "Fields", "Table", 0
            Kill Left$(DatabaseName, InStrRev(DatabaseName, "\")) & "Temp.Mdb"
        End If
        If Err > 0 Then
            LogLine "Failed to create/update " & mcDatabaseName & " " & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
        Else
            LogLine "Created " & mcDatabaseName
        End If
    End If
    Set mdbMain = DBEngine.OpenDatabase(DatabaseName, False, False)
OpenDatabaseError:
    If Err > 0 And Err <> 3265 Then
        LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
    End If
    On Local Error GoTo 0
    
End Sub
Function FindFieldName(r As Recordset, strTest As String) As String
    Dim fldTemp As Field
    If strTest <> "" Then
        For Each fldTemp In r.Fields
            If Right$(fldTemp.Name, Len(strTest)) = strTest Then
                FindFieldName = fldTemp.Name
                Exit For
            End If
        Next
    End If
    Set fldTemp = Nothing
End Function
Function FieldExist(t As TableDef, strField As String) As Integer
    Dim strTemp As String
    
    FieldExist = False
    On Local Error GoTo Einde
    DoEvents
    t.Fields.Refresh
    DoEvents
    If t.Fields.Count > 0 Then
        On Local Error Resume Next
        strTemp = t.Fields(strField).Name
        If Err <> 3265 Then
            Err = 0
            FieldExist = True
        End If
    End If
Einde:
    If Err > 0 And Err <> 3265 Then
        LogLine strField & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
    End If
    On Local Error GoTo 0
End Function
Function IndexExist(t As TableDef, strIndex As String) As Integer
    Dim strTemp As String
    
    IndexExist = False
    On Local Error GoTo Einde
    DoEvents
    t.Indexes.Refresh
    DoEvents
    If t.Indexes.Count > 0 Then
        On Local Error Resume Next
        strTemp = t.Indexes(strIndex).Name
        If Err <> 3265 Then
            Err = 0
            IndexExist = True
        End If
    End If
Einde:
    If Err > 0 And Err <> 3265 Then
        LogLine strIndex & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
    End If
    On Local Error GoTo 0
End Function
Function TableExist(d As DAO.Database, strTable As String) As Integer
    Dim strTemp As String
    
    On Local Error GoTo Einde
    
    TableExist = False
    
    DoEvents
    
    d.TableDefs.Refresh
    
    DoEvents
    
    If d.TableDefs.Count > 0 Then
        On Local Error Resume Next
        strTemp = d.TableDefs(strTable).Name
        If Err <> 3265 Then
            Err = 0
            TableExist = True
        End If
    End If

Einde:
    If Err > 0 And Err <> 3265 Then
        LogLine strTable & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
    End If
    On Local Error GoTo 0
End Function
Function QueryExist(d As DAO.Database, strQuery As String) As Integer
    Dim strTemp As String
    
    QueryExist = False
    On Local Error GoTo Einde
    DoEvents
    d.QueryDefs.Refresh
    DoEvents
    If d.QueryDefs.Count > 0 Then
        On Local Error Resume Next
        strTemp = d.QueryDefs(strQuery).Name
        If Err <> 3265 Then
            Err = 0
            QueryExist = True
        End If
    End If
Einde:
    If Err > 0 And Err <> 3265 Then
        LogLine strQuery & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
    End If
    On Local Error GoTo 0
End Function
Sub AppendDeleteField(tdfTemp As TableDef, strCommand As String, strName As String, Optional varType, Optional varSize)
    On Local Error GoTo Einde
    
    With tdfTemp

        ' Check first to see if the TableDef object is
        ' updatable. If it isn't, control is passed back to
        ' the calling procedure.
        If .Updatable = False Then
            MsgBox "Table " & .Name & " cannot be updated."
            Exit Sub
        End If
        
        ' Depending on the passed data, append or delete a
        ' field to the Fields collection of the specified
        ' TableDef object.
        If strCommand = "APPEND" Then
            If FieldExist(tdfTemp, strName) = False Then
                .Fields.Append .CreateField(strName, varType, varSize)
                LogLine "Field " & strName & " added to " & .Name
            End If
            If .Fields(strName).Type = dbText Or .Fields(strName).Type = dbMemo Then
                If .Fields(strName).AllowZeroLength = False Then
                    .Fields(strName).AllowZeroLength = True
                End If
            End If
        Else
            If strCommand = "DELETE" Then
                .Fields.Delete strName
                LogLine "Field " & strName & " deleted in " & .Name
            End If
        End If

    End With
    
Einde:
    If Err <> 0 Then
        LogLine "Field " & strName & " " & strCommand & " not executed in " & tdfTemp.Name & ": " & Err.Number & " " & Err.Description
    End If

End Sub
Public Sub CreateMdb(strMdbName As String, Optional varLanguage As Variant = dbLangGeneral)
    Dim w As DAO.Workspace
    Dim m As DAO.Database
    
    KillFile strMdbName

    Set w = DBEngine.Workspaces(0)
    Set m = w.CreateDatabase(strMdbName, varLanguage)
    m.Close
    
    If Err > 0 Then
        LogLine "Failed to create " & mcDatabaseName & " " & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
        Err = 0
    Else
        LogLine "Created " & mcDatabaseName
    End If
    
    Set m = Nothing
    Set w = Nothing
    
End Sub
Sub AlterFieldSize(d As DAO.Database, strTableName As String, strFieldName As String, intNewSize As Integer)
    Dim qdfTemp As QueryDef
    Dim tdfTemp As TableDef
    Dim idxTemp As Index
    
    Set tdfTemp = d.TableDefs(strTableName)
    If FieldExist(tdfTemp, strFieldName) = False Then
        AppendDeleteField tdfTemp, "APPEND", strFieldName, dbText, intNewSize
    ElseIf tdfTemp.Fields(strFieldName).Size < intNewSize Then
        With tdfTemp
            ' Add a temporary field to the table.
            If FieldExist(d.TableDefs(strTableName), "AlterTempField") = True Then
                .Fields.Delete "AlterTempField"
            End If
            d.TableDefs.Refresh
            .Indexes.Refresh
            DoEvents
            
            'Wis indexen
            Do While .Indexes.Count > 0
                For Each idxTemp In .Indexes
                    .Indexes.Delete idxTemp.Name
                    .Indexes.Refresh
                Next
            Loop
            d.TableDefs.Refresh
            DoEvents
            
        End With
                
        AppendDeleteField tdfTemp, "APPEND", "AlterTempField", dbText, intNewSize
        
        Set tdfTemp = d.TableDefs(strTableName)
        ' Create a dummy QueryDef object.
        Set qdfTemp = d.CreateQueryDef("", "Select * FROM [Fields]")
        
        ' Copy the data from old field into the new field.
        qdfTemp.SQL = "UPDATE DISTINCTROW [" & strTableName & "] SET AlterTempField = [" & strFieldName & "]"
        qdfTemp.Execute
        
        With tdfTemp
            ' Delete the old field.
            .Fields.Delete strFieldName
            
            ' Rename the temporary field to the old field's name.
            d.TableDefs.Refresh
            .Fields("AlterTempField").Name = strFieldName
            d.TableDefs.Refresh
        End With
        LogLine strTableName & "." & strFieldName & " size: " & Format$(intNewSize)
    End If
    
    Set qdfTemp = Nothing
    Set tdfTemp = Nothing
    Set idxTemp = Nothing
End Sub
Sub CheckField(d As DAO.Database, strTableName As String, strFieldName As String, varType, varSize)
    Dim qdfTemp As QueryDef
    Dim tdfTemp As TableDef
    Dim idxTemp As Index
    
    Set tdfTemp = d.TableDefs(strTableName)
    
    If FieldExist(tdfTemp, strFieldName) = False Then
        AppendDeleteField tdfTemp, "APPEND", strFieldName, varType, varSize
    ElseIf tdfTemp.Fields(strFieldName).Type <> varType Then
        With tdfTemp
            ' Add a temporary field to the table.
            If FieldExist(d.TableDefs(strTableName), "AlterTempField") = True Then
                .Fields.Delete "AlterTempField"
            End If
            d.TableDefs.Refresh
            .Indexes.Refresh
            DoEvents
            
            'Delete indexes
            Do While .Indexes.Count > 0
                For Each idxTemp In .Indexes
                    .Indexes.Delete idxTemp.Name
                    .Indexes.Refresh
                Next
            Loop
            d.TableDefs.Refresh
            DoEvents
            
        End With
                
        AppendDeleteField tdfTemp, "APPEND", "AlterTempField", varType, varSize
        
        Set tdfTemp = d.TableDefs(strTableName)
        ' Create a dummy QueryDef object.
        Set qdfTemp = d.CreateQueryDef("", "Select * FROM [Fields]")
        
        ' Copy the data from old field into the new field.
        qdfTemp.SQL = "UPDATE DISTINCTROW [" & strTableName & "] SET AlterTempField = [" & strFieldName & "]"
        qdfTemp.Execute
        
        With tdfTemp
            ' Delete the old field.
            .Fields.Delete strFieldName
                        
            ' Rename the temporary field to the old field's name.
            d.TableDefs.Refresh
            .Fields("AlterTempField").Name = strFieldName
            d.TableDefs.Refresh
        End With
        LogLine strTableName & "." & strFieldName & " type: " & Format$(varType)
    End If
    
    Set qdfTemp = Nothing
    Set tdfTemp = Nothing
    Set idxTemp = Nothing
End Sub
Public Sub CreateIndex(t As TableDef, strField As String, Optional strIndexName As String)
    Dim idxTemp As Index
    Dim fldTemp As Field
    
    On Local Error GoTo Einde
    
    If strIndexName = "" Then strIndexName = strField
    If IndexExist(t, strIndexName) = False Then
        Set idxTemp = t.CreateIndex(strIndexName)
        Set fldTemp = idxTemp.CreateField(strField)
        idxTemp.Fields.Append fldTemp
        idxTemp.Unique = False
        idxTemp.Primary = False
        t.Indexes.Append idxTemp
        t.Indexes.Refresh
        LogLine "Index " & strIndexName & " added to " & t.Name
        Set idxTemp = Nothing
        Set fldTemp = Nothing
    End If
Einde:
    If Err > 0 Then
        LogLine "Index " & strIndexName & " cannot be added: " & Err.Number & " " & Err.Description
        Err = 0
    End If
    
End Sub
Public Sub CreateTable(d As DAO.Database, strTable As String, strField As String, varFieldType As Variant, intFieldSize As Integer)
    Dim tdfTemp As TableDef
    
    On Local Error GoTo Einde
    
    If TableExist(d, strTable) = False Then
        Set tdfTemp = d.CreateTableDef(strTable)
        tdfTemp.Fields.Append tdfTemp.CreateField(strField, varFieldType, intFieldSize)
        tdfTemp.Fields(strField).AllowZeroLength = True
        d.TableDefs.Append tdfTemp
        d.TableDefs.Refresh
        Set tdfTemp = Nothing
        LogLine "Table " & strTable & " added to " & d.Name
    End If
Einde:
    If Err > 0 Then
        LogLine "Table " & strTable & " cannot be added to " & d.Name & " (" & Err.Number & " " & Err.Description & ")"
        Err = 0
    End If
Exit Sub

    
End Sub
Public Sub AllowZeroLength(t As TableDef)
    Dim fldTemp As Field
    
    On Local Error Resume Next
    For Each fldTemp In t.Fields
        With fldTemp
            If .Type = dbText Or .Type = dbMemo Then
                .AllowZeroLength = True
            End If
        End With
    Next
    Set fldTemp = Nothing
End Sub
Public Function CopyTableBetweenDatabases(strDatabaseFrom As String, strDatabaseTo As String, strTableName As String, strMatchingField As String, intEmptyOldTable As Integer) As Integer
    Dim mdbFrom As DAO.Database
    Dim mdbTo As DAO.Database
    Dim rstFrom As DAO.Recordset
    Dim tdfTo As DAO.TableDef
    Dim fldFrom As DAO.Field
    Dim strTempList As String
    
    On Local Error Resume Next

    If Dir$(strDatabaseFrom) = "" Then
        'database doesn't exist
        CopyTableBetweenDatabases = 1
        Exit Function
    ElseIf Dir$(strDatabaseTo) = "" Then
        'create missing database
        CreateDatabase strDatabaseTo, dbLangGeneral
    End If
    
    Set mdbFrom = DBEngine.OpenDatabase(strDatabaseFrom, False, True)
    Set mdbTo = DBEngine.OpenDatabase(strDatabaseTo, False, False)
    
    DoEvents
    
    If TableExist(mdbFrom, strTableName) = False Then
        'table doesn't exist
        CopyTableBetweenDatabases = 2
        Exit Function
    ElseIf TableExist(mdbTo, strTableName) = False Then
        'just copy complete table
        mdbFrom.Execute ("SELECT * INTO [" & strTableName & "] IN '' [;database=" & strDatabaseTo & ";] FROM [" & strTableName & "]")
    Else
        Set rstFrom = mdbFrom.OpenRecordset("SELECT * FROM [" & strTableName & "]")
        
        'check if all fields exists
        Set tdfTo = mdbTo.TableDefs(strTableName)
        For Each fldFrom In rstFrom.Fields
            If FieldExist(tdfTo, fldFrom.Name) = False Then
                AppendDeleteField tdfTo, "APPEND", fldFrom.Name, fldFrom.Type, fldFrom.Size
            End If
        Next
                
        If intEmptyOldTable = True Then
            mdbTo.Execute ("DELETE * FROM [" & strTableName & "]")
        Else
            'check about matching field
            If FieldExist(tdfTo, strMatchingField) = False Then
                CopyTableBetweenDatabases = 3 'no matching fields available
                Exit Function
            End If
            
            'delete all records with the same matching fields
            Do While Not rstFrom.EOF
                If strTempList = "" Then
                    strTempList = "'" & rstFrom.Fields(strMatchingField) & "'"
                ElseIf InStr(strTempList, "'" & rstFrom.Fields(strMatchingField) & "'") = 0 Then
                    strTempList = strTempList & ",'" & rstFrom.Fields(strMatchingField) & "'"
                End If
                rstFrom.MoveNext
            Loop
            If strTempList <> "" Then
                mdbTo.Execute ("DELETE * FROM [" & strTableName & "] WHERE [" & strMatchingField & "] IN (" & strTempList & ")")
            Else
                CopyTableBetweenDatabases = 4 'no values in matching fields
                mdbTo.Close
                mdbFrom.Close
                Set mdbTo = Nothing
                Set mdbFrom = Nothing
                Exit Function
            End If
        End If
        
        rstFrom.Close
        
        'copy all records
        mdbFrom.Execute ("INSERT INTO [" & strTableName & "] IN '' [;database=" & strDatabaseTo & ";] SELECT * From [" & strTableName & "]")
        
        Set tdfTo = Nothing
        Set rstFrom = Nothing
        Set fldFrom = Nothing
    End If
    
    If Err > 0 Then
        LogLine "Failed to copy from " & strTableName & " " & " " & Err.Source & ": " & Err.Number & ": " & Err.Description
        Err = 0
    Else
        LogLine "Copied " & strTableName
    End If
    
    mdbTo.Close
    mdbFrom.Close
    Set mdbTo = Nothing
    Set mdbFrom = Nothing
    
    
End Function
Public Function DatabaseCompact(strDatabaseName As String, Optional strTempPath As String = "") As Integer
    
    Dim mdbTmp As DAO.Database
    Dim iErrCounter As Integer
    
    
    
    On Local Error GoTo DatabasecompactError
    
    DatabaseCompact = False
    
    If strTempPath = "" Then
        strTempPath = TmpDir
    End If
    If Right$(strTempPath, 1) <> "\" Then
        strTempPath = strTempPath & "\"
    End If
                
  
    If Dir$(strTempPath & App.EXEName & ".Mdb") <> "" Then
        Kill strTempPath & App.EXEName & ".Mdb"
        DoEvents
    End If
    
    DoEvents
    If Err = 0 Then
        
        
        mdbMain.Close
        Set mdbMain = Nothing
        
        DoEvents
        
        DBEngine.CompactDatabase strDatabaseName, strTempPath & App.EXEName & ".Mdb"

        DoEvents
        
        
        If Err = 0 Then
            CopyFile strTempPath & App.EXEName & ".Mdb", strDatabaseName, True
            If Err = 0 Then
                LogLine "Database '" & strDatabaseName & "' compressed"
                DatabaseCompact = True
            Else
                MsgBox Err.Description, vbCritical
                LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
            End If
        Else
            MsgBox Err.Description, vbCritical
            LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
        End If
    Else
        LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
        Err = 0
    End If
    
    On Local Error GoTo 0
    
    Exit Function
    
DatabasecompactError:
    If iErrCounter < 5 Then
        iErrCounter = iErrCounter + 1
        Sleep 1
        Resume
    Else
        MsgBox Err.Description, vbCritical
        LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
    End If
    
End Function
Public Function DatabaseBackup(strDatabaseName As String, strBackupTo As String) As Integer
    Dim strTemp As String
    
    DatabaseBackup = False
    On Local Error Resume Next
    If strBackupTo <> "" Then
        If InStr(strBackupTo, "\") > 0 Then
            If Dir$(Left$(strBackupTo, InStrRev(strBackupTo, "\")), vbDirectory) = "" Then
                MkDir Left$(strBackupTo, InStrRev(strBackupTo, "\"))
            End If
        End If
        If CopyFile(strDatabaseName, strBackupTo, False) = True Then
            DatabaseBackup = True
        End If
    End If
    DoEvents
    On Local Error GoTo 0
End Function

Sub CheckOnIndexes()
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    
    SetMouseHourGlass
    For Each tbl In mdbMain.TableDefs
        If tbl.Attributes = 0 Then
            For Each fld In tbl.Fields
                Select Case fld.Name
                Case "STA", "CODE", "STATUS", "MARKS"
                    CreateIndex tbl, fld.Name
                Case Else
                    If Right$(fld.Name, 2) = "ID" Then
                        CreateIndex tbl, fld.Name
                    End If
                End Select
            Next
        End If
    Next
    SetMouseNormal
    
End Sub
Public Sub CopyField(FieldFrom As Field, FieldTo As Field)
    On Local Error Resume Next
    If FieldTo.Type = dbText Then
         FieldTo.Value = Left$(FieldFrom.Value & "", FieldTo.Size)
    ElseIf FieldTo.Type = dbDate Then
         If FieldFrom.Value Then
            If FieldFrom.Type = dbLong Or FieldFrom.Type = dbInteger Then
                FieldTo.Value = CDate("1-1-" & FieldFrom.Value)
            Else
                FieldTo.Value = CDate(FieldFrom.Value)
            End If
         End If
    Else
         FieldTo.Value = FieldFrom.Value
    End If
    On Local Error GoTo 0
End Sub
Public Function CheckDatabase(cDatabaseName As String) As Integer
    'routines to add missing (=new) fields to database
    
    Dim mdbNew As DAO.Database
    Dim rstNew As DAO.Recordset
    Dim tdfNew As DAO.TableDef
    Dim cLastTable As String
    Dim cTable As String
    Dim vFieldType As Variant
    Dim iFieldSize As Integer
    Dim iTemp As Integer
    Dim iErrorCount As Integer
    
    CheckDatabase = False
    
    On Local Error GoTo CheckDatabaseError
    
    Set mdbNew = DBEngine.OpenDatabase(cDatabaseName, False, False)
    
    cTable = "Fields"
    
    If TableExist(mdbNew, cTable) = False Then
        MsgBox App.EXEName & " could not create a proper database " & cDatabaseName & "."
        Unload frmMain
        End
    End If
    
    'make sure the "warnings" form works:
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Penalties' AND [Field]='PersonID'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Penalties"
            .Fields("Field") = "PersonID"
            .Fields("Label") = "PersonID"
            .Fields("Type") = "CHAR"
            .Fields("Seq") = 12
            .Fields("Comment") = "PersonID of affected rider"
            .Fields("Example") = "XX1234567890"
            .Update
        End With
    End If
    
    'make sure the "warnings" form works:
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Penalties' AND [Field]='Timestamp'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Penalties"
            .Fields("Field") = "Timestamp"
            .Fields("Label") = "Timestamp"
            .Fields("Type") = "DATE"
            .Fields("Seq") = 8
            .Fields("Comment") = "Date and time of creation"
            .Fields("Example") = "-"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Tests' AND [Field]='Removed'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Tests"
            .Fields("Field") = "Removed"
            .Fields("Label") = "Removed"
            .Fields("Type") = "BOOLEAN"
            .Fields("Seq") = 25
            .Fields("Comment") = "True = this test is no longer available"
            .Fields("Example") = "FALSE"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Tests' AND [Field]='WRTest'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Tests"
            .Fields("Field") = "WRTest"
            .Fields("Label") = "WRTest"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 29
            .Fields("Comment") = "Valid FEIF WorldRanking Test"
            .Fields("Example") = "T2"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Tests' AND [Field]='NRTest'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Tests"
            .Fields("Field") = "NRTest"
            .Fields("Label") = "NRTest"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 30
            .Fields("Comment") = "Valid National Ranking Test in several countries"
            .Fields("Example") = "T7"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='TestTimeTables' AND [Field]='ScaleRange'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Tests"
            .Fields("Field") = "ScaleRange"
            .Fields("Label") = "ScaleRange"
            .Fields("Type") = "CURRENCY"
            .Fields("Seq") = 27
            .Fields("Comment") = "Steps in marks in the ranking"
            .Fields("Example") = "0.2"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Entries' AND [Field]='Color'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Entries"
            .Fields("Field") = "Color"
            .Fields("Label") = "Color"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 10
            .Fields("Seq") = 10
            .Fields("Comment") = "Colored band used in group tests and finals"
            .Fields("Example") = "Red"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Entries' AND [Field]='Check'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Entries"
            .Fields("Field") = "Check"
            .Fields("Label") = "Check"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 11
            .Fields("Comment") = "A value other than 0 indicates a mandatory equipment check for this participant"
            .Fields("Example") = "1"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] WHERE [Table]='Marks' AND [Field]='Out'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Marks"
            .Fields("Field") = "Out"
            .Fields("Label") = "Out"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 18
            .Fields("Comment") = "Which mark is taken out when calculating score"
            .Fields("Example") = "2"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Results' AND [Field]='AllTimes'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Results"
            .Fields("Field") = "AllTimes"
            .Fields("Label") = "All Times"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 32
            .Fields("Seq") = 9
            .Fields("Comment") = "Overview of all times for this rider in ascending order"
            .Fields("Example") = "21.12 22.13 23.19 99.99"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Entries' AND [Field]='Nostart'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Entries"
            .Fields("Field") = "NoStart"
            .Fields("Label") = "No Start"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 15
            .Fields("Comment") = "Indicates if partipant will start in next heat (races only)"
            .Fields("Example") = "-1 = No start"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Marks' AND [Field]='Flag'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Marks"
            .Fields("Field") = "Flag"
            .Fields("Label") = "Flag"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 10
            .Fields("Comment") = "Are one or more red flags shown?"
            .Fields("Example") = 0
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Penalties' AND [Field]='Code'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Penalties"
            .Fields("Field") = "Code"
            .Fields("Label") = "Code"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 6
            .Fields("Comment") = "Test"
            .Fields("Example") = "T1"
            .Update
            
            .AddNew
            .Fields("Table") = "Penalties"
            .Fields("Field") = "Status"
            .Fields("Label") = "Status"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 7
            .Fields("Comment") = "Indicates if it was in Preliminary, B-Final or A-Final (in combination with code)"
            .Fields("Example") = "5.50"
            .Update
            
            .AddNew
            .Fields("Table") = "Penalties"
            .Fields("Field") = "Timestamp"
            .Fields("Label") = "Timestamp"
            .Fields("Type") = "Date/Time"
            .Fields("Seq") = 8
            .Fields("Comment") = ""
            .Fields("Example") = ""
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestTimeTables' AND [Field]='Code'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestTimeTables"
            .Fields("Field") = "Code"
            .Fields("Label") = "Code"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 1
            .Fields("Comment") = "Tests having a table to calculate marks from times"
            .Fields("Example") = "PP1"
            .Update
            
            .AddNew
            .Fields("Table") = "TestTimeTables"
            .Fields("Field") = "ScaleFast"
            .Fields("Label") = "ScaleFast"
            .Fields("Type") = "CURRENCY"
            .Fields("Seq") = 2
            .Fields("Comment") = "The fastest time in the scale"
            .Fields("Example") = "5.50"
            .Update
            
            .AddNew
            .Fields("Table") = "TestTimeTables"
            .Fields("Field") = "ScaleSlow"
            .Fields("Label") = "ScaleSlow"
            .Fields("Type") = "CURRENCY"
            .Fields("Seq") = 3
            .Fields("Comment") = "The slowest time in the scale"
            .Fields("Example") = "5.50"
            .Update
            
            .AddNew
            .Fields("Table") = "TestTimeTables"
            .Fields("Field") = "ScaleRange"
            .Fields("Label") = "ScaleRange"
            .Fields("Type") = "CURRENCY"
            .Fields("Seq") = 4
            .Fields("Comment") = "The range of the scale"
            .Fields("Example") = "20"
            .Update
            
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='Code'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Code"
            .Fields("Label") = "Code"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 1
            .Fields("Comment") = "Tests with additional info"
            .Fields("Example") = "F1"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Status"
            .Fields("Label") = "Status"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 2
            .Fields("Comment") = "Preliminary, A-Final or B-Final"
            .Fields("Example") = "0"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Handling"
            .Fields("Label") = "Handling"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 3
            .Fields("Comment") = "Way the test is handled"
            .Fields("Example") = "0"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "BFinal"
            .Fields("Label") = "BFinal"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 4
            .Fields("Comment") = "Highest position in the B-Final"
            .Fields("Example") = "0"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Sponsor"
            .Fields("Label") = "Sponsor"
            .Fields("Type") = "MEMO"
            .Fields("Seq") = 5
            .Fields("Comment") = "Who is the sponsor"
            .Fields("Example") = "Coca-Cola"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Nr"
            .Fields("Label") = "Nr"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 6
            .Fields("Comment") = "Sequence in the event"
            .Fields("Example") = "13"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "SortChar"
            .Fields("Label") = "SortChar"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 7
            .Fields("Comment") = "Character to start sorting start order with"
            .Fields("Example") = "M"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "SortDigit"
            .Fields("Label") = "SortDigit"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 8
            .Fields("Comment") = "Position in string to sort on"
            .Fields("Example") = "2"
            .Update
            
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='CFinal'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "CFinal"
            .Fields("Label") = "CFinal"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 13
            .Fields("Comment") = "Highest position in the C-Final"
            .Fields("Example") = "16"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='Num_J'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J"
            .Fields("Label") = "Num_J"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 9
            .Fields("Comment") = "Number of judges (old style)"
            .Fields("Example") = "5"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='Num_J_0'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_0"
            .Fields("Label") = "Num_J_0"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 10
            .Fields("Comment") = "Number of judges in preliminaries"
            .Fields("Example") = "5"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_1"
            .Fields("Label") = "Num_J_1"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 11
            .Fields("Comment") = "Number of judges in A-Final"
            .Fields("Example") = "5"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_2"
            .Fields("Label") = "Num_J_2"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 12
            .Fields("Comment") = "Number of judges in B-Final"
            .Fields("Example") = "5"
            .Update
            
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='Num_J_1'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_1"
            .Fields("Label") = "Num_J_1"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 10
            .Fields("Comment") = "Number of judges in A-Final"
            .Fields("Example") = "5"
            .Update
            
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_2"
            .Fields("Label") = "Num_J_2"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 11
            .Fields("Comment") = "Number of judges in B-Final"
            .Fields("Example") = "5"
            .Update
            
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='Num_J_3'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_3"
            .Fields("Label") = "Num_J_3"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 12
            .Fields("Comment") = "Number of judges in Final"
            .Fields("Example") = "5"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='TestInfo' AND [Field]='Num_J_4'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_4"
            .Fields("Label") = "Num_J_4"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 12
            .Fields("Comment") = "Number of judges in Final"
            .Fields("Example") = "5"
            .Update
        
            .AddNew
            .Fields("Table") = "TestInfo"
            .Fields("Field") = "Num_J_4"
            .Fields("Label") = "Num_J_4"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 12
            .Fields("Comment") = "Number of judges in Final"
            .Fields("Example") = "5"
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='SectionMarks'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "SectionMarks"
            .Fields("Field") = "STA"
            .Fields("Label") = "STA"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 3
            .Fields("Seq") = 1
            .Fields("Comment") = "linked to table Participants"
            .Fields("Example") = "001"
            .Update
            
            .AddNew
            .Fields("Table") = "SectionMarks"
            .Fields("Field") = "Code"
            .Fields("Label") = "Code"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 2
            .Fields("Comment") = "test code, linked to Tests"
            .Fields("Example") = "F1"
            .Update
            
            .AddNew
            .Fields("Table") = "SectionMarks"
            .Fields("Field") = "Status"
            .Fields("Label") = "Status"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 3
            .Fields("Comment") = "Preliminary, A-Final or B-Final"
            .Fields("Example") = "0"
            .Update
            
            .AddNew
            .Fields("Table") = "SectionMarks"
            .Fields("Field") = "Judge"
            .Fields("Label") = "Judge"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 4
            .Fields("Comment") = "Judge position"
            .Fields("Example") = "1"
            .Update
            
            .AddNew
            .Fields("Table") = "SectionMarks"
            .Fields("Field") = "Section"
            .Fields("Label") = "Section"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 5
            .Fields("Comment") = "identifies the section of the test"
            .Fields("Example") = "1"
            .Update
            
            .AddNew
            .Fields("Table") = "SectionMarks"
            .Fields("Field") = "Mark"
            .Fields("Label") = "Mark"
            .Fields("Type") = "CURRENCY"
            .Fields("Seq") = 6
            .Fields("Comment") = ""
            .Fields("Example") = "7.7"
            .Update
            
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Values'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Values"
            .Fields("Field") = "Code"
            .Fields("Label") = "Code"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 50
            .Fields("Seq") = 1
            .Update
            
            .AddNew
            .Fields("Table") = "Values"
            .Fields("Field") = "Field"
            .Fields("Label") = "Field"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 50
            .Fields("Seq") = 2
            .Update
            
            .AddNew
            .Fields("Table") = "Values"
            .Fields("Field") = "Label"
            .Fields("Label") = "Label"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 50
            .Fields("Seq") = 3
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Forms'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Forms"
            .Fields("Field") = "Title"
            .Fields("Label") = "Title"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 50
            .Fields("Seq") = 1
            .Update
            
            .AddNew
            .Fields("Table") = "Forms"
            .Fields("Field") = "RTFText"
            .Fields("Label") = "RTFText"
            .Fields("Type") = "Memo"
            .Fields("Seq") = 2
            .Update
            
            .AddNew
            .Fields("Table") = "Forms"
            .Fields("Field") = "Owner"
            .Fields("Label") = "Owner"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 50
            .Fields("Seq") = 3
            .Update
        
            .AddNew
            .Fields("Table") = "Forms"
            .Fields("Field") = "Editor"
            .Fields("Label") = "Editor"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 50
            .Fields("Seq") = 4
            .Update
        End With
    End If
    
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Forms' AND [Field]='FormType'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Forms"
            .Fields("Field") = "FormType"
            .Fields("Label") = "FormType"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 5
            .Fields("Comment") = "0 or Null=populated with rider/horse-data, 1=generic form"
            .Fields("Example") = "0"
            .Update
            
        End With
    End If
    
    'make sure the Ontrack data table exists
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [FIELDS] WHERE [Table]='Ontrack'")
    If rstNew.RecordCount = 0 Then
        With rstNew
            .AddNew
            .Fields("Table") = "Ontrack"
            .Fields("Field") = "Code"
            .Fields("Label") = "Code"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 8
            .Fields("Seq") = 1
            .Fields("Comment") = "test code, linked to Tests"
            .Fields("Example") = "T2"
            .Update
            
            .AddNew
            .Fields("Table") = "Ontrack"
            .Fields("Field") = "Position"
            .Fields("Label") = "Position"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 2
            .Fields("Comment") = "participant's position within this group when track is entered"
            .Fields("Example") = "3"
            .Update
            
            .AddNew
            .Fields("Table") = "Ontrack"
            .Fields("Field") = "Section"
            .Fields("Label") = "Section"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 3
            .Fields("Comment") = "current section of test"
            .Fields("Example") = "1"
            .Update
            
            .AddNew
            .Fields("Table") = "Ontrack"
            .Fields("Field") = "STA"
            .Fields("Label") = "STA"
            .Fields("Type") = "CHAR"
            .Fields("Length") = 3
            .Fields("Seq") = 4
            .Fields("Comment") = "linked to table Participants"
            .Fields("Example") = "001"
            .Update
            
            .AddNew
            .Fields("Table") = "Ontrack"
            .Fields("Field") = "Status"
            .Fields("Label") = "Status"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 5
            .Fields("Comment") = "Preliminary, A-Final or B-Final"
            .Fields("Example") = "0"
            .Update
            
            .AddNew
            .Fields("Table") = "Ontrack"
            .Fields("Field") = "Track"
            .Fields("Label") = "Track"
            .Fields("Type") = "INTEGER"
            .Fields("Seq") = 6
            .Fields("Comment") = "number of track"
            .Fields("Example") = "1"
            .Update
        End With
    End If
          
    Set rstNew = mdbNew.OpenRecordset("SELECT * FROM [Fields] ORDER BY [Table],[Field]")
    cLastTable = ""
    If rstNew.RecordCount > 0 Then
        cTable = rstNew.Fields("Table")
        Do While Not rstNew.EOF
            Select Case Left$(rstNew.Fields("Type"), 4)
            Case "BOOL"
                vFieldType = dbBoolean
                iFieldSize = 0
            Case "CURR"
                vFieldType = dbCurrency
                iFieldSize = 0
            Case "DATE"
                vFieldType = dbDate
                iFieldSize = 0
            Case "INTE", "LONG"
                vFieldType = dbLong
                iFieldSize = 0
            Case "MEMO"
                vFieldType = dbMemo
                iFieldSize = 0
            Case Else
                vFieldType = dbText
                If rstNew.Fields("Length") < 1 Or IsNull(rstNew.Fields("Length")) Then
                    iFieldSize = 50
                Else
                    iFieldSize = rstNew.Fields("Length")
                End If
            End Select
            If rstNew.Fields("Table") <> cLastTable Then
                CreateTable mdbNew, rstNew.Fields("Table"), rstNew.Fields("Field"), vFieldType, iFieldSize
                Set tdfNew = mdbNew.TableDefs(rstNew.Fields("Table"))
            End If
            CheckField mdbNew, rstNew.Fields("Table"), rstNew.Fields("Field"), vFieldType, iFieldSize
            If vFieldType = dbText Then
                AlterFieldSize mdbNew, rstNew.Fields("Table"), rstNew.Fields("Field"), iFieldSize
            End If
            
            cLastTable = rstNew.Fields("Table")
            rstNew.MoveNext
        Loop
    End If
    rstNew.Close
    Set rstNew = Nothing
    
CheckDatabaseError:
    mdbNew.Close
    Set mdbNew = Nothing
    
    If Err = 0 Then
        CheckDatabase = True
    ElseIf iErrorCount < 10 Then
        LogLine "Failed to add field to table " & cTable & ": " & Err.Source & ": " & Err.Number & ": " & Err.Description
        Err = 0
        iErrorCount = iErrorCount + 1
        Resume Next
    Else
        LogLine "Error creating/opening database " & cDatabaseName
        Exit Function
    End If
End Function
Sub CheckNumJ()
    Dim rstChk As DAO.Recordset
    Dim rstChk2 As DAO.Recordset
    Dim cQry As String
    
    On Local Error Resume Next
    
    cQry = "SELECT Num_J,Code"
    cQry = cQry & " FROM TestInfo"
    cQry = cQry & " WHERE ISNULL(TestInfo.Num_J)"
    Set rstChk = mdbMain.OpenRecordset(cQry)
    If rstChk.RecordCount > 0 Then
        Do While Not rstChk.EOF
            With rstChk
                cQry = "SELECT Num_J FROM Tests WHERE Code LIKE'" & rstChk.Fields("Code") & "'"
                Set rstChk2 = mdbMain.OpenRecordset(cQry)
                If rstChk2.RecordCount > 0 Then
                    .Edit
                    If IsNull(.Fields("Num_J")) Or .Fields("Num_J") = 0 Then
                        .Fields("Num_J") = rstChk2.Fields("Num_J")
                    End If
                    .Update
                End If
                rstChk2.Close
                .MoveNext
            End With
        Loop
    End If
    
    cQry = "SELECT TestInfo.Num_J, TestInfo.Num_J_0, TestInfo.Num_J_1, TestInfo.Num_J_2"
    cQry = cQry & " FROM TestInfo"
    cQry = cQry & " WHERE ISNULL(TestInfo.Num_J_0)"
    Set rstChk = mdbMain.OpenRecordset(cQry)
    If rstChk.RecordCount > 0 Then
        Do While Not rstChk.EOF
            With rstChk
                .Edit
                If IsNull(.Fields("Num_J_0")) Or .Fields("Num_J_0") = 0 Then
                    .Fields("Num_J_0") = .Fields("Num_J")
                End If
                If IsNull(.Fields("Num_J_1")) Or .Fields("Num_J_1") = 0 Then
                    .Fields("Num_J_1") = .Fields("Num_J")
                End If
                If IsNull(.Fields("Num_J_2")) Or .Fields("Num_J_2") = 0 Then
                    .Fields("Num_J_2") = .Fields("Num_J")
                End If
                If IsNull(.Fields("Num_J_3")) Or .Fields("Num_J_3") = 0 Then
                    .Fields("Num_J_3") = .Fields("Num_J")
                End If
                .Update
                .MoveNext
            End With
        Loop
    End If
    rstChk.Close
    
    Set rstChk2 = Nothing
    Set rstChk = Nothing
    
    On Local Error GoTo 0
End Sub
Sub CreateNewDatabase()
   Dim cTemp As String
   Dim cTemp2 As String
   Dim cMsg As String
   Dim cFipoYear As String
   Dim cFipoVersion As String
   Dim cFipoDate As String
   Dim cWR_Url As String
   Dim cWR2_Url As String
   
   Dim iKey As Integer
   Dim iCounter As Integer
   
   Dim mdbNew As DAO.Database
   Dim tdfNew As DAO.TableDef
   
   Do
        iCounter = iCounter + 1
        cTemp = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Event" & UserName & Format$(iCounter) & ".Mdb"
   Loop While Dir$(cTemp) <> ""
   
   On Local Error Resume Next
   
   With frmMain.CommonDialog1
        .CancelError = True
        .DefaultExt = ".Mdb"
        .DialogTitle = "Select a folder"
        .Filter = "Database|*.Mdb|"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly
        .FileName = cTemp
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        
        If .FileName <> "" Then
            cTemp = .FileName
        End If
    End With
    
    On Local Error GoTo 0
    
    If cTemp = "" Or cTemp = Chr$(27) Then
    Else
        If InStr(cTemp, "\") = 0 And InStr(cTemp, ":") = 0 Then
            cTemp2 = "C:\IceHorse\"
            If Dir$(cTemp2, vbDirectory) = "" Then
                If Err = 0 Then
                    MkDir cTemp2
                    If Err > 0 Then
                        cTemp2 = Environ$("APPDATA") & "\IceHorse\"
                        MkDir cTemp2
                        If Err > 0 Then
                           cTemp2 = ""
                        End If
                    End If
                Else
                    cTemp2 = Environ$("APPDATA") & "\IceHorse\"
                    MkDir cTemp2
                    If Err > 0 Then
                       cTemp2 = ""
                    End If
                End If
                cTemp = cTemp2 & cTemp
            End If
        End If
        
        iKey = MsgBox(Translate("Create new database", mcLanguage) & " '" & cTemp & "'?", vbQuestion + vbYesNo)
        If iKey = vbYes Then
            cFipoDate = GetVariable("FIPO")
            cFipoYear = GetVariable("FIPO Year")
            cFipoVersion = GetVariable("FIPO Version")
            cWR_Url = GetVariable("WR_Url")
            cWR2_Url = GetVariable("WR2_Url")
                   
            mdbMain.Close
            DoEvents
            
            If Dir$(Left$(cTemp, InStrRev(cTemp, "\")), vbDirectory) = "" Then
                MkDir Left$(cTemp, InStrRev(cTemp, "\"))
            End If
            If Dir$(Left$(cTemp, InStrRev(cTemp, "\")) & "Rtf\", vbDirectory) = "" Then
                MkDir Left$(cTemp, InStrRev(cTemp, "\")) & "Rtf"
            End If
            If Dir$(Left$(cTemp, InStrRev(cTemp, "\")) & "Html\", vbDirectory) = "" Then
                MkDir Left$(cTemp, InStrRev(cTemp, "\")) & "Html"
            End If
            
            CreateMdb cTemp
            
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "Fields", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "Values", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "Tests", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "TestSections", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "TestSplits", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "TestTimeTables", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "Combinations", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "CombinationSections", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "CombinationInfo", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "CountryNames", "", False
            CopyTableBetweenDatabases mcDatabaseName, cTemp, "Forms", "", False
            CopyQueriesBetweenDatabases mcDatabaseName, cTemp, "_"
            CopyQueriesBetweenDatabases mcDatabaseName, cTemp, "@"
            CopyQueriesBetweenDatabases mcDatabaseName, cTemp, "#"
                        
            cMsg = Translate("Copy existing participants as well? They might reappear at your next event.", mcLanguage)
            cMsg = cMsg & vbCrLf & Translate("- select 'Yes' to copy riders, horses and startnumbers (to use the same startnumbers again);", mcLanguage)
            cMsg = cMsg & vbCrLf & Translate("- select 'No' to copy riders and horses only (they will all need new startnumbers);", mcLanguage)
            cMsg = cMsg & vbCrLf & Translate("- select 'Cancel' to skip riders and participants.", mcLanguage)
            iKey = MsgBox(cMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1)
            If iKey = vbNo Then
                CopyTableBetweenDatabases mcDatabaseName, cTemp, "Persons", "", False
                CopyTableBetweenDatabases mcDatabaseName, cTemp, "Horses", "", False
            ElseIf iKey = vbYes Then
                CopyTableBetweenDatabases mcDatabaseName, cTemp, "Persons", "", False
                CopyTableBetweenDatabases mcDatabaseName, cTemp, "Horses", "", False
                CopyTableBetweenDatabases mcDatabaseName, cTemp, "Participants", "", False
            End If
            
            CheckDatabase cTemp
            
            mcDatabaseName = cTemp
            WriteIniFile gcIniHorseFile, "Database", "Folder", mcDatabaseName
            WriteIniFile gcIniFile, "Html Files", "Folder", Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Html\"
            WriteIniFile gcIniFile, "Rtf Files", "Folder", Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Rtf\"
            If CheckDatabase(mcDatabaseName) = True Then
                OpenDatabase mcDatabaseName
                SetVariable "ProgramVersion", ""
                SetVariable "FIPO", cFipoDate
                SetVariable "FIPO Year", cFipoYear
                SetVariable "FIPO Version", cFipoVersion
                SetVariable "WR_Url", cWR_Url
                SetVariable "WR2_Url", cWR2_Url
                SetVariable "VersionSwitch", mcVersionSwitch
                SetVariable "Country", mcCountry
                mdbMain.Close
                frmMain.RestartApp
            End If
        End If
    End If
End Sub

Function CreateBackup(d As DAO.Database, Optional cFilename As String) As Integer
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rst As DAO.Recordset
    Dim iTemp As Integer
    Dim iFileNum As Integer
        
    On Local Error Resume Next
    
    If cFilename = "" Then
        cFilename = GetVariable("Backup")
        If cFilename = "" Then
             cFilename = NameOfFile(mcDatabaseName) & ".Bak"
        End If
        
        With frmMain.CommonDialog1
             .CancelError = True
             .DefaultExt = ".Bak"
             .DialogTitle = Translate("Select a folder", mcLanguage)
             .Filter = Translate("Backup", mcLanguage) & "|*.Bak|"
             .FilterIndex = 1
             .Flags = cdlOFNHideReadOnly
             .FileName = cFilename
             .ShowOpen
             If Err = cdlCancel Then
                 Exit Function
             End If
             If .FileName <> "" Then
                 cFilename = NameOfFile(.FileName) & ".Bak"
             End If
         End With
    End If
    
    On Local Error GoTo CreateBackupError
    
    If cFilename = "" Or cFilename = Chr$(27) Then
    Else
        CreateBackup = True
        
        SetMouseHourGlass
        
        SetVariable "Backup", cFilename
        
        ShowProgressbar frmMain, 2, mdbMain.TableDefs.Count
        
        iFileNum = FreeFile
        Open cFilename For Output Access Write Shared As #iFileNum
        For Each tdf In d.TableDefs
            IncreaseProgressbarValue frmMain.ProgressBar1
            If tdf.Attributes = 0 And Left$(tdf.Name, 1) <> "_" Then
                Set rst = mdbMain.OpenRecordset("SELECT * FROM [" & tdf.Name & "]")
                If rst.RecordCount > 0 Then
                    Print #iFileNum, "#" & tdf.Name
                    For Each fld In tdf.Fields
                        Print #iFileNum, fld.Name & vbTab;
                    Next
                    Print #iFileNum, ""
                    Do While Not rst.EOF
                        For iTemp = 0 To rst.Fields.Count - 1
                            If rst.Fields(iTemp).Type = dbCurrency Or rst.Fields(iTemp).Type = dbDouble Or rst.Fields(iTemp).Type = dbSingle Then
                                'solve the problem of decimal comma's
                                Print #iFileNum, Replace(Format$(rst.Fields(iTemp).Value & ""), ",", ".") & vbTab;
                            ElseIf rst.Fields(iTemp).Type = dbBoolean Then
                                If rst.Fields(iTemp).Value = True Then
                                    'get rid of non-English versions of True
                                    Print #iFileNum, "-1" & vbTab;
                                ElseIf rst.Fields(iTemp).Value = False Then
                                    'get rid of non-English versions of False
                                    Print #iFileNum, "0" & vbTab;
                                End If
                            Else
                                Print #iFileNum, Replace(Replace(rst.Fields(iTemp).Value & "", vbTab, ";"), vbCrLf, "|") & vbTab;
                            End If
                        Next iTemp
                        Print #iFileNum, ""
                        rst.MoveNext
                    Loop
                End If
                rst.Close
            End If
        Next tdf
        Close #iFileNum
    End If

CreateBackupError:
    If Err > 0 Then
        CreateBackup = False
        MsgBox cFilename & ": " & Err.Description, vbCritical
    End If
    
    If iFileNum Then
        Close #iFileNum
    End If
    Set tdf = Nothing
    Set fld = Nothing
    Set rst = Nothing
    
    ShowProgressbar frmMain, 2, 0
    
    SetMouseNormal
    
End Function

Sub ReadBackUp(d As DAO.Database, Optional cFilename As String)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rst As DAO.Recordset
    
    Dim cFieldName() As String
    Dim cFieldValue() As String
    Dim iTemp As Integer
    Dim iFileNum As Integer
    Dim iKey As Integer
    
    Dim cTemp As String
    Dim cTableName As String
        
    On Local Error Resume Next
    
    If cFilename = "" Then
        cFilename = GetVariable("Backup")
        With frmMain.CommonDialog1
             .CancelError = True
             .DefaultExt = ".Txt"
             .DialogTitle = Translate("Select a folder", mcLanguage)
             .Filter = Translate("Backup", mcLanguage) & "|*.Bak|"
             .FilterIndex = 1
             .Flags = cdlOFNHideReadOnly
             .FileName = cFilename
             .ShowOpen
             If Err = cdlCancel Then
                 Exit Sub
             End If
             If .FileName <> "" Then
                 cFilename = NameOfFile(.FileName) & ".Bak"
             End If
         End With
    End If
    
    On Local Error GoTo ReadBackupError
    
    If cFilename = "" Or cFilename = Chr$(27) Then
    Else
        iKey = MsgBox(Translate("The backup will overwrite all existing data. Do you want to proceed?", mcLanguage), vbQuestion + vbYesNo)
        If iKey = vbYes Then
            SetMouseHourGlass
            
            SetVariable "Backup", cFilename
            
            
            iFileNum = FreeFile
            Open cFilename For Input Access Read Shared As #iFileNum
            
            ShowProgressbar frmMain, 2, LOF(iFileNum)
            
            Do While Not EOF(iFileNum)
                Line Input #iFileNum, cTemp
                IncreaseProgressbarValue frmMain.ProgressBar1, Loc(iFileNum)
                If Left$(cTemp, 1) = "#" Then
                    cTableName = Trim$(Mid$(cTemp & " ", 2))
                    If Not TableExist(d, cTableName) = True Then
                        MsgBox Translate("Backup is not compatible with current version of " & App.EXEName & ". Missing table:", mcLanguage) & cTableName & ".", vbCritical
                        cTableName = ""
                    Else
                        Set tdf = d.TableDefs(cTableName)
                        Line Input #iFileNum, cTemp
                        cFieldName = Split(cTemp, vbTab)
                        For iTemp = 0 To UBound(cFieldName)
                            If FieldExist(tdf, cFieldName(iTemp)) = False Then
                                cFieldName(iTemp) = ""
                            End If
                        Next iTemp
                        mdbMain.Execute ("DELETE * FROM [" & cTableName & "]")
                        Set rst = mdbMain.OpenRecordset("SELECT * FROM [" & cTableName & "]")
                    End If
                ElseIf cTableName <> "" Then
                    cFieldValue = Split(cTemp, vbTab)
                    rst.AddNew
                    For iTemp = 0 To UBound(cFieldName)
                        If cFieldValue(iTemp) <> "" And cFieldName(iTemp) <> "" Then
                            If rst.Fields(cFieldName(iTemp)).Type = dbDate Then
                                rst.Fields(cFieldName(iTemp)) = CDate(Replace(cFieldValue(iTemp), ".", "-"))
                            ElseIf rst.Fields(cFieldName(iTemp)).Type = dbCurrency Then
                                rst.Fields(cFieldName(iTemp)) = Val(Replace(cFieldValue(iTemp), ",", "."))
                            Else
                                rst.Fields(cFieldName(iTemp)) = cFieldValue(iTemp)
                            End If
                        End If
                    Next iTemp
                    rst.Update
                End If
            Loop
            rst.Close
            mdbMain.Close
            frmMain.RestartApp
        End If
    End If
    
ReadBackupError:
    If Err > 0 Then
        MsgBox cFilename & ": " & Err.Description, vbCritical
    End If
    
    If iFileNum Then
        Close #iFileNum
    End If
    Set tdf = Nothing
    Set fld = Nothing
    Set rst = Nothing
    
    ShowProgressbar frmMain, 2, 0
    
    SetMouseNormal

End Sub
Sub CopyQueriesBetweenDatabases(strDatabaseFrom As String, strDatabaseTo As String, strQueryName As String)
    Dim qdfFrom As DAO.QueryDef
    Dim qdfTo As DAO.QueryDef
    Dim mdbFrom As DAO.Database
    Dim mdbTo As DAO.Database
    
    On Local Error Resume Next
    
    Set mdbFrom = DBEngine.OpenDatabase(strDatabaseFrom, False, True)
    Set mdbTo = DBEngine.OpenDatabase(strDatabaseTo, False, False)
        
    For Each qdfFrom In mdbFrom.QueryDefs
        If strQueryName <> "" Then
            If Left$(qdfFrom.Name, Len(strQueryName)) = strQueryName Then
                Set qdfTo = New DAO.QueryDef
                qdfTo.Name = qdfFrom.Name
                qdfTo.SQL = qdfFrom.SQL
                mdbTo.QueryDefs.Append qdfTo
            End If
        Else
            Set qdfTo = New DAO.QueryDef
            qdfTo.Name = qdfFrom.Name
            qdfTo.SQL = qdfFrom.SQL
            mdbTo.QueryDefs.Append qdfTo
        End If
    Next
            
    mdbTo.Close
    mdbFrom.Close
    Set mdbTo = Nothing
    Set mdbFrom = Nothing
    
    On Local Error GoTo 0
    
End Sub
