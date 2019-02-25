Attribute VB_Name = "modIceLogDB"
' IceLogDB.bas
' modIceLogDB: Write log data to an SQLite database.
' Copyright (C) Lutz Lesener 2010
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

Public sqliteLogDB As cConnection
Public strLogDBName As String

Function OpenLogDB() As Boolean
    Set sqliteLogDB = New_c.Connection
    
    'Put name of log database together:
    If strLogDBName = "" Then
        strLogDBName = GetSpecialFolderLocation(CSIDL_APPDATA) & "\IceHorse\logdb3.db3"
        Debug.Print "LogDB: " & strLogDBName
    End If
    
    'Does the database exist?
    If Dir$(strLogDBName) = "" Then
        sqliteLogDB.CreateNewDB strLogDBName
        sqliteLogDB.Execute GetTextFromFile(App.Path & "\logdb.sql")
        'sqliteLogDB.Execute GetTextFromFile(GetSpecialFolderLocation(CSIDL_APPDATA) & "\IceHorse\logdb.sql")
    Else
        sqliteLogDB.OpenDB strLogDBName
    End If

    OpenLogDB = True
End Function

Function WriteLogDBStart(tevent As String, tcode As String, tstatus As Integer, tsta As String, PPos As Integer, PGruppe As Integer, zeittext As String) As Integer
    Dim SQL As String
    Dim myset As cRecordset
    On Local Error GoTo WriteLogDBStartErr
    
    Dim mdbset As DAO.Recordset
    Dim myRiderID As String, myHorseID As String
    Dim myRider As String, myHorse As String
    Dim myPosition As Integer, mygroup As Integer, mycolor As String, myrr As Integer
    
    SQL = "SELECT Participants.STA, Persons.Name_First, Persons.Name_Last, Persons.FEIFID As PID, Horses.FEIFID As HID, Horses.Name_Horse "
    SQL = SQL & "FROM (Participants INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) INNER JOIN Persons ON Participants.PersonID = Persons.PersonID "
    SQL = SQL & "WHERE (((Participants.STA)='" & tsta & "'));"
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        myRiderID = mdbset.Fields("pid") & ""
        myRider = mdbset.Fields("name_first") & " " & mdbset.Fields("name_last") & ""
        myHorseID = mdbset.Fields("hid") & ""
        myHorse = mdbset.Fields("name_horse") & ""
    Else
        myRiderID = ""
        myRider = ""
        myHorseID = ""
        myHorse = ""
    End If
    mdbset.Close
    Set mdbset = Nothing
    
    'Versuchen die Gruppe und Position zu finden:
    SQL = "SELECT * FROM entries "
    SQL = SQL & "WHERE code='" & tcode & "' AND status=" & tstatus
    SQL = SQL & "AND sta='" & tsta & "'"
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        myPosition = mdbset.Fields("position")
        mygroup = mdbset.Fields("group")
        myrr = mdbset.Fields("rr")
    Else
        myPosition = PPos
        mygroup = PGruppe
        myrr = 0
    End If
    mdbset.Close
    Set mdbset = Nothing
    'Werte umkopieren:
    PPos = myPosition
    PGruppe = mygroup
    
    'Schritt 1: Falls vorhanden, diesen Start aus der LogDB löschen:
    SQL = "DELETE FROM startlists WHERE competition='" & tevent & "' AND testcode='" & tcode & "' AND teststatus=" & tstatus & " AND sta='" & tsta & "'"
    sqliteLogDB.Execute SQL
    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    
    'Schritt 2: Start eintragen:
    SQL = "INSERT INTO startlists (competition, testcode, teststatus, sta, startgroup, position, message, riderid, rider, horseid, horse) "
    SQL = SQL & "VALUES ('" & tevent & "', '" & tcode & "', " & tstatus & ", '" & tsta & "', " & PGruppe & ", " & PPos & ",' " & zeittext & "', "
    SQL = SQL & "'" & myRiderID & "', '" & myRider & "', '" & myHorseID & "', '" & myHorse & "');"
    sqliteLogDB.Execute SQL
        
    WriteLogDBStart = 1
    Exit Function
    
WriteLogDBStartErr:
    WriteLogDBStart = 0
    Exit Function
End Function
Function WriteLogDBMarks(tevent As String, tcode As String, tstatus As Integer, tsta As String, taction As Integer, tmessage As String, tsection As Integer) As Boolean
    Dim SQL As String
    Dim myset As cRecordset
    On Local Error GoTo WriteLogDBMarksErr
    
    Dim mdbset As DAO.Recordset
    Dim myRiderID As String, myHorseID As String
    Dim myRider As String, myHorse As String
    SQL = "SELECT Participants.STA, Persons.Name_First, Persons.Name_Last, Persons.FEIFID As PID, Horses.FEIFID As HID, Horses.Name_Horse "
    SQL = SQL & "FROM (Participants INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) INNER JOIN Persons ON Participants.PersonID = Persons.PersonID "
    SQL = SQL & "WHERE (((Participants.STA)='" & tsta & "'));"
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        myRiderID = mdbset.Fields("pid") & ""
        myRider = mdbset.Fields("name_first") & " " & mdbset.Fields("name_last") & ""
        myHorseID = mdbset.Fields("hid") & ""
        myHorse = mdbset.Fields("name_horse") & ""
    Else
        myRiderID = ""
        myRider = ""
        myHorseID = ""
        myHorse = ""
    End If
    mdbset.Close
    Set mdbset = Nothing
    
    If Left(tmessage, 2) = "> " Then
        tmessage = Right(tmessage, Len(tmessage) - 2)
    End If
    
    'Zunächst auch hier checken, ob exakt dieses Resultat schon existiert:
    SQL = "SELECT * FROM results WHERE "
    SQL = SQL & "competition='" & tevent & "' AND "
    SQL = SQL & "testcode='" & tcode & "' AND "
    SQL = SQL & "teststatus=" & tstatus & " AND "
    SQL = SQL & "testsection=" & tsection & " AND "
    SQL = SQL & "sta='" & tsta & "' AND "
    SQL = SQL & "testaction=" & taction & " AND "
    SQL = SQL & "message='" & tmessage & "'"
    Set myset = sqliteLogDB.OpenRecordset(SQL)
    If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    
    'nur neue Infos schreiben:
    If myset.RecordCount = 0 Then
        SQL = "INSERT INTO results (competition, testcode, teststatus, testsection, sta, testaction, message, riderid, rider, horseid, horse) "
        SQL = SQL & "VALUES ('" & tevent & "', '" & tcode & "', " & tstatus & ", " & tsection & ", '" & tsta & "', " & taction & ",'" & tmessage & "', "
        SQL = SQL & "'" & myRiderID & "', '" & myRider & "', '" & myHorseID & "', '" & myHorse & "');"
    
        sqliteLogDB.Execute SQL
    End If

    WriteLogDBMarks = True
    Exit Function
    
WriteLogDBMarksErr:
    WriteLogDBMarks = False
    Exit Function
End Function
'
Function WriteLogDBMarks2(tevent As String, tcode As String, tstatus As Integer, tsta As String) As Boolean

    Dim SQL As String
    Dim myset As cRecordset
    Dim mdbset As DAO.Recordset
    Dim mymark(5) As Currency
    Dim myscore As Currency
    Dim mysection As Integer
    Dim myRider As String, myRiderID As String
    Dim myHorse As String, myHorseID As String
    
    On Local Error GoTo WriteLogDBMarksErr2
    
    
    'Get the marks/times:
    SQL = "SELECT Marks.Code, Marks.Mark1, Marks.Mark2, Marks.Mark3, Marks.Mark4, Marks.Mark5, Marks.Score, Marks.Section, Marks.Status, Marks.STA, Horses.Name_Horse, Persons.Name_Last, Persons.Name_First, Persons.FEIFID, Horses.FEIFID "
    SQL = SQL & "FROM Persons INNER JOIN ((Participants INNER JOIN Marks ON Participants.STA = Marks.STA) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) ON Persons.PersonID = Participants.PersonID "
    SQL = SQL & "WHERE Marks.Code='" & tcode & "' AND Marks.Status=" & tstatus & " "
    SQL = SQL & "AND Marks.STA='" & tsta & "'"
    
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        While Not mdbset.EOF
            'Copy recordset data to variables:
            mymark(1) = IIf(IsNull(mdbset.Fields("mark1")), -1, mdbset.Fields("mark1"))
            mymark(2) = IIf(IsNull(mdbset.Fields("mark2")), -1, mdbset.Fields("mark2"))
            mymark(3) = IIf(IsNull(mdbset.Fields("mark3")), -1, mdbset.Fields("mark3"))
            mymark(4) = IIf(IsNull(mdbset.Fields("mark4")), -1, mdbset.Fields("mark4"))
            mymark(5) = IIf(IsNull(mdbset.Fields("mark5")), -1, mdbset.Fields("mark5"))
            myscore = mdbset.Fields("score")
            mysection = mdbset.Fields("section")
            myRider = mdbset.Fields("name_first") & " " & mdbset.Fields("name_last")
            myRiderID = mdbset.Fields("persons.feifid") & ""
            myHorse = mdbset.Fields("name_horse") & ""
            myHorseID = mdbset.Fields("horses.feifid") & ""
            
            'Second, check if exactly this result is already in the database.
            'We do this to keep the internet traffic as low as possible.
            SQL = "SELECT * FROM results WHERE "
            SQL = SQL & "competition=" & SanitizeLogDBString(tevent) & " AND "
            SQL = SQL & "testcode=" & SanitizeLogDBString(tcode) & " AND "
            SQL = SQL & "teststatus=" & tstatus & " AND "
            SQL = SQL & "testsection=" & mysection & " AND "
            SQL = SQL & "sta=" & SanitizeLogDBString(tsta) & " AND "
            SQL = SQL & "rider=" & SanitizeLogDBString(myRider) & " AND "
            SQL = SQL & "riderid=" & SanitizeLogDBString(myRiderID) & " AND "
            SQL = SQL & "horse=" & SanitizeLogDBString(myHorse) & " AND "
            SQL = SQL & "horseid=" & SanitizeLogDBString(myHorseID) & " AND "
            SQL = SQL & "score=" & Replace(Format(myscore, "0.00"), ",", ".") & " AND "
            SQL = SQL & "mark1=" & Replace(Format(mymark(1), "0.00"), ",", ".") & " AND "
            SQL = SQL & "mark2=" & Replace(Format(mymark(2), "0.00"), ",", ".") & " AND "
            SQL = SQL & "mark3=" & Replace(Format(mymark(3), "0.00"), ",", ".") & " AND "
            SQL = SQL & "mark4=" & Replace(Format(mymark(4), "0.00"), ",", ".") & " AND "
            SQL = SQL & "mark5=" & Replace(Format(mymark(5), "0.00"), ",", ".")
                        
            Set myset = sqliteLogDB.OpenRecordset(SQL)
            If Err Then MsgBox Err.Description: Err.Clear: Exit Function
            
            'nur neue Infos schreiben:
            If myset.RecordCount = 0 Then
                SQL = "INSERT INTO results (competition, testcode, teststatus, testsection, sta, riderid, rider, horseid, horse, score, mark1, mark2, mark3, mark4, mark5) "
                SQL = SQL & "VALUES ("
                SQL = SQL & SanitizeLogDBString(tevent) & ", "
                SQL = SQL & SanitizeLogDBString(tcode) & ", "
                SQL = SQL & tstatus & ", "
                SQL = SQL & mysection & ", "
                SQL = SQL & SanitizeLogDBString(tsta) & ", "
                SQL = SQL & SanitizeLogDBString(myRiderID) & ", "
                SQL = SQL & SanitizeLogDBString(myRider) & ", "
                SQL = SQL & SanitizeLogDBString(myHorseID) & ", "
                SQL = SQL & SanitizeLogDBString(myHorse) & ", "
                SQL = SQL & Replace(Format(myscore, "0.00"), ",", ".") & ", "
                SQL = SQL & Replace(Format(mymark(1), "0.00"), ",", ".") & ", "
                SQL = SQL & Replace(Format(mymark(2), "0.00"), ",", ".") & ", "
                SQL = SQL & Replace(Format(mymark(3), "0.00"), ",", ".") & ", "
                SQL = SQL & Replace(Format(mymark(4), "0.00"), ",", ".") & ", "
                SQL = SQL & Replace(Format(mymark(5), "0.00"), ",", ".")
                SQL = SQL & ")"
                sqliteLogDB.Execute SQL
            End If
            
            mdbset.MoveNext
        Wend
    End If
    
    
    mdbset.Close
    Set mdbset = Nothing
    
    

    WriteLogDBMarks2 = True
    Exit Function
    
WriteLogDBMarksErr2:
    WriteLogDBMarks2 = False
    Exit Function
End Function
Function SanitizeLogDBString(data As String) As String
    SanitizeLogDBString = "'" & Replace(data, "'", "") & "'"
End Function
'DelLogDBCOnfMarks deletes the passed test from the log db.
Function DelLogDBConfMarks(tcompetition As String, tcode As String, tstatus As Integer) As Boolean
    Dim SQL As String
    On Local Error GoTo DelLogDBConfMarksErr
    SQL = "DELETE FROM confresults"
    SQL = SQL & " WHERE competition = " & SanitizeLogDBString(tcompetition)
    SQL = SQL & " AND testcode = " & SanitizeLogDBString(tcode)
    SQL = SQL & " AND teststatus = " & tstatus
    sqliteLogDB.Execute SQL
    DelLogDBConfMarks = True
    Exit Function
    
DelLogDBConfMarksErr:
    DelLogDBConfMarks = False
    Exit Function
End Function
'WriteLogDBConfMarks saves a printed result list entry to the log db.
Function WriteLogDBConfMarks2(tcompetition As String, tcode As String, twr As String, tstatus As Integer, cmstatus As String) As Boolean
    Dim SQL As String
    Dim myset As cRecordset
    Dim sstatus As Integer
    On Local Error GoTo WriteLogDBConfMarksErr2
    
    
    
    
    If Left(tmessage, 2) = "> " Then
        tmessage = Right(tmessage, Len(tmessage) - 2)
    End If
    
    'Set detailed status of this result list:
    sstatus = 0
    If cmstatus = "Provisional Result List" Then sstatus = 1
    If cmstatus = "Final Result List" Then sstatus = 2
    If Left(cmstatus, 19) = "Revised Result List" Then sstatus = 3
    
    Dim mdbset As DAO.Recordset
    
    'Let's try to find the judges:
    Dim myjudge(5) As String
    Dim myjudgeid(5) As String
    Dim zaehler As Integer
    SQL = "SELECT TestJudges.Code, TestJudges.Status, TestJudges.Position, TestJudges.JudgeId, Persons.Name_First, Persons.Name_Last "
    SQL = SQL & "FROM Persons INNER JOIN TestJudges ON Persons.PersonID = TestJudges.JudgeId "
    SQL = SQL & "WHERE TestJudges.Code='" & tcode & "' "
    SQL = SQL & "AND TestJudges.Status=" & tstatus & " ORDER BY position ASC"
    zaehler = 0
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        While Not mdbset.EOF
            zaehler = zaehler + 1
            myjudge(zaehler) = mdbset.Fields("name_first") & " " & mdbset.Fields("name_last")
            myjudgeid(zaehler) = mdbset.Fields("judgeid") & ""
            mdbset.MoveNext
        Wend
    End If
    mdbset.Close
    Set mdbset = Nothing
    
    'Get the result list, shall we?
    SQL = "SELECT Results.Code, Results.Status, Results.STA, Results.Position, Results.Disq, Results.Score, Horses.FEIFID, Horses.Name_Horse, Persons.Name_First, Persons.Name_Last, Persons.FEIFID "
    SQL = SQL & "FROM Persons INNER JOIN (Horses INNER JOIN (Results INNER JOIN Participants ON Results.STA = Participants.STA) ON Horses.HorseID = Participants.HorseID) ON Persons.PersonID = Participants.PersonID "
    SQL = SQL & "WHERE Results.Code='" & tcode & "' "
    SQL = SQL & "AND Results.Status=" & tstatus
    
    Dim myPosition As Integer
    Dim myDisq As Integer
    Dim myscore As Currency
    Dim mySta As String
    Dim myRider As String
    Dim myRiderID As String
    Dim myHorse As String
    Dim myHorseID As String
    Dim singlemarks As Boolean
    
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        While Not mdbset.EOF
            'Process each result:
            
            'First, copy the values we need into variables:
            mySta = mdbset.Fields("sta") & ""
            myPosition = mdbset.Fields("position")
            myDisq = mdbset.Fields("disq")
            myscore = mdbset.Fields("score")
            myRider = mdbset.Fields("name_first") & " " & mdbset.Fields("name_last")
            myRiderID = mdbset.Fields("persons.feifid") & ""
            myHorse = mdbset.Fields("name_horse") & ""
            myHorseID = mdbset.Fields("horses.feifid") & ""
            
            'Second, check if exactly this result is already in the database.
            'We do this to keep the internet traffic as low as possible.
            SQL = "SELECT * FROM confresults WHERE "
            SQL = SQL & "competition=" & SanitizeLogDBString(tcompetition) & " AND "
            SQL = SQL & "testcode=" & SanitizeLogDBString(tcode) & " AND "
            SQL = SQL & "wrcode=" & SanitizeLogDBString(twr) & " AND "
            SQL = SQL & "teststatus=" & tstatus & " AND "
            SQL = SQL & "sta=" & SanitizeLogDBString(mySta) & " AND "
            SQL = SQL & "rider=" & SanitizeLogDBString(myRider) & " AND "
            SQL = SQL & "riderid=" & SanitizeLogDBString(myRiderID) & " AND "
            SQL = SQL & "horse=" & SanitizeLogDBString(myHorse) & " AND "
            SQL = SQL & "horseid=" & SanitizeLogDBString(myHorseID) & " AND "
            SQL = SQL & "score=" & Replace(Format(myscore, "0.00"), ",", ".") & " AND "
            SQL = SQL & "position=" & myPosition & " AND "
            SQL = SQL & "disq=" & myDisq & " AND "
            SQL = SQL & "judge1=" & SanitizeLogDBString(myjudge(1)) & " AND "
            SQL = SQL & "judge2=" & SanitizeLogDBString(myjudge(2)) & " AND "
            SQL = SQL & "judge3=" & SanitizeLogDBString(myjudge(3)) & " AND "
            SQL = SQL & "judge4=" & SanitizeLogDBString(myjudge(4)) & " AND "
            SQL = SQL & "judge5=" & SanitizeLogDBString(myjudge(5)) & " AND "
            SQL = SQL & "judgeid1=" & SanitizeLogDBString(myjudgeid(1)) & " AND "
            SQL = SQL & "judgeid2=" & SanitizeLogDBString(myjudgeid(2)) & " AND "
            SQL = SQL & "judgeid3=" & SanitizeLogDBString(myjudgeid(3)) & " AND "
            SQL = SQL & "judgeid4=" & SanitizeLogDBString(myjudgeid(4)) & " AND "
            SQL = SQL & "judgeid5=" & SanitizeLogDBString(myjudgeid(5))
                        
            Set myset = sqliteLogDB.OpenRecordset(SQL)
            If Err Then MsgBox Err.Description: Err.Clear: Exit Function
    
            'nur neue Infos schreiben:
            If myset.RecordCount = 0 Then
                SQL = "INSERT INTO confresults (competition, testcode, wrcode, teststatus, sta, riderid, rider, horseid, horse, score, position, disq, substatus, judge1, judge2, judge3, judge4, judge5, judgeid1, judgeid2, judgeid3, judgeid4, judgeid5) "
                SQL = SQL & "VALUES ("
                SQL = SQL & SanitizeLogDBString(tcompetition) & ", "
                SQL = SQL & SanitizeLogDBString(tcode) & ", "
                SQL = SQL & SanitizeLogDBString(twr) & ", "
                SQL = SQL & tstatus & ", "
                SQL = SQL & SanitizeLogDBString(mySta) & ", "
                SQL = SQL & SanitizeLogDBString(myRiderID) & ", "
                SQL = SQL & SanitizeLogDBString(myRider) & ", "
                SQL = SQL & SanitizeLogDBString(myHorseID) & ", "
                SQL = SQL & SanitizeLogDBString(myHorse) & ", "
                SQL = SQL & Replace(Format(myscore, "0.00"), ",", ".") & ", "
                SQL = SQL & myPosition & ", "
                SQL = SQL & myDisq & ", "
                SQL = SQL & sstatus & ", "
                SQL = SQL & SanitizeLogDBString(myjudge(1)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudge(2)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudge(3)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudge(4)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudge(5)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudgeid(1)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudgeid(2)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudgeid(3)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudgeid(4)) & ", "
                SQL = SQL & SanitizeLogDBString(myjudgeid(5))
                SQL = SQL & ")"
                sqliteLogDB.Execute SQL
            End If
            
            'Make sure we have all the single marks:
            singlemarks = WriteLogDBMarks2(tcompetition, tcode, tstatus, mySta)
            
            mdbset.MoveNext
        Wend
            
    End If
                        

    mdbset.Close
    Set mdbset = Nothing
    

    WriteLogDBConfMarks2 = True
    Exit Function
    
WriteLogDBConfMarksErr2:
    WriteLogDBConfMarks2 = False
    Exit Function
End Function
'WriteLogDBConfMarks saves a printed result list entry to the log db.
Function WriteLogDBConfMarks(tcompetition As String, tcode As String, twr As String, tstatus As Integer, tsta As String, triderid As String, trider As String, thorseid As String, thorse As String, tmessage As String, tscore As Double, tposition As Integer, tdisq As Integer, tsection As Integer, cmstatus As String, richter As String) As Boolean
    Dim SQL As String
    Dim myset As cRecordset
    Dim sstatus As Integer
    On Local Error GoTo WriteLogDBConfMarksErr
    
    If Left(tmessage, 2) = "> " Then
        tmessage = Right(tmessage, Len(tmessage) - 2)
    End If
    
    'Set detailed status of this result list:
    sstatus = 0
    If cmstatus = "Provisional Result List" Then sstatus = 1
    If cmstatus = "Final Result List" Then sstatus = 2
    If Left(cmstatus, 19) = "Revised Result List" Then sstatus = 3
    
    'Fetch FEIF-IDs from database if possible:
    Dim mdbset As DAO.Recordset
    Dim myRiderID As String, myHorseID As String
    Dim myRider As String, myHorse As String
    SQL = "SELECT Participants.STA, Persons.Name_First, Persons.Name_Last, Persons.FEIFID As PID, Horses.FEIFID As HID, Horses.Name_Horse "
    SQL = SQL & "FROM (Participants INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) INNER JOIN Persons ON Participants.PersonID = Persons.PersonID "
    SQL = SQL & "WHERE (((Participants.STA)='" & tsta & "'));"
    Set mdbset = mdbMain.OpenRecordset(SQL)
    If mdbset.RecordCount > 0 Then
        myRiderID = mdbset.Fields("pid") & ""
        myRider = mdbset.Fields("name_first") & " " & mdbset.Fields("name_last") & ""
        myHorseID = mdbset.Fields("hid") & ""
        myHorse = mdbset.Fields("name_horse") & ""
    Else
        myRiderID = ""
        myRider = ""
        myHorseID = ""
        myHorse = ""
    End If
    mdbset.Close
    Set mdbset = Nothing
    
    'Infos schreiben:
    SQL = "INSERT INTO confresults (competition, testcode, wrcode, teststatus, testsection, sta, riderid, rider, horseid, horse, message, score, position, disq, substatus, judges) "
    SQL = SQL & "VALUES ("
    SQL = SQL & SanitizeLogDBString(tcompetition) & ", "
    SQL = SQL & SanitizeLogDBString(tcode) & ", "
    SQL = SQL & SanitizeLogDBString(twr) & ", "
    SQL = SQL & tstatus & ", "
    SQL = SQL & tsection & ", "
    SQL = SQL & SanitizeLogDBString(tsta) & ", "
    SQL = SQL & SanitizeLogDBString(myRiderID) & ", "
    SQL = SQL & SanitizeLogDBString(trider) & ", "
    SQL = SQL & SanitizeLogDBString(myHorseID) & ", "
    SQL = SQL & SanitizeLogDBString(thorse) & ", "
    SQL = SQL & SanitizeLogDBString(tmessage) & ", "
    SQL = SQL & Replace(Format(tscore, "0.00"), ",", ".") & ", "
    SQL = SQL & tposition & ", "
    SQL = SQL & tdisq & ", "
    SQL = SQL & sstatus & ", "
    SQL = SQL & SanitizeLogDBString(richter)
    SQL = SQL & ")"
    sqliteLogDB.Execute SQL


    WriteLogDBConfMarks = True
    Exit Function
    
WriteLogDBConfMarksErr:
    WriteLogDBConfMarks = False
    Exit Function
End Function
Private Function GetTextFromFile(FileName$) As String
    Dim FNr&: FNr = FreeFile
    Open FileName For Binary Access Read As FNr
    GetTextFromFile = Space(LOF(FNr))
    Get FNr, , GetTextFromFile: Close FNr
End Function

