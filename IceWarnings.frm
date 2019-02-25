VERSION 5.00
Begin VB.Form frmWarnings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Warnings"
   ClientHeight    =   4950
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbTest 
      DataField       =   "Test"
      DataSource      =   "dtaPenalties"
      Height          =   315
      Left            =   1800
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox cmbMeasure 
      DataField       =   "Measure"
      DataSource      =   "dtaPenalties"
      Height          =   315
      Left            =   1800
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   240
      TabIndex        =   13
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Save"
      Height          =   300
      Left            =   1560
      TabIndex        =   12
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   3000
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   300
      Left            =   4440
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Data dtaPenalties 
      Caption         =   "Warnings"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Width           =   5295
   End
   Begin VB.TextBox txtComments 
      DataField       =   "Comments"
      DataSource      =   "dtaPenalties"
      Height          =   525
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox txtCause 
      DataField       =   "Cause"
      DataSource      =   "dtaPenalties"
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblTimestamp 
      Caption         =   "Label1"
      DataField       =   "Timestamp"
      DataSource      =   "dtaPenalties"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label lblTimeStampLabel 
      Caption         =   "Timestamp:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblTestLabel 
      Caption         =   "Test:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblResponsiblePersonLabel 
      Caption         =   "Responsible:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblResponsiblePerson 
      Caption         =   "Label1"
      DataField       =   "responsibleperson"
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label lblCommentsLabel 
      Caption         =   "Comments:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblCauseLabel 
      Caption         =   "Cause:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblMeasureLabel 
      Caption         =   "Measure:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblAffectedPersonLabel 
      Caption         =   "Affected person:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblAffectedPerson 
      Caption         =   "Label1"
      DataField       =   "warnedperson"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmWarnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' IceWarnings.frm
' frmWarnings: Handling disciplinary measures with IceTest
' Copyright (C) Lutz Lesener 2007
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
Private Sub dtaPenalties_Reposition()
    Dim mysql As String
    Dim rstPersons As Recordset
    
    'are there warnings in the database?
    If dtaPenalties.Recordset.RecordCount = 0 Then
        lblAffectedPerson.Caption = ""
        lblResponsiblePerson.Caption = ""
        Exit Sub
    End If
    
    'Retrieve full names of affected persons (culprit and judge):
    mysql = "SELECT * FROM persons WHERE personid='" & dtaPenalties.Recordset("personid") & "' OR personid='" & dtaPenalties.Recordset("responsible_id") & "';"
    
    Set rstPersons = mdbMain.OpenRecordset(mysql)
    
    While Not rstPersons.EOF
        If rstPersons("personid") = dtaPenalties.Recordset("personid") Then
            lblAffectedPerson.Caption = rstPersons("Name_last") & ", " & rstPersons("Name_first")
        End If
        
        If rstPersons("personid") = dtaPenalties.Recordset("responsible_id") Then
            lblResponsiblePerson.Caption = rstPersons("Name_last") & ", " & rstPersons("Name_first")
        End If
        rstPersons.MoveNext
    Wend
    
    rstPersons.Close
    Set rstPersons = Nothing
End Sub
Private Sub Form_Load()
    Dim myquery As String
    Dim rstTemp As Recordset
    ReadFormPosition Me
    
    ChangeFontSize Me, msFontSize
    
    TranslateControls Me
    
    'Populate combobox with measures:
    cmbMeasure.AddItem Translate("Elimination from a class", mcLanguage)
    cmbMeasure.AddItem Translate("Warning not to be published", mcLanguage)
    cmbMeasure.AddItem Translate("Warning to be published", mcLanguage)
    cmbMeasure.AddItem Translate("Elimination from an event", mcLanguage)
    
    'Populate combobox with tests:
    myquery = "SELECT Tests.* FROM Tests INNER JOIN TestInfo ON Tests.Code=TestInfo.Code WHERE TestInfo.Nr>0 AND (Removed=False or ISNULL(Removed)) ORDER BY TestInfo.Nr"
    Set rstTemp = mdbMain.OpenRecordset(myquery)
    While Not rstTemp.EOF
        cmbTest.AddItem rstTemp("Code")
        rstTemp.MoveNext
    Wend
    rstTemp.Close
    Set rstTemp = Nothing
    
    
    'retrieve existing penalties:
    myquery = "SELECT * FROM penalties ORDER BY timestamp DESC;"
      
    With dtaPenalties
        .DatabaseName = mcDatabaseName
        .RecordSource = myquery
        .Refresh
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdUpdate_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
  WriteFormPosition Me
  Screen.MousePointer = vbDefault
End Sub
Private Sub cmdAdd_Click()
    Dim cTemp As String
    Dim iTemp As Integer
    Dim cSta As String
    Dim iPosition As Integer
    Dim rstTemp As Recordset
    Dim cQry As String
    Dim cJudge As String
    Dim cCulpritID As String
    Dim cCulprit As String
   
    'On Error GoTo AddErr

    'Look up the rider who received the warning:
    cTemp = InputBox$(Translate("Search for", mcLanguage), Translate("Rider receiving the warning?", mcLanguage))
   
    cQry = "SELECT Participants.Sta "
    cQry = cQry & " & '  -  ' & Persons.Name_First"
    cQry = cQry & " & ' ' & Persons.Name_Last"
    cQry = cQry & " & ' - ' & Horses.Name_Horse as cList"
    cQry = cQry & " FROM (Participants"
    cQry = cQry & " INNER JOIN Persons"
    cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
    cQry = cQry & " INNER JOIN Horses"
    cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
    cQry = cQry & " WHERE Participants.Sta & ' ' & Persons.Name_First & ' ' & Persons.Name_Last & ' - ' & Horses.Name_Horse LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
    cQry = cQry & " ORDER BY Participants.Sta"
   
    With frmToolBox
        .intChecked = True
        .strQry = cQry
        .Caption = Translate("Searching", mcLanguage) & " '" & cTemp & "' "
        .Show 1, Me
    End With
   
    If frmMain.Tempvar = "" Then
        Exit Sub
    Else
        'use the first selected participant (potential problem!)
        cSta = Left$(frmMain.Tempvar, 3)
        frmMain.Tempvar = ""
    End If
   
    'Look up the judge who gave the warning:
    cTemp = InputBox$(Translate("Search for", mcLanguage), Translate("Responsible judge?", mcLanguage))

    cQry = "SELECT Persons.PersonId & ' - ' & Persons.Name_First & ' ' & Persons.Name_Last  AS cList FROM persons "
    cQry = cQry & "WHERE Persons.Name_First & ' ' & Persons.Name_Last LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
   
    With frmToolBox
        .intChecked = True
        .strQry = cQry
        .Caption = Translate("Searching", mcLanguage) & " '" & cTemp & "' "
        .Show 1, Me
    End With
   
    If frmMain.Tempvar = "" Then
        Exit Sub
    Else
        'use the first selected person (potential problem!)
        cJudge = Trim(Left(frmMain.Tempvar, InStr(frmMain.Tempvar, "-") - 2))
    End If
    
    
    'Ok, culprit and judge are identified now.
    'Retrieve the ID and full name of the rider according to the start number:
    cQry = "SELECT Participants.STA, Persons.PersonID, Persons.Name_Last, Persons.Name_First "
    cQry = cQry & "FROM Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID "
    cQry = cQry & "WHERE Participants.STA='" & cSta & "';"
    
    Set rstTemp = mdbMain.OpenRecordset(cQry)
    If rstTemp.RecordCount > 0 Then
        cCulpritID = rstTemp("PersonID")
        cCulprit = rstTemp("Name_first") & " " & rstTemp("Name_last")
    End If
    rstTemp.Close
    Set rstTemp = Nothing
    
    'Create new entry:
    Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM penalties;")
    With rstTemp
        .AddNew
        .Fields("STA") = cSta
        .Fields("PersonID") = cCulpritID
        .Fields("Responsible_ID") = cJudge
        .Fields("Timestamp") = Now()
        .Update
    End With
    
    rstTemp.Close
    Set rstTemp = Nothing
    
    dtaPenalties.Refresh
    dtaPenalties.Recordset.MoveFirst
    
    Exit Sub
    

AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
    Dim iKey As Integer
    
    On Error GoTo DeleteErr
  
    iKey = MsgBox(Translate("Delete this warning?", mcLanguage), vbYesNo + vbQuestion)
    
    If iKey = vbYes Then
        dtaPenalties.Recordset.Delete
        dtaPenalties.Refresh
        dtaPenalties.Recordset.MoveFirst
        MsgBox Translate("Selected warning has been deleted.", mcLanguage)
    End If
    
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  
  dtaPenalties.UpdateRecord
    
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

