Attribute VB_Name = "modIceHorse"
' Functions related to horses

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
Function CreateHorseId() As String
    Dim rstId As Recordset
    Dim iCounter As Integer
    Dim cId As String
    
    Static iPrevCounter
    iCounter = iPrevCounter
    Do
        iCounter = iCounter + 1
        cId = "XX" & Format$(Now, "YYMMDD") & Format$(iCounter, "0000")
        Set rstId = mdbMain.OpenRecordset("SELECT HorseId from Horses WHERE HorseId LIKE '" & cId & "'")
    Loop While rstId.RecordCount > 0
    rstId.Close
    iPrevCounter = iCounter
    Set rstId = Nothing
    CreateHorseId = cId
End Function

