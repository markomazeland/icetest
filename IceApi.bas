Attribute VB_Name = "modIceApi"
'This module contains public domain api calls

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

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETTABSTOPS = &H192
Public Const MAX_strPath = 260

Public Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type


Sub ReadIniFile(strIniFileName As String, strSection As String, strLabel As String, strValue As String, Optional intPersonal As Integer = 0, Optional intSpecial As Integer = 0)
   'read value from ini file
    Dim lngTemp As Long
    Dim intTemp As Integer
    
    If strIniFileName = "" Then
       strIniFileName = App.EXEName + ".Ini"
    ElseIf InStr(strIniFileName, "\") = 0 And InStr(strIniFileName, ".") = 0 Then
       strIniFileName = strIniFileName + ".Ini"
    End If
    
    strValue = Space$(250)
    If intPersonal = 0 Then
        lngTemp = GetPrivateProfileString(strSection, strLabel, strValue, strValue, CLng(Len(strValue)), strIniFileName)
    Else
        lngTemp = GetPrivateProfileString(strSection & "_" & UserName, strLabel, strValue, strValue, CLng(Len(strValue)), strIniFileName)
    End If
    Do While Left$(strValue, 1) = Chr$(0)
        strValue = Mid$(strValue, 2)
    Loop
    strValue = StripNull(strValue)
    If intSpecial = 0 Then
        intTemp = InStr(strValue, ";")
        If intTemp > 0 Then
            strValue = RTrim$(Left$(strValue, intTemp - 1))
        End If
    End If
    strValue = Trim$(strValue)
    If strValue = "" And intPersonal <> 0 Then
        ReadIniFile strIniFileName, strSection, strLabel, strValue, 0, intSpecial
    End If
End Sub
Sub WriteIniFile(strIniFileName As String, strSection As String, strLabel As String, strValue As String, Optional intPersonal As Integer = 0)
   'write value to ini file
    Dim lngTemp As Long
    
    strValue = Trim$(strValue)
    If strIniFileName = "" Then
       strIniFileName = App.EXEName + ".Ini"
    ElseIf InStr(strIniFileName, "\") = 0 And InStr(strIniFileName, ".") = 0 Then
       strIniFileName = strIniFileName + ".Ini"
    End If
    If intPersonal = 0 Then
        lngTemp = WritePrivateProfileString(strSection, strLabel, strValue, strIniFileName)
    Else
        lngTemp = WritePrivateProfileString(strSection & "_" & UserName, strLabel, strValue, strIniFileName)
    End If
End Sub
Public Function UserName() As String
   'name of the logged in user
   Dim strTemp As String
   
   Static cStaticUserName As String
   If cStaticUserName = "" Then
        strTemp = String(255, 0)
        If GetUserName(strTemp, Len(strTemp)) Then
           cStaticUserName = StripNull(strTemp)
        End If
    End If
    UserName = cStaticUserName
End Function
Public Function StripNull(strLine As String) As String
    'strips null chars from strings
    Dim intTemp As Integer
    
    intTemp = InStr(strLine, Chr$(0))
    If intTemp > 0 Then
      StripNull = Left$(strLine, intTemp - 1)
    Else
      StripNull = strLine
    End If
End Function
Public Function WinDir() As String
    'name of the windows directory
    Dim strTmpPath As String * 512
    Dim lngRet As Long
    
    lngRet = GetWindowsDirectory(strTmpPath, 512)
    If (lngRet > 0 And lngRet < 512) Then
       WinDir = Left$(strTmpPath, InStr(strTmpPath, vbNullChar) - 1)
       If Right$(WinDir, 1) <> "\" Then
            WinDir = WinDir & "\"
       End If
    End If
End Function


Public Function MachineName(Optional intGetNew As Integer = False) As String
    'determines the name of the PC/workstation
    Dim strTemp As String
    Dim lngRet As Long
   
    Static strStaticMachineName As String
    
    If intGetNew = True Then
         strStaticMachineName = ""
    End If
    If strStaticMachineName = "" Then
         strTemp = Space(256)
         lngRet = Len(strTemp)
         If GetComputerName(strTemp, lngRet) <> 0 And lngRet > 0 Then
             strStaticMachineName = StripNull(strTemp)
         End If
         If Environ$("WINSTATIONNAME") <> "" Then
            strStaticMachineName = strStaticMachineName & "_" & Environ$("WINSTATIONNAME")
         End If
         If Environ$("CLIENTNAME") <> "" Then
            strStaticMachineName = strStaticMachineName & "_" & Environ$("CLIENTNAME")
         End If
         If Environ$("SESSIONNAME") <> "" Then
            strStaticMachineName = strStaticMachineName & "_" & Environ$("SESSIONNAME")
         End If
         strStaticMachineName = Replace(Replace(Replace(Replace(strStaticMachineName, "*", "_"), "?", "_"), "#", "_"), "'", "")
         strTemp = ""
    End If
    
    MachineName = Right$(strStaticMachineName, 255)
End Function

Public Function TmpDir() As String
   'looks for default temp dir
    Dim strTmpPath As String
    Dim lngRet As Long
    Dim intFileNum As Integer
    
    Static strStaticTmpDir As String
    
    On Local Error GoTo TmpDirError
    
    If strStaticTmpDir = "" Then
        TmpDir = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Temp\"
        If Dir$(TmpDir, vbDirectory) = "" Then
            If Err > 0 Then
                MsgBox Translate("Cannot access", mcLanguage) & " " & TmpDir, vbCritical
                TmpDir = ""
            Else
                MkDir TmpDir
            End If
           
            intFileNum = FreeFile
            Open TmpDir & App.EXEName & ".$$$" For Output As #intFileNum
            Print #intFileNum, App.EXEName
            Close #intFileNum
            Kill TmpDir & App.EXEName & ".$$$"
        End If
        strStaticTmpDir = TmpDir
    Else
        TmpDir = strStaticTmpDir
    End If
    
TmpDirError:
    If Err > 0 Then
        MsgBox App.EXEName & " doesn't have " & " the right access to the directory for temporary files '" & TmpDir & "' for user " & UserName & "." & vbCrLf & "(" & Err.Description & ")", vbCritical
        End
    End If
End Function

Public Function PickDirFromTree(lHwnd As Long, Title As String) As String
    'Opens a form with a treeview with folders
    
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = Title
    
    With tBrowseInfo
        .hwndOwner = lHwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
        sBuffer = Space(MAX_strPath)
        SHGetPathFromIDList lpIDList, sBuffer
        If Right$(sBuffer, 1) <> "\" Then
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1) & "\"
        Else
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        End If
        PickDirFromTree = sBuffer
    End If

End Function

Public Function ResourceToDisk(ByVal intResourceNum As Integer, ByVal strResourceType As String, ByVal strDestFileName As String) As Long
    '=============================================
    'Saves a resource item to disk
    'Returns 0 on success, error number on failure
    '=============================================
    
    
    Dim bytResourceData()   As Byte
    Dim intFileNumOut         As Integer
    
    On Local Error GoTo ResourceToDisk_err
    
    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(intResourceNum, strResourceType)
    
    'Get Free File Handle
    intFileNumOut = FreeFile
    
    'Open the output file
    Open strDestFileName For Binary Access Write Shared As #intFileNumOut
        
    'Write the resource to the file
    Put #intFileNumOut, , bytResourceData
    
    'Close the file
    Close #intFileNumOut
    
    'Return 0 for success
    ResourceToDisk = 0
    
    Exit Function
    
ResourceToDisk_err:
    'Return error number
    ResourceToDisk = Err.Number
End Function


