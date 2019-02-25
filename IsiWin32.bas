Attribute VB_Name = "modWin32"
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

'code for Win32 file operations:
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS, sets dialog title
End Type

Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_APPDATA = &H1A

Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'GetWindow and SetWindowLong are required to display a floating windows that stays always on top:
Public Declare Function GetWindow& Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long)
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Const CP_ACP = 0
Public Const CP_UTF8 = 65001
Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&
Public Const FO_COPY = &H2 ' Copy File/Folder
Public Const FO_DELETE = &H3 ' Delete File/Folder
Public Const FO_MOVE = &H1 ' Move File/Folder
Public Const FO_RENAME = &H4 ' Rename File/Folder
Public Const FOF_ALLOWUNDO = &H40 ' Allow to undo rename, delete ie sends to recycle bin
Public Const FOF_FILESONLY = &H80  ' Only allow files
Public Const FOF_NOCONFIRMATION = &H10  ' No File Delete or Overwrite Confirmation Dialog
Public Const FOF_SILENT = &H4 ' No copy/move dialog
Public Const FOF_SIMPLEPROGRESS = &H100 ' Does not display file names
Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)

' Reg Key ROOT Types...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2
Public Const REG_DWORD = 4                      ' 32-bit number

Public Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Public Const gREGVALSYSINFOLOC = "MSINFO"
Public Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Public Const gREGVALSYSINFO = "PATH"


Public Const LOCALE_ILANGUAGE             As Long = &H1    'language id
Public Const LOCALE_SLANGUAGE             As Long = &H2    'localized name of lang
Public Const LOCALE_SENGLANGUAGE          As Long = &H1001 'English name of lang
Public Const LOCALE_SABBREVLANGNAME       As Long = &H3    'abbreviated lang name
Public Const LOCALE_SNATIVELANGNAME       As Long = &H4    'native name of lang
Public Const LOCALE_ICOUNTRY              As Long = &H5    'country code
Public Const LOCALE_SCOUNTRY              As Long = &H6    'localized name of country
Public Const LOCALE_SENGCOUNTRY           As Long = &H1002 'English name of country
Public Const LOCALE_SABBREVCTRYNAME       As Long = &H7    'abbreviated country name
Public Const LOCALE_SNATIVECTRYNAME       As Long = &H8    'native name of country
Public Const LOCALE_SINTLSYMBOL           As Long = &H15   'intl monetary symbol
Public Const LOCALE_IDEFAULTLANGUAGE      As Long = &H9    'def language id
Public Const LOCALE_IDEFAULTCOUNTRY       As Long = &HA    'def country code
Public Const LOCALE_IDEFAULTCODEPAGE      As Long = &HB    'def oem code page
Public Const LOCALE_IDEFAULTANSICODEPAGE  As Long = &H1004 'def ansi code page
Public Const LOCALE_IDEFAULTMACCODEPAGE   As Long = &H1011 'def mac code page
Public Const LOCALE_IMEASURE              As Long = &HD     '0 = metric, 1 = US

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
'#if(WINVER >=  &H0400)
Public Const LOCALE_SISO639LANGNAME       As Long = &H59   'ISO abbreviated language name
Public Const LOCALE_SISO3166CTRYNAME      As Long = &H5A   'ISO abbreviated country name
'#endif /* WINVER >= as long = &H0400 */

'#if(WINVER >=  &H0500)
Public Const LOCALE_SNATIVECURRNAME        As Long = &H1008 'native name of currency
Public Const LOCALE_IDEFAULTEBCDICCODEPAGE As Long = &H1012 'default ebcdic code page
Public Const LOCALE_SSORTNAME              As Long = &H1013 'sort name
'#endif /* WINVER >=  &H0500 */

Public Declare Function GetThreadLocale Lib "kernel32" () As Long
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
Public Function GetSpecialFolderLocation(CSIDL As Long) As String
Const S_OK = 0
Const MAX_PATH = 260
Dim spath As String
Dim pidl As Long

'fill the idl structure with the specified folder item
If SHGetSpecialFolderLocation(0, CSIDL, pidl) = S_OK Then

'if the pidl is returned, initialize
'and get the path from the id list
spath = Space$(MAX_PATH)

If SHGetPathFromIDList(ByVal pidl, ByVal spath) Then
    'return the path
    If Len(spath) Then GetSpecialFolderLocation = Left(spath, InStr(spath, Chr$(0)) - 1)
End If
'free the pidl
CoTaskMemFree pidl
End If

End Function
Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function
Public Function WindowIsOpen(pstrWindow As String) As Boolean
    WindowIsOpen = (FindWindow(vbNullString, pstrWindow) <> 0)
End Function
Public Sub GetDrives(strDrives As String, strDriveType As String)
   'lists available drive letters by type
    Dim lngTemp As Long
    Dim intTemp As Integer
    Dim strTemp As String
    Dim strTemp2 As String
    
    strTemp = Space$(255)
    lngTemp = GetLogicalDriveStrings(255, strTemp)
    strTemp = Trim$(strTemp)
    
    'GetDriveType results in drive types:
    '0 = unknown
    '1 = FD
    '2 = removable
    '3 = HD
    '4 = netwerk
    '5 = CD
    '6 = ram
    
    strDrives = ""
    If strDriveType <> "" And strDriveType <> "0" Then
        Do While strTemp <> ""
            intTemp = InStr(strTemp, Chr$(0))
            If intTemp = 0 Then
                strTemp2 = strTemp
                strTemp = ""
            Else
                strTemp2 = Left$(strTemp, intTemp - 1)
                strTemp = Trim$(Mid$(strTemp & " ", intTemp + 1))
            End If
                
            If InStr(strDriveType, Format$(GetDriveType(strTemp2))) Then
                strDrives = strDrives & strTemp2 & " "
            End If
        Loop
        strDrives = Trim$(strDrives)
    Else
        strDrives = Replace$(strTemp, Chr$(0), " ")
    End If
End Sub
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function



