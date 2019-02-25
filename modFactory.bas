Attribute VB_Name = "modFactory"
'This module (modFactory) is basically an Instantiation-
'helper for all the RichClient3-Classes, since it delivers
'small helper-objects from cFactory for usage in your App.
'More than that it can also ensure regfree-loading
'of the Toolset-Classes over DirectCOM.dll-GetInstance().

'If the functionality of this module is not enough
'for your purposes - e.g. if you want more flexibility
'with the Toolset-Paths, then please take a look at
'Ulrich Korndoerfers enhanced Loader-Module, which offers
'a few more "degrees of freedom" than this module here.
'Ulrichs module is called modRC3.bas and is included
'in the original Toolset-Folder. Please read his
'extensive comments in modRC3.bas, on how to use
'his module.

'Ok, what do we want to achieve here... once we have
'created the small cFactory-Object (possible also regfree),
'we can instantiate all the other RichClient-Classes
'with its help, so the GetInstance-Call (which is the
'"real workhorse" for the regfree-loading) is usually only
'needed once against cFactory (though one can of course also
'instantiate the Factory "normally" with VBs 'New'-Operator).
'This small "FrameWork-EntryPoint" currently contains only two
'small "Sub-Members" (Sub-Objects). One is the cConstructor-
'Object, which "knows" and supports instantiation of all
'RichClient3-Classes and also offers optional Init-
'Parameters for most of them - the other Object-Member in
'cFactory is the regfree-Object, which offers the same
'GetInstanceEx-Call as we use here "directly" in this module,
'as a regfree.GetInstanceEx-Method, usable then per Intellisense
'throughout your App, to e.g. instantiate other COM-Objects
'from your own AX-Dlls regfree - but the regfree-Object also
'includes methods, which allow you to create and manage
'COM-Objects on their own threads - and that also in a
'regfree manner.

'But as already said, the Factory-object can be instantiated
'also from a normally registered RichClient - in either case
'the contained cConstructor-object (exposed by a Public
'Property New_c here) should be useful in your App, since
'it offers the already mentioned Optional Init-Params...

'So instead of writing e.g.:

'Dim Cnn As cConnection 'define a SQLite-Connection
'Set Cnn = New cConnection 'normal instantiation per 'New'
'Cnn.OpenDB "c:\SomeFolder\Some.db"

'you could write using the cConstructor (callable per New_c):

'Dim Cnn As cConnection 'define a SQLite-Connection
'Set Cnn = New_c.Connection("c:\SomeFolder\Some.db")

'and save some lines of code this way, but also keep
'your App prepared for regfree-loading, without replacing
'too many "normal" VB-instantiation-calls of the
'RichClient3-Classes later on, in case you are not
'making use of this module (or the New_c - Constructor)
'from the "very beginning".


'------------------- Implementation ----------------------
Option Explicit

'Specifiy your Path for Regfree-Loading here in the following Const.
'As long as you start with only a Backslash, then the App.Path is used
'and e.g. a "\Bin\" would then resolve to App.Path\Bin\ and looks
'for the Toolset-Dlls there - but you are also free, to specify an
'absolute path as e.g. "\\NetworkShare\RichCient3\" or "c:\RichClient3\"
'if you prefer to use it this way which could be useful for testing...
Private Const ToolSetPath$ = "\" 'a single Backslash resolves to App.Path

'this is basically a replacement to CreateObject(), but working regfree
'and as said above, the same Call is also available from inside the
'"Object-Model" per regfree.GetInstance(), once you have created the
'EntryPoint-object of the Framework (the cFactory-Instance).
'We therefore declare it here just as 'Private'
Private Declare Function GetInstanceEx Lib "DirectCom" _
                        (StrPtr_FName As Long, StrPtr_ClassName As Long, _
                         Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
Private Declare Function GetInstanceOld Lib "DirectCom" Alias "GETINSTANCE" _
                        (FName As String, ClassName As String) As Object
Private Declare Function GETINSTANCELASTERROR Lib "DirectCom" () As String

'used, to preload DirectCOM.dll from a given Path, before we try our calls
Private Declare Function LoadLibraryW Lib "kernel32.dll" _
                        (ByVal LibFilePath As Long) As Long


'deliver the constructor-helper from the factory
Public Property Get New_c() As cConstructor
  Set New_c = dhF.C
End Property

'deliver the regfree-"namespace" from the factory
Public Property Get regfree() As cRegFree
  Set regfree = dhF.regfree
End Property

Public Property Get dhF() As cFactory
Static F As cFactory, hLib As Long
  'if we already have an instance, we pass it and return immediately
  If Not F Is Nothing Then
    Set dhF = F
    Exit Property
  End If
  
  On Error GoTo ErrHandler
  
  If RunningInIde Then
    Set F = New cFactory '"normal" instancing, using VBs 'New'-Operator
    Set dhF = F
  Else 'try regfree factory-instancing
    Dim RegFreePath As String
    RegFreePath = ToolSetPath
    If Left$(RegFreePath, 1) = "\" And Left$(RegFreePath, 2) <> "\\" Then
      RegFreePath = App.Path & RegFreePath 'use expansion to the App.Path
    End If
    If Right$(RegFreePath, 1) <> "\" Then RegFreePath = RegFreePath & "\"

    If hLib = 0 Then
      If Not FileExists(RegFreePath & "DirectCOM.dll") Then
        Err.Raise vbObjectError, , "DirectCOM.dll not found in Folder:" _
                  & vbCrLf & RegFreePath
      End If
      hLib = LoadLibraryW(StrPtr(RegFreePath & "DirectCOM.dll")) '<-preload
    End If
    Set F = GetInstance(RegFreePath & "dhRichClient3.dll", "cFactory", True)
    Set dhF = F
  End If
Exit Property

ErrHandler:
  If MsgBox(Err.Description, vbYesNo, "Make an attempt with a registered Factory?") = vbYes Then
    Set F = New cFactory '"normal" instancing, using VBs 'New'-Operator
    Set dhF = F
  End If
End Property

Private Function RunningInIde() As Boolean
Static Done As Boolean, Result As Boolean
  If Not Done Then
    Done = True
    On Error Resume Next
    Debug.Print 1 / 0
    Result = Err: Err.Clear
  End If
  RunningInIde = Result
End Function

Private Function FileExists(ByRef FileName As String) As Boolean
  On Error Resume Next
    FileExists = ((GetAttr(FileName) And vbDirectory) <> vbDirectory)
  Err.Clear
End Function

'The new GetInstance-Wrapper-Proc, which is using the new DirectCOM.dll (March 2009 and newer)
'with the new Unicode-capable GetInstanceEx-Call (which now supports the AlteredSearchPath-Flag as well) -
'If you omit that optional param or set it to True, then LoadLibraryExW is used with the appropriate
'Flag. If the Param was set to False, then the behaviour is the same as with the former
'DirectCOM.dll-GETINSTANCE-Call - only that LoadLibraryW is used instead of LoadLibraryA.
'This routine also tries a fallback to the former DirectCOM.dll-GETINSTANCE-Call, in case
'you are using it against an older version of this small regfree-helper-lib.
Private Function GetInstance(DllFileName As String, ClassName As String, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
  On Error Resume Next
    Set GetInstance = GetInstanceEx(StrPtr(DllFileName), StrPtr(ClassName), UseAlteredSearchPath)
  If Err.Number = 453 Then 'GetInstanceEx not available, probably an older DirectCOM.dll...
    Err.Clear
    Set GetInstance = GetInstanceOld(DllFileName, ClassName) 'so let's try the older GETINSTANCE-call
  End If
  If Err Then
    Dim Error As String
    Error = Err.Description
    On Error GoTo 0: Err.Raise vbObjectError, , Error
  Else
    If GetInstance Is Nothing Then
      On Error GoTo 0: Err.Raise vbObjectError, , GETINSTANCELASTERROR()
    End If
  End If
End Function

