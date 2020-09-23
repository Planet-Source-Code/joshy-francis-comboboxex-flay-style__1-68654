VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXPStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function ActivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" (ByVal hwnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Private Declare Function DeactivateWindowTheme Lib "uxtheme" Alias "SetWindowTheme" (ByVal hwnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long

Sub UpdateManiFest()
Dim ManiFest As String
    ManiFest = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & _
    Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & _
    "?><assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & _
    Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & _
    "><assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & _
    " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & _
    "Template.WindowsXP.Theme" & Chr(34) & " type=" & Chr(34) & "win32" & _
    Chr(34) & " /> <description>For use with VB6.</description> <dependency>" & _
    " <dependentAssembly> <assemblyIdentity type=" & Chr(34) & "win32" & _
    Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & _
    Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & _
    " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & _
    " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & _
    " language=" & Chr(34) & "*" & Chr(34) & " /> </dependentAssembly>" & _
    " </dependency> </assembly>"
Dim sPath As String, Pid As Long
    sPath = App.Path
        If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    sPath = sPath & App.EXEName & ".exe.manifest"
        If IsFile(sPath) = False Then
            Open sPath For Output As 1
                Print #1, ManiFest
            Close 1
                SetAttr sPath, vbSystem + vbHidden
        End If
ManiFest = ""
End Sub
Public Function IsDir(d) As Boolean
On Error GoTo E
        IsDir = False
    RmDir d
        MkDir d
        IsDir = True
Exit Function
E:
'    Debug.Print Err.Description
    If Err.Description = "Path not found" Then
        IsDir = False
    Else
        IsDir = True
    End If
End Function
Public Function IsFile(F) As Boolean
On Error GoTo E
    IsFile = False
        If InStr(F, ":") = 0 Then
            Exit Function
        End If
        If InStr(F, "\") = 0 Then
            Exit Function
        End If
Dim N As Integer
    N = FreeFile
        Open F For Input As N
        Close N
        Reset
    IsFile = True
Exit Function
E:
    Reset
IsFile = False
End Function

Sub EnableTheme(ByVal hwnd As Long)
    Call ActivateWindowTheme(hwnd)
End Sub

Sub DisableTheme(ByVal hwnd As Long)
    Call DeactivateWindowTheme(hwnd)
End Sub

Sub InitTheme()
    UpdateManiFest
Call InitCommonControls
End Sub
