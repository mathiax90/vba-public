VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinRegHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'props
Public AppGuid As String
'const
Const DefaultSection As String = "HKEY_CURRENT_USER\Software\Microsoft\Office\MacroOptions\"
'helpers and common objs
Private WS As New WshShell
Private StringHelper As New StringHelperClass
Private PathHelper As New PathHelperClass

Public Sub SetAppOption(key As String, value As String)
    Call SetOption(AppGuid, key, value)
End Sub

Function GetAppOption(key As String)
    GetAppOption = GetOption(AppGuid, key)
End Function

Public Sub SetOption(appPath As String, key As String, value As String)
    Dim path As String
    path = PathHelper.Combine(DefaultSection, appPath)
    path = PathHelper.Combine(path, key)
    WS.RegWrite path, value, "REG_SZ"
End Sub

Function GetOption(appPath As String, key As String) As String
    Dim path As String
    path = PathHelper.Combine(DefaultSection, appPath)
    path = PathHelper.Combine(path, key)
    GetOption = WS.RegRead(path)
    Exit Function
exceptionHandler:
    GetOption = ""
End Function
