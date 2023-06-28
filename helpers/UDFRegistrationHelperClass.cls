VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UDFRegistrationHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Name As String
Public Description As String
Public Category As String
Public Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10 As String

Public Sub Register()
Dim val As String
Dim ArgumentDesctiptions(1 To 10) As String

If Arg1 <> "" Then ArgumentDesctiptions(1) = Arg1
If Arg2 <> "" Then ArgumentDesctiptions(2) = Arg2
If Arg3 <> "" Then ArgumentDesctiptions(3) = Arg3
If Arg4 <> "" Then ArgumentDesctiptions(4) = Arg4
If Arg5 <> "" Then ArgumentDesctiptions(5) = Arg5
If Arg6 <> "" Then ArgumentDesctiptions(6) = Arg6
If Arg7 <> "" Then ArgumentDesctiptions(7) = Arg7
If Arg8 <> "" Then ArgumentDesctiptions(8) = Arg8
If Arg9 <> "" Then ArgumentDesctiptions(9) = Arg9
If Arg10 <> "" Then ArgumentDesctiptions(10) = Arg10


Application.MacroOptions _
  Macro:=Name, _
  Description:=Description, _
  Category:=Category, _
  ArgumentDescriptions:=ArgumentDesctiptions
End Sub

