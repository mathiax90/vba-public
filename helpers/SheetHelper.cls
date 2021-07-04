VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Function SheetExists(shName As String, Wb As Workbook) As Boolean
    Dim Sh As Worksheet
    On Error Resume Next
    Set Sh = Wb.Sheets(shName)
    On Error GoTo 0
    If Not Sh Is Nothing Then
        SheetExists = True
    End If
End Function
