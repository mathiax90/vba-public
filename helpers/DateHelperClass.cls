VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Function MonthNameToMonth(monthName As String) As Integer
    On Error GoTo Exception
    MonthNameToMonth = Format(CDate(monthName & "/1 0"), "m") * 1&
    On Error GoTo 0
    Exit Function
Exception:
    MonthNameToMonth = 0
    Debug.Print "DateHelper.MonthNameToMonth: Error while parse monthName to monthNumber"
End Function


Function FormatYYYYMMDD(inval As String, Optional fmt As String) As String
    If fmt = "" Then fmt = "dd.mm.yyyy"
    
    inval = Left(Trim(inval), 10)
    If inval <> "" Then
        FormatYYYYMMDD = Format(CDate(inval), fmt)
    Else
        FormatYYYYMMDD = ""
    End If
End Function
