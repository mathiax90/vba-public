Attribute VB_Name = "DateHelperTest"
Option Explicit
Dim DateHelper As New DateHelperClass
Sub MonthNameToMonthTest()
    Dim mName As String
    Dim i As Integer
    For i = 1 To 12
        mName = monthName(i)
        Debug.Print mName
        Debug.Print DateHelper.MonthNameToMonth(mName)
    Next i
End Sub

Sub FormatYYYYMMDD()
    Dim dateStr As String
    dateStr = "2020-01-31"
    Debug.Assert DateHelper.FormatYYYYMMDD(dateStr) = "31.01.2020"
End Sub
