Attribute VB_Name = "DateHelperTest"
Option Explicit
Sub MonthNameToMonthTest()
    Dim DateHelper As New DateHelperClass
    Dim mName As String
    Dim i As Integer
    For i = 1 To 12
        mName = monthName(i)
        Debug.Print mName
        Debug.Print DateHelper.MonthNameToMonth(mName)
    Next i
End Sub

