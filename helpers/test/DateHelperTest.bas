Attribute VB_Name = "DateHelperTest"
Option Explicit
Sub MonthNameToMonthTest()
    Dim DtHelper As New DateHelper
    Dim mName As String
    Dim i As Integer
    For i = 1 To 12
        mName = monthName(i)
        Debug.Print mName
        Debug.Print DtHelper.MonthNameToMonth(mName)
    Next i
End Sub

