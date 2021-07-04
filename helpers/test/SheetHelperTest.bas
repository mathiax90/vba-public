Attribute VB_Name = "SheetHelperTest"
Option Explicit
Sub SheetExistsTest()
    Dim Wb As Workbook
    Set Wb = ActiveWorkbook
    Dim ShHelper As New SheetHelper
    If ShHelper.SheetExists("Лист1", Wb) Then
        Debug.Print "Sheet exists"
    Else
        Debug.Print "Sheet doesn't exists"
    End If
End Sub
