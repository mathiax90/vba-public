Attribute VB_Name = "WinRegHelperClassTest"
Option Explicit
Dim WinRegHelper As New WinRegHelperClass
Dim Errors As New Collection

Sub RunAllTests()
    Dim error As Variant
    Set Errors = New Collection
    WinRegHelper.AppGuid = "WinRegHelperClassTest"
    
    Call SetAppOptionTest
    Call GetAppOptionTest
    
    For Each error In Errors
        Debug.Print error
    Next
End Sub

Sub SetAppOptionTest()
    WinRegHelper.SetAppOption "option1", "option1 value"
End Sub

Sub GetAppOptionTest()
    Dim option1 As String
    option1 = WinRegHelper.GetAppOption("option1")
    If option1 <> "option1 value" Then Errors.Add "GetAppOptionTest error"
End Sub

