Attribute VB_Name = "StringHelperClassTest"
Option Explicit
'use StringHelperClass
Dim Helper As New StringHelperClass
Dim Errors As New Collection

Sub RunAllTests()
    Dim error As Variant
    Set Errors = New Collection
    
    CapitalizeTest
    CamelCaseBySymbolTest
    CamelCaseBySymbolTestWithPrefix
    PascalCaseBySymbolTest
    PascalCaseBySymbolTestWithPrefix
    StartsWithTest
    EndWith_test
    RepeatTest
    For Each error In Errors
        Debug.Print error
    Next
End Sub
Sub StartsWithTest()
    If Not Helper.StartsWith("asd", "as") Then
        Errors.Add ("StartsWithTest Error")
    End If
    
    If Helper.StartsWith("asd", "ad") Then
        Errors.Add ("StartsWithTest Error")
    End If
    
End Sub

Sub CapitalizeTest()
    Dim inval As String
    inval = "hello"
    If Helper.CapitalizeFirstLetter(inval) <> "Hello" Then
        Errors.Add ("CapitalizeTest Error")
    End If
End Sub

Sub CamelCaseBySymbolTest()
    Dim inval As String
    inval = "SCHET_SLUCH"
    If Helper.CamelCaseBySymbol(inval, "_") <> "schetSluch" Then
        Errors.Add ("CamelCaseBySymbolTest Error")
    End If
End Sub

Sub CamelCaseBySymbolTestWithPrefix()
    Dim inval As String
    inval = "T_SCHET_SLUCH"
    If Helper.CamelCaseBySymbol(inval, "_", "T_") <> "schetSluch" Then
        Errors.Add ("CamelCaseBySymbolTestWithPrefix Error")
    End If
End Sub

Sub PascalCaseBySymbolTest()
    Dim inval As String
    inval = "SCHET_SLUCH"
    'Debug.Print Helper.PascalCaseBySymbol(inval, "_")
    If Helper.PascalCaseBySymbol(inval, "_") <> "SchetSluch" Then
        Errors.Add ("PascalCaseBySymbolTest Error")
    End If
End Sub

Sub PascalCaseBySymbolTestWithPrefix()
    Dim inval As String
    inval = "T_SCHET_SLUCH"
    'Debug.Print Helper.PascalCaseBySymbol(inval, "_", "T_")
    If Helper.PascalCaseBySymbol(inval, "_", "T_") <> "SchetSluch" Then
        Errors.Add ("PascalCaseBySymbolTestWithPrefix Error")
    End If
End Sub

Sub Contains_test()
    Dim stack As String
    Dim needle As String
    
    stack = "123"
    needle = ""
        
    If Helper.Contains(stack, needle) Then
        Debug.Print stack & " contains " & needle
    Else
        Debug.Print "needle not found"
    End If
End Sub

Sub EndWith_test()
    Dim haystack As String
    Dim needle As String
    
    haystack = "prvetvet"
    needle = "vet"
      
    If Not Helper.EndWith(haystack, needle) Then Errors.Add "EndWith_test error"
    
End Sub


Sub RepeatTest()
    If Helper.Repeat("a", 5) = "aaaaa" Then
        Errors.Add "RepeatTest Pass"
    Else
        Errors.Add "RepeatTest Fail"
    End If
End Sub
