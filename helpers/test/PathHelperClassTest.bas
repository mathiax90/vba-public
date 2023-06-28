Attribute VB_Name = "PathHelperClassTest"
Option Explicit
'use StringHelperClass
Dim Helper As New PathHelperClass
Dim Errors As New Collection

Sub RunAllTests()
    Dim error As Variant
    Set Errors = New Collection
       
    Call CombineTest
    
    For Each error In Errors
        Debug.Print error
    Next
End Sub
Sub CombineTest()
    Dim path1 As String, path2 As String
    path1 = "\\c:\dir1\\"
    path2 = "\\dir2\\"
    Debug.Print Helper.Combine(path1, path2)
End Sub

Sub NormalizeTest()
    Dim path1 As String, path2 As String
    path1 = "\\c:\dir1\\"
    
    Debug.Print Helper.Normalize(path1, "\")
End Sub
