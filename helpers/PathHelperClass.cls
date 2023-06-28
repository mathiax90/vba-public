VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private StringHelper As New StringHelperClass
Const LoopLimit As Integer = 20
Function Combine(path1 As String, path2 As String, Optional ByVal separator As String = "\")
    Combine = Normalize(path1, separator) & separator & Normalize(path2, separator)
End Function

'убирает разделитель в начале и конце пути
Public Function Normalize(path As String, separator As String)
    Dim res As String
    Dim i As Integer
    res = path
        
    i = 1
    Do
        If i > LoopLimit Then Exit Do
        If StringHelper.EndWith(res, separator) Then
            res = Left(res, Len(res) - Len(separator))
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    
    i = 1
    Do
        If i > LoopLimit Then Exit Do
        If StringHelper.StartsWith(res, separator) Then
            res = Right(res, Len(res) - Len(separator))
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    
    Normalize = res
End Function
