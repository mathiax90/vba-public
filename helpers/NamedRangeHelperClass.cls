VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NamedRangeHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Function RangeExists(rangeName As String, Optional sh As Worksheet) As Boolean
    Dim rng As Range
    On Error Resume Next
    
    If sh Is Nothing Then
        Set rng = Range(rangeName)
    Else
        Set rng = sh.Range(rangeName)
    End If
        
    RangeExists = Err.number = 0
End Function
