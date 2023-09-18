VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function length(a As Variant) As Long
 If IsEmpty(a) Then
      length = 0
   Else
      length = UBound(a) - LBound(a) + 1
   End If
End Function


Public Function Contains(CellRef As String, needle As String, Optional CaseSensetive As Boolean = True) As Boolean
    If CaseSensetive Then
        CellRef = LCase(CellRef)
        needle = LCase(needle)
    End If
    If needle = "" Then
        Err.Raise number:=vbObjectError + 513, _
            Description:="Needle can't be empty string"
    End If
    Dim index As Integer
    index = InStr(CellRef, needle)
    If index > 0 Then
        Contains = True
    Else
        Contains = False
    End If
End Function
