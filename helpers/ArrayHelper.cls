VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function Length(a As Variant) As Long
 If IsEmpty(a) Then
      Length = 0
   Else
      Length = UBound(a) - LBound(a) + 1
   End If
End Function
