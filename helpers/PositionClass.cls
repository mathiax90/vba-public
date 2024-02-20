VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PositionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Top As Double
Public Left As Double

Public Function CreatePosition(t As Double, l As Double) As PositionClass
    Dim s As New PositionClass
    s.Top = t
    s.Left = l
    Set CreatePosition = s
End Function

