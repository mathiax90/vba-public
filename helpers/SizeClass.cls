VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SizeClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Width As Double
Public Height As Double

Public Function CreateSize(h As Double, w As Double) As SizeClass
    Dim s As New SizeClass
    s.Width = w
    s.Height = h
    Set CreateSize = s
End Function

