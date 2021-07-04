VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Sub BordersAll(Rng As Range)
    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
End Sub

Sub BordersNone(Rng As Range)
    With Rng.Borders
        .LineStyle = xlNone
'        .Color = vbBlack
'        .Weight = xlThin
    End With
End Sub



