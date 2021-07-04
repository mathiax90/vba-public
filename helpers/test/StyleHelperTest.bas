Attribute VB_Name = "StyleHelperTest"
Option Explicit
Sub BordersAllTest()
    Dim Style As New StyleHelper
    Style.BordersAll ActiveSheet.Cells(1, 1)
    'Style.BordersNone ActiveSheet.Cells(1, 1)
End Sub
