Attribute VB_Name = "StyleHelperTest"
Option Explicit
Sub BordersAllTest()
    Dim StyleHelper As New StyleHelperClass
    StyleHelper.BordersAll ActiveSheet.Cells(1, 1)
    'Style.BordersNone ActiveSheet.Cells(1, 1)
End Sub

Sub BordersNoneTest()
    Dim StyleHelper As New StyleHelperClass
    StyleHelper.BordersNone ActiveSheet.Cells(1, 1)
    'Style.BordersNone ActiveSheet.Cells(1, 1)
End Sub
hn?        ?n?xn?        ?n??n?        ?n??n?        ?n??n?        ?n??n?        ?n??n?