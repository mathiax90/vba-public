Attribute VB_Name = "ArrayHelperTest"
Option Explicit
Sub LengthTest()
    Dim testArray(0 To 3) As Long
    Dim ArrayHelper As New ArrayHelperClass
    Debug.Print ArrayHelper.Length(testArray)
End Sub
