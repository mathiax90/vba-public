Attribute VB_Name = "RegexHelperTest"
Sub ReplaceReplaceTest()
    Dim str As String
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    
    str = "������  �������� ������ �� ��������      ������� � �������   ��������"
    Debug.Print RegexHelper.Replace(str, "\s{2,}", " ", True, True, True)
End Sub


Sub RemoveDoubleSpaceTest()
    Dim str As String
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    str = "������  �������� ������ �� ��������      ������� � �������   ��������"
    Debug.Print RegexHelper.RemoveDoubleSpace(str)
End Sub

Sub RemoveEmptyLinesTest()
    Dim str As String
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    str = "������" & Chr(13)
    str = str & " " & Chr(13)
    str = str & Chr(13)
    Debug.Print "|" & RegexHelper.RemoveEmptyLines(str) & "|"
End Sub
