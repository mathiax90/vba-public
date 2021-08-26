Attribute VB_Name = "RegexHelperTest"
Sub ReplaceTest()
    Debug.Print "ReplaceTest"
    
    Dim str As String
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    
    str = "������  �������� ������ �� ��������      ������� � �������   ��������"
    Debug.Print RegexHelper.Replace(str, "\s{2,}", " ", True, True, True)
End Sub


Sub RemoveDoubleSpaceTest()
    Debug.Print "RemoveDoubleSpaceTest"
    Dim str As String
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    str = "������  �������� ������ �� ��������      ������� � �������   ��������"
    Debug.Print RegexHelper.RemoveDoubleSpace(str)
End Sub

Sub RemoveEmptyLinesTest()
    Debug.Print "RemoveEmptyLinesTest"
    Dim str As String
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    str = "������" & Chr(13)
    str = str & " " & Chr(13)
    str = str & Chr(13)
    Debug.Print "|" & RegexHelper.RemoveEmptyLines(str) & "|"
End Sub

Sub HasFirstMatchTest()
    Debug.Print "HasFirstMatchTest"
    Dim str As String
    Dim pattern As String
    
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    str = "21 ���. 00 ���. ��� 33 ���. 40 ���."
    pattern = "\d{1,}\D{0,}���\.\D{0,}\d{1,}\D{0,}���\."
    
    Dim match As String
    
    If RegexHelper.HasFirstMatch(str, pattern, match) Then
        Debug.Print "|" & match & "|"
    Else
        Debug.Print "no match"
    End If
End Sub

Sub HasMatchTest()
    Debug.Print "HasMatchTest"

    Dim str As String
    Dim pattern As String
    
    Dim RegexHelper As RegexHelperClass
    Set RegexHelper = New RegexHelperClass
    str = "21 ���. 00 ���. ��� 33 ���. 40 ���."
    pattern = "\d{1,}\D{0,}���\.\D{0,}\d{1,}\D{0,}���\."
    
    Dim match As Object
    
    If RegexHelper.HasMatch(str, pattern, match) Then
        For i = 0 To match.Count - 1
            Debug.Print match(i)
        Next i
    Else
        Debug.Print "no match"
    End If
End Sub


