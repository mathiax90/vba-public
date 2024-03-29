VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegexHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'����� �� ����������� ��������� � ������.
Public Function Replace(str As String, _
       Pattern As String, _
       Replacement As String, _
       Optional IsGlobal As Boolean = True, _
       Optional IgnoreCase As Boolean = False, _
       Optional Multiline As Boolean = False) _
       As String
       
 '���� ������ �� ������
 If Not str Like "" And Not Pattern Like "" Then

  Dim RegExp As Object '��� ����������� ���������
  Set RegExp = CreateObject("VBScript.RegExp")

  With RegExp
   '��� ���������� ��� ������ ������?
   .Global = IsGlobal
   .IgnoreCase = IgnoreCase '������� �������?
   '������������ �������� �����?
   .Multiline = Multiline
   .Pattern = Pattern '���������
  End With

  '�����/��������
  On Error Resume Next
  Replace = RegExp.Replace(str, Replacement)
  Set RegExp = Nothing '������� ������

 End If
End Function

'����� �� ����������� ��������� � ������.
Public Function HasFirstMatch(str As String, _
       Pattern As String, _
       ByRef FirstMatch As String, _
       Optional IsGlobal As Boolean = True, _
       Optional IgnoreCase As Boolean = False, _
       Optional Multiline As Boolean = False) _
       As Boolean
    HasFirstMatch = False
    '���� ������ �� ������
    If Not str Like "" And Not Pattern Like "" Then
    
        Dim RegExp As Object '��� ����������� ���������
        Set RegExp = CreateObject("VBScript.RegExp")
        
        With RegExp
        '��� ���������� ��� ������ ������?
        .Global = IsGlobal
        .IgnoreCase = IgnoreCase '������� �������?
        '������������ �������� �����?
        .Multiline = Multiline
        .Pattern = Pattern '���������
        End With
        
        '�����/��������
        On Error Resume Next
        Set Matches = RegExp.Execute(str)
        If Matches.Count >= 1 Then
            HasFirstMatch = True
            FirstMatch = CStr(Matches(0))
        End If
        Set RegExp = Nothing '������� ������
    End If
End Function

'����� �� ����������� ��������� � ������.
Public Function HasMatch(str As String, _
       Pattern As String, _
       ByRef Matches As Object, _
       Optional IsGlobal As Boolean = True, _
       Optional IgnoreCase As Boolean = False, _
       Optional Multiline As Boolean = False) _
       As Boolean
    HasMatch = False
    '���� ������ �� ������
    
    If Not str Like "" And Not Pattern Like "" Then
    
        Dim RegExp As Object '��� ����������� ���������
        Set RegExp = CreateObject("VBScript.RegExp")
        
        With RegExp
        '��� ���������� ��� ������ ������?
        .Global = IsGlobal
        .IgnoreCase = IgnoreCase '������� �������?
        '������������ �������� �����?
        .Multiline = Multiline
        .Pattern = Pattern '���������
        End With
        
        '�����/��������
        On Error Resume Next
        Set Matches = RegExp.Execute(str)
        If Matches.Count >= 1 Then
            HasMatch = True
        End If
        Set RegExp = Nothing '������� ������
    
    End If
End Function

Function RemoveDoubleSpace(str As String) As String
    RemoveDoubleSpace = Replace(str, "\s{2,}", " ", True, True, True)
End Function

Function RemoveEmptyLines(str As String) As String
    RemoveEmptyLines = Replace(str, "\s{0,}((\r\n){2,}|\n{2,}|\r{2,})", Chr(13), True, True, True)
End Function

Function ReplaceNewLineSymbolWithChr13(str As String) As String
    ReplaceNewLineSymbolWithChr13 = Replace(str, "(\r\n|\n|\n\r)", Chr(13), True, True, True)
End Function

