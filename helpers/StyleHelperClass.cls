VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StyleHelperClass"
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



���
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
  RegExpFindReplace = RegExp.Replace(str, Replace)
  Set RegExp = N          