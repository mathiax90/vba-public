Attribute VB_Name = "OpenDialogHelperTest"
Option Explicit

Sub SelectFilesTest()
    Dim SelectedFiles As Object
    Dim i As Integer
    Dim OpenDialogHelper As New OpenDialogHelperClass
    If Not OpenDialogHelper.SelectFiles(SelectedFiles, Array("xlsx", "xls")) Then
        MsgBox "����� �� �������"
        Exit Sub
    End If
    
    For i = 1 To SelectedFiles.Count
        Debug.Print SelectedFiles.Item(i)
    Next i
End Sub

Sub SelectFileTest()
    Dim selectedFile As String
    Dim OpenDialogHelper As New OpenDialogHelperClass
    If Not OpenDialogHelper.SelectFile(selectedFile, Array("xlsx", "xls")) Then
        MsgBox "����� �� �������"
        Exit Sub
    End If
    Debug.Print selectedFile
End Sub

Sub SelectFolderTest()
    Dim selectedFolder As String
    Dim OpenDialogHelper As New OpenDialogHelperClass
    
    If Not OpenDialogHelper.SelectFolder(selectedFolder) Then
        MsgBox "����� �� �������"
        Exit Sub
    End If
Debug.Print selectedFolder

End Sub
    ��0                                                                                                      �����  � �                    ����    �!            �                                                                                                                            zC  zC�   �           @�D"3              �                                                                          /3                                                                              