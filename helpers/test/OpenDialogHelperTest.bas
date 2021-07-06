Attribute VB_Name = "OpenDialogHelperTest"
Option Explicit

Sub SelectFilesTest()
    Dim SelectedFiles As Object
    Dim i As Integer
    Dim OpenDialog As New OpenDialogHelper
    If Not OpenDialog.SelectFiles(SelectedFiles, Array("xlsx", "xls")) Then
        MsgBox "Файлы не выбраны"
        Exit Sub
    End If
    
    For i = 1 To SelectedFiles.Count
        Debug.Print SelectedFiles.Item(i)
    Next i
End Sub

Sub SelectFileTest()
    Dim selectedFile As String
    Dim OpenDialog As New OpenDialogHelper
    If Not OpenDialog.SelectFile(selectedFile, Array("xlsx", "xls")) Then
        MsgBox "Файлы не выбраны"
        Exit Sub
    End If
    Debug.Print selectedFile
End Sub

Sub SelectFolderTest()
    Dim selectedFolder As String
    Dim OpenDialog As New OpenDialogHelper
    
    If Not OpenDialog.SelectFolder(selectedFolder) Then
        MsgBox "Файлы не выбраны"
        Exit Sub
    End If
Debug.Print selectedFolder

End Sub
