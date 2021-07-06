VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OpenDialogHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Function SelectFiles(ByRef SelectedFiles As Object, Optional fileType As Variant, Optional InitialDir As String = "c:\") As Boolean
    Dim ArrHelper As New ArrayHelper
    Dim Fd As Office.FileDialog
    Set Fd = Application.FileDialog(msoFileDialogFilePicker)
    With Fd
        .Filters.Clear
        If ArrHelper.Length(fileType) > 1 Then .Filters.Add "Тип файлов", "*." & Join(fileType, ", *.")
        .Title = "Choose an Excel file"
        .AllowMultiSelect = True
        .InitialFileName = InitialDir
        If .Show = True Then
            Set SelectedFiles = .SelectedItems
        End If
    End With
    
    If SelectedFiles Is Nothing Then
        SelectFiles = False
    Else
        SelectFiles = True
    End If
End Function

Function SelectFile(ByRef selectedFile As String, Optional fileType As Variant, Optional InitialDir As String = "c:\") As Boolean
    Dim ArrHelper As New ArrayHelper
    Dim Fd As Office.FileDialog
    Set Fd = Application.FileDialog(msoFileDialogFilePicker)
    With Fd
        .Filters.Clear
        If ArrHelper.Length(fileType) Then .Filters.Add "Тип файлов", "*." & Join(fileType, ", *.")
        .Title = "Choose an Excel file"
        .AllowMultiSelect = False
        .InitialFileName = InitialDir
        If .Show = True Then
            selectedFile = .SelectedItems(1)
        End If
    End With
    
    If IsEmpty(selectedFile) Or selectedFile = "" Then
        SelectFile = False
    Else
        SelectFile = True
    End If
End Function

Function SelectFolder(ByRef selectedFolder As String, Optional InitialDir As String = "c:\") As Boolean
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            selectedFolder = .SelectedItems(1)
            SelectFolder = True
        Else
            SelectFolder = False
        End If
    End With
End Function


