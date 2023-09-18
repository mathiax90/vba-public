VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ArrayHelper As New ArrayHelperClass

Public Function StartsWith(haystack As String, needle As String) As Boolean
    Dim index As Long
    index = InStr(haystack, needle)
    If index = 1 Then
        StartsWith = True
    Else
        StartsWith = False
    End If
End Function

Public Function EndWith(haystack As String, needle As String) As Boolean
    Dim index As Long
    index = InStrRev(haystack, needle)
    If Len(haystack) + 1 - index - Len(needle) = 0 Then
        EndWith = True
    Else
        EndWith = False
    End If
End Function

Public Function Contains(CellRef As String, needle As String, Optional CaseSensetive As Boolean = True) As Boolean
    If CaseSensetive Then
        CellRef = LCase(CellRef)
        needle = LCase(needle)
    End If
    If needle = "" Then
        Err.Raise number:=vbObjectError + 513, _
            Description:="Needle can't be empty string"
    End If
    Dim index As Integer
    index = InStr(CellRef, needle)
    If index > 0 Then
        Contains = True
    Else
        Contains = False
    End If
End Function

Public Function CapitalizeFirstLetter(CellRef As String)
    CellRef = Trim(CellRef)
    If Len(CellRef) <> 0 Then
        CapitalizeFirstLetter = UCase(Left(CellRef, 1)) & LCase(Right(CellRef, Len(CellRef) - 1))
    Else
        CapitalizeFirstLetter = ""
    End If
End Function

Public Function CamelCaseBySymbol(CellRef As String, Symbol As String, Optional Prefix As String = "")
    Dim arraySplit() As String
    Dim val, outval As String
    Dim i As Integer
    
    CellRef = Trim(CellRef)
    If Trim(Prefix) <> "" Then
        CellRef = Replace(CellRef, Trim(Prefix), "", 1, 1)
    End If
    
    
    arraySplit = Split(CellRef, Symbol)
    outval = LCase(arraySplit(0))
    
    For i = 1 To ArrayHelper.length(arraySplit) - 1
        outval = outval & CapitalizeFirstLetter(Trim(arraySplit(i)))
    Next
        
    CamelCaseBySymbol = outval
End Function

Public Function PascalCaseBySymbol(CellRef As String, Symbol As String, Optional Prefix As String = "")
    Dim arraySplit() As String
    Dim val, outval As String
    Dim i As Integer
    
    CellRef = Trim(CellRef)
    If Trim(Prefix) <> "" Then
        CellRef = Replace(CellRef, Trim(Prefix), "", 1, 1)
    End If
    
    arraySplit = Split(CellRef, Symbol)
    
    For i = 0 To ArrayHelper.length(arraySplit) - 1
        outval = outval & CapitalizeFirstLetter(Trim(arraySplit(i)))
    Next
        
    PascalCaseBySymbol = outval
End Function


Public Function IsNullOrEmpty(value As String)
    IsNullOrEmpty = (Trim(value & vbNullString) = vbNullString)
End Function

Public Function Repeat(value As String, number As Integer)
    Dim i As Integer
    
    For i = 1 To number
        Repeat = Repeat & value
    Next
End Function

Public Function Coalesce(ParamArray Fields() As Variant) As Variant

    Dim v As Variant

    For Each v In Fields
        If "" & v <> "" Then
            Coalesce = v
            Exit Function
        End If
    Next
    Coalesce = ""

End Function

Public Function JoinStringCollection(c As Collection, delim As String)
    Dim item As Variant
    Dim length As Long
    Dim r As String
    
    For Each item In c
        If Trim(CStr(item)) <> "" Then
            r = Trim(r) & CStr(item) & delim
        End If
    Next
        If Len(r) > 0 Then r = Left(r, Len(r) - 1)
    JoinStringCollection = r
End Function

