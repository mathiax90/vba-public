VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartHelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Sub SetPosAndSize(chrt As chart, pos As PositionClass, s As SizeClass, retryCount As Integer)

RetryChartDimensions:
    On Error GoTo ErrChartDimensions
    With chrt.PlotArea
        .Width = s.Width
        .Height = s.Height
        .Left = pos.Left
        .Top = pos.Top
    End With
    On Error GoTo 0
    
    ' More code
    ' more code
    
    Exit Sub 'gracefully exit this subroutine before the error-handling.'
    
ErrChartDimensions:
    Err.Clear
    'Pause before setting the PlotArea dimensions:'
    Dim wtTime As Double
    Dim startTime As Long
    
    'A maximum 3 second delay should be more than enough time.
    If wtTime < retryCount Then
        wtTime = wtTime + 0.5
        startTime = Timer
        While Timer < startTime + wtTime
            DoEvents
        Wend
    End If
    Resume RetryChartDimensions

End Sub
