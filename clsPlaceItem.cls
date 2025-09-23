Option Explicit
Public filePath As String
Public FullPath As String
Public Config As String
Public Count As Long
Public WidthIn As Double
Public HeightIn As Double
Public ThicknessIn As Double
'codex/fix-compile-error-at-thinaxisindex-5tl2al
Public thinAxis As Long

Private Sub Class_Initialize()
    thinAxis = -1
End Sub

Public Property Get ThinAxisIndex() As Long
    ThinAxisIndex = thinAxis
End Property

Public Property Let ThinAxisIndex(ByVal value As Long)
    thinAxis = value
 main
End Property


