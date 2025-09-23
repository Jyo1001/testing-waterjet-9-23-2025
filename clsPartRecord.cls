Option Explicit
Public DisplayName As String
Public FullPath As String
Public Config As String
Public Qty As Long
Public ThickIn As Double
Public BBoxX As Double
Public BBoxY As Double
Public BBoxZ As Double
Public ThinAxisIndex As Long

Private Sub Class_Initialize()
    ThinAxisIndex = -1
End Sub

' codex/fix-compile-error-at-thinaxisindex-5tl2al
Public Property Get thinAxis() As Long
    thinAxis = ThinAxisIndex
End Property

Public Property Let thinAxis(ByVal value As Long)
    ThinAxisIndex = value
End Property
