Option Explicit

' Uses globals from the module:
'   g_AllParts As Collection (of clsPartRecord)
'   g_SelectedIndices As Collection
'   g_GapIn As Double
'   g_UserCancelled As Boolean

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim pr As clsPartRecord

    lstParts.Clear
    lstParts.MultiSelect = fmMultiSelectExtended
    lstParts.IntegralHeight = False

    If (g_AllParts Is Nothing) Or g_AllParts.Count = 0 Then
        lstParts.AddItem "(no parts found)"
        Exit Sub
    End If

    For i = 1 To g_AllParts.Count
        Set pr = g_AllParts(i)
        lstParts.AddItem BuildDisplayText(pr) _
            & "  | Qty: " & pr.Qty _
            & " | Thick: " & Format$(pr.ThickIn, "0.000") & " in"
        lstParts.Selected(lstParts.ListCount - 1) = True
    Next

    If g_GapIn <= 0# Then g_GapIn = 0.125
    txtGap.value = Format$(g_GapIn, "0.###")
End Sub

Private Sub cmdBuild_Click()
    Dim i As Long, v As Double

    If IsNumeric(Replace(txtGap.value, ",", ".")) Then
        v = CDbl(Replace(txtGap.value, ",", "."))
        If v > 0# Then g_GapIn = v
    End If

    Set g_SelectedIndices = New Collection
    For i = 0 To lstParts.ListCount - 1
        If lstParts.Selected(i) Then g_SelectedIndices.Add (i + 1)
    Next

    g_UserCancelled = (g_SelectedIndices.Count = 0)
    Me.Hide
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Long
    For i = 0 To lstParts.ListCount - 1
        lstParts.Selected(i) = True
    Next
End Sub

Private Sub cmdDeselectAll_Click()
    Dim i As Long
    For i = 0 To lstParts.ListCount - 1
        lstParts.Selected(i) = False
    Next
End Sub

Private Sub cmdCancel_Click()
    g_UserCancelled = True
    Set g_SelectedIndices = Nothing
    Me.Hide
End Sub

Private Sub txtGap_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim v As Double
    If IsNumeric(Replace(txtGap.value, ",", ".")) Then
        v = CDbl(Replace(txtGap.value, ",", "."))
        If v > 0# Then txtGap.value = Format$(v, "0.###")
    End If
End Sub

Private Sub lstParts_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdBuild_Click
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        g_UserCancelled = True
        Set g_SelectedIndices = Nothing
    End If
End Sub





