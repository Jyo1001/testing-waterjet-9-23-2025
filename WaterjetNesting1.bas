Option Explicit

' ========= Units =========
Private Const M_TO_IN As Double = 39.3700787401575
Private Const IN_TO_M As Double = 0.0254

' ========= OPTIONAL: force a specific .DRWDOT (drawing template) =========
' Set to your Drawing.DRWDOT so drawings are consistent (change if needed)
Private Const DRAWING_TEMPLATE_OVERRIDE As String = _
    "I:\_6.SolidWorks Documents\SolidWorksSetUp\drawing templates\Drawing.DRWDOT"

' ========= Your .SLDDRT sheet format (border/title block) =========
Private Const SHEET_FORMAT_PATH As String = _
    "I:\_6.SolidWorks Documents\SolidWorksSetUp\drawing templates\BLANK.slddrt"

' ========= Orientation tolerances (inches) =========
Private Const ORIENTATION_AXIS_TOL_IN As Double = 0.01
Private Const ORIENTATION_THICKNESS_TOL_IN As Double = 0.01
Private Const ORIENTATION_AXIS_ALIGNMENT_TOL As Double = 0.001
Private Const ORIENTATION_TOP_PLANE_TOL_IN As Double = 0.001

' ========= Globals (used by frmNest) =========
Public g_SelectedIndices As Collection
Public g_GapIn As Double
Public g_UserCancelled As Boolean
Public g_AllParts As Collection   ' of clsPartRecord

' cached SolidWorks handle for math utilities
Private g_swApp As SldWorks.SldWorks

' for pinpointing fatal locations
Private g_LastStep As String

' ========= Logger =========
Private Sub LogMessage(msg As String, Optional showPopup As Boolean = False)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & msg
    If showPopup Then MsgBox msg, vbExclamation
End Sub

' Dump network drive mappings for debugging ("I:" issues, etc.)
Private Sub LogDriveMappings()
    On Error Resume Next
    Dim net As Object: Set net = CreateObject("WScript.Network")
    Dim col As Object: Set col = net.EnumNetworkDrives
    Dim i As Long
    For i = 0 To col.Count - 1 Step 2
        LogMessage "[DRIVE] " & col.Item(i) & " -> " & col.Item(i + 1)
    Next
    On Error GoTo 0
End Sub

' =========================
'        ENTRY POINT
' =========================
Sub Waterjet_Nesting_Workflow()
    On Error GoTo ohno
    g_LastStep = "[ENTRY]"

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swAsm As SldWorks.AssemblyDoc

    Set swApp = Application.SldWorks
    Set g_swApp = swApp
    Set swModel = swApp.ActiveDoc
    If swModel Is Nothing Then
        LogMessage "No active document.", True: Exit Sub
    End If
    If swModel.GetType <> swDocASSEMBLY Then
        LogMessage "Active document is not an assembly.", True: Exit Sub
    End If
    Set swAsm = swModel

    g_LastStep = "[ENTRY] ResolveAllLightWeight"
    TryResolveAllLightweight swModel, swAsm

    ' 1) Collect parts
    g_LastStep = "[COLLECT] start"
    Set g_AllParts = New Collection
    CollectAssemblyParts swAsm, swModel, g_AllParts
    If g_AllParts.Count = 0 Then
        LogMessage "No parts found in assembly.", True: Exit Sub
    End If

    ' 2) User form (fresh instance each run)
    g_UserCancelled = False
    If g_GapIn <= 0# Then g_GapIn = 0.125
    Set g_SelectedIndices = Nothing

    DumpAllPartsForUI

    On Error Resume Next
    Unload frmNest
    On Error GoTo 0
    Load frmNest
    frmNest.Show
    If g_UserCancelled Then Exit Sub
    If g_SelectedIndices Is Nothing Or g_SelectedIndices.Count = 0 Then
        LogMessage "No parts selected.", True: Exit Sub
    End If
    Unload frmNest

    ' 3) Filter selection
    g_LastStep = "[FILTER] selected"
    Dim filtered As New Collection, i As Long
    For i = 1 To g_SelectedIndices.Count
        filtered.Add g_AllParts(g_SelectedIndices(i))
    Next

    ' 4) Group by thickness
    g_LastStep = "[GROUP] by thickness"
    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim thkKey As Long, pr As clsPartRecord
    For i = 1 To filtered.Count
        Set pr = filtered(i)
        thkKey = CLng(pr.ThickIn * 1000# + 0.5)
        If Not groups.Exists(thkKey) Then Set groups(thkKey) = New Collection
        groups(thkKey).Add pr
    Next
    If groups.Count = 0 Then
        LogMessage "No groups created.", True: Exit Sub
    End If

    ' 5) Output folder
    g_LastStep = "[OUTPUT] folder"
    Dim asmPath As String: asmPath = swModel.GetPathName
    If Len(asmPath) = 0 Then
        LogMessage "Assembly must be saved before running.", True: Exit Sub
    End If
    Dim outFolder As String: outFolder = GetParentFolder(asmPath) & "\For waterjet cutting"
    EnsureFolder outFolder

    ' 6) Templates
    g_LastStep = "[TEMPLATES] fetch"
    Dim asmTpl As String, drwTplDefault As String
    asmTpl = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)
    drwTplDefault = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateDrawing)
    If Len(asmTpl) = 0 Then
        LogMessage "Set default assembly template in SolidWorks Options.", True: Exit Sub
    End If

    ' 7) Process each thickness group
    Dim k As Variant
    For Each k In groups.Keys
        Dim thkIn As Double: thkIn = CDbl(k) / 1000#
        Dim niceName As String: niceName = Format(thkIn, "0.###") & " in thick sheet"
        LogMessage "Processing group: " & niceName

        Dim groupFolder As String
        groupFolder = outFolder & "\" & SanitizeFileName(niceName)
        EnsureFolder groupFolder
        LogMessage "[FOLDER] Output for group -> " & groupFolder

        ' Build items list
        g_LastStep = "[GROUP] MakePlacementList"
        Dim items As Collection: Set items = MakePlacementList(groups(k))
        If items.Count = 0 Then
            LogMessage "No placeable items in group " & niceName
            GoTo NextGroup
        End If

        ' Create nesting assembly
        g_LastStep = "[NEWDOC] NewDocument"
        Dim nestAsmModel As SldWorks.ModelDoc2
        Set nestAsmModel = swApp.NewDocument(asmTpl, 0, 0, 0)
        If nestAsmModel Is Nothing Then
            LogMessage "Failed to create assembly for " & niceName, True
            GoTo NextGroup
        End If
        If nestAsmModel.GetType <> swDocASSEMBLY Then
            LogMessage "Template mismatch: assembly template is not .asmdot", True
            nestAsmModel.Quit: GoTo NextGroup
        End If
        Dim nestAsm As SldWorks.AssemblyDoc: Set nestAsm = nestAsmModel

        ' ---- Force IPS units on the new assembly
        ForceUnitsIPS nestAsmModel

        ' Save unique (silent)
        g_LastStep = "[SAVE] SaveAs4"
        Dim baseAsmPath As String: baseAsmPath = groupFolder & "\" & SanitizeFileName(niceName) & ".SLDASM"
        Dim targetAsmPath As String: targetAsmPath = UniqueTargetPath(baseAsmPath)
        Dim e As Long, w As Long
        nestAsmModel.SaveAs4 targetAsmPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, e, w
        LogMessage "[SAVE] SaveAs4 err=" & e & " warn=" & w & " -> " & targetAsmPath
        If e <> 0 Then
            LogMessage "[ERROR] Aborting group due to SaveAs4 failure for " & niceName, True
            nestAsmModel.Quit
            GoTo NextGroup
        End If

        ' Emit quantity report alongside assembly/DXF outputs
        Dim qtyReportPath As String
        qtyReportPath = Replace$(targetAsmPath, ".SLDASM", ".txt")
        WriteQuantityReportForGroup groups(k), qtyReportPath

        ' Place parts (coordinate-based, explicit config)
        g_LastStep = "[PLACE] begin"
        PlaceItemsGrid nestAsm, items, g_GapIn

        ' Save after placement
        g_LastStep = "[SAVE] post-place"
        Dim errCode As Long: nestAsmModel.Save3 swSaveAsOptions_Silent, errCode, 0
        LogMessage "[SAVE] Save3 after placement err=" & errCode
        If errCode <> 0 Then
            LogMessage "[ERROR] Aborting DXF export due to Save3 failure for " & niceName, True
            nestAsmModel.Quit
            GoTo NextGroup
        End If

        ' Export top-view-only DXF at 1:1
        g_LastStep = "[DXF] export"
        Dim dxfPath As String: dxfPath = Replace$(targetAsmPath, ".SLDASM", ".DXF")
        ExportModelTopDXF swApp, drwTplDefault, targetAsmPath, dxfPath

        ' Export individual DXFs for each unique part in this thickness group
        ExportIndividualPartDXFs swApp, drwTplDefault, groups(k), groupFolder

NextGroup:
    Next

    LogMessage "Waterjet nesting complete. Output: " & outFolder, True
    Exit Sub

ohno:
    LogMessage "Fatal error at " & g_LastStep & ": " & Err.Description, True
End Sub

' =========================
'  ASSEMBLY-LEVEL LW RESOLVE
' =========================
Private Sub TryResolveAllLightweight(swModel As SldWorks.ModelDoc2, swAsm As SldWorks.AssemblyDoc)
    On Error Resume Next
    CallByName swAsm, "ResolveAllLightWeightComponents", VbMethod, False
    CallByName swAsm, "ResolveAllLightWeightComponents3", VbMethod, True
    Dim ext As Object: Set ext = swModel.Extension
    If Not ext Is Nothing Then
        CallByName ext, "ResolveAllLightWeightComponents", VbMethod, True
        CallByName ext, "ResolveAllLightWeightComponents2", VbMethod, True
    End If
    swModel.EditRebuild3
    On Error GoTo 0
End Sub

' =========================
'  COLLECT PARTS (distinct instances)
' =========================
Private Sub CollectAssemblyParts(swAsm As SldWorks.AssemblyDoc, _
                                 swAsmModel As SldWorks.ModelDoc2, _
                                 ByRef outParts As Collection)

    Dim comps As Variant: comps = swAsm.GetComponents(True)
    If IsEmpty(comps) Then
        LogMessage "No components returned by GetComponents.": Exit Sub
    End If

    Dim asmFolder As String: asmFolder = GetParentFolder(swAsmModel.GetPathName)

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(comps) To UBound(comps)
        Dim c As SldWorks.Component2: Set c = comps(i)
        If c Is Nothing Then LogMessage "Skip: null component.": GoTo cont

        Dim sup As Long: sup = c.GetSuppression2
        If sup = swComponentSuppressionState_e.swComponentSuppressed Then
            LogMessage "Skip: suppressed " & c.Name2: GoTo cont
        End If

        Dim pth As String: pth = c.GetPathName
        Dim cfg As String: cfg = c.ReferencedConfiguration

        If Len(pth) = 0 Or InStr(1, pth, ".SLD", vbTextCompare) = 0 Then
            Dim md2 As SldWorks.ModelDoc2: Set md2 = c.GetModelDoc2
            If md2 Is Nothing Then
                EnsureResolved c
                Set md2 = c.GetModelDoc2
            End If
            If Not md2 Is Nothing Then pth = md2.GetPathName
        End If

        If Len(pth) = 0 Then
            Dim md3 As SldWorks.ModelDoc2: Set md3 = c.GetModelDoc2
            If Not md3 Is Nothing Then pth = EnsureExternalPathForVirtual(md3, asmFolder, c.Name2)
            If Len(pth) = 0 Then
                LogMessage "Skip: virtual part failed to export " & c.Name2
                GoTo cont
            End If
        End If

        ' bounding box
        Dim dxIn As Double, dyIn As Double, dzIn As Double
        Dim md4 As SldWorks.ModelDoc2: Set md4 = c.GetModelDoc2
        If Not TryGetBBoxInches_ComponentOnly(c, dxIn, dyIn, dzIn) Then
            EnsureResolved c
            Set md4 = c.GetModelDoc2
            If md4 Is Nothing Then
                LogMessage "Skip: unresolved (no ModelDoc2) " & c.Name2: GoTo cont
            End If
            If Not TryGetBBoxInches(c, dxIn, dyIn, dzIn, md4) Then
                LogMessage "Skip: bbox invalid for " & pth: GoTo cont
            End If
        End If
        If md4 Is Nothing Then Set md4 = c.GetModelDoc2
        Dim thinIdxModel As Long
        thinIdxModel = DetermineThinAxisIndex(md4, dxIn, dyIn, dzIn)

        Dim key As String: key = UCase$(pth) & "::" & UCase$(cfg)
        If Not dict.Exists(key) Then
            Dim rec As clsPartRecord
            Set rec = New clsPartRecord
            rec.FullPath = pth
            rec.Config = cfg
            rec.DisplayName = BuildDisplayText(rec)
            rec.Qty = 1
            rec.BBoxX = dxIn: rec.BBoxY = dyIn: rec.BBoxZ = dzIn
            rec.ThickIn = Round(Min3(dxIn, dyIn, dzIn), 3)
            rec.ThinAxisIndex = thinIdxModel
            dict.Add key, rec
            LogMessage "[COLLECT] " & rec.DisplayName & "  path=" & pth
        Else
            Dim r As clsPartRecord: Set r = dict(key)
            r.Qty = r.Qty + 1
            If r.ThinAxisIndex < 0 And thinIdxModel >= 0 Then r.ThinAxisIndex = thinIdxModel
        End If
cont:
    Next i

    Dim kk As Variant
    For Each kk In dict.Keys
        outParts.Add dict(kk)
    Next

    Dim j As Long
    For j = 1 To outParts.Count
        LogMessage "[UI] " & j & " -> " & outParts(j).FullPath
    Next
    LogMessage "Collected " & outParts.Count & " unique parts."
End Sub

Private Function TryGetBBoxInches_ComponentOnly(ByVal c As Object, _
                                                 ByRef dxIn As Double, _
                                                 ByRef dyIn As Double, _
                                                 ByRef dzIn As Double) As Boolean
    On Error Resume Next
    Dim v As Variant: v = SafeGetBox(c)
    If IsValidBox(v) Then
        dxIn = Abs(CDbl(v(3)) - CDbl(v(0))) * M_TO_IN
        dyIn = Abs(CDbl(v(4)) - CDbl(v(1))) * M_TO_IN
        dzIn = Abs(CDbl(v(5)) - CDbl(v(2))) * M_TO_IN
        TryGetBBoxInches_ComponentOnly = True
    Else
        TryGetBBoxInches_ComponentOnly = False
    End If
    On Error GoTo 0
End Function

Private Sub EnsureResolved(ByVal c As Object)
    On Error Resume Next
    Const swCompResolved As Long = 2
    CallByName c, "SetSuppression2", VbMethod, swCompResolved, 2, Nothing
    CallByName c, "SetLightWeightToResolved", VbMethod, True
    CallByName c, "SetLightWeightToResolved2", VbMethod, True
    On Error GoTo 0
End Sub

Private Function ComponentIsFixed(comp As SldWorks.Component2) As Boolean
    On Error Resume Next
    If comp Is Nothing Then Exit Function

    Dim fixedState As Variant
    fixedState = CallByName(comp, "IsFixed2", VbMethod)
    If Err.Number <> 0 Then
        Err.Clear
        fixedState = CallByName(comp, "IsFixed", VbMethod)
    End If

    If IsError(fixedState) Or IsNull(fixedState) Then
        ComponentIsFixed = False
    Else
        ComponentIsFixed = CBool(fixedState)
    End If
    On Error GoTo 0
End Function

Private Sub EnsureComponentIsFloat(comp As SldWorks.Component2, asm As SldWorks.AssemblyDoc)
    On Error Resume Next
    If comp Is Nothing Then Exit Sub
    If asm Is Nothing Then Exit Sub

    If ComponentIsFixed(comp) Then
        CallByName comp, "Select2", VbMethod, False, 0
        CallByName asm, "EditFloat", VbMethod
        CallByName asm, "ClearSelection2", VbMethod, True
        LogMessage "[PLACE] Floated fixed component before orientation: " & comp.Name2
    End If
    On Error GoTo 0
End Sub

Private Sub FixComponentInAssembly(comp As SldWorks.Component2, asm As SldWorks.AssemblyDoc)
    On Error Resume Next
    If comp Is Nothing Then Exit Sub
    If asm Is Nothing Then Exit Sub

    If Not ComponentIsFixed(comp) Then
        CallByName comp, "Select2", VbMethod, False, 0
        CallByName asm, "EditFix", VbMethod
        CallByName asm, "ClearSelection2", VbMethod, True
    End If
    On Error GoTo 0
End Sub

Private Function EnsureExternalPathForVirtual(md As SldWorks.ModelDoc2, _
                                              ByVal suggestFolder As String, _
                                              ByVal baseName As String) As String
    On Error Resume Next
    Dim outDir As String: outDir = suggestFolder & "\Extracted Virtual Parts"
    EnsureFolder outDir
    Dim outPath As String: outPath = outDir & "\" & SanitizeFileName(baseName) & ".SLDPRT"
    outPath = UniqueTargetPath(outPath)
    Dim e As Long, w As Long
    md.SaveAs4 outPath, swSaveAsCurrentVersion, swSaveAsOptions_Silent, e, w
    If e = 0 Then EnsureExternalPathForVirtual = outPath Else EnsureExternalPathForVirtual = ""
    On Error GoTo 0
End Function

Private Function SafeGetBox(ByVal c As Object) As Variant
    On Error Resume Next
    Dim v As Variant
    v = CallByName(c, "GetBox", VbMethod)
    If IsValidBox(v) Then SafeGetBox = v: GoTo done
    Err.Clear: v = CallByName(c, "GetBox", VbMethod, False)
    If IsValidBox(v) Then SafeGetBox = v: GoTo done
    Err.Clear: v = CallByName(c, "GetBox", VbMethod, True)
    If IsValidBox(v) Then SafeGetBox = v
done:
    On Error GoTo 0
End Function

' ========= explicit AddComponent5(x,y,z,config) path =========
Private Function SafeAddComponent(ByVal asmDoc As Object, _
                                  ByVal filePath As String, _
                                  ByVal cfg As String, _
                                  ByVal xM As Double, ByVal yM As Double, ByVal zM As Double) _
                                  As SldWorks.Component2
    On Error Resume Next
    Dim r As Object

    Set r = CallByName(asmDoc, "AddComponent5", VbMethod, filePath, 0, cfg, xM, yM, zM)
    If r Is Nothing Then
        Set r = CallByName(asmDoc, "AddComponent3", VbMethod, filePath, xM, yM, zM)
        If r Is Nothing Then
            Set r = CallByName(asmDoc, "AddComponent2", VbMethod, filePath, xM, yM, zM)
        End If
    End If
    On Error GoTo 0

    If Not r Is Nothing Then
        Dim gotPath As String
        On Error Resume Next
        gotPath = r.GetPathName
        On Error GoTo 0
        If Len(gotPath) > 0 And StrComp(UCase$(gotPath), UCase$(filePath), vbTextCompare) <> 0 Then
            LogMessage "[WARN] Added a different file than requested: " & gotPath & " vs " & filePath
        End If
        Set SafeAddComponent = r
    Else
        Set SafeAddComponent = Nothing
    End If
End Function

' =========================
'  NESTING / PLACEMENT (no transforms)
' =========================
Private Sub PlaceItemsGrid(nestAsm As SldWorks.AssemblyDoc, _
                           items As Collection, _
                           GapIn As Double)

    Dim gapM As Double: gapM = GapIn * IN_TO_M
    Dim cursorX As Double, cursorY As Double, rowH As Double
    Dim targetRowWidthM As Double: targetRowWidthM = 60# * IN_TO_M

    Dim i As Long, n As Long
    For i = 1 To items.Count
        Dim pi As clsPlaceItem: Set pi = items(i)
        If Len(pi.FullPath) = 0 Then
            LogMessage "Skip placement: empty file path for " & pi.Config
            GoTo nextItem
        End If

        Dim placements As Long: placements = 1
        If pi.Count > 1 Then
            LogMessage "[INFO] Qty " & pi.Count & " requested for " & GetFileName(pi.FullPath) & " (" & pi.Config & ") - placing single instance"
        End If

        For n = 1 To placements
            Dim wM As Double: wM = pi.WidthIn * IN_TO_M
            Dim hM As Double: hM = pi.HeightIn * IN_TO_M

            If cursorX > 0 And (cursorX + wM) > targetRowWidthM Then
                cursorX = 0
                cursorY = cursorY + rowH + gapM
                rowH = 0
            End If

            g_LastStep = "[PLACE] AddComponent5(x,y,z)"
            Dim comp As SldWorks.Component2
            Set comp = SafeAddComponent(nestAsm, pi.FullPath, pi.Config, cursorX, cursorY, 0#)
            If comp Is Nothing Then
                LogMessage "AddComponent failed: " & pi.FullPath & " (" & pi.Config & ")", True
            Else
                g_LastStep = "[PLACE] orient component"
                OrientComponentForNesting nestAsm, comp, pi
            End If

            cursorX = cursorX + wM + gapM
            If hM > rowH Then rowH = hM
        Next n
nextItem:
    Next i

    g_LastStep = "[PLACE] ForceRebuild3"
    nestAsm.ForceRebuild3 False
End Sub

' Orient each newly inserted part so the thinnest axis aligns with the assembly Top plane
' codex/fix-compile-error-at-thinaxisindex-5tl2al
Private Sub OrientComponentForNesting(nestAsm As SldWorks.AssemblyDoc, _
                                      comp As SldWorks.Component2, _
                                      pi As clsPlaceItem)

    On Error Resume Next
    Const AXIS_ALIGN_EPS As Double = 0.000001
    Const TOP_PLANE_GAP_TOL_M As Double = ORIENTATION_TOP_PLANE_TOL_IN * IN_TO_M
    If comp Is Nothing Then Exit Sub

    EnsureResolved comp
    EnsureComponentIsFloat comp, nestAsm

    Dim partDoc As SldWorks.ModelDoc2
    Set partDoc = comp.GetModelDoc2

    Dim hasLargestFace As Boolean
    Dim largestFaceNormal(0 To 2) As Double
    Dim largestFaceArea As Double
    If Not partDoc Is Nothing Then
        hasLargestFace = TryGetLargestPlanarFaceNormal(partDoc, _
            largestFaceNormal(0), largestFaceNormal(1), largestFaceNormal(2), largestFaceArea)
    End If

    Dim partLabel As String
    partLabel = GetFileName(pi.FullPath) & " (" & pi.Config & ")"

    Dim thinAxisIdx As Long
    thinAxisIdx = pi.thinAxis
    If thinAxisIdx < 0 Or thinAxisIdx > 2 Then
        Dim compDx As Double, compDy As Double, compDz As Double
        If TryGetBBoxInches_ComponentOnly(comp, compDx, compDy, compDz) Then
            thinAxisIdx = IndexOfMin3(compDx, compDy, compDz)
        Else
            If partDoc Is Nothing Then Set partDoc = comp.GetModelDoc2
            If Not partDoc Is Nothing Then
                thinAxisIdx = DetermineThinAxisIndex(partDoc, compDx, compDy, compDz)
                If Not hasLargestFace Then
                    hasLargestFace = TryGetLargestPlanarFaceNormal(partDoc, _
                        largestFaceNormal(0), largestFaceNormal(1), largestFaceNormal(2), largestFaceArea)
                End If
            End If
        End If
        If thinAxisIdx >= 0 And thinAxisIdx <= 2 Then
            pi.thinAxis = thinAxisIdx
        Else
            LogMessage "[WARN] Unable to resolve thin axis for " & pi.FullPath & _
                " (" & pi.Config & ")"
        End If
    End If

    Dim baseTransform As SldWorks.MathTransform
    Set baseTransform = comp.Transform2
    If baseTransform Is Nothing Then
        LogMessage "[WARN] Orientation skipped (no transform) for " & pi.FullPath & " (" & pi.Config & ")"

        Exit Sub
    End If

    Dim baseData As Variant: baseData = baseTransform.ArrayData
    If IsEmpty(baseData) Or UBound(baseData) < 14 Then
        LogMessage "[WARN] Orientation skipped (no transform data) for " & pi.FullPath & " (" & pi.Config & ")"
        Exit Sub
    End If

    Dim mathUtil As SldWorks.MathUtility
    If g_swApp Is Nothing Then Set g_swApp = Application.SldWorks
    Set mathUtil = g_swApp.GetMathUtility
    If mathUtil Is Nothing Then
        LogMessage "[WARN] Orientation skipped (no MathUtility) for " & pi.FullPath & " (" & pi.Config & ")"
        Exit Sub
    End If

    Dim baseRot As Variant
    baseRot = ExtractRotationMatrix(baseData)
    If IsEmpty(baseRot) Then
        LogMessage "[WARN] Orientation skipped (invalid base rotation) for " & pi.FullPath & " (" & pi.Config & ")"
        Exit Sub
    End If

    Dim bestMatrix As Variant: bestMatrix = baseRot
    Dim bestScore As Double: bestScore = 1000000000#
    Dim bestFound As Boolean

    Dim bestAxisMatrix As Variant
    Dim bestAxisAlign As Double: bestAxisAlign = -1#
    Dim bestAxisPlanar As Double: bestAxisPlanar = 1000000000#
    Dim bestAxisFound As Boolean

    Dim bestFaceMatrix As Variant
    Dim bestFaceAlignment As Double: bestFaceAlignment = -1#
    Dim bestFacePlanar As Double: bestFacePlanar = 1000000000#
    Dim bestFaceAxisComponent As Double: bestFaceAxisComponent = -2#
    Dim bestFaceFound As Boolean
codex/fix-orientation-of-assembly-part-bsaaop

    Dim candidateMatrices As Collection
    Set candidateMatrices = BuildOrientationCandidateMatrices(baseRot, thinAxisIdx, hasLargestFace, largestFaceNormal)
    If candidateMatrices Is Nothing Then
        Set candidateMatrices = New Collection
        candidateMatrices.Add baseRot
    End If

    Dim newR As Variant
    For Each newR In candidateMatrices


 main
        Dim candidateTransform As SldWorks.MathTransform
        Set candidateTransform = CreateTransformFromMatrix(baseData, newR, mathUtil)
        If candidateTransform Is Nothing Then GoTo nextRot

        comp.SetTransformAndSolve2 candidateTransform

        Dim hasScore As Boolean
        Dim score As Double
        Dim axisPlanar As Double
        Dim axisAlign As Double
        Dim axisOk As Boolean

        If thinAxisIdx >= 0 And thinAxisIdx <= 2 Then
            axisOk = EvaluateThinAxisAlignment(newR, thinAxisIdx, axisPlanar, axisAlign)
            If axisOk Then
                Dim betterAxis As Boolean
                betterAxis = (axisAlign > bestAxisAlign + AXIS_ALIGN_EPS)
                If Not betterAxis Then
                    If Abs(axisAlign - bestAxisAlign) <= AXIS_ALIGN_EPS Then
                        betterAxis = (axisPlanar < bestAxisPlanar - AXIS_ALIGN_EPS)
                    End If
                End If

                If betterAxis Then
                    bestAxisMatrix = newR
                    bestAxisAlign = axisAlign
                    bestAxisPlanar = axisPlanar
                    bestAxisFound = True
                End If
                score = OrientationMatrixScore(axisPlanar, axisAlign)
                hasScore = True
            End If
        End If

        If hasLargestFace Then
            Dim facePlanarErr As Double, faceAxisComponent As Double
            Dim faceAlignment As Double
            faceAlignment = EvaluateFaceAlignment(newR, largestFaceNormal(0), largestFaceNormal(1), _
                largestFaceNormal(2), facePlanarErr, faceAxisComponent)
            If faceAlignment > 0# Then
                Dim betterFace As Boolean
                betterFace = (faceAlignment > bestFaceAlignment + AXIS_ALIGN_EPS)
                If Not betterFace Then
                    If Abs(faceAlignment - bestFaceAlignment) <= AXIS_ALIGN_EPS Then
                        betterFace = (facePlanarErr < bestFacePlanar - AXIS_ALIGN_EPS)
                        If Not betterFace Then
                            betterFace = (faceAxisComponent > bestFaceAxisComponent + AXIS_ALIGN_EPS)
                        End If
                    End If
                End If

                If betterFace Or Not bestFaceFound Then
                    bestFaceMatrix = newR
                    bestFaceAlignment = faceAlignment
                    bestFacePlanar = facePlanarErr
                    bestFaceAxisComponent = faceAxisComponent
                    bestFaceFound = True
                End If
            End If
        End If

        If Not hasScore Then
            Dim isZThin As Boolean, thicknessDiff As Double
            Dim measured As Double, zDelta As Double
            Dim planeGapTmp As Double

            If EvaluateOrientationMetrics(comp, pi, isZThin, thicknessDiff, measured, zDelta, planeGapTmp) Then
                score = OrientationCandidateScoreFromBBox(isZThin, thicknessDiff, zDelta)
                hasScore = True
            End If
        End If

        If hasScore Then
            If (Not bestFound) Or score < bestScore Then
                bestMatrix = newR
                bestScore = score
                bestFound = True
            End If
        End If
'codex/fix-compile-error-at-thinaxisindex-5tl2al
nextRot:
    Next newR

    Dim finalMatrix As Variant
    Dim usedFaceMatrix As Boolean

    If bestFaceFound And bestFaceAlignment >= 1# - ORIENTATION_AXIS_ALIGNMENT_TOL Then
        finalMatrix = bestFaceMatrix
        usedFaceMatrix = True
    ElseIf bestAxisFound Then
        finalMatrix = bestAxisMatrix
    ElseIf bestFound Then
        finalMatrix = bestMatrix
    Else
        finalMatrix = baseRot
    End If

    Dim finalTransform As SldWorks.MathTransform
    Set finalTransform = CreateTransformFromMatrix(baseData, finalMatrix, mathUtil)
    If finalTransform Is Nothing Then
        Set finalTransform = baseTransform
        finalMatrix = baseRot
    End If

    Dim applyOk As Variant
    applyOk = comp.SetTransformAndSolve2(finalTransform)
    If VarType(applyOk) = vbBoolean Then
        If Not CBool(applyOk) Then
            LogMessage "[WARN] Failed to apply orientation transform for " & pi.FullPath & _
                " (" & pi.Config & ")"
        End If
    End If

    Dim planeShiftIn As Double
    planeShiftIn = 0#

    Dim finalBox As Variant
    finalBox = SafeGetBox(comp)
    If IsValidBox(finalBox) Then
        Dim zMinM As Double
        zMinM = CDbl(finalBox(2))
        If Abs(zMinM) > TOP_PLANE_GAP_TOL_M Then
            Dim finalData As Variant
            finalData = finalTransform.ArrayData
            If Not IsEmpty(finalData) And UBound(finalData) >= 14 Then
                finalData(14) = CDbl(finalData(14)) - zMinM
                Dim adjustedTransform As SldWorks.MathTransform
                Set adjustedTransform = mathUtil.CreateTransform(finalData)
                If Not adjustedTransform Is Nothing Then
                    Set finalTransform = adjustedTransform
                    comp.SetTransformAndSolve2 finalTransform
                    planeShiftIn = Abs(zMinM) * M_TO_IN
                    finalBox = SafeGetBox(comp)
                End If
            End If
        End If
    End If

    Dim matrixEvalOk As Boolean
    Dim finalPlanarErr As Double, finalAxisAlign As Double
    If thinAxisIdx >= 0 And thinAxisIdx <= 2 Then
        matrixEvalOk = EvaluateThinAxisAlignment(finalMatrix, thinAxisIdx, finalPlanarErr, finalAxisAlign)
    End If

    Dim finalFaceAlignment As Double
    Dim finalFacePlanar As Double
    Dim finalFaceAxisComponent As Double
    Dim finalFaceOk As Boolean

    Dim finalIsZThin As Boolean, finalDiff As Double
    Dim finalMeasured As Double, finalZDelta As Double
    Dim measurementOk As Boolean
    Dim finalPlaneGapIn As Double

    measurementOk = EvaluateOrientationMetrics(comp, pi, finalIsZThin, finalDiff, finalMeasured, finalZDelta, finalPlaneGapIn)
    If Not measurementOk Then
        nestAsm.EditRebuild3
        measurementOk = EvaluateOrientationMetrics(comp, pi, finalIsZThin, finalDiff, finalMeasured, finalZDelta, finalPlaneGapIn)
    End If

    If hasLargestFace Then
        finalFaceAlignment = EvaluateFaceAlignment(finalMatrix, largestFaceNormal(0), largestFaceNormal(1), _
            largestFaceNormal(2), finalFacePlanar, finalFaceAxisComponent)
        finalFaceOk = (finalFaceAlignment >= 1# - ORIENTATION_AXIS_ALIGNMENT_TOL)
    Else
        finalFaceOk = True
    End If

    Dim axisAlignedOk As Boolean: axisAlignedOk = True
    Dim planeGapOk As Boolean
    planeGapOk = (measurementOk And finalPlaneGapIn <= ORIENTATION_TOP_PLANE_TOL_IN)

    If thinAxisIdx >= 0 Then
        If matrixEvalOk Then
            axisAlignedOk = (finalAxisAlign >= 1# - ORIENTATION_AXIS_ALIGNMENT_TOL And finalPlanarErr <= ORIENTATION_AXIS_ALIGNMENT_TOL)
        ElseIf measurementOk Then
            axisAlignedOk = finalIsZThin
        Else
            axisAlignedOk = False
        End If
    End If

    Dim faceAlignedOk As Boolean
    faceAlignedOk = finalFaceOk

    Dim orientationAligned As Boolean
    orientationAligned = (axisAlignedOk And planeGapOk And faceAlignedOk)

    Dim faceStatusMsg As String
    If hasLargestFace Then
        faceStatusMsg = "; largest face |Z|=" & Format$(finalFaceAlignment, "0.000")
        If Not finalFaceOk Then faceStatusMsg = faceStatusMsg & " (NOT aligned)"
    End If

    If planeShiftIn > 0# Then
        LogMessage "[PLACE] Shifted " & partLabel & " by " & Format$(planeShiftIn, "0.###") & " in to rest on Top plane"
    End If

    If usedFaceMatrix And hasLargestFace Then
        Dim faceAreaIn2 As Double
        faceAreaIn2 = largestFaceArea * (M_TO_IN * M_TO_IN)
        LogMessage "[PLACE] Using largest face alignment for " & partLabel & _
            " (|Z|=" & Format$(finalFaceAlignment, "0.000") & ", area=" & _
            Format$(faceAreaIn2, "0.0") & " in^2)"
    End If

    If orientationAligned Then
        If measurementOk Then
            Dim thicknessMsg As String
            thicknessMsg = "thickness " & Format$(finalMeasured, "0.###") & " in"
            If pi.ThicknessIn > 0# Then
                thicknessMsg = thicknessMsg & " (? " & Format$(finalDiff, "0.###") & " in)"
            End If
            Dim thicknessWarn As Boolean
            thicknessWarn = (pi.ThicknessIn > 0# And Abs(finalDiff) > ORIENTATION_THICKNESS_TOL_IN)

            If matrixEvalOk Then
                LogMessage "[CHECK] Orientation OK for " & partLabel & ": thin axis -> Top (|Z|=" & _
                    Format$(finalAxisAlign, "0.000") & ", planar=" & Format$(finalPlanarErr, "0.000") & ", plane gap=" & _
                    Format$(finalPlaneGapIn, "0.000") & " in); " & thicknessMsg & faceStatusMsg
            Else
                LogMessage "[CHECK] Orientation OK for " & partLabel & ": " & thicknessMsg & _
                    " (bbox ?Z=" & Format$(finalZDelta, "0.###") & " in, plane gap=" & _
                    Format$(finalPlaneGapIn, "0.000") & " in)" & faceStatusMsg
            End If

            If thicknessWarn Then
                LogMessage "[WARN] Thickness mismatch after orientation for " & partLabel & _
                    ": expected " & Format$(pi.ThicknessIn, "0.###") & " in, measured " & _
                    Format$(finalMeasured, "0.###") & " in"
            End If
        ElseIf matrixEvalOk Then
            LogMessage "[CHECK] Orientation OK for " & partLabel & ": thin axis -> Top (|Z|=" & _
                Format$(finalAxisAlign, "0.000") & ", planar=" & Format$(finalPlanarErr, "0.000") & _
                ", plane gap=" & Format$(finalPlaneGapIn, "0.000") & " in); thickness check unavailable" & faceStatusMsg
        Else
            LogMessage "[CHECK] Orientation OK for " & partLabel & ": verification limited (plane gap=" & _
                Format$(finalPlaneGapIn, "0.000") & " in)" & faceStatusMsg
        End If
    Else
        If Not planeGapOk And measurementOk Then
            LogMessage "[ERROR] Part not resting on Top plane for " & partLabel & _
                ": gap=" & Format$(finalPlaneGapIn, "0.###") & " in"
        End If

        If Not faceAlignedOk And hasLargestFace Then
            LogMessage "[ERROR] Largest face not aligned with Top plane for " & partLabel & _
                ": |Z|=" & Format$(finalFaceAlignment, "0.000")
        End If

        If Not axisAlignedOk Then
            If matrixEvalOk Then
                LogMessage "[ERROR] Thin axis misaligned for " & partLabel & ": |Z|=" & _
                    Format$(finalAxisAlign, "0.000") & ", planar=" & Format$(finalPlanarErr, "0.000")
            ElseIf measurementOk Then
                LogMessage "[ERROR] Thin axis not aligned with Top plane for " & partLabel & _
                    "; Z delta=" & Format$(finalZDelta, "0.###") & " in"
            Else
                LogMessage "[WARN] Unable to verify orientation for " & pi.FullPath & " (" & pi.Config & ")"
            End If
        ElseIf planeGapOk Then
            If faceAlignedOk Then
                LogMessage "[WARN] Orientation check inconclusive for " & partLabel
            End If
        Else
            LogMessage "[WARN] Unable to verify orientation for " & pi.FullPath & " (" & pi.Config & ")"
        End If
    End If

    On Error GoTo 0

    FixComponentInAssembly comp, nestAsm
End Sub

Private Function OrientationCandidateScoreFromBBox(isZThin As Boolean, _
                                                   thicknessDiff As Double, _
                                                   zDelta As Double) As Double
    Dim penalty As Double: penalty = Abs(thicknessDiff) + zDelta
    If Not isZThin Then penalty = penalty + 1000#
    OrientationCandidateScoreFromBBox = penalty
End Function

Private Function BuildOrientationCandidateRotations() As Collection
    Dim result As New Collection
    Dim qx As Long, qy As Long, qz As Long
    For qz = 0 To 3
        Dim rotZ As Variant
        rotZ = QuarterTurnMatrix(2, qz)
        For qy = 0 To 3
            Dim rotY As Variant
            rotY = QuarterTurnMatrix(1, qy)
            Dim rotZY As Variant
            rotZY = MultiplyMatrix3x3(rotZ, rotY)
            For qx = 0 To 3
                Dim rotX As Variant
                rotX = QuarterTurnMatrix(0, qx)
                Dim combined As Variant
                combined = MultiplyMatrix3x3(rotZY, rotX)
                result.Add combined
            Next qx
        Next qy
    Next qz
    Set BuildOrientationCandidateRotations = result
End Function

Private Function QuarterTurnMatrix(axisIndex As Long, quarterTurns As Long) As Variant
    Dim rot(0 To 2, 0 To 2) As Double
    Dim qt As Long: qt = ((quarterTurns Mod 4) + 4) Mod 4

    Select Case axisIndex
        Case 0 ' X axis
            Select Case qt
                Case 0
                    rot(0, 0) = 1#: rot(1, 1) = 1#: rot(2, 2) = 1#
                Case 1 ' +90
                    rot(0, 0) = 1#: rot(1, 2) = -1#: rot(2, 1) = 1#
                Case 2 ' 180
                    rot(0, 0) = 1#: rot(1, 1) = -1#: rot(2, 2) = -1#
                Case 3 ' -90
                    rot(0, 0) = 1#: rot(1, 2) = 1#: rot(2, 1) = -1#
            End Select

        Case 1 ' Y axis
            Select Case qt
                Case 0
                    rot(0, 0) = 1#: rot(1, 1) = 1#: rot(2, 2) = 1#
                Case 1 ' +90
                    rot(0, 2) = 1#: rot(1, 1) = 1#: rot(2, 0) = -1#
                Case 2 ' 180
                    rot(0, 0) = -1#: rot(1, 1) = 1#: rot(2, 2) = -1#
                Case 3 ' -90
                    rot(0, 2) = -1#: rot(1, 1) = 1#: rot(2, 0) = 1#
            End Select

        Case Else ' Z axis
            Select Case qt
                Case 0
                    rot(0, 0) = 1#: rot(1, 1) = 1#: rot(2, 2) = 1#
                Case 1 ' +90
                    rot(0, 1) = -1#: rot(1, 0) = 1#: rot(2, 2) = 1#
                Case 2 ' 180
                    rot(0, 0) = -1#: rot(1, 1) = -1#: rot(2, 2) = 1#
                Case 3 ' -90
                    rot(0, 1) = 1#: rot(1, 0) = -1#: rot(2, 2) = 1#
            End Select
    End Select

    QuarterTurnMatrix = rot
End Function

Private Function MultiplyMatrix3x3(a As Variant, b As Variant) As Variant
    Dim res(0 To 2, 0 To 2) As Double
    Dim i As Long, j As Long, k As Long
    For i = 0 To 2
        For j = 0 To 2
            Dim sum As Double: sum = 0#
            For k = 0 To 2
                sum = sum + CDbl(a(i, k)) * CDbl(b(k, j))
            Next k
            If Abs(sum) < 0.000000000001 Then sum = 0#
            res(i, j) = sum
        Next j
    Next i
    MultiplyMatrix3x3 = res
End Function

Private Sub MultiplyMatrixVector3x3(mat As Variant, x As Double, y As Double, z As Double, _
                                    ByRef outX As Double, ByRef outY As Double, ByRef outZ As Double)
    outX = 0#: outY = 0#: outZ = 0#
    On Error Resume Next
    If IsEmpty(mat) Then Exit Sub
    outX = CDbl(mat(0, 0)) * x + CDbl(mat(0, 1)) * y + CDbl(mat(0, 2)) * z
    outY = CDbl(mat(1, 0)) * x + CDbl(mat(1, 1)) * y + CDbl(mat(1, 2)) * z
    outZ = CDbl(mat(2, 0)) * x + CDbl(mat(2, 1)) * y + CDbl(mat(2, 2)) * z
    On Error GoTo 0
End Sub

Private Function IdentityMatrix3() As Variant
    Dim m(0 To 2, 0 To 2) As Double
    m(0, 0) = 1#: m(1, 1) = 1#: m(2, 2) = 1#
    IdentityMatrix3 = m
End Function

Private Function MatrixKey(mat As Variant) As String
    If IsEmpty(mat) Then Exit Function
    Dim key As String
    Dim i As Long, j As Long
    For i = 0 To 2
        For j = 0 To 2
            key = key & "|" & Format$(CDbl(mat(i, j)), "0.000000")
        Next j
    Next i
    MatrixKey = key
End Function

Private Sub AddMatrixCandidate(ByRef col As Collection, ByRef seen As Object, candidate As Variant)
    If col Is Nothing Then Exit Sub
    If IsEmpty(candidate) Then Exit Sub
    Dim key As String: key = MatrixKey(candidate)
    If Len(key) = 0 Then Exit Sub
    If seen Is Nothing Then
        col.Add candidate
    ElseIf Not seen.Exists(key) Then
        seen.Add key, True
        col.Add candidate
    End If
End Sub

Private Function BuildAlignmentMatrixForVectors(srcX As Double, srcY As Double, srcZ As Double, _
                                                dstX As Double, dstY As Double, dstZ As Double) As Variant
    Const EPS As Double = 0.0000000001

    Dim srcMag As Double
    srcMag = Sqr(srcX * srcX + srcY * srcY + srcZ * srcZ)
    If srcMag <= EPS Then Exit Function

    Dim dstMag As Double
    dstMag = Sqr(dstX * dstX + dstY * dstY + dstZ * dstZ)
    If dstMag <= EPS Then Exit Function

    Dim ax As Double, ay As Double, az As Double
    ax = srcX / srcMag
    ay = srcY / srcMag
    az = srcZ / srcMag

    Dim bx As Double, by As Double, bz As Double
    bx = dstX / dstMag
    by = dstY / dstMag
    bz = dstZ / dstMag

    Dim dotProd As Double
    dotProd = ax * bx + ay * by + az * bz
    If dotProd > 1# Then dotProd = 1#
    If dotProd < -1# Then dotProd = -1#

    Dim axisX As Double, axisY As Double, axisZ As Double
    axisX = ay * bz - az * by
    axisY = az * bx - ax * bz
    axisZ = ax * by - ay * bx

    Dim s As Double
    s = Sqr(axisX * axisX + axisY * axisY + axisZ * axisZ)

    Dim result(0 To 2, 0 To 2) As Double

    If s <= EPS Then
        If dotProd >= 1# - 0.0000001 Then
            BuildAlignmentMatrixForVectors = IdentityMatrix3()
            Exit Function
        End If

        Dim px As Double, py As Double, pz As Double
        px = 0#: py = -az: pz = ay
        Dim pMag As Double: pMag = Sqr(px * px + py * py + pz * pz)
        If pMag <= EPS Then
            px = -az: py = ax: pz = 0#
            pMag = Sqr(px * px + py * py + pz * pz)
        End If
        If pMag <= EPS Then
            px = 1#: py = 0#: pz = 0#: pMag = 1#
        End If
        px = px / pMag: py = py / pMag: pz = pz / pMag

        result(0, 0) = 2# * px * px - 1#
        result(0, 1) = 2# * px * py
        result(0, 2) = 2# * px * pz
        result(1, 0) = 2# * py * px
        result(1, 1) = 2# * py * py - 1#
        result(1, 2) = 2# * py * pz
        result(2, 0) = 2# * pz * px
        result(2, 1) = 2# * pz * py
        result(2, 2) = 2# * pz * pz - 1#
        BuildAlignmentMatrixForVectors = result
        Exit Function
    End If

    axisX = axisX / s
    axisY = axisY / s
    axisZ = axisZ / s

    Dim sinTheta As Double: sinTheta = s
    Dim cosTheta As Double: cosTheta = dotProd
    Dim oneMinusCos As Double: oneMinusCos = 1# - cosTheta

    result(0, 0) = cosTheta + axisX * axisX * oneMinusCos
    result(0, 1) = axisX * axisY * oneMinusCos - axisZ * sinTheta
    result(0, 2) = axisX * axisZ * oneMinusCos + axisY * sinTheta
    result(1, 0) = axisY * axisX * oneMinusCos + axisZ * sinTheta
    result(1, 1) = cosTheta + axisY * axisY * oneMinusCos
    result(1, 2) = axisY * axisZ * oneMinusCos - axisX * sinTheta
    result(2, 0) = axisZ * axisX * oneMinusCos - axisY * sinTheta
    result(2, 1) = axisZ * axisY * oneMinusCos + axisX * sinTheta
    result(2, 2) = cosTheta + axisZ * axisZ * oneMinusCos

    BuildAlignmentMatrixForVectors = result
End Function

Private Sub AddAlignmentBasedCandidates(result As Collection, seen As Object, baseRot As Variant, _
                                        sourceX As Double, sourceY As Double, sourceZ As Double)
    Const MIN_MAG As Double = 0.000000001
    Dim mag As Double
    mag = Sqr(sourceX * sourceX + sourceY * sourceY + sourceZ * sourceZ)
    If mag <= MIN_MAG Then Exit Sub

    Dim targetIndex As Long
    For targetIndex = 0 To 1
        Dim targetZ As Double
        If targetIndex = 0 Then targetZ = 1# Else targetZ = -1#

        Dim align As Variant
        align = BuildAlignmentMatrixForVectors(sourceX, sourceY, sourceZ, 0#, 0#, targetZ)
        If Not IsEmpty(align) Then
            Dim aligned As Variant
            aligned = MultiplyMatrix3x3(align, baseRot)
            AddMatrixCandidate result, seen, aligned

            Dim q As Long
            For q = 1 To 3
                Dim zRot As Variant
                zRot = QuarterTurnMatrix(2, q)
                Dim spun As Variant
                spun = MultiplyMatrix3x3(zRot, aligned)
                AddMatrixCandidate result, seen, spun
            Next q
        End If
    Next targetIndex
End Sub

Private Function BuildOrientationCandidateMatrices(baseRot As Variant, _
                                                   thinAxisIdx As Long, _
                                                   hasLargestFace As Boolean, _
                                                   ByRef largestFaceNormal() As Double) As Collection

    Dim result As New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    AddMatrixCandidate result, seen, baseRot

    Dim baseCandidates As Collection
    Set baseCandidates = BuildOrientationCandidateRotations()

    Dim rot As Variant
    For Each rot In baseCandidates
        Dim newR As Variant
        newR = MultiplyMatrix3x3(baseRot, rot)
        AddMatrixCandidate result, seen, newR
    Next rot

    Dim assemblyVec(0 To 2) As Double

    If thinAxisIdx >= 0 And thinAxisIdx <= 2 Then
        Dim axisVec(0 To 2) As Double
        axisVec(0) = 0#: axisVec(1) = 0#: axisVec(2) = 0#
        axisVec(thinAxisIdx) = 1#
        MultiplyMatrixVector3x3 baseRot, axisVec(0), axisVec(1), axisVec(2), _
            assemblyVec(0), assemblyVec(1), assemblyVec(2)
        AddAlignmentBasedCandidates result, seen, baseRot, assemblyVec(0), assemblyVec(1), assemblyVec(2)
    End If

    If hasLargestFace Then
        MultiplyMatrixVector3x3 baseRot, largestFaceNormal(0), largestFaceNormal(1), largestFaceNormal(2), _
            assemblyVec(0), assemblyVec(1), assemblyVec(2)
        AddAlignmentBasedCandidates result, seen, baseRot, assemblyVec(0), assemblyVec(1), assemblyVec(2)
    End If

    Set BuildOrientationCandidateMatrices = result
End Function

Private Function ExtractRotationMatrix(baseData As Variant) As Variant
    If IsEmpty(baseData) Then Exit Function
    If UBound(baseData) < 10 Then Exit Function

    Dim rot(0 To 2, 0 To 2) As Double
    rot(0, 0) = CDbl(baseData(0))
    rot(1, 0) = CDbl(baseData(1))
    rot(2, 0) = CDbl(baseData(2))
    rot(0, 1) = CDbl(baseData(4))
    rot(1, 1) = CDbl(baseData(5))
    rot(2, 1) = CDbl(baseData(6))
    rot(0, 2) = CDbl(baseData(8))
    rot(1, 2) = CDbl(baseData(9))
    rot(2, 2) = CDbl(baseData(10))

    ExtractRotationMatrix = rot
End Function

Private Function EvaluateThinAxisAlignment(rot As Variant, _
                                           ThinAxisIndex As Long, _
                                           ByRef planarError As Double, _
                                           ByRef axisAlignment As Double) As Boolean
    On Error Resume Next
    If IsEmpty(rot) Then Exit Function
    If ThinAxisIndex < 0 Or ThinAxisIndex > 2 Then Exit Function

    Dim vx As Double, vy As Double, vz As Double
    vx = CDbl(rot(0, ThinAxisIndex))
    vy = CDbl(rot(1, ThinAxisIndex))
    vz = CDbl(rot(2, ThinAxisIndex))

    planarError = Sqr(vx * vx + vy * vy)
    axisAlignment = Abs(vz)
    EvaluateThinAxisAlignment = True
    On Error GoTo 0
End Function

Private Function EvaluateFaceAlignment(rot As Variant, _
                                       normalX As Double, _
                                       normalY As Double, _
                                       normalZ As Double, _
                                       ByRef planarError As Double, _
                                       ByRef axisComponent As Double) As Double
    On Error Resume Next
    If IsEmpty(rot) Then Exit Function

    Dim ax As Double, ay As Double, az As Double
    ax = CDbl(rot(0, 0)) * normalX + CDbl(rot(0, 1)) * normalY + CDbl(rot(0, 2)) * normalZ
    ay = CDbl(rot(1, 0)) * normalX + CDbl(rot(1, 1)) * normalY + CDbl(rot(1, 2)) * normalZ
    az = CDbl(rot(2, 0)) * normalX + CDbl(rot(2, 1)) * normalY + CDbl(rot(2, 2)) * normalZ

    planarError = Sqr(ax * ax + ay * ay)
    axisComponent = az

    Dim mag As Double: mag = Sqr(ax * ax + ay * ay + az * az)
    If mag > 0# Then
        EvaluateFaceAlignment = Abs(az / mag)
    End If
    On Error GoTo 0
End Function

Private Function OrientationMatrixScore(planarError As Double, axisAlignment As Double) As Double
    OrientationMatrixScore = planarError * 1000# + (1# - axisAlignment)
End Function

Private Function TryGetLargestPlanarFaceNormal(partDoc As SldWorks.ModelDoc2, _
                                               ByRef normalX As Double, _
                                               ByRef normalY As Double, _
                                               ByRef normalZ As Double, _
                                               ByRef faceArea As Double) As Boolean
    On Error Resume Next

    If partDoc Is Nothing Then GoTo done
    If partDoc.GetType <> swDocPART Then GoTo done

    Dim part As SldWorks.PartDoc
    Set part = partDoc
    If part Is Nothing Then GoTo done

    Dim faces As Variant
    faces = part.GetFaces
    If IsEmpty(faces) Then GoTo done
    If Not IsArray(faces) Then GoTo done

    Dim bestArea As Double: bestArea = 0#
    Dim i As Long
    For i = LBound(faces) To UBound(faces)
        Dim face As SldWorks.Face2
        Set face = faces(i)
        If Not face Is Nothing Then
            Dim surf As SldWorks.Surface
            Set surf = face.GetSurface
            If Not surf Is Nothing Then
                If CBool(surf.IsPlane) Then
                    Dim area As Double: area = Abs(face.GetArea)
                    If area > bestArea Then
                        Dim params As Variant
                        params = surf.PlaneParams
                        If IsArray(params) Then
                            If UBound(params) >= 5 Then
                                Dim nx As Double: nx = CDbl(params(3))
                                Dim ny As Double: ny = CDbl(params(4))
                                Dim nz As Double: nz = CDbl(params(5))
                                Dim mag As Double: mag = Sqr(nx * nx + ny * ny + nz * nz)
                                If mag > 0# Then
                                    nx = nx / mag
                                    ny = ny / mag
                                    nz = nz / mag
                                    bestArea = area
                                    normalX = nx
                                    normalY = ny
                                    normalZ = nz
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If bestArea > 0# Then
        faceArea = bestArea
        TryGetLargestPlanarFaceNormal = True
    End If

done:
    On Error GoTo 0
End Function

Private Function CreateTransformFromMatrix(baseData As Variant, _
                                          newR As Variant, _
                                          mathUtil As SldWorks.MathUtility) As SldWorks.MathTransform
    On Error Resume Next
    If mathUtil Is Nothing Then Exit Function
    If IsEmpty(baseData) Then Exit Function
    If IsEmpty(newR) Then Exit Function
    If UBound(baseData) < 14 Then Exit Function

    Dim arr(0 To 15) As Double
    arr(0) = CDbl(newR(0, 0))
    arr(1) = CDbl(newR(1, 0))
    arr(2) = CDbl(newR(2, 0))
    arr(3) = 0#
    arr(4) = CDbl(newR(0, 1))
    arr(5) = CDbl(newR(1, 1))
    arr(6) = CDbl(newR(2, 1))
    arr(7) = 0#
    arr(8) = CDbl(newR(0, 2))
    arr(9) = CDbl(newR(1, 2))
    arr(10) = CDbl(newR(2, 2))
    arr(11) = 0#
    arr(12) = CDbl(baseData(12))
    arr(13) = CDbl(baseData(13))
    arr(14) = CDbl(baseData(14))
    arr(15) = 1#

    Set CreateTransformFromMatrix = mathUtil.CreateTransform(arr)
    On Error GoTo 0
End Function

Private Function EvaluateOrientationMetrics(comp As SldWorks.Component2, _
                                            pi As clsPlaceItem, _
                                            ByRef isZThin As Boolean, _
                                            ByRef thicknessDiff As Double, _
                                            ByRef measuredThickness As Double, _
                                            ByRef zAxisDelta As Double, _
                                            ByRef planeGapIn As Double) As Boolean
    On Error Resume Next

    EnsureResolved comp

    Dim box As Variant
    box = SafeGetBox(comp)
    If Not IsValidBox(box) Then
        EvaluateOrientationMetrics = False
        On Error GoTo 0
        Exit Function
    End If

    Dim minX As Double: minX = CDbl(box(0))
    Dim minY As Double: minY = CDbl(box(1))
    Dim minZ As Double: minZ = CDbl(box(2))
    Dim maxX As Double: maxX = CDbl(box(3))
    Dim maxY As Double: maxY = CDbl(box(4))
    Dim maxZ As Double: maxZ = CDbl(box(5))

    Dim spanXIn As Double: spanXIn = Abs(maxX - minX) * M_TO_IN
    Dim spanYIn As Double: spanYIn = Abs(maxY - minY) * M_TO_IN
    Dim spanZIn As Double: spanZIn = Abs(maxZ - minZ) * M_TO_IN

    Dim minDim As Double: minDim = Min3(spanXIn, spanYIn, spanZIn)
    measuredThickness = minDim
    zAxisDelta = Abs(spanZIn - minDim)
    isZThin = (zAxisDelta <= ORIENTATION_AXIS_TOL_IN)

    planeGapIn = Abs(minZ) * M_TO_IN

    If pi.ThicknessIn > 0# Then
        thicknessDiff = Abs(minDim - pi.ThicknessIn)
    Else
        thicknessDiff = 0#
    End If

    EvaluateOrientationMetrics = True
    On Error GoTo 0
End Function

' =========================
'        DXF EXPORT (Top-only, 1:1)
' =========================
Private Sub ExportModelTopDXF(swApp As SldWorks.SldWorks, _
                              drwTplDefault As String, _
                              modelRefPath As String, _
                              outDXF As String)

    On Error Resume Next

    ' 1) Choose drawing template
    Dim drwTplToUse As String
    drwTplToUse = Trim$(DRAWING_TEMPLATE_OVERRIDE)
    If Len(drwTplToUse) > 0 Then
        If Dir(drwTplToUse) = "" Then
            LogMessage "[DXF] Drawing template override not found: " & drwTplToUse & " (falling back to default)."
            drwTplToUse = drwTplDefault
        Else
            LogMessage "[DXF] Using drawing template override: " & drwTplToUse
        End If
    Else
        drwTplToUse = drwTplDefault
    End If

    LogMessage "[DXF] Preparing drawing for " & modelRefPath & " -> " & outDXF

    g_LastStep = "[DXF] NewDocument"
    Dim drw As SldWorks.ModelDoc2: Set drw = swApp.NewDocument(drwTplToUse, 0, 0, 0)
    If drw Is Nothing Then
        LogMessage "Failed to open drawing template for DXF export: " & outDXF, True
        Exit Sub
    End If
    If drw.GetType <> swDocDRAWING Then
        LogMessage "Template mismatch: drawing template is not .drwdot", True
        drw.Quit: Exit Sub
    End If

    ' ---- Force IPS immediately on the new drawing
    ForceUnitsIPS drw

    Dim dd As SldWorks.DrawingDoc: Set dd = drw

    ' 2) Apply sheet format if reachable
    Dim fmt As String: fmt = ResolveSheetFormatPath()
    If Len(fmt) > 0 Then
        LogMessage "[DXF] Applying sheet format: " & fmt
        ApplySheetFormat dd, fmt
        ' Some templates flip units back; re-assert IPS after applying SLDDRT
        ForceUnitsIPS drw
    Else
        LogMessage "[DXF] Sheet format not found: " & SHEET_FORMAT_PATH & " (using default)."
        LogDriveMappings
    End If

    ' 3) Create ONLY a Top view, at 1:1 (scale parameter = 1#)
    g_LastStep = "[DXF] CreateDrawViewFromModelView3(*Top)"
    Dim topV As SldWorks.View
    Set topV = dd.CreateDrawViewFromModelView3(modelRefPath, "*Top", 0.3, 0.22, 1#)
    If topV Is Nothing Then
        Set topV = dd.CreateDrawViewFromModelView3(modelRefPath, "Top", 0.3, 0.22, 1#)
    End If
    If topV Is Nothing Then
        LogMessage "Could not create Top view for " & modelRefPath, True
        drw.Quit
        Exit Sub
    End If

    ' Force 1:1 via the correct property (some versions ignore the ctor scale)
    On Error Resume Next
    topV.ScaleDecimal = 1#
    On Error GoTo 0

    ' 4) Remove any other views
    DeleteAllViewsExcept dd, topV.Name

    ' 5) Save DXF
    g_LastStep = "[DXF] SaveAs4"
    Dim errs As Long, warns As Long
    drw.SaveAs4 outDXF, swSaveAsCurrentVersion, swSaveAsOptions_Silent, errs, warns
    If errs <> 0 Then LogMessage "DXF export error code: " & errs & " for " & outDXF
    drw.Quit

    On Error GoTo 0
End Sub

' ---- Helpers for DXF/Sheet Format ----

' Single, definitive implementation (do not duplicate)
Private Function ResolveSheetFormatPath() As String
    Dim p As String: p = Trim$(SHEET_FORMAT_PATH)
    If Len(p) = 0 Then Exit Function

    LogMessage "[DXF] Checking sheet format path: " & p
    If Dir$(p) <> "" Then
        ResolveSheetFormatPath = p
    Else
        ResolveSheetFormatPath = ""   ' fall back to template default
    End If
End Function

' Apply .slddrt to current sheet
Private Sub ApplySheetFormat(dd As SldWorks.DrawingDoc, fmtPath As String)
    On Error Resume Next
    Dim sh As SldWorks.Sheet: Set sh = dd.GetCurrentSheet
    If sh Is Nothing Then Exit Sub

    CallByName sh, "SetTemplateName2", VbMethod, fmtPath
    CallByName sh, "ReloadTemplate", VbMethod, True

    ' Fallback for older versions: reinforce via SetupSheet5
    Dim nm As String: nm = CallByName(sh, "GetName", VbMethod)
    If Len(nm) > 0 Then
        CallByName dd, "SetupSheet5", VbMethod, nm, fmtPath, 0, 0#, 0#, 1#, 1#, False, "", 0#, 0#
    End If
    On Error GoTo 0
End Sub

' Delete all model views except one
Private Sub DeleteAllViewsExcept(dd As SldWorks.DrawingDoc, keepName As String)
    On Error Resume Next
    Dim sheetView As SldWorks.View: Set sheetView = dd.GetFirstView
    If sheetView Is Nothing Then Exit Sub

    Dim names As Collection: Set names = New Collection
    Dim v As SldWorks.View: Set v = sheetView.GetNextView
    Do While Not v Is Nothing
        If StrComp(v.Name, keepName, vbTextCompare) <> 0 Then
            names.Add v.Name
        End If
        Set v = v.GetNextView
    Loop

    Dim i As Long
    For i = 1 To names.Count
        If Not CallByName(dd, "DeleteView", VbMethod, names(i)) Then
            dd.ActivateView names(i)
            Dim md As SldWorks.ModelDoc2: Set md = dd
            md.SelectByID2 names(i), "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
            md.DeleteSelection2 0
        End If
    Next
    On Error GoTo 0
End Sub

' =========================
'     UNITS: force IPS
' =========================
Private Sub ForceUnitsIPS(md As SldWorks.ModelDoc2)
    On Error Resume Next

    ' A) Preferred: document-level prefs via ModelDocExtension
    Dim ext As SldWorks.ModelDocExtension
    Set ext = md.Extension
    If Not ext Is Nothing Then
        ext.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swUnitSystem, _
            swUnitSystem_e.swUnitSystem_IPS

        ext.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swUnitsLinear, _
            swLengthUnit_e.swINCHES

        ext.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swUnitsAngular, _
            swAngleUnit_e.swDEGREES

        ext.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swUnitsLinearDecimalPlaces, 3
        ext.SetUserPreferenceIntegerValue _
            swUserPreferenceIntegerValue_e.swUnitsAngularDecimalPlaces, 2
    End If

    ' B) Supplemental: legacy setters (use 0 for "decimal" display)
    md.SetUnits swINCHES, 0, 0, 3, False
    md.SetAngularUnits swDEGREES, 0, 0, 2

    md.ForceRebuild3 False
    On Error GoTo 0
End Sub

' =========================
'          HELPERS
' =========================
Private Function TryGetBBoxInches(ByVal c As Object, _
                                  ByRef dxIn As Double, _
                                  ByRef dyIn As Double, _
                                  ByRef dzIn As Double, _
                                  Optional ByVal md As Object = Nothing) As Boolean
    On Error Resume Next
    Dim v As Variant

    v = SafeGetBox(c)
    If IsValidBox(v) Then GoTo hasBox

    If md Is Nothing Then Set md = CallByName(c, "GetModelDoc2", VbMethod)
    If Not md Is Nothing Then
        v = CallByName(md, "GetBox", VbMethod)
        If IsValidBox(v) Then GoTo hasBox

        Dim ext As Object: Set ext = CallByName(md, "Extension", VbGet)
        If Not ext Is Nothing Then
            v = CallByName(ext, "GetBox", VbMethod)
            If IsValidBox(v) Then GoTo hasBox
        End If

        Dim bodies As Variant
        bodies = CallByName(md, "GetBodies2", VbMethod, 0, True) ' 0=Solid
        If IsArray(bodies) Then
            Dim haveAny As Boolean
            Dim minx As Double, miny As Double, minz As Double
            Dim maxx As Double, maxy As Double, maxz As Double
            minx = 1E+99: miny = 1E+99: minz = 1E+99
            maxx = -1E+99: maxy = -1E+99: maxz = -1E+99

            Dim i As Long, bb As Variant
            For i = LBound(bodies) To UBound(bodies)
                bb = CallByName(bodies(i), "GetBodyBox", VbMethod)
                If IsValidBox(bb) Then
                    haveAny = True
                    If CDbl(bb(0)) < minx Then minx = CDbl(bb(0))
                    If CDbl(bb(1)) < miny Then miny = CDbl(bb(1))
                    If CDbl(bb(2)) < minz Then minz = CDbl(bb(2))
                    If CDbl(bb(3)) > maxx Then maxx = CDbl(bb(3))
                    If CDbl(bb(4)) > maxy Then maxy = CDbl(bb(4))
                    If CDbl(bb(5)) > maxz Then maxz = CDbl(bb(5))
                End If
            Next
            If haveAny Then
                v = Array(minx, miny, minz, maxx, maxy, maxz)
                GoTo hasBox
            End If
        End If
    End If

    TryGetBBoxInches = False
    On Error GoTo 0
    Exit Function

hasBox:
    dxIn = Abs(CDbl(v(3)) - CDbl(v(0))) * M_TO_IN
    dyIn = Abs(CDbl(v(4)) - CDbl(v(1))) * M_TO_IN
    dzIn = Abs(CDbl(v(5)) - CDbl(v(2))) * M_TO_IN
    TryGetBBoxInches = (dxIn > 0# Or dyIn > 0# Or dzIn > 0#)
    On Error GoTo 0
End Function

Private Function IsValidBox(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If Not IsArray(v) Then Exit Function
    If UBound(v) < 5 Then Exit Function
    Dim i As Long
    For i = 0 To 5
        If Not IsNumeric(v(i)) Then Exit Function
    Next i
    IsValidBox = True
End Function

Private Function DetermineThinAxisIndex(md As SldWorks.ModelDoc2, _
                                        ByVal fallbackX As Double, _
                                        ByVal fallbackY As Double, _
                                        ByVal fallbackZ As Double) As Long
    Dim dx As Double, dy As Double, dz As Double
    If Not md Is Nothing Then
        If TryGetBBoxInches(md, dx, dy, dz, md) Then
            DetermineThinAxisIndex = IndexOfMin3(dx, dy, dz)
            Exit Function
        End If
    End If
    DetermineThinAxisIndex = IndexOfMin3(fallbackX, fallbackY, fallbackZ)
End Function

Private Function IndexOfMin3(a As Double, b As Double, c As Double) As Long
    Dim ax As Double: ax = Abs(a)
    Dim ay As Double: ay = Abs(b)
    Dim az As Double: az = Abs(c)
    If ax <= 0# And ay <= 0# And az <= 0# Then
        IndexOfMin3 = -1
        Exit Function
    End If

    Dim idx As Long: idx = 0
    Dim minVal As Double: minVal = ax
    If ay < minVal Then
        minVal = ay
        idx = 1
    End If
    If az < minVal Then idx = 2
    IndexOfMin3 = idx
End Function

Private Function Min3(a As Double, b As Double, c As Double) As Double
    Dim m As Double: m = a
    If b < m Then m = b
    If c < m Then m = c
    Min3 = m
End Function

Public Function GetFileName(p As String) As String
    Dim i As Long: i = InStrRev(p, "\")
    If i > 0 Then GetFileName = Mid$(p, i + 1) Else GetFileName = p
End Function

Public Function GetFileBaseName(p As String) As String
    Dim nm As String: nm = GetFileName(p)
    Dim dotPos As Long: dotPos = InStrRev(nm, ".")
    If dotPos > 0 Then
        GetFileBaseName = Left$(nm, dotPos - 1)
    Else
        GetFileBaseName = nm
    End If
End Function

Private Function GetParentFolder(p As String) As String
    Dim i As Long: i = InStrRev(p, "\")
    If i > 0 Then GetParentFolder = Left$(p, i - 1) Else GetParentFolder = CurDir$
End Function

Private Sub EnsureFolder(f As String)
    If Dir(f, vbDirectory) = "" Then MkDir f
End Sub

Private Function SanitizeFileName(s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, CStr(bad(i)), "_")
    Next
    SanitizeFileName = s
End Function

Private Function BuildModelReferenceForDrawing(modelPath As String, cfg As String) As String
    Dim p As String: p = Trim$(modelPath)
    If Len(p) = 0 Then Exit Function

    Dim c As String: c = Trim$(cfg)
    If Len(c) > 0 Then
        BuildModelReferenceForDrawing = p & "@" & c
    Else
        BuildModelReferenceForDrawing = p
    End If
End Function

Private Function UniqueTargetPath(ByVal desired As String) As String
    Dim p As String, e As String, n As String, idx As Long
    p = desired
    If Dir(p) = "" Then UniqueTargetPath = p: Exit Function

    Dim dotPos As Long: dotPos = InStrRev(p, ".")
    If dotPos = 0 Then
        n = p: e = ""
    Else
        n = Left$(p, dotPos - 1): e = Mid$(p, dotPos)
    End If

    idx = 2
    Do
        p = n & " (" & idx & ")" & e
        idx = idx + 1
    Loop While Dir(p) <> ""

    UniqueTargetPath = p
End Function

Private Function ShortFolder(ByVal p As String) As String
    Dim dirOnly As String: dirOnly = GetParentFolder(p)
    Dim i As Long: i = InStrRev(dirOnly, "\")
    If i > 0 Then ShortFolder = Mid$(dirOnly, i + 1) Else ShortFolder = dirOnly
End Function

Public Function BuildDisplayText(pr As clsPartRecord) As String
    BuildDisplayText = GetFileName(pr.FullPath) & " (" & pr.Config & ")  [" & ShortFolder(pr.FullPath) & "]"
End Function

Private Sub DumpAllPartsForUI()
    Dim i As Long
    For i = 1 To g_AllParts.Count
        Dim pr As clsPartRecord: Set pr = g_AllParts(i)
        LogMessage "[UI] " & i & " -> " & pr.FullPath
    Next
End Sub

' ========= PUBLIC (used by entry point) =========
Public Function MakePlacementList(thkGroup As Collection) As Collection
    Dim L As New Collection
    Dim i As Long
    For i = 1 To thkGroup.Count
        Dim pr As clsPartRecord: Set pr = thkGroup(i)

        Dim pi As New clsPlaceItem
        pi.FullPath = pr.FullPath
        pi.Config = pr.Config
        pi.Count = pr.Qty

        Dim dimsOriginal(0 To 2) As Double
        dimsOriginal(0) = Abs(pr.BBoxX)
        dimsOriginal(1) = Abs(pr.BBoxY)
        dimsOriginal(2) = Abs(pr.BBoxZ)

'codex/fix-compile-error-at-thinaxisindex-5tl2al
        Dim thinIdx As Long: thinIdx = pr.ThinAxisIndex
        If thinIdx < 0 Or thinIdx > 2 Then
            thinIdx = IndexOfMin3(dimsOriginal(0), dimsOriginal(1), dimsOriginal(2))
        End If
        pi.thinAxis = thinIdx

        If pr.ThickIn > 0# Then
            pi.ThicknessIn = pr.ThickIn
        ElseIf thinIdx >= 0 And thinIdx <= 2 Then
            pi.ThicknessIn = dimsOriginal(thinIdx)
        Else
            pi.ThicknessIn = Min3(dimsOriginal(0), dimsOriginal(1), dimsOriginal(2))
        End If
        If pi.ThicknessIn <= 0# Then pi.ThicknessIn = 0.01

        Dim dims(0 To 2) As Double
        Dim j As Long, k As Long, tmp As Double
        For j = 0 To 2
            dims(j) = dimsOriginal(j)
        Next j
        For j = 0 To 1
            For k = j + 1 To 2
                If dims(k) > dims(j) Then
                    tmp = dims(j)
                    dims(j) = dims(k)
                    dims(k) = tmp
                End If
            Next k
        Next j

        pi.WidthIn = dims(0)
        pi.HeightIn = dims(1)
        If pi.WidthIn <= 0# Then
            If pr.ThickIn > 0# Then
                pi.WidthIn = pr.ThickIn
            Else
                pi.WidthIn = 0.01
            End If
            LogMessage "[WARN] Width fallback for " & pr.FullPath & " (" & pr.Config & ")"
        End If
        If pi.HeightIn <= 0# Then
            If pr.ThickIn > 0# Then
                pi.HeightIn = pr.ThickIn
            Else
                pi.HeightIn = 0.01
            End If
            LogMessage "[WARN] Height fallback for " & pr.FullPath & " (" & pr.Config & ")"
        End If

        L.Add pi
    Next
    Set MakePlacementList = L
End Function

Private Sub WriteQuantityReportForGroup(thkGroup As Collection, reportPath As String)
    On Error GoTo fail

    Dim fnum As Integer
    fnum = FreeFile
    Open reportPath For Output As #fnum
    Print #fnum, "Part,Configuration,Quantity"

    Dim i As Long
    For i = 1 To thkGroup.Count
        Dim pr As clsPartRecord: Set pr = thkGroup(i)
        Print #fnum, GetFileName(pr.FullPath) & "," & pr.Config & "," & CStr(pr.Qty)
    Next i

    Close #fnum
    LogMessage "[TXT] Wrote quantity report -> " & reportPath
    On Error GoTo 0
    Exit Sub

' codex/fix-compile-error-at-thinaxisindex-5tl2al
fail:
    Dim errMsg As String: errMsg = Err.Description
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
    On Error GoTo 0
    LogMessage "[WARN] Failed to write quantity report: " & reportPath & " (" & errMsg & ")", True
End Sub

' Export DXFs for each unique part in a thickness group into the specified folder
Private Sub ExportIndividualPartDXFs(swApp As SldWorks.SldWorks, _
                                     drwTplDefault As String, _
                                     thkGroup As Collection, _
                                     groupFolder As String)

    If thkGroup Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To thkGroup.Count
        Dim pr As clsPartRecord: Set pr = thkGroup(i)
        If pr Is Nothing Then GoTo cont
        If Len(Trim$(pr.FullPath)) = 0 Then
            LogMessage "[DXF] Skip: empty path for part index " & CStr(i)
            GoTo cont
        End If

        Dim baseName As String
        baseName = GetFileBaseName(pr.FullPath)
        If Len(Trim$(pr.Config)) > 0 Then baseName = baseName & " - " & pr.Config

        Dim partDxfBase As String
        partDxfBase = groupFolder & "\" & SanitizeFileName(baseName) & ".DXF"
        Dim partDxfPath As String
        partDxfPath = UniqueTargetPath(partDxfBase)

        Dim modelRef As String
        modelRef = BuildModelReferenceForDrawing(pr.FullPath, pr.Config)
        If Len(modelRef) = 0 Then
            LogMessage "[DXF] Skip: unable to determine model reference for " & pr.DisplayName
            GoTo cont
        End If

        LogMessage "[DXF] Exporting part -> " & partDxfPath
        ExportModelTopDXF swApp, drwTplDefault, modelRef, partDxfPath
cont:
    Next i
End Sub


