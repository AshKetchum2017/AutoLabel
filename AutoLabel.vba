Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String


Private Const REG_APP_NAME As String = "RinCorelMacros"
Private Const REG_SECTION_NAME As String = "AutoLabel"
Private Const LABEL_TAG_X As Double = 306.2
Private Const LABEL_TAG_Y As Double = 8.8
Private Const UNI_LABEL_GAP_X As Double = 3#
Private Const UNI_LABEL_GAP_Y As Double = 0#
Private Const TAG_Y_OFFSET As Double = 0#
Private Const LABEL_HI_DIE_X As Double = 136#
Private Const LABEL_HI_DIE_Y As Double = 6.2
Private Const HI_DIE_GAP_X As Double = 5#
Private Const HI_DIE_GAP_Y As Double = 0.5
Private Const HI_DIE_Y_OFFSET As Double = 0#
Private Const HI_KISS_GAP_X As Double = 0#
Private Const HI_KISS_GAP_Y As Double = 10#
Private Const KISS_LABEL_TAG_X As Double = 315.2
Private Const KISS_LABEL_TAG_Y As Double = 3.8
Private Const KISS_LABEL_GAP_X As Double = 0#
Private Const KISS_LABEL_GAP_Y As Double = 0.5
Private Const MIN_FONT_SIZE As Long = 4
Private Const MAX_FONT_SIZE As Long = 24
Private Const UNI_LABEL_FONT_SIZE As Single = 8!
Private Const HI_DIE_FONT_SIZE As Single = 12!
Private Const HI_KISS_FONT_SIZE As Single = 12!
Private Const KISS_LABEL_FONT_SIZE As Single = 4!
Private Const PREVIEW_FONT_SIZE As Single = 10!

Private Sub chkMasterPage_Click()

End Sub

Private Sub cmdSaveConfig_Click()

    SaveCurrentConfig
    MsgBox "Pengaturan Auto Label berhasil disimpan.", vbInformation

End Sub

Private Sub UserForm_Initialize()

    Dim fontSize As Long

    cmbFontSize.Clear
    For fontSize = MIN_FONT_SIZE To MAX_FONT_SIZE
        cmbFontSize.AddItem CStr(fontSize)
    Next fontSize

    If optHiDie.value Then
        cmbFontSize.Text = CStr(HI_DIE_FONT_SIZE)
    Else
        cmbFontSize.Text = CStr(UNI_LABEL_FONT_SIZE)
    End If
    LoadSavedConfig
    UpdatePreview

End Sub

Private Sub LoadSavedConfig()

    chkMasterPage.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkMasterPage", CStr(chkMasterPage.value)))
    chkUppercase.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkUppercase", CStr(chkUppercase.value)))
    chkBold.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkBold", CStr(chkBold.value)))
    chkItalic.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkItalic", CStr(chkItalic.value)))
    chkSmallCaps.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkSmallCaps", CStr(chkSmallCaps.value)))
    chkColorize.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkColorize", CStr(chkColorize.value)))
    chkFraction.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkFraction", CStr(chkFraction.value)))
    chkEmDash.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "chkEmDash", CStr(chkEmDash.value)))
    optKissLabel.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "optKissLabel", CStr(optKissLabel.value)))
    optHiDie.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "optHiDie", CStr(optHiDie.value)))
    optHiKiss.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "optHiKiss", CStr(optHiKiss.value)))
    optUniLabel.value = CBool(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "optUniLabel", CStr(optUniLabel.value)))

    LoadSavedFontSize

End Sub

Private Sub LoadSavedFontSize()

    Dim savedFontSize As Single

    savedFontSize = Val(GetSetting(REG_APP_NAME, REG_SECTION_NAME, "cmbFontSize", cmbFontSize.Text))
    If savedFontSize >= MIN_FONT_SIZE And savedFontSize <= MAX_FONT_SIZE Then
        cmbFontSize.Text = CStr(savedFontSize)
    End If

End Sub

Private Sub SaveCurrentConfig()

    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkMasterPage", CStr(chkMasterPage.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkUppercase", CStr(chkUppercase.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkBold", CStr(chkBold.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkItalic", CStr(chkItalic.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkSmallCaps", CStr(chkSmallCaps.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkColorize", CStr(chkColorize.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkFraction", CStr(chkFraction.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "chkEmDash", CStr(chkEmDash.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "optKissLabel", CStr(optKissLabel.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "optHiDie", CStr(optHiDie.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "optHiKiss", CStr(optHiKiss.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "optUniLabel", CStr(optUniLabel.value)
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "cmbFontSize", Trim$(cmbFontSize.Text)

End Sub

Private Sub chkBold_Click()

UpdatePreview

End Sub

Private Sub chkColorize_Click()

    UpdatePreview

End Sub

Private Sub chkEmDash_Click()

    UpdatePreview

End Sub

Private Sub chkFraction_Click()

    UpdatePreview

End Sub

Private Sub chkItalic_Click()

    UpdatePreview

End Sub

Private Sub chkSmallCaps_Click()

    UpdatePreview

End Sub

Private Sub chkUppercase_Click()

    UpdatePreview

End Sub

Private Sub cmbFontSize_Change()

    UpdatePreview

End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdSubmit_Click()

    Dim s As Shape
    Dim labelText As String
    Dim fontSize As Single
    Dim previousUnit As cdrUnit
    Dim targetLayer As Layer
    Dim textX As Double
    Dim textY As Double
    Dim textAlignment As Long

    labelText = GetLabelText()
    If Len(labelText) = 0 Then
        txtText.SetFocus
        Exit Sub
    End If

    fontSize = SubmitFontSize()
    textX = SubmitX()
    textY = SubmitY()
    textAlignment = SubmitAlignment()

    On Error GoTo SubmitError
    previousUnit = ActiveDocument.Unit
    ActiveDocument.Unit = cdrMillimeter
    Set targetLayer = GetSubmitLayer()

    Set s = targetLayer.CreateArtisticText( _
        textX, _
        textY, _
        labelText, _
        cdrLanguageNone, _
        cdrCharSetMixed, _
        txtText.Font.Name, _
        fontSize, _
        IIf(chkBold.value, cdrTrue, cdrFalse), _
        IIf(chkItalic.value, cdrTrue, cdrFalse), _
        cdrMixedFontLine, _
        textAlignment)

    ApplyLabelTextFormatting s, labelText
    ApplySubmitPosition s

    ActiveDocument.Unit = previousUnit

    Unload Me
    Exit Sub

SubmitError:
    ActiveDocument.Unit = previousUnit
    MsgBox "Auto Label gagal dibuat: " & Err.Description, vbExclamation

End Sub

Private Function GetSubmitLayer() As Layer

    If chkMasterPage.value Then
        Set GetSubmitLayer = GetOrCreatePageLayer(ActiveDocument.MasterPage, "Layer 1")
    Else
        Set GetSubmitLayer = GetOrCreatePageLayer(ActivePage, "Layer 1")
    End If

End Function

Private Function GetOrCreatePageLayer(ByVal targetPage As Page, ByVal layerName As String) As Layer

    Dim targetLayer As Layer

    On Error Resume Next
    Set targetLayer = targetPage.Layers(layerName)
    On Error GoTo 0

    If targetLayer Is Nothing Then
        Set targetLayer = targetPage.CreateLayer(layerName)
    End If

    Set GetOrCreatePageLayer = targetLayer

End Function

Private Sub lblFontSize_Click()

End Sub

Private Sub optHiDie_Click()

    cmbFontSize.Text = CStr(HI_DIE_FONT_SIZE)
    UpdatePreview

End Sub

Private Sub optHiKiss_Click()

    cmbFontSize.Text = CStr(HI_KISS_FONT_SIZE)
    UpdatePreview

End Sub

Private Sub optUniLabel_Click()

    cmbFontSize.Text = CStr(UNI_LABEL_FONT_SIZE)
    UpdatePreview

End Sub

Private Sub optKissLabel_Click()

    cmbFontSize.Text = CStr(KISS_LABEL_FONT_SIZE)
    UpdatePreview

End Sub

Private Sub txtText_Change()

    UpdatePreview

End Sub

Private Sub UpdatePreview()

    With txtText
        .Font.Bold = chkBold.value
        .Font.Italic = chkItalic.value
        .Font.Size = PREVIEW_FONT_SIZE
    End With

End Sub

Private Function GetLabelText() As String

    Dim value As String

    value = txtText.Text

    If chkEmDash.value Then
        value = Replace(value, " - ", " " & ChrW$(8212) & " ", 1, 1, vbTextCompare)
    End If

    If chkUppercase.value Or optKissLabel.value Then
        value = UCase$(value)
    End If

    GetLabelText = Trim$(value)

End Function

Private Function SubmitFontSize() As Single

    If optKissLabel.value Then
        SubmitFontSize = KISS_LABEL_FONT_SIZE
    Else
        SubmitFontSize = SelectedFontSize()
    End If

End Function

Private Function FindAnchor( _
    ByVal layerName As String, _
    ByVal anchorName As String) As Shape

    Dim lyr As Layer
    Dim shp As Shape
    Dim wasEditable As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    On Error Resume Next
    Set lyr = ActivePage.Layers(layerName)
    On Error GoTo 0

    If lyr Is Nothing Then Exit Function

    wasEditable = lyr.Editable
    On Error GoTo RestoreLayer
    lyr.Editable = True

    For Each shp In lyr.Shapes

        If LCase$(shp.Name) = LCase$(anchorName) Then
            Set FindAnchor = shp
            Exit For
        End If

    Next shp

RestoreLayer:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source
    On Error Resume Next
    lyr.Editable = wasEditable
    On Error GoTo 0

    If errorNumber <> 0 Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If

End Function

Private Function FindHiDieAnchor() As Shape

    Dim lyr As Layer
    Dim shp As Shape
    Dim wasEditable As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String

    On Error Resume Next
    Set lyr = ActivePage.Layers("Regmark")
    On Error GoTo 0

    If lyr Is Nothing Then Exit Function

    wasEditable = lyr.Editable
    On Error GoTo RestoreLayer
    lyr.Editable = True
    
    For Each shp In lyr.Shapes

        If shp.Type = cdrRectangleShape Then
            If shp.SizeWidth >= 100# _
               And shp.SizeWidth <= 150# _
               And shp.SizeHeight >= 0.8# _
               And shp.SizeHeight <= 1.2# Then

                Set FindHiDieAnchor = shp
                Exit For

            End If
        End If

    Next shp

RestoreLayer:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source
    On Error Resume Next
    lyr.Editable = wasEditable
    On Error GoTo 0

    If errorNumber <> 0 Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If

End Function

Private Function FindHiKissAnchor() As Shape

    Dim lyr As Layer
    Dim shp As Shape
    Dim wasEditable As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim errorSource As String
    Dim matchingCount As Long
    Dim firstMatch As Shape
    Dim shapeIndex As Long

    On Error Resume Next
    Set lyr = ActivePage.Layers("Regmark")
    On Error GoTo 0

    If lyr Is Nothing Then Exit Function

    wasEditable = lyr.Editable
    On Error GoTo RestoreLayer
    lyr.Editable = True
    
    For shapeIndex = lyr.Shapes.Count To 1 Step -1
        Set shp = lyr.Shapes(shapeIndex)

        If shp.Type = cdrRectangleShape Then
            If shp.SizeWidth = 3# _
               And shp.SizeHeight = 3# Then

                matchingCount = matchingCount + 1
                If matchingCount = 1 Then
                    Set firstMatch = shp
                ElseIf matchingCount = 2 Then
                    Set FindHiKissAnchor = shp
                    Exit For
                End If
            End If
        End If

    Next shapeIndex

RestoreLayer:
    errorNumber = Err.Number
    errorDescription = Err.Description
    errorSource = Err.Source
    On Error Resume Next
    lyr.Editable = wasEditable
    On Error GoTo 0

    If FindHiKissAnchor Is Nothing Then
        Set FindHiKissAnchor = firstMatch
    End If

    If errorNumber <> 0 Then
        Err.Raise errorNumber, errorSource, errorDescription
    End If

End Function

Private Function SubmitX() As Double
    Dim anchor As Shape

    If optKissLabel.Value Then
        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            SubmitX = KISS_LABEL_TAG_X
        Else
            SubmitX = anchor.LeftX + KISS_LABEL_GAP_X
        End If

    ElseIf optHiDie.Value Then
        Set anchor = FindHiDieAnchor()

        If anchor Is Nothing Then
            SubmitX = LABEL_HI_DIE_X
        Else
            ' 5 mm dari sisi kanan anchor
            SubmitX = anchor.RightX + HI_DIE_GAP_X
        End If

    ElseIf optHiKiss.Value Then
        Set anchor = FindHiKissAnchor()

        If anchor Is Nothing Then
            SubmitX = LABEL_TAG_X
        Else
            ' 5 mm dari sisi kanan anchor
            SubmitX = anchor.LeftX + HI_KISS_GAP_X
        End If
        
    Else
        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            SubmitX = LABEL_TAG_X
        Else
            SubmitX = anchor.LeftX + UNI_LABEL_GAP_X
        End If
    End If
End Function

Private Function SubmitY() As Double
    Dim anchor As Shape

    If optKissLabel.value Then
        Set anchor = FindAnchor("scpro2_printonly", "filename")
        
        If anchor Is Nothing Then
            SubmitY = KISS_LABEL_TAG_Y
        Else
            SubmitY = anchor.BottomY - KISS_LABEL_GAP_Y + TAG_Y_OFFSET
        End If

    ElseIf optHiDie.value Then
        Set anchor = FindHiDieAnchor()

        If anchor Is Nothing Then
            SubmitY = LABEL_HI_DIE_Y
        Else
            SubmitY = anchor.BottomY + HI_DIE_GAP_Y + HI_DIE_Y_OFFSET
        End If
    
    ElseIf optHiKiss.value Then
        Set anchor = FindHiKissAnchor()

        If anchor Is Nothing Then
            SubmitY = LABEL_TAG_Y
        Else
            SubmitY = anchor.BottomY + HI_KISS_GAP_Y + TAG_Y_OFFSET
        End If

    Else
        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            SubmitY = LABEL_TAG_Y
        Else
            SubmitY = anchor.BottomY + UNI_LABEL_GAP_Y + TAG_Y_OFFSET
        End If
    End If

End Function

Private Function SubmitAlignment() As Long

    If optKissLabel.value Or optHiDie.value Or optHiKiss.value Then
        SubmitAlignment = cdrLeftAlignment
    Else
        SubmitAlignment = cdrRightAlignment
    End If

End Function

Private Sub SetTextBaselineX(ByVal textShape As Shape, ByVal targetX As Double)

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    textShape.Text.Story.Paragraphs.Last.Baselines.GetBoundingBox _
        x, y, w, h

    textShape.Move targetX - x, 0#

End Sub

Private Sub SetTextBaselineY(ByVal textShape As Shape, ByVal targetY As Double)

    Dim x As Double
    Dim y As Double
    Dim w As Double
    Dim h As Double

    textShape.Text.Story.Paragraphs.Last.Baselines.GetBoundingBox _
        x, y, w, h

    textShape.Move 0#, targetY - y

End Sub

Private Sub ApplySubmitPosition(ByVal labelShape As Shape)
    Dim anchor As Shape

    If optKissLabel.Value Then

        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            labelShape.LeftX = KISS_LABEL_TAG_X
            labelShape.BottomY = KISS_LABEL_TAG_Y
        Else
            labelShape.LeftX = anchor.LeftX + KISS_LABEL_GAP_X
            labelShape.TopY = anchor.BottomY - KISS_LABEL_GAP_Y + TAG_Y_OFFSET
        End If

    ElseIf optHiDie.Value Then

        Set anchor = FindHiDieAnchor()

        If anchor Is Nothing Then
            labelShape.LeftX = LABEL_HI_DIE_X
        Else
            labelShape.LeftX = anchor.RightX + HI_DIE_GAP_X
        End If

        If anchor Is Nothing Then
            SetTextBaselineY labelShape, LABEL_HI_DIE_Y + HI_DIE_Y_OFFSET
        Else
            SetTextBaselineY labelShape, anchor.BottomY - HI_DIE_GAP_Y + HI_DIE_Y_OFFSET
        End If

    ElseIf optHiKiss.Value Then

        Set anchor = FindHiKissAnchor()
        labelShape.Rotate -90#

        If anchor Is Nothing Then
            labelShape.LeftX = LABEL_TAG_X
            SetTextBaselineX labelShape, LABEL_TAG_X
            labelShape.BottomY = LABEL_TAG_Y
        Else
            SetTextBaselineX labelShape, anchor.LeftX + HI_KISS_GAP_X
            labelShape.BottomY = anchor.BottomY + HI_KISS_GAP_Y
        End If

    Else

        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            labelShape.RightX = LABEL_TAG_X
            SetTextBaselineY labelShape, LABEL_TAG_Y + TAG_Y_OFFSET
        Else
            labelShape.RightX = anchor.LeftX - UNI_LABEL_GAP_X
            SetTextBaselineY labelShape, anchor.BottomY + UNI_LABEL_GAP_Y + TAG_Y_OFFSET
        End If


    End If
End Sub

Private Sub ApplyLabelTextFormatting(ByVal labelShape As Shape, ByVal labelText As String)

    If chkSmallCaps.value Then
        ApplySmallCaps labelShape, labelText
    End If

    If chkFraction.value Then
        ApplyFractions labelShape, labelText
    End If

    If chkColorize.value Then
        ApplyColorize labelShape, labelText
    Else
        ApplyTextFill labelShape, 1, Len(labelText), 0, 0, 0, 100
    End If

End Sub

Private Sub ApplyColorize(ByVal labelShape As Shape, ByVal labelText As String)

    Dim tagStart As Long
    Dim tagEnd As Long
    Dim customerStart As Long
    Dim customerEnd As Long
    Dim dashPosition As Long
    Dim dayStart As Long
    Dim dayEnd As Long
    Dim slashPosition As Long
    Dim monthStart As Long
    Dim monthEnd As Long
    Dim compactDateStart As Long
    Dim compactDateEnd As Long
    Dim markerStart As Long
    Dim markerEnd As Long
    Dim serialStart As Long
    Dim serialEnd As Long

    FindLeadingFpTag labelText, tagStart, tagEnd
    If tagStart > 0 Then
        ApplyTextFill labelShape, tagStart, tagEnd, 0, 60, 100, 0
    End If

    If Not FindDateParts(labelText, dayStart, dayEnd, slashPosition, monthStart, monthEnd) Then
        If FindCompactDateParts(labelText, compactDateStart, compactDateEnd) Then
            customerStart = CustomerNameStart(labelText)
            dashPosition = SeparatorPositionBeforeDate(labelText, compactDateStart)
            customerEnd = LastNonSpaceBefore(labelText, dashPosition)
            ApplyTextFill labelShape, customerStart, customerEnd, 0, 0, 0, 100
            ApplyTextFill labelShape, dashPosition, dashPosition, 0, 100, 100, 0
            ApplyTextFill labelShape, compactDateStart, compactDateEnd, 100, 100, 0, 0
            Exit Sub
        End If

        customerStart = CustomerNameStart(labelText)
        customerEnd = CustomerNameEnd(labelText)
        If tagStart = 0 Then
            ApplyTextFill labelShape, customerStart, customerEnd, 0, 100, 0, 0
        Else
            ApplyTextFill labelShape, customerStart, customerEnd, 0, 0, 0, 100
        End If
        Exit Sub
    End If

    customerStart = CustomerNameStart(labelText)
    dashPosition = SeparatorPositionBeforeDate(labelText, dayStart)
    If dashPosition > 0 Then
        customerEnd = LastNonSpaceBefore(labelText, dashPosition)
        ApplyTextFill labelShape, dashPosition, dashPosition, 0, 100, 100, 0
    Else
        customerEnd = LastNonSpaceBefore(labelText, dayStart)
    End If

    ApplyTextFill labelShape, customerStart, customerEnd, 0, 0, 0, 100
    ApplyTextFill labelShape, dayStart, dayEnd, 100, 0, 100, 0
    ApplyTextFill labelShape, slashPosition, slashPosition, 100, 100, 0, 0
    ApplyTextFill labelShape, monthStart, monthEnd, 100, 0, 100, 0

    FindPostDateMarker labelText, monthEnd, markerStart, markerEnd
    If markerStart > 0 Then
        ApplyTextFill labelShape, markerStart, markerEnd, 0, 100, 0, 0
        FindSerialNumber labelText, markerEnd, serialStart, serialEnd
    Else
        FindSerialNumber labelText, monthEnd, serialStart, serialEnd
    End If

    If serialStart > 0 Then
        ApplyTextFill labelShape, serialStart, serialEnd, 100, 0, 0, 0
    End If

End Sub

Private Sub ApplySmallCaps(ByVal labelShape As Shape, ByVal labelText As String)

    Dim tagStart As Long
    Dim tagEnd As Long
    Dim customerStart As Long
    Dim customerEnd As Long
    Dim dayStart As Long
    Dim dayEnd As Long
    Dim slashPosition As Long
    Dim monthStart As Long
    Dim monthEnd As Long
    Dim markerStart As Long
    Dim markerEnd As Long

    customerEnd = CustomerNameEnd(labelText)
    If customerEnd <= 0 Then
        Exit Sub
    End If

    FindLeadingFpTag labelText, tagStart, tagEnd
    If tagStart > 0 Then
        ApplyLowercaseSmallCapsToRange labelShape, tagStart, tagEnd
        customerStart = FirstNonSpacePosition(labelText, tagEnd + 1)
    Else
        customerStart = FirstNonSpacePosition(labelText, 1)
    End If

    ApplySmallCapsToRange labelShape, customerStart, customerEnd

    If FindDateParts(labelText, dayStart, dayEnd, slashPosition, monthStart, monthEnd) Then
        FindPostDateMarker labelText, monthEnd, markerStart, markerEnd
        If markerStart > 0 Then
            ApplyLowercaseSmallCapsToRange labelShape, markerStart, markerEnd
        End If
    End If

End Sub

Private Sub ApplySmallCapsToRange(ByVal labelShape As Shape, ByVal rangeStart As Long, ByVal rangeEnd As Long)

    If rangeStart <= 0 Or rangeEnd < rangeStart Then
        Exit Sub
    End If

    labelShape.Text.Story.Range(rangeStart - 1, rangeEnd).Case = cdrSmallCapsFontCase

End Sub

Private Sub ApplyLowercaseSmallCapsToRange(ByVal labelShape As Shape, ByVal rangeStart As Long, ByVal rangeEnd As Long)

    Dim targetRange As TextRange

    If rangeStart <= 0 Or rangeEnd < rangeStart Then
        Exit Sub
    End If

    Set targetRange = labelShape.Text.Story.Range(rangeStart - 1, rangeEnd)
    targetRange.Text = LCase$(targetRange.Text)
    targetRange.Case = cdrSmallCapsFontCase

End Sub

Private Sub ApplyTextFill(ByVal labelShape As Shape, ByVal rangeStart As Long, ByVal rangeEnd As Long, ByVal cyan As Long, ByVal magenta As Long, ByVal yellow As Long, ByVal black As Long)

    If rangeStart <= 0 Or rangeEnd < rangeStart Then
        Exit Sub
    End If

    labelShape.Text.Story.Range(rangeStart - 1, rangeEnd).Fill.ApplyUniformFill CreateCMYKColor(cyan, magenta, yellow, black)

End Sub

Private Sub FindLeadingFpTag(ByVal labelText As String, ByRef tagStart As Long, ByRef tagEnd As Long)

    Dim position As Long

    position = FirstNonSpacePosition(labelText, 1)
    If position = 0 Then
        Exit Sub
    End If

    If UCase$(Mid$(labelText, position, 4)) = "[FP]" Then
        tagStart = position
        tagEnd = position + 3
    End If

End Sub

Private Function CustomerNameStart(ByVal labelText As String) As Long

    Dim tagStart As Long
    Dim tagEnd As Long

    FindLeadingFpTag labelText, tagStart, tagEnd
    If tagEnd > 0 Then
        CustomerNameStart = FirstNonSpacePosition(labelText, tagEnd + 1)
    Else
        CustomerNameStart = FirstNonSpacePosition(labelText, 1)
    End If

End Function

Private Sub ApplyFractions(ByVal labelShape As Shape, ByVal labelText As String)

    Dim fractionStart As Long
    Dim fractionEnd As Long
    Dim fractionRange As TextRange

    If Not FindFirstFractionRange(labelText, fractionStart, fractionEnd) Then
        Exit Sub
    End If

    Set fractionRange = labelShape.Text.Story.Range(fractionStart - 1, fractionEnd)
    fractionRange.SetOpenTypeFeature "frac", 1

End Sub

Private Function CustomerNameEnd(ByVal labelText As String) As Long

    Dim delimiterPosition As Long
    Dim emDashPosition As Long
    Dim slashPosition As Long
    Dim position As Long

    delimiterPosition = InStr(1, labelText, " - ", vbTextCompare)
    If delimiterPosition > 1 Then
        CustomerNameEnd = delimiterPosition - 1
        Exit Function
    End If

    emDashPosition = InStr(1, labelText, " " & ChrW$(8212) & " ", vbTextCompare)
    If emDashPosition > 1 Then
        CustomerNameEnd = emDashPosition - 1
        Exit Function
    End If

    slashPosition = InStr(1, labelText, "/", vbTextCompare)
    If slashPosition <= 1 Then
        CustomerNameEnd = Len(labelText)
        Exit Function
    End If

    position = slashPosition - 1
    Do While position > 0 And Mid$(labelText, position, 1) Like "[0-9]"
        position = position - 1
    Loop

    Do While position > 0 And Mid$(labelText, position, 1) = " "
        position = position - 1
    Loop

    If position > 0 And Mid$(labelText, position, 1) = "-" Then
        position = position - 1
        Do While position > 0 And Mid$(labelText, position, 1) = " "
            position = position - 1
        Loop
    End If

    CustomerNameEnd = position

End Function

Private Function FindCompactDateParts(ByVal labelText As String, ByRef dateStart As Long, ByRef dateEnd As Long) As Boolean

    Dim position As Long

    position = Len(labelText)
    Do While position > 0 And Mid$(labelText, position, 1) = " "
        position = position - 1
    Loop

    dateEnd = position
    dateStart = dateEnd - 5
    If dateStart <= 0 Then
        Exit Function
    End If

    For position = dateStart To dateEnd
        If Not Mid$(labelText, position, 1) Like "[0-9]" Then
            Exit Function
        End If
    Next position

    FindCompactDateParts = (SeparatorPositionBeforeDate(labelText, dateStart) > 0)

End Function

Private Function SeparatorPositionBeforeDate(ByVal labelText As String, ByVal dateStart As Long) As Long

    Dim position As Long
    Dim currentChar As String

    position = dateStart - 1
    Do While position > 0 And Mid$(labelText, position, 1) = " "
        position = position - 1
    Loop

    If position <= 0 Then
        Exit Function
    End If

    currentChar = Mid$(labelText, position, 1)
    If currentChar = "-" Or currentChar = ChrW$(8212) Then
        SeparatorPositionBeforeDate = position
    End If

End Function

Private Function FindDateParts(ByVal labelText As String, ByRef dayStart As Long, ByRef dayEnd As Long, ByRef slashPosition As Long, ByRef monthStart As Long, ByRef monthEnd As Long) As Boolean

    Dim dateStart As Long
    Dim dateEnd As Long

    If Not FindFirstFractionRange(labelText, dateStart, dateEnd) Then
        Exit Function
    End If

    slashPosition = InStr(dateStart, labelText, "/", vbTextCompare)
    If slashPosition <= dateStart Or slashPosition >= dateEnd Then
        Exit Function
    End If

    dayStart = dateStart
    dayEnd = slashPosition - 1
    monthStart = slashPosition + 1
    monthEnd = dateEnd
    FindDateParts = True

End Function

Private Sub FindPostDateMarker(ByVal labelText As String, ByVal dateEnd As Long, ByRef markerStart As Long, ByRef markerEnd As Long)

    Dim position As Long

    position = FirstNonSpacePosition(labelText, dateEnd + 1)
    If position = 0 Or position + 2 > Len(labelText) Then
        Exit Sub
    End If

    If Mid$(labelText, position, 1) = "(" And _
       UCase$(Mid$(labelText, position + 1, 1)) Like "[A-Z]" And _
       Mid$(labelText, position + 2, 1) = ")" Then
        markerStart = position
        markerEnd = position + 2
    End If

End Sub

Private Sub FindSerialNumber(ByVal labelText As String, ByVal searchAfter As Long, ByRef serialStart As Long, ByRef serialEnd As Long)

    Dim position As Long

    position = FirstNonSpacePosition(labelText, searchAfter + 1)
    If position = 0 Then
        Exit Sub
    End If

    If Not Mid$(labelText, position, 1) Like "[0-9]" Then
        Exit Sub
    End If

    serialStart = position
    Do While position <= Len(labelText) And Mid$(labelText, position, 1) Like "[0-9]"
        position = position + 1
    Loop
    serialEnd = position - 1

End Sub

Private Function FirstNonSpacePosition(ByVal labelText As String, ByVal startPosition As Long) As Long

    Dim position As Long

    position = startPosition
    Do While position <= Len(labelText) And Mid$(labelText, position, 1) = " "
        position = position + 1
    Loop

    If position <= Len(labelText) Then
        FirstNonSpacePosition = position
    End If

End Function

Private Function LastNonSpaceBefore(ByVal labelText As String, ByVal beforePosition As Long) As Long

    Dim position As Long

    position = beforePosition - 1
    Do While position > 0 And Mid$(labelText, position, 1) = " "
        position = position - 1
    Loop

    LastNonSpaceBefore = position

End Function

Private Function FindFirstFractionRange(ByVal labelText As String, ByRef rangeStart As Long, ByRef rangeEnd As Long) As Boolean

    Dim slashPosition As Long
    Dim startPosition As Long
    Dim endPosition As Long

    slashPosition = InStr(1, labelText, "/", vbTextCompare)

    Do While slashPosition > 0
        startPosition = slashPosition - 1
        Do While startPosition > 0 And Mid$(labelText, startPosition, 1) Like "[0-9]"
            startPosition = startPosition - 1
        Loop
        startPosition = startPosition + 1

        endPosition = slashPosition + 1
        Do While endPosition <= Len(labelText) And Mid$(labelText, endPosition, 1) Like "[0-9]"
            endPosition = endPosition + 1
        Loop
        endPosition = endPosition - 1

        If startPosition < slashPosition And endPosition > slashPosition Then
            rangeStart = startPosition
            rangeEnd = endPosition
            FindFirstFractionRange = True
            Exit Function
        End If

        slashPosition = InStr(slashPosition + 1, labelText, "/", vbTextCompare)
    Loop

End Function

Private Function SelectedFontSize() As Single

    Dim value As Single

    value = Val(cmbFontSize.Text)
    If value <= 0 Then
        value = UNI_LABEL_FONT_SIZE
    End If
    If value < MIN_FONT_SIZE Then
        value = MIN_FONT_SIZE
    End If
    If value > MAX_FONT_SIZE Then
        value = MAX_FONT_SIZE
    End If

    SelectedFontSize = value

End Function

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
