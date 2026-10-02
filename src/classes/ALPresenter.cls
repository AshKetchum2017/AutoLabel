Option Explicit

' Import as a Class Module named ALPresenter in the Auto Label GMS.
Private pView As Object
Private pModel As ALInputModel
Private pOptions As ALLabelOptions

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
Private Const OP_ONE As String = "Saiful"
Private Const OP_TWO As String = "Hafizh"
Private Const OP_THREE As String = "Dori"
Private Const OP_FOUR As String = "Reni"

Public Sub Submit(ByVal view As Object, Optional ByVal fromBehavior As Boolean = False)
    Dim doc As Document
    Dim labelShape As Shape
    Dim labelText As String
    Dim mainText As String
    Dim fontSize As Single
    Dim previousUnit As cdrUnit
    Dim hasPreviousUnit As Boolean
    Dim targetLayer As Layer
    Dim textX As Double
    Dim textY As Double
    Dim textAlignment As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo SubmitError
    Set pView = view
    Set pOptions = pView.ALReadOptions()
    Set pModel = New ALInputModel
    pModel.Mode = pOptions.Mode
    Set doc = ActiveDocument
    mainText = pModel.Resolve(pOptions.InputText, doc.FilePath, pOptions.Fraction, pOptions.EmDash)
    If pOptions.EmDash Then
        mainText = Replace(mainText, " - ", " " & ChrW$(8212) & " ", 1, 1, vbTextCompare)
    End If
    mainText = Trim$(mainText)
    If pOptions.HiDie Then mainText = TitleCaseCustomer(mainText)
    If pOptions.Uppercase Or pOptions.KissLabel Then mainText = UCase$(mainText)

    If Len(mainText) = 0 Then
        If fromBehavior Then Err.Raise 5, "ALPresenter.Submit", "Teks label kosong setelah token diselesaikan."
        pView.ALFocusInput
        GoTo Done
    End If

    labelText = mainText
    If pModel.HasToken And pOptions.KissLabel And Len(pModel.partNumber) > 0 Then
        labelText = labelText & vbCrLf & "PT." & pModel.partNumber
    End If

    fontSize = pOptions.fontSize
    textX = SubmitX()
    textY = SubmitY()
    textAlignment = SubmitAlignment()

    previousUnit = doc.Unit
    hasPreviousUnit = True
    doc.Unit = cdrMillimeter
    Set targetLayer = GetSubmitLayer()
    Set labelShape = targetLayer.CreateArtisticText( _
        textX, _
        textY, _
        labelText, _
        cdrLanguageNone, _
        cdrCharSetMixed, _
        pOptions.FontName, _
        fontSize, _
        IIf(pOptions.Bold, cdrTrue, cdrFalse), _
        IIf(pOptions.Italic, cdrTrue, cdrFalse), _
        cdrMixedFontLine, _
        textAlignment)

    If pOptions.Mode = "bp" Then ApplySubmitPosition labelShape
    ApplyLabelTextFormatting labelShape, mainText, Len(labelText)
    If pOptions.Mode = "normal" Then ApplySubmitPosition labelShape
    doc.Unit = previousUnit
    hasPreviousUnit = False
    pView.ALClose
    GoTo Done

SubmitError:
    errorNumber = Err.number
    errorSource = Err.source
    errorDescription = Err.description
    If hasPreviousUnit Then
        On Error Resume Next
        doc.Unit = previousUnit
        On Error GoTo 0
    End If
    If fromBehavior Then
        Set pModel = Nothing
        Set pOptions = Nothing
        Set pView = Nothing
        Err.Raise errorNumber, errorSource, errorDescription
    End If
    MsgBox "Auto Label gagal dibuat." & vbCrLf & _
        "Number : " & CStr(errorNumber) & vbCrLf & _
        "Source : " & errorSource & vbCrLf & _
        "Error  : " & errorDescription, _
        vbExclamation
    On Error Resume Next
    If Not pView Is Nothing Then pView.ALFocusInput
    On Error GoTo 0

Done:
    Set pModel = Nothing
    Set pOptions = Nothing
    Set pView = Nothing
End Sub

Private Function TitleCaseCustomer(ByVal labelText As String) As String
    Dim customerStart As Long
    Dim customerEnd As Long
    Dim customerText As String

    customerStart = CustomerNameStart(labelText)
    customerEnd = CustomerNameEnd(labelText)
    TitleCaseCustomer = labelText
    If customerStart <= 0 Or customerEnd < customerStart Then Exit Function

    customerText = Mid$(labelText, customerStart, customerEnd - customerStart + 1)
    customerText = StrConv(LCase$(customerText), vbProperCase)
    TitleCaseCustomer = Left$(labelText, customerStart - 1) & customerText & Mid$(labelText, customerEnd + 1)
End Function

Private Function GetSubmitLayer() As Layer

    If pOptions.MasterPage Then
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
    errorNumber = Err.number
    errorDescription = Err.description
    errorSource = Err.source
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
               And shp.SizeWidth <= 200# _
               And shp.SizeHeight >= 0.7 _
               And shp.SizeHeight <= 1.3 Then

                Set FindHiDieAnchor = shp
                Exit For

            End If
        End If

    Next shp

RestoreLayer:
    errorNumber = Err.number
    errorDescription = Err.description
    errorSource = Err.source
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
    errorNumber = Err.number
    errorDescription = Err.description
    errorSource = Err.source
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

    If pOptions.KissLabel Then
        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            SubmitX = KISS_LABEL_TAG_X
        Else
            SubmitX = anchor.LeftX + KISS_LABEL_GAP_X
        End If

    ElseIf pOptions.HiDie Then
        Set anchor = FindHiDieAnchor()

        If anchor Is Nothing Then
            SubmitX = LABEL_HI_DIE_X
        Else
            ' 5 mm dari sisi kanan anchor
            SubmitX = anchor.RightX + HI_DIE_GAP_X
        End If

    ElseIf pOptions.HiKiss Then
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

    If pOptions.KissLabel Then
        Set anchor = FindAnchor("scpro2_printonly", "filename")
        
        If anchor Is Nothing Then
            SubmitY = KISS_LABEL_TAG_Y
        Else
            SubmitY = anchor.BottomY - KISS_LABEL_GAP_Y + TAG_Y_OFFSET
        End If

    ElseIf pOptions.HiDie Then
        Set anchor = FindHiDieAnchor()

        If anchor Is Nothing Then
            SubmitY = LABEL_HI_DIE_Y
        Else
            SubmitY = anchor.BottomY + HI_DIE_GAP_Y + HI_DIE_Y_OFFSET
        End If
    
    ElseIf pOptions.HiKiss Then
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

    If pOptions.KissLabel Or pOptions.HiDie Or pOptions.HiKiss Then
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

    If pOptions.KissLabel Then

        Set anchor = FindAnchor("scpro2_printonly", "filename")

        If anchor Is Nothing Then
            labelShape.LeftX = KISS_LABEL_TAG_X
            labelShape.BottomY = KISS_LABEL_TAG_Y
        Else
            labelShape.LeftX = anchor.LeftX + KISS_LABEL_GAP_X
            labelShape.TopY = anchor.BottomY - KISS_LABEL_GAP_Y + TAG_Y_OFFSET
        End If

    ElseIf pOptions.HiDie Then

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

    ElseIf pOptions.HiKiss Then

        Set anchor = FindHiKissAnchor()
        If pOptions.Mode = "bp" Then
            labelShape.Rotate 90#
        Else
            labelShape.Rotate -90#
        End If

        If anchor Is Nothing Then
            labelShape.LeftX = LABEL_TAG_X
            SetTextBaselineX labelShape, LABEL_TAG_X
            labelShape.BottomY = LABEL_TAG_Y
        Else
            If pOptions.Mode = "bp" Then
                SetTextBaselineX labelShape, anchor.RightX + HI_KISS_GAP_X
            Else
                SetTextBaselineX labelShape, anchor.LeftX + HI_KISS_GAP_X
            End If
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

Private Sub ApplyLabelTextFormatting(ByVal labelShape As Shape, ByVal labelText As String, ByVal fullTextLength As Long)

    If fullTextLength > Len(labelText) Then
        ' Start with black for the entire story; only the main line is colorized below.
        labelShape.Text.Story.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
    End If

    If pOptions.SmallCaps Then
        ApplySmallCaps labelShape, labelText
    End If

    If pOptions.Fraction Then
        ApplyFractions labelShape, labelText
    End If

    If pOptions.Mode = "bp" Then
        ApplyTextFill labelShape, 1, Len(labelText), 0, 0, 0, 100
        If pOptions.Colorize Then ApplyOperatorFormatting labelShape, labelText
    ElseIf pOptions.Colorize Then
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

Private Sub ApplyOperatorFormatting(ByVal labelShape As Shape, ByVal labelText As String)
    Dim dateStart As Long, dateEnd As Long, separatorPosition As Long
    Dim customerStart As Long, customerEnd As Long
    Dim custCy As Long, custMg As Long, custYl As Long, custBk As Long

    If Not ParseDateGrammar(labelText, dateStart, dateEnd, separatorPosition) Then Exit Sub

    Select Case Trim$(pOptions.OperatorName)
        Case OP_ONE: custMg = 100: custYl = 100
        Case OP_TWO: custBk = 100
        Case OP_THREE: custMg = 100
        Case OP_FOUR: custCy = 100: custMg = 100
        Case Else: Exit Sub
    End Select

    customerStart = CustomerNameStart(labelText)
    If separatorPosition > 0 Then
        customerEnd = LastNonSpaceBefore(labelText, separatorPosition)
    Else
        customerEnd = LastNonSpaceBefore(labelText, dateStart)
    End If

    ApplyTextFill labelShape, customerStart, customerEnd, custCy, custMg, custYl, custBk
    If separatorPosition > 0 Then _
        ApplyTextFill labelShape, separatorPosition, separatorPosition, 0, 0, 0, 100
    ApplyTextFill labelShape, dateStart, dateEnd, 0, 0, 0, 100
End Sub

Private Function ParseDateGrammar( _
    ByVal labelText As String, _
    ByRef dateStart As Long, _
    ByRef dateEnd As Long, _
    ByRef separatorPosition As Long) As Boolean

    Dim position As Long, firstSlashPosition As Long, secondSlashPosition As Long

    dateEnd = Len(labelText)
    Do While dateEnd > 0 And Mid$(labelText, dateEnd, 1) = " "
        dateEnd = dateEnd - 1
    Loop
    If dateEnd <= 0 Then Exit Function

    If dateEnd >= 8 And Mid$(labelText, dateEnd - 5, 1) = "/" Then
        dateStart = dateEnd - 7
    ElseIf dateEnd >= 5 And Mid$(labelText, dateEnd - 2, 1) = "/" Then
        dateStart = dateEnd - 4
    Else
        dateStart = dateEnd - 5
    End If
    If dateStart <= 0 Then Exit Function

    firstSlashPosition = InStr(dateStart, labelText, "/", vbTextCompare)
    If firstSlashPosition = 0 Then
        separatorPosition = dateStart - 1
        Do While separatorPosition > 0 And Mid$(labelText, separatorPosition, 1) = " "
            separatorPosition = separatorPosition - 1
        Loop
        If (separatorPosition <= 0 Or _
            (Mid$(labelText, separatorPosition, 1) <> "-" And _
             Mid$(labelText, separatorPosition, 1) <> ChrW$(8212))) And _
           dateStart = dateEnd - 5 Then
            dateStart = dateEnd - 3
        End If
    End If
    If dateStart <= 0 Then Exit Function

    firstSlashPosition = InStr(dateStart, labelText, "/", vbTextCompare)
    If firstSlashPosition > 0 And firstSlashPosition < dateEnd Then
        secondSlashPosition = InStr(firstSlashPosition + 1, labelText, "/", vbTextCompare)
        If secondSlashPosition > 0 Then
            If secondSlashPosition >= dateEnd Or _
               secondSlashPosition - firstSlashPosition <> 3 Or _
               dateEnd - secondSlashPosition <> 3 Then Exit Function
        ElseIf firstSlashPosition - dateStart <> 2 Or dateEnd - firstSlashPosition <> 2 Then
            Exit Function
        End If
        For position = dateStart To dateEnd
            If Mid$(labelText, position, 1) <> "/" And _
               Not Mid$(labelText, position, 1) Like "[0-9]" Then Exit Function
        Next position
    Else
        If firstSlashPosition > 0 Then Exit Function
        If dateEnd - dateStart + 1 <> 4 And dateEnd - dateStart + 1 <> 6 Then Exit Function
        For position = dateStart To dateEnd
            If Not Mid$(labelText, position, 1) Like "[0-9]" Then Exit Function
        Next position
    End If

    separatorPosition = 0
    If dateStart > 1 Then
        separatorPosition = dateStart - 1
        Do While separatorPosition > 0 And Mid$(labelText, separatorPosition, 1) = " "
            separatorPosition = separatorPosition - 1
        Loop
        If separatorPosition > 0 Then
            If Mid$(labelText, separatorPosition, 1) <> "-" And _
               Mid$(labelText, separatorPosition, 1) <> ChrW$(8212) Then separatorPosition = 0
        End If
    End If
    If separatorPosition > 0 Then
        If LastNonSpaceBefore(labelText, separatorPosition) <= 0 Then Exit Function
    End If
    ParseDateGrammar = True
End Function

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
    Dim separatorPosition As Long, secondSlashPosition As Long

    If pOptions.Mode = "bp" Then
        If Not ParseDateGrammar(labelText, dateStart, dateEnd, separatorPosition) Then Exit Function
    Else
        If Not FindFirstFractionRange(labelText, dateStart, dateEnd) Then Exit Function
    End If

    slashPosition = InStr(dateStart, labelText, "/", vbTextCompare)
    If slashPosition <= dateStart Or slashPosition >= dateEnd Then
        Exit Function
    End If

    dayStart = dateStart
    dayEnd = slashPosition - 1
    monthStart = slashPosition + 1
    monthEnd = dateEnd
    If pOptions.Mode = "bp" Then
        secondSlashPosition = InStr(slashPosition + 1, labelText, "/", vbTextCompare)
        If secondSlashPosition > 0 And secondSlashPosition < dateEnd Then _
            monthEnd = secondSlashPosition - 1
    End If
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
