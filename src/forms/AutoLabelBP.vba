Option Explicit

' VBE: UserForm AutoLabelWizardBP needs TextBox (Name) = txbLabel.
' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String

Private Const REG_APP_NAME As String = "RinCorelMacros"
Private Const REG_SECTION_NAME As String = "AutoLabelBP"
Private Const MIN_FONT_SIZE As Long = 4
Private Const MAX_FONT_SIZE As Long = 24
Private Const UNI_LABEL_FONT_SIZE As Single = 8!
Private Const HI_DIE_FONT_SIZE As Single = 12!
Private Const HI_KISS_FONT_SIZE As Single = 12!
Private Const KISS_LABEL_FONT_SIZE As Single = 4!
Private Const PREVIEW_FONT_SIZE As Single = 10!
Private Const OP_ONE As String = "Saiful"
Private Const OP_TWO As String = "Hafizh"
Private Const OP_THREE As String = "Dori"
Private Const OP_FOUR As String = "Reni"

Private Sub chkMasterPage_Click()

End Sub

Private Sub cmdSaveConfig_Click()

    SaveCurrentConfig
    MsgBox "Pengaturan Auto Label BP berhasil disimpan.", vbInformation

End Sub

Private Sub UserForm_Initialize()

    Dim fontSize As Long

    cmbFontSize.Clear
    For fontSize = MIN_FONT_SIZE To MAX_FONT_SIZE
        cmbFontSize.AddItem CStr(fontSize)
    Next fontSize

    cmbOperator.Clear
    cmbOperator.AddItem OP_ONE
    cmbOperator.AddItem OP_TWO
    cmbOperator.AddItem OP_THREE
    cmbOperator.AddItem OP_FOUR
    cmbOperator.ListIndex = 0

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
    cmbOperator.Text = GetSetting(REG_APP_NAME, REG_SECTION_NAME, "cmbOperator", cmbOperator.Text)

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
    SaveSetting REG_APP_NAME, REG_SECTION_NAME, "cmbOperator", Trim$(cmbOperator.Text)

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

Private Sub cmbOperator_Change()

End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdSubmit_Click()
    Dim presenter As ALPresenter

    Set presenter = New ALPresenter
    presenter.Submit Me
End Sub

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

Private Sub txbLabel_Change()

    UpdatePreview

End Sub

Private Sub UpdatePreview()

    With txbLabel
        .Font.Bold = chkBold.value
        .Font.Italic = chkItalic.value
        .Font.Size = PREVIEW_FONT_SIZE
    End With

End Sub

Private Function SubmitFontSize() As Single

    If optKissLabel.value Then
        SubmitFontSize = KISS_LABEL_FONT_SIZE
    Else
        SubmitFontSize = SelectedFontSize()
    End If

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

Public Function ALReadOptions() As ALLabelOptions
    Dim options As ALLabelOptions

    Set options = New ALLabelOptions
    options.Mode = "bp"
    options.InputText = txbLabel.Text
    options.FontName = txbLabel.Font.Name
    options.FontSize = SubmitFontSize()
    options.OperatorName = cmbOperator.Text
    options.MasterPage = chkMasterPage.Value
    options.Uppercase = chkUppercase.Value
    options.Bold = chkBold.Value
    options.Italic = chkItalic.Value
    options.SmallCaps = chkSmallCaps.Value
    options.Colorize = chkColorize.Value
    options.Fraction = chkFraction.Value
    options.EmDash = chkEmDash.Value
    options.KissLabel = optKissLabel.Value
    options.HiDie = optHiDie.Value
    options.HiKiss = optHiKiss.Value
    options.UniLabel = optUniLabel.Value
    Set ALReadOptions = options
End Function

Public Sub ALFocusInput()
    txbLabel.SetFocus
End Sub

Public Sub ALClose()
    Unload Me
End Sub

' MacroBehavior adapter for AutoLabelWizardBP.
Public Sub MRBehaviorValue(ByVal target As String, ByVal value As Variant)
    Select Case LCase$(target)
        Case "txblabel": txbLabel.Text = CStr(value)
        Case "cmbfontsize"
            If VarType(value) = vbString Then
                cmbFontSize.Text = CStr(value)
            Else
                cmbFontSize.Text = Trim$(Str$(CDbl(value)))
            End If
        Case "cmboperator": cmbOperator.Text = CStr(value)
        Case "chkmasterpage": chkMasterPage.Value = CBool(value)
        Case "chkuppercase": chkUppercase.Value = CBool(value)
        Case "chkbold": chkBold.Value = CBool(value)
        Case "chkitalic": chkItalic.Value = CBool(value)
        Case "chksmallcaps": chkSmallCaps.Value = CBool(value)
        Case "chkcolorize": chkColorize.Value = CBool(value)
        Case "chkfraction": chkFraction.Value = CBool(value)
        Case "chkemdash": chkEmDash.Value = CBool(value)
        Case "optkisslabel"
            optKissLabel.Value = CBool(value)
            If optKissLabel.Value Then optKissLabel_Click
        Case "opthidie"
            optHiDie.Value = CBool(value)
            If optHiDie.Value Then optHiDie_Click
        Case "opthikiss"
            optHiKiss.Value = CBool(value)
            If optHiKiss.Value Then optHiKiss_Click
        Case "optunilabel"
            optUniLabel.Value = CBool(value)
            If optUniLabel.Value Then optUniLabel_Click
        Case Else: Err.Raise 5, "AutoLabelWizardBP.MRBehaviorValue", "Target tidak terdaftar: " & target
    End Select
    UpdatePreview
End Sub

Public Function MRBehaviorReadValue(ByVal target As String) As Variant
    Select Case LCase$(target)
        Case "cmbfontsize": MRBehaviorReadValue = cmbFontSize.Text
        Case "cmboperator": MRBehaviorReadValue = cmbOperator.Text
        Case "chkmasterpage": MRBehaviorReadValue = chkMasterPage.Value
        Case "chkuppercase": MRBehaviorReadValue = chkUppercase.Value
        Case "chkbold": MRBehaviorReadValue = chkBold.Value
        Case "chkitalic": MRBehaviorReadValue = chkItalic.Value
        Case "chksmallcaps": MRBehaviorReadValue = chkSmallCaps.Value
        Case "chkcolorize": MRBehaviorReadValue = chkColorize.Value
        Case "chkfraction": MRBehaviorReadValue = chkFraction.Value
        Case "chkemdash": MRBehaviorReadValue = chkEmDash.Value
        Case "optkisslabel": MRBehaviorReadValue = optKissLabel.Value
        Case "opthidie": MRBehaviorReadValue = optHiDie.Value
        Case "opthikiss": MRBehaviorReadValue = optHiKiss.Value
        Case "optunilabel": MRBehaviorReadValue = optUniLabel.Value
        Case Else: Err.Raise 5, "AutoLabelWizardBP.MRBehaviorReadValue", "Default tidak tersedia: " & target
    End Select
End Function

Public Sub MRBehaviorSubmit()
    Dim presenter As ALPresenter
    Set presenter = New ALPresenter
    presenter.Submit Me, True
End Sub

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
