Option Explicit

' Class Module (Name) = ALBehaviorContract. Pure semantic validation.
Private pText As String
Private pAt As Long

Public Function Validate(ByVal script As String) As MRBehaviorBlock
    Dim parser As New MRBehaviorParser
    Dim blocks As Collection, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction
    Dim stageSeen As Boolean, labelSet As Boolean, submitted As Boolean

    pText = script
    pAt = 1
    Set blocks = parser.Parse(script, Array("AutoLabel"))
    Set block = blocks(1)

    For Each ins In block.Instructions
        pAt = ins.Position
        If submitted Then Fail "Tidak boleh ada instruksi setelah @cmdSubmit."
        If StrComp(ins.FormName, "AutoLabelWizard", vbTextCompare) <> 0 Then _
            Fail "UserForm tidak terdaftar: " & ins.FormName

        If ins.Kind = "stage" Then
            If stageSeen Then Fail "AutoLabelWizard hanya boleh mempunyai satu tahap."
            stageSeen = True
        ElseIf ins.Kind = "action" Then
            If StrComp(ins.Target, "cmdSubmit", vbTextCompare) <> 0 Then _
                Fail "Action tidak terdaftar: AutoLabelWizard.@" & ins.Target
            If Not labelSet Then Fail "Isi txbLabel sebelum @cmdSubmit."
            submitted = True
        Else
            ValidateValue ins
            If StrComp(ins.Target, "txbLabel", vbTextCompare) = 0 And ins.Kind = "string" Then _
                labelSet = True
        End If
    Next ins

    pAt = Len(pText) + 1
    If Not submitted Then Fail "Akhiri behavior AutoLabel dengan @cmdSubmit."
    Set Validate = block
End Function

Private Sub ValidateValue(ByVal ins As MRBehaviorInstruction)
    Dim target As String, raw As String, fontSize As Double
    Dim i As Long, digits As Long, dots As Long

    target = LCase$(ins.Target)
    Select Case target
        Case "txblabel"
            If ins.Kind = "skip" Then Exit Sub
            If ins.Kind <> "string" Then Fail "txbLabel harus berupa string dalam quote."
            If Len(Trim$(CStr(ins.Value))) = 0 Then Fail "txbLabel tidak boleh kosong."
        Case "cmbfontsize"
            If ins.Kind = "skip" Or ins.Kind = "default" Then Exit Sub
            If ins.Kind <> "number" And ins.Kind <> "string" Then _
                Fail "cmbFontSize harus berupa angka 4 sampai 24."
            If ins.Kind = "string" Then
                raw = CStr(ins.Value)
                For i = 1 To Len(raw)
                    If Mid$(raw, i, 1) >= "0" And Mid$(raw, i, 1) <= "9" Then
                        digits = digits + 1
                    ElseIf Mid$(raw, i, 1) = "." Then
                        dots = dots + 1
                    Else
                        Fail "cmbFontSize harus berupa angka 4 sampai 24."
                    End If
                Next i
                If digits = 0 Or dots > 1 Then Fail "cmbFontSize harus berupa angka 4 sampai 24."
                fontSize = Val(raw)
            Else
                fontSize = CDbl(ins.Value)
            End If
            If fontSize < 4 Or fontSize > 24 Then Fail "cmbFontSize harus berupa angka 4 sampai 24."
        Case "chkmasterpage", "chkuppercase", "chkbold", "chkitalic", _
             "chksmallcaps", "chkcolorize", "chkfraction", "chkemdash", _
             "optkisslabel", "opthidie", "opthikiss", "optunilabel"
            If ins.Kind = "skip" Or ins.Kind = "default" Then Exit Sub
            If ins.Kind <> "boolean" Then Fail ins.Target & " harus True atau False."
        Case Else
            Fail "Target tidak terdaftar: AutoLabelWizard." & ins.Target
    End Select
End Sub

Private Sub Fail(ByVal message As String)
    Dim prefix As String, line As Long, column As Long, lastBreak As Long

    prefix = Left$(pText, pAt - 1)
    line = UBound(Split(prefix, vbLf)) + 1
    lastBreak = InStrRev(prefix, vbLf)
    column = pAt - lastBreak
    Err.Raise 5, "ALBehaviorContract", "Semantic AutoLabel baris " & CStr(line) & _
        ", kolom " & CStr(column) & ": " & message
End Sub
