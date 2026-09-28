Option Explicit

' Class Module (Name) = ALBehaviorSession. Keeps one Auto Label form instance per step.
Private pMenu As Object
Private pDefaults As Object

Public Sub Start(ByVal script As String, ByVal observer As Object, ByVal token As String)
    Dim contract As New ALBehaviorContract, block As MRBehaviorBlock
    Dim ins As MRBehaviorInstruction, existing As Object, value As Variant
    Dim operation As String, number As Long, source As String, description As String
    Dim formName As String

    On Error GoTo Failed
    operation = "Memvalidasi kontrak AutoLabel"
    Set block = contract.Validate(script)
    If StrComp(block.MacroID, "AutoLabelBP", vbTextCompare) = 0 Then
        formName = "AutoLabelWizardBP"
    Else
        formName = "AutoLabelWizard"
    End If

    operation = "Memeriksa " & formName & " yang sudah terbuka"
    For Each existing In VBA.UserForms
        If StrComp(TypeName(existing), formName, vbTextCompare) = 0 Then _
            Err.Raise 5, "ALBehaviorSession", "Tutup " & formName & " yang sudah terbuka sebelum menjalankan behavior."
    Next existing
    Set existing = Nothing

    operation = "Membuka " & formName
    Set pMenu = VBA.UserForms.Add(formName)
    pMenu.MRBindRunner observer, token
    pMenu.Show vbModeless
    operation = "Membaca nilai awal kontrol AutoLabel"
    SnapshotDefaults block

    For Each ins In block.Instructions
        operation = ins.FormName & " / " & ins.Target & " (posisi " & CStr(ins.Position) & ")"
        Select Case ins.Kind
            Case "stage", "skip"
                ' The initialized form keeps its saved value for omitted/skip targets.
            Case "action"
                pMenu.MRBehaviorSubmit
            Case Else
                value = ins.Value
                If ins.Kind = "default" Then value = pDefaults(LCase$(ins.Target))
                pMenu.MRBehaviorValue ins.Target, value
        End Select
    Next ins

    Set pMenu = Nothing
    Set pDefaults = Nothing
    Exit Sub

Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    ' A failed Submit leaves the visible form available for manual inspection.
    On Error Resume Next
    If Not pMenu Is Nothing Then
        pMenu.MRDetachRunner
        If Not pMenu.Visible Then Unload pMenu
    End If
    Set pMenu = Nothing
    Set pDefaults = Nothing
    On Error GoTo 0
    Err.Raise number, "ALBehaviorSession", operation & vbCrLf & _
        "Source asli: " & source & vbCrLf & description
End Sub

Private Sub SnapshotDefaults(ByVal block As MRBehaviorBlock)
    Dim ins As MRBehaviorInstruction, target As String

    Set pDefaults = CreateObject("Scripting.Dictionary")
    For Each ins In block.Instructions
        If ins.Kind = "default" Then
            target = LCase$(ins.Target)
            If Not pDefaults.Exists(target) Then _
                pDefaults.Add target, pMenu.MRBehaviorReadValue(ins.Target)
        End If
    Next ins
End Sub
