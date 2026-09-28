Option Explicit

' Import as a Class Module named ALInputModel in each Auto Label GMS.
Private pMode As String
Private pHasToken As Boolean
Private pPartNumber As String

Public Property Let Mode(ByVal value As String)
    Select Case LCase$(Trim$(value))
        Case "normal", "bp"
            pMode = LCase$(Trim$(value))
        Case Else
            Err.Raise 5, "ALInputModel", "Mode Auto Label tidak dikenal: " & value
    End Select
End Property

Public Property Get Mode() As String
    Mode = pMode
End Property

Public Property Get HasToken() As Boolean
    HasToken = pHasToken
End Property

Public Property Get PartNumber() As String
    PartNumber = pPartNumber
End Property

Public Function Resolve(ByVal templateText As String, ByVal sourceDirectory As String, _
    Optional ByVal useFraction As Boolean = False, Optional ByVal useEmDash As Boolean = False) As String
    Dim position As Long
    Dim customerName As String
    Dim result As String
    Dim token As String
    Dim dayText As String, monthText As String, yearText As String
    Dim separator As String
    Dim customerReady As Boolean, dateReady As Boolean

    If Len(pMode) = 0 Then Err.Raise 5, "ALInputModel", "Mode Auto Label belum ditentukan."
    pHasToken = False
    pPartNumber = vbNullString
    If useFraction Then
        separator = "/"
    ElseIf useEmDash Then
        separator = "-"
    End If

    position = 1
    Do While position <= Len(templateText)
        If Mid$(templateText, position, 2) = "/*" Then
            Select Case True
                Case Mid$(templateText, position, 3) = "/*~"
                    token = "/*~"
                    If Not customerReady Then
                        customerName = CustomerFromDirectory(sourceDirectory)
                        customerReady = True
                    End If
                    result = result & customerName
                Case Mid$(templateText, position, 8) = "/*ddmmyy"
                    token = "/*ddmmyy"
                Case Mid$(templateText, position, 6) = "/*ddmm"
                    token = "/*ddmm"
                Case Mid$(templateText, position, 4) = "/*dd"
                    token = "/*dd"
                Case Mid$(templateText, position, 4) = "/*mm"
                    token = "/*mm"
                Case Mid$(templateText, position, 4) = "/*yy"
                    token = "/*yy"
                Case Else
                    Err.Raise 5, "ALInputModel", "Token tidak dikenal pada posisi " & CStr(position) & "."
            End Select

            If token <> "/*~" Then
                If Not dateReady Then
                    SourceDateParts sourceDirectory, dayText, monthText, yearText
                    dateReady = True
                End If
                Select Case token
                    Case "/*ddmmyy": result = result & dayText & separator & monthText & separator & Right$(yearText, 2)
                    Case "/*ddmm": result = result & dayText & separator & monthText
                    Case "/*dd": result = result & dayText
                    Case "/*mm": result = result & monthText
                    Case "/*yy": result = result & Right$(yearText, 2)
                End Select
            End If
            pHasToken = True
            position = position + Len(token)
        Else
            result = result & Mid$(templateText, position, 1)
            position = position + 1
        End If
    Loop

    Resolve = result
End Function

Private Function CustomerFromDirectory(ByVal sourceDirectory As String) As String
    Dim folderPath As String
    Dim folderName As String
    Dim separatorPosition As Long

    folderPath = sourceDirectory
    Do While Len(folderPath) > 0 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    If Len(folderPath) = 0 Or Right$(folderPath, 1) = ":" Then
        Err.Raise 5, "ALInputModel", "Simpan dokumen CDR di dalam folder pelanggan sebelum memakai token /*~."
    End If

    separatorPosition = InStrRev(folderPath, "\")
    If separatorPosition = 0 Then
        Err.Raise 5, "ALInputModel", "Folder pelanggan untuk token /*~ tidak ditemukan."
    End If
    folderName = Mid$(folderPath, separatorPosition + 1)

    If TryPartNumber(folderName, pPartNumber) Then
        folderPath = Left$(folderPath, separatorPosition - 1)
        separatorPosition = InStrRev(folderPath, "\")
        If separatorPosition = 0 Then
            Err.Raise 5, "ALInputModel", "Folder pelanggan sebelum folder PART tidak ditemukan."
        End If
        folderName = Mid$(folderPath, separatorPosition + 1)
    End If

    CustomerFromDirectory = LCase$(Trim$(folderName))
    If Len(CustomerFromDirectory) = 0 Then
        Err.Raise 5, "ALInputModel", "Nama folder pelanggan untuk token /*~ kosong."
    End If
End Function

Private Sub SourceDateParts(ByVal sourceDirectory As String, _
    ByRef dayText As String, ByRef monthText As String, ByRef yearText As String)
    Dim fso As Object
    Dim dayPath As String, monthPath As String
    Dim dayName As String, monthName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    dayPath = sourceDirectory
    Do While Len(dayPath) > 3 And Right$(dayPath, 1) = "\"
        dayPath = Left$(dayPath, Len(dayPath) - 1)
    Loop

    Do While Len(dayPath) > 0
        monthPath = fso.GetParentFolderName(dayPath)
        dayName = fso.GetFileName(dayPath)
        monthName = fso.GetFileName(monthPath)
        If IsDayFolder(dayName) And IsMonthFolder(monthName) Then
            yearText = SourceYear(fso.GetParentFolderName(monthPath))
            dayText = Right$("0" & CStr(CLng(dayName)), 2)
            monthText = Left$(monthName, 2)
            Exit Sub
        End If
        If StrComp(dayPath, monthPath, vbTextCompare) = 0 Then Exit Do
        dayPath = monthPath
    Loop

    Err.Raise 5, "ALInputModel", "Folder tanggal dan bulan sumber tidak ditemukan untuk token tanggal."
End Sub

Private Function IsMonthFolder(ByVal folderName As String) As Boolean
    Dim months As Variant, index As Long

    months = Split("JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER", "|")
    For index = 0 To UBound(months)
        If StrComp(folderName, Right$("0" & CStr(index + 1), 2) & ". " & months(index), vbTextCompare) = 0 Then
            IsMonthFolder = True
            Exit Function
        End If
    Next index
End Function

Private Function IsDayFolder(ByVal folderName As String) As Boolean
    Dim index As Long

    If Len(folderName) < 1 Or Len(folderName) > 2 Then Exit Function
    For index = 1 To Len(folderName)
        If InStr(1, "0123456789", Mid$(folderName, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsDayFolder = (CLng(folderName) >= 1 And CLng(folderName) <= 31)
End Function

Private Function SourceYear(ByVal ancestorPath As String) As String
    Dim fso As Object
    Dim currentPath As String, folderName As String, yearText As String
    Dim pass As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' As in AutoSNC: prefer a four-digit year folder, then scan folder names.
    For pass = 1 To 2
        currentPath = ancestorPath
        Do While Len(currentPath) > 0
            If Left$(currentPath, 2) = "\\" And InStr(3, currentPath, "\") = 0 Then Exit Do
            folderName = fso.GetFileName(currentPath)
            If pass = 1 Then
                If IsFourDigitYear(folderName) Then
                    SourceYear = folderName
                    Exit Function
                End If
            Else
                yearText = YearInFolderName(folderName)
                If Len(yearText) > 0 Then
                    SourceYear = yearText
                    Exit Function
                End If
            End If
            currentPath = fso.GetParentFolderName(currentPath)
        Loop
    Next pass
    Err.Raise 5, "ALInputModel", "Tahun sumber tidak ditemukan untuk token tanggal."
End Function

Private Function IsFourDigitYear(ByVal value As String) As Boolean
    Dim index As Long

    If Len(value) <> 4 Then Exit Function
    For index = 1 To 4
        If InStr(1, "0123456789", Mid$(value, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsFourDigitYear = (Left$(value, 1) <> "0")
End Function

Private Function YearInFolderName(ByVal folderName As String) As String
    Dim index As Long, digits As String, character As String

    For index = 1 To Len(folderName) + 1
        character = Mid$(folderName & " ", index, 1)
        If InStr(1, "0123456789", character, vbBinaryCompare) > 0 Then
            digits = digits & character
        Else
            If IsFourDigitYear(digits) Then
                If Len(YearInFolderName) > 0 And YearInFolderName <> digits Then
                    Err.Raise 5, "ALInputModel", "Tahun ambigu pada folder " & folderName & "."
                End If
                YearInFolderName = digits
            End If
            digits = vbNullString
        End If
    Next index
End Function

Private Function TryPartNumber(ByVal folderName As String, ByRef partNumber As String) As Boolean
    Dim numberText As String
    Dim position As Long

    If StrComp(Left$(folderName, 4), "PART", vbTextCompare) <> 0 Then Exit Function
    numberText = Mid$(folderName, 5)
    If Len(numberText) = 0 Then
        Err.Raise 5, "ALInputModel", "Nomor pada folder PART kosong."
    End If
    If Left$(numberText, 1) <> " " And Not Left$(numberText, 1) Like "[0-9]" Then Exit Function

    numberText = Trim$(numberText)
    If Len(numberText) = 0 Then
        Err.Raise 5, "ALInputModel", "Nomor pada folder PART kosong."
    End If
    For position = 1 To Len(numberText)
        If Not Mid$(numberText, position, 1) Like "[0-9]" Then
            Err.Raise 5, "ALInputModel", "Folder PART harus diikuti angka: " & folderName
        End If
    Next position

    Do While Len(numberText) > 1 And Left$(numberText, 1) = "0"
        numberText = Mid$(numberText, 2)
    Loop
    partNumber = numberText
    TryPartNumber = True
End Function
