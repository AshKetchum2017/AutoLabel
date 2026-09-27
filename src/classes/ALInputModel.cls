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

Public Function Resolve(ByVal templateText As String, ByVal sourceDirectory As String) As String
    Dim position As Long
    Dim customerName As String
    Dim result As String

    If Len(pMode) = 0 Then Err.Raise 5, "ALInputModel", "Mode Auto Label belum ditentukan."
    pHasToken = False
    pPartNumber = vbNullString

    position = 1
    Do While position <= Len(templateText)
        If Mid$(templateText, position, 2) = "/*" Then
            If Mid$(templateText, position, 3) <> "/*~" Then
                Err.Raise 5, "ALInputModel", "Token tidak dikenal pada posisi " & CStr(position) & "."
            End If

            If Not pHasToken Then customerName = CustomerFromDirectory(sourceDirectory)
            result = result & customerName
            pHasToken = True
            position = position + 3
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
