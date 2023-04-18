'MdlExtrairZero'

Public Function Extrai_Zeros(Cpf_CNPJ As String)

Dim i        As Integer
Dim num_pos  As Integer

i = 0
For i = 1 To Len(Cpf_CNPJ)
    If InStr(1, Cpf_CNPJ, "0") > 0 Then
        If Left(Cpf_CNPJ, 1) = 0 Then
            num_pos = IIf(i = 1, 1, i - 1)
            Cpf_CNPJ = Right(Cpf_CNPJ, Len(Cpf_CNPJ) - 1)
        End If
    End If
    i = i + 1
Next i

Extrai_Zeros = Cpf_CNPJ

End Function

