'MdlSemCarater'

Option Compare Database

Public Function SemCaracterEspecial(ByVal palavra As String) As String
Dim cont As Integer
SemCaracterEspecial = palavra
For cont = 1 To Len(SemCaracterEspecial)
  Select Case Mid(SemCaracterEspecial, cont, 1)
    Case "Ç"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "C" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
    Case "ç"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "c" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
    Case "á", "à", "ã", "â", "ª"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "a" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
    Case "Á", "À", "Ã", "Â"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "A" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
    Case "É", "È", "Ê"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "E" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
    Case "é", "è", "ê"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "e" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "Í", "Ì", "Î"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "I" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "í", "ì", "î"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "i" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "Ó", "Ò", "Õ", "Ô"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "O" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "ó", "ò", "õ", "ô", "º", "°"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "o" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "Ú", "Ù", "Û", "Ü"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "U" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "ú", "ù", "û", "ü"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "u" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "@", "#", "$", "%", "$", "&", "§"
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "o" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case "", "", ""
      SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "o" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
     Case Else
  End Select
Next
End Function

