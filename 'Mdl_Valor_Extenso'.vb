'Mdl_Valor_Extenso'

Option Compare Database

Function Extenso_Valor(valor As Double) As String
    
    Dim strMoeda    As String
    Dim cents       As Variant
    Dim decimalSep  As String
    
    '   Se o valor for igual ou maior que 1 quatrilhao
    '   passar erro e sair da funcao
    If valor > 999999999999999# Then
        Extenso_Valor = "Valor excede 999.999.999.999.999"
        Exit Function
    End If
    
    '   Se valor for igual a 1, a unidade está no singular
    If arredBaixo(valor) = 1 Then
    '   a string da moeda no singular
        strMoeda = " real"
    '   Se for maior que 1 a unidade está no plural
    ElseIf arredBaixo(valor) > 1 Then
            strMoeda = " reais"
    End If
    
    '   Remove os centavos
    cents = valor - arredBaixo(valor)
    '   Remove os centavos do valor
    valor = valor - CDbl(cents)
    '   Passo o extenso dos centavos
    cents = centavos(CDbl(cents) * 100)

    '   Caso a string seja diferente de branco e valor seja maior ou igual a 1
    If cents <> "" And valor >= 1 Then
    '   acrescentar uma vírgula antes do extenso
        cents = " e " & cents
    End If
    
    '   Iniciar o processo de conversao dos valores longos
    strMoeda = Trim(Trilhoes(valor)) & strMoeda & cents
    strMoeda = Replace(strMoeda, ", e", " e")
    strMoeda = Replace(strMoeda, ", r", " r")
    
    If Left(strMoeda, 2) = "e " Then
        strMoeda = Mid(strMoeda, 3, Len(strMoeda))
    End If
    
    vzz = "00000000000000000000"
    vtam = Len(Trim(Mid(Trim(valor), 2, 100)))
    
    If Right(vzz + vzz + vzz + vzz, vtam) = Mid(Trim(valor), 2, 100) And InStr(UCase(strMoeda), UCase("es ")) > 0 Then
        vetor = Split(strMoeda, " ")
        vtrocar = vetor(UBound(vetor))
        strMoeda = Replace(strMoeda, vtrocar, "de " + vtrocar)
    End If
    
    strMoeda = StrConv(strMoeda, vbProperCase)
    strMoeda = Replace(strMoeda, "E", "e")
    
    Extenso_Valor = strMoeda
    
End Function

Private Function centavos(valor As Double) As String

    Dim dezena      As Integer
    Dim unidade     As Integer
    
    '   Passa o valor para base decimal
    valor = Round(CDbl(valor / 100), 2)
    
    '   Se for um centavo, escrever valor e sair da funcao
    If valor = 0.01 Then
        centavos = "um centavo"
        Exit Function
    End If
    
    '   Repassa valor para dezenas
    valor = valor * 100
    
    '   Se nao houver dezenas no valor passado
    If dezenas(valor) = "" Then
    '   a string centavos fica em branco
        centavos = ""
    Else
    '   caso contrário, passar extenso das dezenas e concatenar
    '   com a palavra centavos
        centavos = dezenas(valor) & " centavos"
    End If
    
End Function



Private Function unidades(unidade As Double) As String
    Dim unid(9)
    ' Define as unidades a serem usadas
    unid(1) = "um"
    unid(2) = "dois"
    unid(3) = "três"
    unid(4) = "quatro"
    unid(5) = "cinco"
    unid(6) = "seis"
    unid(7) = "sete"
    unid(8) = "oito"
    unid(9) = "nove"

    ' Retorna a string referente a unidade passada para
    ' esta funcao
    unidades = Trim(unid(unidade))
End Function

Private Function dezenas(dezena As Double) As String

    Dim dezes(9)
    Dim dez(9)
    Dim intDezena       As Double
    Dim intUnidade      As Double
    Dim tmpStr          As String
    
'   Define as dezenas a serem utilizadas
    dezes(1) = "onze"
    dezes(2) = "doze"
    dezes(3) = "treze"
    dezes(4) = "quatorze"
    dezes(5) = "quinze"
    dezes(6) = "dezesseis"
    dezes(7) = "dezessete"
    dezes(8) = "dezoito"
    dezes(9) = "dezenove"

    dez(1) = "dez"
    dez(2) = "vinte"
    dez(3) = "trinta"
    dez(4) = "quarenta"
    dez(5) = "cinquenta"
    dez(6) = "sessenta"
    dez(7) = "setenta"
    dez(8) = "oitenta"
    dez(9) = "noventa"

'   Calcula o inteiro da dezena
    intDezena = Int(dezena / 10)
    
'   Calcula o inteiro da unidade
    intUnidade = dezena Mod 10
    
'   Se o inteiro da dezena for zero
    If intDezena = 0 Then
'       dezenas sao iguais as unidades
        dezenas = unidades(intUnidade)
        Exit Function
    Else:
'       caso contrário, é igual a dez
        dezenas = dez(intDezena)
    End If
    
'   Se o inteiro da dezena for igual a 1 e
'   o inteiro da unidade for zero, os valores estao
'   entre 11 e 19
    If (intDezena = 1 And intUnidade > 0) Then
        dezenas = dezes(intUnidade)
    Else
'   Caso contrário, valor está entre 20 e 90 inclusive
        If (intDezena > 1 And intUnidade > 0) Then
'           Concatena a string da dezena com a string da unidade
            dezenas = dezenas & " e " & unidades(intUnidade)
        End If
    End If
    dezenas = dezenas

End Function

Private Function centenas(centena As Double) As String
    
    Dim tmpCento      As Double
    Dim tmpDez        As Double
    Dim tmpUni        As Double
    Dim tmpUniMod     As Double
    Dim tmpModDez     As Double
    Dim centoString   As String
    Dim cento(9)
    
'   Define as centenas
    cento(1) = "cento"
    cento(2) = "duzentos"
    cento(3) = "trezentos"
    cento(4) = "quatrocentos"
    cento(5) = "quinhentos"
    cento(6) = "seiscentos"
    cento(7) = "setecentos"
    cento(8) = "oitocentos"
    cento(9) = "novecentos"

'   Calcula o inteiro da centena
    tmpCento = Int(centena / 100)
    
'   Calcula a parte da dezena
    tmpDez = centena - (tmpCento * 100)
    
'   Calcula o inteiro da unidade
    tmpUni = Int(tmpDez / 10)
    
'   Calcula o resto da unidade
    tmpUniMod = tmpUni Mod 10
    
'   Calcula o resto da dezena
    tmpModDez = tmpDez Mod 10
    
'   Se centena for cem, definir string como "cem " e sair
    If centena = 100 Then
        centoString = "cem "
    Else
'   Caso contrário definir a string da centena
        centoString = cento(tmpCento)
    End If
    
'   Avalia se a unidade é maior ou igual a zero, se o resto da unidade é igual ou
'   maior que zero, se a dezena é maior ou igual a um e se a centena é igual ou
'   maior que 1. Se forem verdadeiros; entao, adicionar " e " a string da centena
    If (tmpUni >= 0 And tmpUniMod >= 0 And tmpDez >= 1 And tmpCento >= 1) Then
        centoString = centoString & " e "
    End If
    
'   Concatena a string do cento com a string da dezena
    centenas = Trim(centoString & dezenas(tmpDez))
    
End Function

Private Function milhares(milhar As Double) As String

    Dim tmpMilhar      As Double
    Dim tmpCento       As Double
    Dim milString      As String

'   Calcula o inteiro da milhar
    tmpMilhar = Int(milhar / 1000)
    
'   Calcula o cento dentro da milhar
    tmpCento = milhar - (tmpMilhar * 1000)
    
'   Se milhar for zero, entao a string da milhar fica em branco
    If tmpMilhar = 0 Then milString = ""
    
'   Se for igual a 1, entao
 '   If '(tmpMilhar = 1) Then
'       string da milhar é igual a unidade e "mil"
        'milString = unidades(tmpMilhar) & "um mil "
'       se maior que 1 e menor que dez, string igual a unidades
    If (tmpMilhar >= 1 And tmpMilhar < 10) Then
            milString = unidades(tmpMilhar) & " mil, "
'           Se for entre 10 e 100, entao string igual a dezenas
            ElseIf (tmpMilhar >= 10 And tmpMilhar < 100) Then
                milString = dezenas(tmpMilhar) & " mil, "
'               Se for entre 100 e 1000, entao igual string centenas
                ElseIf (tmpMilhar >= 100 And tmpMilhar < 1000) Then
                    milString = centenas(tmpMilhar) & " mil, "
    End If
    'If tmpCento = 1 Then milString = " e "
    If (tmpCento >= 1 And tmpCento <= 100) Then milString = milString & "e "
    milhares = Trim(milString & centenas(tmpCento))
    
End Function

Private Function milhoes(milhao As Double) As String

'   Ver comentários para milhares acima
    Dim tmpMilhao      As Double
    Dim tmpMilhares    As Double
    Dim miString       As String
    
    tmpMilhao = Int(milhao / 1000000)
    tmpMilhares = milhao - (tmpMilhao * 1000000)
    
    If tmpMilhao = 0 Then miString = ""
    If (tmpMilhao = 1) Then
        miString = unidades(tmpMilhao) & " milhão, "
        ElseIf (tmpMilhao > 1 And tmpMilhao < 10) Then
            miString = unidades(tmpMilhao) & " milhões, "
            ElseIf (tmpMilhao >= 10 And tmpMilhao < 100) Then
                miString = dezenas(tmpMilhao) & " milhões, "
                ElseIf (tmpMilhao >= 100 And tmpMilhao < 1000) Then
                    miString = centenas(tmpMilhao) & " milhões, "
    End If
    
    If milhao = 1000000# Then miString = "um milhão de "
    milhoes = Trim(miString & milhares(tmpMilhares))
    
End Function

Private Function bilhoes(bilhao As Double) As String

'   Ver comentários para milhares acima
    Dim tmpBilhao     As Double
    Dim tmpMilhao       As Double
    'Dim tmpMilhoes      As Double
    Dim biString       As String

    tmpBilhao = Int(bilhao / 1000000000)
    tmpMilhao = bilhao - (tmpBilhao * 1000000000)
    
    If (tmpBilhao = 1) Then
        biString = unidades(tmpBilhao) & " bilhão, "
        ElseIf (tmpBilhao > 1 And tmpBilhao < 10) Then
            biString = unidades(tmpBilhao) & " bilhões, "
            ElseIf (tmpBilhao >= 10 And tmpBilhao < 100) Then
                biString = dezenas(tmpBilhao) & " bilhões, "
                ElseIf (tmpBilhao >= 100 And tmpBilhao < 1000) Then
                    biString = centenas(tmpBilhao) & " bilhões, "
    End If
    
    If bilhao = 1000000000# Then biString = "um bilhão de "
    bilhoes = Trim(biString & milhoes(tmpMilhao))
    
End Function

Private Function Trilhoes(Trilhao As Double) As String

    ' Ver comentários para milhares acima
    Dim tmpTrilhao     As Double
    Dim tmpBilhao       As Double
    Dim triString       As String

    tmpTrilhao = Int(Trilhao / 1000000000000#)
    tmpBilhao = Trilhao - (tmpTrilhao * 1000000000000#)
    
    If (tmpTrilhao = 1) Then
        triString = unidades(tmpTrilhao) & " trilhão, "
        ElseIf (tmpTrilhao > 1 And tmpTrilhao < 10) Then
            triString = unidades(tmpTrilhao) & " trilhões, "
            ElseIf (tmpTrilhao >= 10 And tmpTrilhao < 100) Then
                triString = dezenas(tmpTrilhao) & " trilhões, "
                ElseIf (tmpTrilhao >= 100 And tmpTrilhao < 1000) Then
                    triString = centenas(tmpTrilhao) & " trilhões, "
    End If
    
    If Trilhao = 1000000000000# Then triString = "um trilhão de "
    Trilhoes = Trim(triString & bilhoes(tmpBilhao))
    
End Function

Function arredBaixo(valor)

    Dim tmpValor
    tmpValor = Round(CDbl(Right(Round(valor, 2) * 100, 2)) / 100, 2)
    arredBaixo = Round(Round(valor, 2) - tmpValor, 0)
    
End Function

Function Extenso_Taxa(taxa As Double) As String
    
    Dim strMoeda    As String
    Dim decimais       As Variant
    Dim decimalSep  As String
    Dim isInteiro As Boolean
    
    '   Identifica os decimais
    decimais = taxa - arredBaixo(taxa)
    '   Remove os decimais da taxa
    taxa = taxa - CDbl(decimais)
    '   Passa o extenso dos decimais
    decimais = porcentos(CDbl(decimais) * 100)
    
    '   Verifica se a taxa apresentada é um número inteiro
    If decimais = 0 Or decimais = "" Then isInteiro = True

    '   Caso a string seja diferente de branco e taxa seja maior ou igual a 1
    If decimais <> "" And taxa >= 1 Then
    '   acrescentar uma vírgula antes do extenso
        decimais = " e " & decimais
    End If
    
    '   Iniciar o processo de conversao das taxas
    strMoeda = Trim(Trilhoes(taxa)) & strMoeda & decimais
    strMoeda = Replace(strMoeda, ", e", " e")
    strMoeda = Replace(strMoeda, ", r", " r")
    
    If Left(strMoeda, 2) = "e " Then
        strMoeda = Mid(strMoeda, 3, Len(strMoeda))
    End If
    
    vzz = "00000000000000000000"
    vtam = Len(Trim(Mid(Trim(taxa), 2, 100)))
    
    If Right(vzz + vzz + vzz + vzz, vtam) = Mid(Trim(taxa), 2, 100) And InStr(UCase(strMoeda), UCase("es ")) > 0 Then
        vetor = Split(strMoeda, " ")
        vtrocar = vetor(UBound(vetor))
        strMoeda = Replace(strMoeda, vtrocar, "de " + vtrocar)
    End If
    
    If taxa < 1 Then strMoeda = "Zero e " & strMoeda
    If isInteiro Then strMoeda = strMoeda & " por cento"
    
    strMoeda = StrConv(strMoeda, vbProperCase)
    strMoeda = Replace(strMoeda, "E", "e")
    strMoeda = Replace(strMoeda, "Por Cento", "por cento")
    
    Extenso_Taxa = strMoeda
    
End Function

Private Function porcentos(valor As Double) As String

    Dim dezena      As Integer
    Dim unidade     As Integer
    
    '   Passa o valor para base decimal
    valor = Round(CDbl(valor / 100), 2)
    
    '   Repassa valor para dezenas
    valor = valor * 100
    
    '   Se nao houver dezenas no valor passado
    If dezenas(valor) = "" Then
    '   a string porcentos fica em branco
        porcentos = ""
    Else
    '   caso contrário, passar extenso das dezenas e concatenar
    '   com a palavra por cento
        porcentos = dezenas(valor) & " por cento"
    End If
    
End Function