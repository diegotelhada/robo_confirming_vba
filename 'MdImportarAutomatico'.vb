'MdImportarAutomatico'

Option Compare Database

Sub ImportarArqOped() 'Direto da Rede

Dim Linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
Dim Dataarq As Date
Dim FSO As New FileSystemObject
Dim arq As File
Dim Contador As String
Dim TbDados As Recordset
Dim DataVenNF As Date
Dim DataOp As Date
Dim Juros As Double
Dim Custo As Double
Dim Spread As Double
Dim SpreadAnual As Double
Dim SpreadBanco As Double
Dim RateSpread As Double
Dim SpreadClte As Double
Dim ValorOp As Currency
Dim Valornom As Currency
Dim Valorabat As Currency
Dim Valoracres As Currency
Dim Valorpg As Currency
Dim ValorJuros As Currency
Dim valoriof As Currency
Dim Valorliq As Currency
Dim ValorCusto As Currency
Dim ReceitaBanco As Currency
Dim ReceitaClte As Currency
Dim Valortco As Currency
Dim valorttr As Currency
Dim DiaSemana As String
Dim TbDados1 As Recordset
'Dim ARQDATA As String

Set TbDados1 = CurrentDb.OpenRecordset("TblCalendario", dbOpenDynaset)

DiaSemana = UCase(WeekdayName(Weekday(Date), True))   'Retorna o nome do dia da semana abreviado

If UCase(DiaSemana) = "SEG" Then
ArqData = Date - 3
Else
ArqData = Date - 1
End If
'Encontra o Ultimo Dia util
TbDados1.FindFirst "Data_dia like '*" & ArqData & "'"
    
     If TbDados1.NoMatch = False Then
        Do While True
            If TbDados1!TIPO <> "UTIL" Then: ArqData = ArqData - 1

            If TbDados1!TIPO = "UTIL" Then: Exit Do
        Loop
     End If

TbDados1.Close

Contador = 0

Dia = Left(ArqData, 2)
Mes = Mid(ArqData, 4, 2)
Ano = Right(ArqData, 2)
Data = Dia & "/" & Mes & "/" & Ano

Set TbDados = CurrentDb.OpenRecordset("TblArqoped", dbOpenDynaset)

        TbDados.FindFirst "ArqData like '*" & ArqData & "'"
             
             If TbDados.NoMatch = False Then
             MsgBox "Arquivo ja Importado!", vbCritical, "Atenção"
             GoTo Fim
             End If

        DataPesq = Format(ArqData, "DDMMYY")
        'DataPesq = "300113"
        
        
        File = "\\saont46\apps2\Confirming\ArquivosYC\ARQOPED" & DataPesq & ".TXT"

Dataarq = Data

Set Caminho = FSO.GetFile(File)

    Open Caminho For Input As #1      'Abre Arquivo texto

        Do While Not EOF(1)
  
        Line Input #1, FileBuffer
        
        FileBuffer = SemCaracterEspecial(FileBuffer)
        
    TbDados.AddNew
     
        TbDados!ArqData = Dataarq ' Data do Arquivo
     
        TbDados!Segmento = Trim(Mid(FileBuffer, 1, 4)) 'Segmento
        
        'If Trim(Mid(FileBuffer, 5, 1)) = ";" Then
        TbDados!Nome_Ancora = Trim(Mid(FileBuffer, 6, 30)) 'Nome do Ancora
        'End If
     
        'If Trim(Mid(FileBuffer, 36, 1)) = ";" Then ' Numero do Banco do Ancora
        TbDados!Banco_Ancora = Trim(Mid(FileBuffer, 37, 4))
        'End If
     
        'If Trim(Mid(FileBuffer, 41, 1)) = ";" Then 'Numero da Agencia do Ancora
        TbDados!Agencia_Ancora = Trim(Mid(FileBuffer, 42, 4))
        'End If
     
        'If Trim(Mid(FileBuffer, 46, 1)) = ";" Then 'Numero do convenio do Ancora
        TbDados!Convenio_Ancora = Trim(Mid(FileBuffer, 47, 12))
        'End If
      
        'If Trim(Mid(FileBuffer, 59, 1)) = ";" Then ' CNPJ do Ancora
        TbDados!Cnpj_Ancora = Extrai_Zeros(Trim(Mid(FileBuffer, 60, 15)))
        'End If
      
        'If Trim(Mid(FileBuffer, 75, 1)) = ";" Then ' Nome do fornecedor
        TbDados!Nome_Fornecedor = SemCaracterEspecial(Trim(Mid(FileBuffer, 76, 40)))
        'End If
     
        'If Trim(Mid(FileBuffer, 116, 1)) = ";" Then ' CNPJ do fornecedor
        TbDados!Cnpj_Fornecedor = Extrai_Zeros(Trim(Mid(FileBuffer, 117, 15)))
        'End If
           
        'If Trim(Mid(FileBuffer, 132, 1)) = ";" Then ' Cod da Operação
        TbDados!Cod_Oper = Trim(Mid(FileBuffer, 133, 15))
        'End If
        
        'If Trim(Mid(FileBuffer, 148, 1)) = ";" Then ' Modalidade da operação "2 = Cessão de Crédito"
        TbDados!Modalidade_Oper = Trim(Mid(FileBuffer, 149, 1))
        'End If
        
        'If Trim(Mid(FileBuffer, 150, 1)) = ";" Then ' Tipo da Liquidação =  "1 - Crédito em C/C"  "2 - DOC" "3 - TED CIP" "4 - TED STR"
        TbDados!Tipo_Liq = Trim(Mid(FileBuffer, 151, 1))
        'End If
        
        'If Trim(Mid(FileBuffer, 152, 1)) = ";" Then ' Banco do Remetente
        TbDados!Banco_Remet = Trim(Mid(FileBuffer, 153, 4))
        'End If
        
        'If Trim(Mid(FileBuffer, 157, 1)) = ";" Then ' Agencia Remetente
        TbDados!Agencia_Remet = Trim(Mid(FileBuffer, 158, 4))
        'End If
        
        'If Trim(Mid(FileBuffer, 162, 1)) = ";" Then ' Conta Remetente
        TbDados!Conta_Remet = Trim(Mid(FileBuffer, 163, 12))
        'End If
        
        'If Trim(Mid(FileBuffer, 175, 1)) = ";" Then ' Tipo de Pagamento
        TbDados!Tipo_Pag = Trim(Mid(FileBuffer, 176, 1))
        'End If
        
        'If Trim(Mid(FileBuffer, 177, 1)) = ";" Then ' Banco favorecido
        TbDados!Banco_Fav = Extrai_Zeros(Trim(Mid(FileBuffer, 178, 5)))
        'End If
        
        'If Trim(Mid(FileBuffer, 183, 1)) = ";" Then ' Agencia favorecido
        TbDados!Agencia_Fav = Trim(Mid(FileBuffer, 184, 5))
        'End If
        
        'If Trim(Mid(FileBuffer, 189, 1)) = ";" Then ' Conta favorecido
        TbDados!Conta_Fav = Trim(Mid(FileBuffer, 190, 13))
        'Ed If
        
        'If Trim(Mid(FileBuffer, 203, 1)) = ";" Then ' Data da Operação
        DataOp = Trim(Mid(FileBuffer, 204, 10))
        TbDados!Data_op = DataOp
        'End If
       
       ' If Trim(Mid(FileBuffer, 214, 1)) = ";" Then ' Data Final
        TbDados!Data_Final = Trim(Mid(FileBuffer, 215, 10))
        'End If
    
        'If Trim(Mid(FileBuffer, 225, 1)) = ";" Then ' Prazo Médio
        TbDados!Prazo_Medio = Trim(Mid(FileBuffer, 226, 3))
        'End If
        
        'If Trim(Mid(FileBuffer, 229, 1)) = ";" Then ' Juros
        Juros = Trim(Mid(FileBuffer, 230, 13))
        TbDados!Juros = Juros / 10000000
        'End If
        
        'If Trim(Mid(FileBuffer, 243, 1)) = ";" Then ' Custo
        Custo = Trim(Mid(FileBuffer, 244, 13))
        TbDados!Custo = Custo / 10000000
        'End If
        
        'If Trim(Mid(FileBuffer, 257, 1)) = ";" Then ' Spread
        Spread = Trim(Mid(FileBuffer, 258, 13))
        TbDados!Spread = Spread / 10000000
        'End If
        
        'If Trim(Mid(FileBuffer, 271, 1)) = ";" Then ' Spread Anual
        SpreadAnual = Trim(Mid(FileBuffer, 272, 13))
        TbDados!Spread_Anual = SpreadAnual / 10000000
       ' End If
        
        'If Trim(Mid(FileBuffer, 285, 1)) = ";" Then ' Valor da Operação
        ValorOp = Trim(Mid(FileBuffer, 286, 17))
        TbDados!Valor_OP = ValorOp / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 303, 1)) = ";" Then ' Valor da TCO
        Valortco = Trim(Mid(FileBuffer, 304, 17))
        TbDados!Valor_TCO = Valortco / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 321, 1)) = ";" Then ' Valor da TTR
        valorttr = Trim(Mid(FileBuffer, 322, 17))
        TbDados!Valot_TTR = valorttr / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 339, 1)) = ";" Then ' Numero da Nota
        TbDados!Compromisso = Trim(Mid(FileBuffer, 340, 30))
        'End If
        
        'If Trim(Mid(FileBuffer, 370, 1)) = ";" Then ' Data de vencimento da Nota
        DataVenNF = Trim(Mid(FileBuffer, 371, 10))
        TbDados!Data_Venc = DataVenNF
        'End If
        
        'If Trim(Mid(FileBuffer, 381, 1)) = ";" Then ' Valor nominal da Nota
        Valornom = Trim(Mid(FileBuffer, 382, 17))
        TbDados!Valor_Nom = Valornom / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 399, 1)) = ";" Then ' Valor abatimento da Nota
        Valorabat = Trim(Mid(FileBuffer, 400, 17))
        TbDados!Valor_Abat = Valorabat / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 417, 1)) = ";" Then ' Valor acrescido da Nota
        Valoracres = Trim(Mid(FileBuffer, 418, 17))
        TbDados!Valor_Acres = Valoracres / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 435, 1)) = ";" Then ' Valor de pagamento da Nota
        Valorpg = Trim(Mid(FileBuffer, 436, 17))
        TbDados!Valor_Pagmto = Valorpg / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 453, 1)) = ";" Then ' Valor do Juros
        ValorJuros = Trim(Mid(FileBuffer, 454, 17))
        TbDados!Valor_Juros = ValorJuros / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 471, 1)) = ";" Then ' Valor do IOF
        valoriof = Trim(Mid(FileBuffer, 472, 17))
        TbDados!Valor_IOF = valoriof / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 489, 1)) = ";" Then ' Valor Liquido da Nota
        Valorliq = Trim(Mid(FileBuffer, 490, 17))
        TbDados!Valor_Liquido = Valorliq / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 507, 1)) = ";" Then ' Valor Custo
        ValorCusto = Trim(Mid(FileBuffer, 508, 17))
        TbDados!Valor_Custo = ValorCusto / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 525, 1)) = ";" Then ' Spread Banco
        SpreadBanco = Trim(Mid(FileBuffer, 526, 13))
        TbDados!Spread_Banco = SpreadBanco / 10000000
        'End If
        
        'If Trim(Mid(FileBuffer, 539, 1)) = ";" Then ' Receita Banco
        ReceitaBanco = Trim(Mid(FileBuffer, 540, 17))
        TbDados!Receita_Banco = ReceitaBanco / 100
        'End If
        
        'If Trim(Mid(FileBuffer, 557, 1)) = ";" Then ' PREMIO
        TbDados!Tp_Apur_prem = Trim(Mid(FileBuffer, 558, 1))
        'End If
        
        'If Trim(Mid(FileBuffer, 559, 1)) = ";" Then ' PREMIO
        TbDados!Tp_Rem_prem = Trim(Mid(FileBuffer, 560, 1))
        'End If
        
        'If Trim(Mid(FileBuffer, 561, 1)) = ";" Then ' PREMIO
        TbDados!Tp_Pgto_prem = Trim(Mid(FileBuffer, 562, 1))
        'End If
        
        'If Trim(Mid(FileBuffer, 563, 1)) = ";" Then ' PREMIO
        TbDados!Dt_pfto_Prem = Trim(Mid(FileBuffer, 564, 10))
        'End If
        
        'If Trim(Mid(FileBuffer, 574, 1)) = ";" Then ' PREMIO
        TbDados!Cod_Bco_Prem = Trim(Mid(FileBuffer, 575, 4))
        'End If
        
        'If Trim(Mid(FileBuffer, 579, 1)) = ";" Then ' PREMIO
        TbDados!Cod_Age_Prem = Trim(Mid(FileBuffer, 580, 4))
        'End If
        
        'If Trim(Mid(FileBuffer, 584, 1)) = ";" Then ' PREMIO
        TbDados!Cod_Conta_prem = Trim(Mid(FileBuffer, 585, 12))
        'End If
        
        'If Trim(Mid(FileBuffer, 597, 1)) = ";" Then ' PREMIO
        RateSpread = Trim(Mid(FileBuffer, 598, 13))
        TbDados!Rate_Spread = RateSpread / 10000000
        'End If
        
        'If Trim(Mid(FileBuffer, 611, 1)) = ";" Then ' PREMIO
        SpreadClte = Trim(Mid(FileBuffer, 612, 13))
        TbDados!Spread_Clte = SpreadClte / 10000000
        'End If
        
        'If Trim(Mid(FileBuffer, 625, 1)) = ";" Then ' PREMIO
        ReceitaClte = Trim(Mid(FileBuffer, 626, 17))
        TbDados!Receita_Clte = ReceitaClte / 100
        'End If
    
        TbDados!Prazo_NF = DataVenNF - DataOp
        
        Contador = Contador + 1
        
     TbDados.Update
        
        If Contador = 116 Then
        MsgBox ""
        End If
           
     Loop
     
     
    
Close #1


MsgBox "IMPORTADO COM SUCESSO", , "ATENÇÃO"

Fim:

End Sub

