'Mdl_Liquidacoes'

Option Compare Database
'===========================================================================================================================
'MODULO: Mdl_Liquidacoes
'IMPORTAR: LIQUIDADOS e BOLETOS
'GERAR: BOLETOS NAO LIQUIDADOS e TITULOS EM TEIMOSINHA
'ULTIMA ATUALIZAÇÃO: 27/04/2017
'USUARIO: MARCELO HENRIQUE DE SOUZA
'===========================================================================================================================
Function EmailLiquidacoes(Tipo, Assunto, Arquivo)

    'Criar Objeto Outlook
    Set sbObj = New Scripting.FileSystemObject
    Set olapp = CreateObject("Outlook.Application")
    Set oitem = olapp.CreateItem(0)
        
        'Atributos do Email(Remetente, Assunto, Destino, Copia)
        oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
        oitem.Subject = ("RELATORIO DE " & Assunto & " - CONFIRMING")
        oitem.BCC = "wellington.da.silva@santander.com.br;mlalmeida@santander.com.br;guialmeida@santander.com.br;beasilva@santander.com.br;rmmosti@santander.com.br;jorge.junior@santander.com.br;lfrossi@santander.com.br;emanuela.conceicao@santander.com.br"
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGOAtualizado.jpg"
                                        
            'Corpo do Email
            Corpo1 = "Prezados,"
            Relatorio = "RELATORIO DE " & Assunto
            Corpo21 = " referente as operações do ultimo dia útil."
            Corpo3 = ""
            Confirming = ""
            Corpo31 = ""
            Fones = ""
            Corpo4 = "Atenciosamente."
            Assinatura1 = "Manufatura"
            Assinatura2 = "Contratos e Minutarias"
            Assinatura3 = "Confirming"
            Assinatura4 = "Rua Amador Bueno, 474"
            Assinatura5 = "CEP: 04752-005  São Paulo-SP"
            Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
            Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
            Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
                                            
                'Valida tipo de email
                If Tipo = "SUCESSO" Then
                    Corpo2 = "Segue anexo o "
                    oitem.Attachments.Add Arquivo
                Else
                    Corpo2 = "Hoje não teremos o "
                End If

                oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Calibri <BR>" & Corpo1 & "<BR/>" & _
                "<BR>" & Corpo2 & "<B>" & Relatorio & "</B>" & Corpo21 & "<BR>" & "<BR>" & Corpo4 & "<BR><BR><BR>" & _
                " <img src=" & Assinatura & " height=100 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 3" & "<BR>" & _
                "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 2 <BR>" & Assinatura3 & _
                "<BR/>" & Assinatura4 & "<BR/>" & Assinatura5 & "<BR/></FONT><FONT COLOR = BLACK FACE = Calibri Size = 1 <BR><I>" & Assinatura6 & _
                "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"
            
            oitem.Send

End Function
Sub ImportarBoletos()

    Dim FSO As New FileSystemObject, arq As File, File As String, valor As Currency
    Dim Contador As String, TbBoletos As Recordset, TbArq As Recordset, Dataout As Date
    Dim linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
        
        'Abrir base de dados
        Call AbrirBDBoletos
        
        'Pesq ultima dia util
        ARQDATA = UltimoDiaUtil: DataPesq = Format(ARQDATA, "DDMMYY")
        
        'Arquivo para importar
        File = "\\saont46\apps2\Confirming\ArquivosYC\YCXYK-BOLETOS-GERADOS" & DataPesq & ".TXT"
        ''Novo Caminho
        File = "\\fscore02\apps2\Confirming\ArquivosYC\YCXYK-BOLETOS-GERADOS" & DataPesq & ".TXT"
        
            'Valida arquivo na maquina
            sFname = "C:\temp\BOLETOS.TXT"
            If (Dir(sFname) <> "") Then
                Kill sFname
            End If
                       
            'Tirar caracteres especiais
            Call ReplaceEXE(File)
            Call TratarArquivoDeBoletos(File)
            
                'Abrir tabelas
                Set Caminho = FSO.GetFile(File)
                Set TbArq = BDB.OpenRecordset("BOLETOS")
                Set TbBoletos = BDB.OpenRecordset("Tbl_boletos", dbOpenDynaset)
                
                'Importar Arquivo
                Do While TbArq.EOF = False
                    If TbArq!INSTRUCAO = "REGISTRO DE BOLETO" Then
                        TbBoletos.AddNew
                            TbBoletos!ARQDATA = ARQDATA
                            TbBoletos!Agencia = Mid(TbArq!Convenio, 5, 4)
                            TbBoletos!Convenio = Right(TbArq!Convenio, 12)
                            TbBoletos![nome_convenio] = TbArq![NOME DO CONVENIO]
                            TbBoletos!Tipo = TbArq!TPDOC
                            TbBoletos!Cnpj_Fornecedor = TbArq!NDOCUMENTO
                            TbBoletos!Nome_Fornecedor = TbArq![NOME/RAZAO SOCIAL]
                            TbBoletos!COD_OPERACAO = TbArq!OPERACAO
                            TbBoletos!Compromisso = TbArq!Compromisso
                            TbBoletos!VALOR_COMPROMISSO = TbArq![VALOR COMPROMISSO]
                            TbBoletos!Data_Venc = TbArq![DATA VENC]
                            TbBoletos!NUM_TITULO = TbArq![NUM TIT DESC *]
                        TbBoletos.Update
                       End If
                    TbArq.MoveNext
                Loop
        
        'Fechar Arquivos
        TbArq.Close
        TbBoletos.Close

End Sub
Sub ImportarBaixados()

    Dim Contador As String, TbLiquidados, TbArq As Recordset, Dataout As Date
    Dim FSO As New FileSystemObject, arq As File, File As String, valor As Currency
    Dim linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
        
        'Abrir base de dados
        Call AbrirBDBoletos
        
        'Pesquisar ultimo dia util
        ARQDATA = UltimoDiaUtil(): DataPesq = Format(ARQDATA, "DDMMYY")
                            
        'Arquivo para importação
        File = "\\saont46\apps2\Confirming\ArquivosYC\ARQCOMPBAIXADOYD" & DataPesq & ".TXT"
        ''Novo Caminho
        File = "\\fscore02\apps2\Confirming\ArquivosYC\ARQCOMPBAIXADOYD" & DataPesq & ".TXT"
        
            
        For Each Files In FSO.GetFolder("\\D1691641\Publica\Confirming\LIQUIDACOES\").Files
            
            File = Files

            'Valida arquivo na maquina
            sFname = "C:\temp\ARQCOMPBAIXADO.TXT"
            If (Dir(sFname) <> "") Then
                Kill sFname
            End If
            
            'Tirar caracteres especiais
            Call ReplaceEXE(File)
            Call TratarArquivoDeBaixados(File)
                
                Set arq = FSO.GetFile(sFname)
                
                If arq.Size > 0 Then
                    BDB.Execute ("INSERT INTO Tbl_Liquidados ( ARQDATA, Agencia, Convenio, NOME_CONVENIO, CNPJ_FORNECEDOR, NOME_FORNCEDOR, COD_OPERACAO, Compromisso, NUM_TITULO, VALOR_COMPROMISSO, DIFERENÇA, DATA_LIQ ) SELECT Right([BAIXADOYD]![DataBaixa],2) & '/' & Mid([BAIXADOYD]![DataBaixa],6,2) & '/' & Left([BAIXADOYD]![DataBaixa],4) AS Arqdata, BAIXADOYD.Agencia, BAIXADOYD.Convenio, BAIXADOYD.NomeAncora, BAIXADOYD.CNPJFornecedor, BAIXADOYD.NomeFornecedor, BAIXADOYD.CodOperacao, BAIXADOYD.Compromisso, BAIXADOYD.NumTitulo, BAIXADOYD.valorOperacao, BAIXADOYD.Juros, Right([BAIXADOYD]![DataBaixa],2) & '/' & Mid([BAIXADOYD]![DataBaixa],6,2) & '/' & Left([BAIXADOYD]![DataBaixa],4) AS Baixa FROM BAIXADOYD;")
                End If
            Debug.Print Files
        Next Files

End Sub
Sub ImportarLiquidados()
    
    Dim Contador As String, TbLiquidados, TbArq As Recordset, Dataout As Date
    Dim FSO As New FileSystemObject, arq As File, File As String, valor As Currency
    Dim linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
        
        'Abrir base de dados
        Call AbrirBDBoletos
        
        'Pesquisar ultimo dia util
        ARQDATA = UltimoDiaUtil(): DataPesq = Format(ARQDATA, "DDMMYY")
                            
        'Arquivo para importação
        File = "\\saont46\apps2\Confirming\ArquivosYC\YCXYD-LIQUIDACOES" & DataPesq & ".TXT"
           
        ''Novo Caminho
        File = "\\fscore02\apps2\Confirming\ArquivosYC\YCXYD-LIQUIDACOES" & DataPesq & ".TXT"
        
            'For Each Files In FSO.GetFolder("\\D1691641\Publica\Confirming\LIQUIDACOES\").Files
            
            'File = Files
            
            'Valida arquivo na maquina
            sFname = "C:\temp\LIQUIDACOES.TXT"
            If (Dir(sFname) <> "") Then
                Kill sFname
            End If
            
            'Tirar caracteres especiais
            Call ReplaceEXE(File)
            Call TratarArquivoDeLiquidados(File)
                
                'Abrir tabelas
                Set TbLiquidados = BDB.OpenRecordset("Tbl_Liquidados", dbOpenDynaset)
                Set TbArq = BDB.OpenRecordset("LIQUIDACOES")
                
                'Importar Arquivo
                Do While TbArq.EOF = False
                        TbLiquidados.AddNew
                            TbLiquidados!ARQDATA = Format(TbArq![DATA LIQ], "dd/mm/yyyy")
                            TbLiquidados!Agencia = Mid(TbArq!Convenio, 5, 4)
                            TbLiquidados!Convenio = Right(TbArq!Convenio, 12)
                            TbLiquidados!nome_convenio = SemCaracterEspecial(TbArq![NOME DO CONVENIO])
                            TbLiquidados!Tipo = TbArq!TPDOC
                            TbLiquidados!Cnpj_Fornecedor = Extrai_Zeros(TbArq!NDOCUMENTO)
                            TbLiquidados!NOME_FORNCEDOR = SemCaracterEspecial(TbArq![NOME/RAZAO SOCIAL])
                            TbLiquidados!COD_OPERACAO = TbArq!OPERACAO
                            TbLiquidados!Compromisso = TbArq!Compromisso
                            If IsNumeric(TbArq![VALOR PAGO]) Then: TbLiquidados!VALOR_PAGO = TbArq![VALOR PAGO]
                            If IsNumeric(TbArq![VALOR NOMINAL]) Then: TbLiquidados!VALOR_COMPROMISSO = TbArq![VALOR NOMINAL]
                            If IsNumeric(TbArq!DIFERENCA) Then: TbLiquidados!DIFERENÇA = TbArq!DIFERENCA
                            TbLiquidados!DATA_LIQ = Format(TbArq![DATA LIQ], "dd/mm/yyyy")
                            TbLiquidados!NUM_TITULO = TbArq![NUM TIT DESC *]
                        TbLiquidados.Update
                    TbArq.MoveNext
                Loop
            'Fechar Arquivos
            'Debug.Print Files
            TbArq.Close
        'Next Files
End Sub
Sub AtualizarBoletos()

    Call AbrirBDBoletos

        BDB.Execute ("UPDATE Tbl_boletos INNER JOIN Tbl_Liquidados ON Tbl_boletos.NUM_TITULO = Tbl_Liquidados.NUM_TITULO SET Tbl_boletos.SITUACAO = 'LIQUIDADO';")

End Sub
Sub BoletosNaoLiquidados()

    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim ObjPlan1Excel As Object, linha As Double, Ret As String
    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
    
        'Abrir base de dados
        Call AbrirBDBoletos
            
            'Gera o dia da semana
            DiaSemana = UCase(WeekdayName(Weekday(Date), True)): ARQDATA = UltimoDiaUtil(): DataPesq = Format(ARQDATA, "mm/dd/yyyy")

            'Pesquisar boletos não liquidados
            If DiaSemana = "TER" Then
                DtFim = ARQDATA: DtFim = DtFim - 2: ARQDATA = DtFim: DataPesq = Format(ARQDATA, "mm/dd/yyyy")
                Set TbDados = BDB.OpenRecordset("SELECT Tbl_Boletos.ARQDATA, Tbl_Boletos.AGENCIA, Tbl_Boletos.CONVENIO, Tbl_Boletos.NOME_CONVENIO, Tbl_Boletos.TIPO, Tbl_Boletos.CNPJ_FORNECEDOR, Tbl_Boletos.NOME_FORNECEDOR, Tbl_Boletos.COD_OPERACAO, Tbl_Boletos.COMPROMISSO, Tbl_Boletos.VALOR_COMPROMISSO, Tbl_Boletos.DATA_VENC, Tbl_Boletos.NUM_TITULO, 'NÃO LIQUIDADO' AS SITU FROM Tbl_Boletos WHERE (((Tbl_Boletos.NOME_CONVENIO) Not Like '*NESTLE*' And (Tbl_Boletos.NOME_CONVENIO) Not Like '*CHOCOLATES GAROTO*') AND ((Tbl_Boletos.DATA_VENC)>=#" & DataPesq & "# And (Tbl_Boletos.DATA_VENC)<Date()) AND ((Tbl_Boletos.SITUACAO) Is Null)) ORDER BY Tbl_Boletos.NOME_CONVENIO;", dbOpenDynaset)
            Else
                Set TbDados = BDB.OpenRecordset("SELECT Tbl_Boletos.ARQDATA, Tbl_Boletos.AGENCIA, Tbl_Boletos.CONVENIO, Tbl_Boletos.NOME_CONVENIO, Tbl_Boletos.TIPO, Tbl_Boletos.CNPJ_FORNECEDOR, Tbl_Boletos.NOME_FORNECEDOR, Tbl_Boletos.COD_OPERACAO, Tbl_Boletos.COMPROMISSO, Tbl_Boletos.VALOR_COMPROMISSO, Tbl_Boletos.DATA_VENC, Tbl_Boletos.NUM_TITULO, 'NÃO LIQUIDADO' AS SITU FROM Tbl_Boletos WHERE (((Tbl_Boletos.NOME_CONVENIO) Not Like '*NESTLE*' And (Tbl_Boletos.NOME_CONVENIO) Not Like '*CHOCOLATES GAROTO*') AND ((Tbl_Boletos.DATA_VENC)=#" & DataPesq & "#) AND ((Tbl_Boletos.SITUACAO) Is Null)) ORDER BY Tbl_Boletos.NOME_CONVENIO;", dbOpenDynaset)
            End If
            
            'Valida se existe boletos não liquidados
            If TbDados.EOF = False Then

                'Criar objeto Excel para abrir mascara
                Set ObjExcel = CreateObject("EXCEL.application")
                ObjExcel.Workbooks.Open FileName:=Caminho & "Boletos.xlsx", ReadOnly:=True
                Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
        
                    linha = 9: LinhaSitu = linha
                    ObjPlan1Excel.Range("D4") = ARQDATA
                    ObjPlan1Excel.Range("D4").NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
                    UltimaLinha = TbDados.RecordCount
                    UltimaLinha = UltimaLinha + 8
                    TbDados.Close
                    
                    'Formatando planilha
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":M" & UltimaLinha).Font.Size = 8
                    ObjPlan1Excel.Range("A" & linha & ":A" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("B" & linha & ":B" & UltimaLinha).NumberFormat = "00000"
                    ObjPlan1Excel.Range("E" & linha & ":G" & UltimaLinha).NumberFormat = "00000"
                    ObjPlan1Excel.Range("J" & linha & ":J" & UltimaLinha).Style = "Currency"
                    ObjPlan1Excel.Range("K" & linha & ":K" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("L" & linha & ":L" & UltimaLinha).NumberFormat = "00000"
                    ObjPlan1Excel.Columns("A:M").Select
                    ObjPlan1Excel.Columns.AutoFit
                    ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
                    ObjPlan1Excel.Range("C2").Select
                    
                    'Ajustando nome do relatorio
                    Nome = "Relatorio de Boletos Nao Liquidados - " & Format(ARQDATA, "ddmmyy")
                    Nome = Trata_NomeArquivo(Nome)

                    'Verifica se ja existe este relatorio na rede, se sim, deletar.
                    sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                    If (Dir(sFname) <> "") Then
                        Kill sFname
                    End If
                
                'Salvar relatorio na rede
                ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                ObjExcel.activeworkbook.Close SaveChanges:=False
                ObjExcel.Quit
                
                'Camimho completo do arquivo para envio
                File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"

                'Função para envio do Relatorio
                Call EmailLiquidacoes("SUCESSO", "BOLETOS NÃO LIQUIDADOS", File)
            Else
                'Função para envio de negativa
                Call EmailLiquidacoes("ERRO", "BOLETOS NÃO LIQUIDADOS", "")
            End If

End Sub
Sub Teimosinha()

    Dim Nome As String, TbDados As Recordset
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim DtInicio As Date, DtFim As Date, Ret As String
    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
        
        'Abrir base de dados
        Call AbrirBDBoletos
            
            'Verifica o dia da semana
            DiaSemana = UCase(WeekdayName(Weekday(Date), True))
            
            'limpar tabela temporaria de Debito em Conta
            BDB.Execute ("DELETE Tbl_Temp_DebConta.* FROM Tbl_Temp_DebConta;")
            
            'Pesquisar ultimo dia util
            ARQDATA = UltimoDiaUtil(): DataPesq = Format(ARQDATA, "mm/dd/yyyy")
            
            'Inserir na tabela temporaria os operações em Teimosinha
            If DiaSemana = "TER" Then
                DtFim = ARQDATA: DtFim = DtFim - 2: ARQDATA = DtFim: DataPesq = Format(ARQDATA, "mm/dd/yyyy")
                BDB.Execute ("INSERT INTO Tbl_Temp_DebConta ( AGENCIA, CONVENIO, NOME_CONVENIO, CNPJ_FORNECEDOR, NOME_FORNECEDOR, COD_OPERACAO, COMPROMISSO, DATA_VENC, VALOR_COMPROMISSO, CODOPER ) SELECT TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Fornecedor, TblArqoped.Nome_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, Tbl_Boletos.COD_OPERACAO" _
                & " FROM TblArqoped LEFT JOIN Tbl_Boletos ON (TblArqoped.Agencia_Ancora = Tbl_Boletos.AGENCIA) AND (TblArqoped.Convenio_Ancora = Tbl_Boletos.CONVENIO) AND (TblArqoped.Cnpj_Fornecedor = Tbl_Boletos.CNPJ_FORNECEDOR) AND (TblArqoped.Cod_Oper = Tbl_Boletos.COD_OPERACAO) AND (TblArqoped.Compromisso = Tbl_Boletos.COMPROMISSO) AND (TblArqoped.Valor_Nom = Tbl_Boletos.VALOR_COMPROMISSO) WHERE (((TblArqoped.Data_Venc)>=#" & DataPesq & "# And (TblArqoped.Data_Venc)<Date()) AND ((Tbl_Boletos.COD_OPERACAO) Is Null));")
            Else
                BDB.Execute ("INSERT INTO Tbl_Temp_DebConta ( AGENCIA, CONVENIO, NOME_CONVENIO, CNPJ_FORNECEDOR, NOME_FORNECEDOR, COD_OPERACAO, COMPROMISSO, DATA_VENC, VALOR_COMPROMISSO, CODOPER ) SELECT TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Fornecedor, TblArqoped.Nome_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, Tbl_Boletos.COD_OPERACAO" _
                & " FROM TblArqoped LEFT JOIN Tbl_Boletos ON (TblArqoped.Valor_Nom = Tbl_Boletos.VALOR_COMPROMISSO) AND (TblArqoped.Compromisso = Tbl_Boletos.COMPROMISSO) AND (TblArqoped.Cod_Oper = Tbl_Boletos.COD_OPERACAO) AND (TblArqoped.Cnpj_Fornecedor = Tbl_Boletos.CNPJ_FORNECEDOR) AND (TblArqoped.Convenio_Ancora = Tbl_Boletos.CONVENIO) AND (TblArqoped.Agencia_Ancora = Tbl_Boletos.AGENCIA) WHERE (((TblArqoped.Data_Venc)=#" & DataPesq & "#) AND ((Tbl_Boletos.COD_OPERACAO) Is Null));")
            End If
            
            'Pesquisa as operações em teimosinha para gerar o relatorio
            If DiaSemana = "TER" Then
                DtFim = ARQDATA: DtFim = DtFim - 2: ARQDATA = DtFim: DataPesq = Format(ARQDATA, "mm/dd/yyyy")
                Set TbDados = BDB.OpenRecordset("SELECT Tbl_Temp_DebConta.AGENCIA, Tbl_Temp_DebConta.CONVENIO, Tbl_Temp_DebConta.NOME_CONVENIO, Tbl_Temp_DebConta.CNPJ_FORNECEDOR, Tbl_Temp_DebConta.NOME_FORNECEDOR, Tbl_Temp_DebConta.COD_OPERACAO, Tbl_Temp_DebConta.COMPROMISSO, Tbl_Temp_DebConta.VALOR_COMPROMISSO, Tbl_Temp_DebConta.DATA_VENC, Tbl_Liquidados.COD_OPERACAO FROM Tbl_Temp_DebConta LEFT JOIN Tbl_Liquidados ON (Tbl_Temp_DebConta.AGENCIA = Tbl_Liquidados.AGENCIA) AND (Tbl_Temp_DebConta.CONVENIO = Tbl_Liquidados.CONVENIO) AND (Tbl_Temp_DebConta.CNPJ_FORNECEDOR = Tbl_Liquidados.CNPJ_FORNECEDOR) AND (Tbl_Temp_DebConta.COD_OPERACAO = Tbl_Liquidados.COD_OPERACAO) AND (Tbl_Temp_DebConta.COMPROMISSO = Tbl_Liquidados.COMPROMISSO) AND (Tbl_Temp_DebConta.VALOR_COMPROMISSO = Tbl_Liquidados.VALOR_COMPROMISSO) WHERE (((Tbl_Temp_DebConta.DATA_VENC)>=#" & DataPesq & "# And (Tbl_Temp_DebConta.DATA_VENC)<Date()) AND ((Tbl_Liquidados.COD_OPERACAO) Is Null));", dbOpenDynaset)
            Else
                Set TbDados = BDB.OpenRecordset("SELECT Tbl_Temp_DebConta.AGENCIA, Tbl_Temp_DebConta.CONVENIO, Tbl_Temp_DebConta.NOME_CONVENIO, Tbl_Temp_DebConta.CNPJ_FORNECEDOR, Tbl_Temp_DebConta.NOME_FORNECEDOR, Tbl_Temp_DebConta.COD_OPERACAO, Tbl_Temp_DebConta.COMPROMISSO, Tbl_Temp_DebConta.VALOR_COMPROMISSO, Tbl_Temp_DebConta.DATA_VENC, Tbl_Liquidados.COD_OPERACAO FROM Tbl_Temp_DebConta LEFT JOIN Tbl_Liquidados ON (Tbl_Temp_DebConta.VALOR_COMPROMISSO = Tbl_Liquidados.VALOR_COMPROMISSO) AND (Tbl_Temp_DebConta.COMPROMISSO = Tbl_Liquidados.COMPROMISSO) AND (Tbl_Temp_DebConta.COD_OPERACAO = Tbl_Liquidados.COD_OPERACAO) AND (Tbl_Temp_DebConta.CNPJ_FORNECEDOR = Tbl_Liquidados.CNPJ_FORNECEDOR) AND (Tbl_Temp_DebConta.CONVENIO = Tbl_Liquidados.CONVENIO) AND (Tbl_Temp_DebConta.AGENCIA = Tbl_Liquidados.AGENCIA) WHERE (((Tbl_Temp_DebConta.DATA_VENC)=#" & DataPesq & "#) AND ((Tbl_Liquidados.COD_OPERACAO) Is Null));", dbOpenDynaset)
            End If
            
            'Valida se existe operação em teimosinha
            If TbDados.EOF = False Then
                
                'Criar objeto excel para abrir a Mascara do relatorio
                Set ObjExcel = CreateObject("EXCEL.application")
                ObjExcel.Workbooks.Open FileName:=Caminho & "Teimosinha.xlsx", ReadOnly:=True
                Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")

                    linha = 9: LinhaSitu = linha
    
                    ObjPlan1Excel.Range("F4") = UltimoDiaUtil()
                    ObjPlan1Excel.Range("F4").NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
                    UltimaLinha = TbDados.RecordCount
                    UltimaLinha = UltimaLinha + 8
                    
                    'Formtatando Relatorio
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":I" & UltimaLinha).Font.Size = 8
                    ObjPlan1Excel.Range("I" & linha & ":I" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).NumberFormat = "00000"
                    ObjPlan1Excel.Range("F" & linha & ":G" & UltimaLinha).NumberFormat = "00000"
                    ObjPlan1Excel.Range("H" & linha & ":H" & UltimaLinha).Style = "Currency"
                    ObjPlan1Excel.Range("D" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
                    ObjPlan1Excel.Columns("A:I").Select
                    ObjPlan1Excel.Columns.AutoFit
                    ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
                    ObjPlan1Excel.Range("C2").Select

                    'Ajustar o nome do Relatorio
                    Nome = "Relatorio de Titulos em Teimosinha - " & Format(UltimoDiaUtil(), "ddmmyy")
                    Nome = Trata_NomeArquivo(Nome)
                    TbDados.Close
                    
                    'Verifica se ja existe este relatorio na rede, se sim, deletar.
                    sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                    If (Dir(sFname) <> "") Then
                        Kill sFname
                    End If
                
                    vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
                    'ObjPlan1Excel.SaveAs FileName:=sFname
                    ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
                    ObjExcel.activeworkbook.Close SaveChanges:=False
                    
                    ObjExcel.Quit
                    
                    Dim FSO As New FileSystemObject
                    FSO.MoveFile vCaminhoLocal, sFname
                
                'Salvar Relatorio na rede
                'ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                'ObjExcel.activeworkbook.Close SaveChanges:=False
                'ObjExcel.Quit
                
                'Camimho completo do arquivo para envio
                File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"

                'Função para enviar o relatorio
                Call EmailLiquidacoes("SUCESSO", "TITULOS EM TEIMOSINHA", File)
            Else
                'Função para enviar email de negativa
                Call EmailLiquidacoes("ERRO", "TITULOS EM TEIMOSINHA", "")
            End If

End Sub


