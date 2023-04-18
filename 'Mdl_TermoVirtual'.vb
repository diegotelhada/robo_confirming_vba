'Mdl_TermoVirtual'

Option Compare Database
Sub ContratoMae()
    
    'Função para importar as bases
        Call ImportarBases
        
End Sub
Function AtualizaFielDeposPrincial()
    
    'Abrir banco de dados principal
    Call AbrirBDRelatorios
    
    'Abrir banco de dados Termo Virtual
    Call AbrirDBTVirtual

    'Limpar Tabela
    BDREL.Execute ("DELETE TblFielDepositario.* FROM TblFielDepositario;")

    'Atualizar Tabela
    
    DBTVirtual.Execute ("INSERT INTO TblFielDepositario ( Banco, AGENCIA, Convenio, SubProduto, [Nome do convênio], [Dt cadastro], Situação, Bloqueio, Ambiente, Pre_Aprovado, Fiel_Depositario ) IN '\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb' SELECT TbL_Temp_Convenios.BCO, TbL_Temp_Convenios.AGENCIA, TbL_Temp_Convenios.NR_CONVENIO, TbL_Temp_Convenios.SUBPRODU, TbL_Temp_Convenios.NOME_CONVENIO, TbL_Temp_Convenios.DT_CAD, TbL_Temp_Convenios.STATUS_CONV, TbL_Temp_Convenios.TIP_BLOQUEIO, TbL_Temp_Convenios.AMBIENTE, TbL_Temp_Convenios.PRE_APROV, TbL_Temp_Convenios.FIEL_DPOS FROM TbL_Temp_Convenios;")
    
    'Atualiza tabela de convenios loquidados
    DBTVirtual.Execute ("INSERT INTO Tbl_Convenio_liquidacao ( Agencia, Convenio, Nome_Ancora, CNPJ_Ancora, FormaLiquidacao ) SELECT Convenios.[AG# ], Format([Convenios]![NR CONVENIO],'000000000000') AS Convenio, Convenios.[NOME CONVENIO                 ], Convenios.[CPF/CNPJ       ], Convenios.[FORMA LIQUIDACAO                  ] FROM Convenios;")
    
    'Atualiza o tipo de liquidacao
    DBTVirtual.Execute ("UPDATE Tbl_Convenio_liquidacao SET Tbl_Convenio_liquidacao.Tipo = IIf(Left([Tbl_Convenio_liquidacao]![FormaLiquidacao],5)='CONTA','TEIMOSINHA',IIf(Left([Tbl_Convenio_liquidacao]![FormaLiquidacao],6)='BOLETO','BOLETO','DESCONHECIDO'));")

End Function
Function PesqNomeAncora(Agencia, Convenio)
    
    Dim TbNome As Recordset
    
        Set TbNome = DBTVirtual.OpenRecordset("SELECT TblClientes.Convenio_Ancora, TblClientes.Agencia_Ancora, TblClientes.Nome_Ancora FROM TblClientes WHERE (((TblClientes.Convenio_Ancora)='" & Format(Convenio, "000000000000") & "') AND ((TblClientes.Agencia_Ancora)=" & Agencia & "));", dbOpenDynaset)
            
            If TbNome.EOF = False Then: PesqNomeAncora = ReplaceString(TbNome!Nome_Ancora)

End Function
Function EnviarArquivo(Caminho As String, Email As String, nomeFornecedor As String, NomeAncora As String)

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

        'Linha1 = "Prezado(a) Fornecedor,"
        'linha2 = "Segue anexo ": linha21 = "contrato": linha211 = " mãe para as operações de Confirming®. Favor prosseguir da seguinte forma:"
        'linha3 = "•  Pagina 2: ": linha31 = "preencher item 2.1.2 (pessoas autorizadas a fazer operações)"
        'linha4 = "•  Pagina 4: ": linha41 = " assinar no campo reservado para sua empresa (não é necessário preencher/assinar o campo de testemunha)"
        'linha5 = "•  Todas as páginas (inclusive os anexos): ": linha51 = " rubricar"
        
        'Imagem = "\\Saont46\apps2\Confirming\PROJETORELATORIOS\Logo.jpg"
        
        'Assinatura01 = "Após devidamente assinado/preenchido, favor "
        'Assinatura1 = "enviar todas as páginas do contrato para o endereço: VIA: CORREIO(A.R), Sedex ou Portador."
        
        'Assinatura12 = "BANCO SANTANDER S/A- CASA1"
        'Assinatura13 = "Depto CONFIRMING"
        
        'Assinatura2 = "Rua Amador Bueno, 474 - Santo Amaro"
        'Assinatura3 = "Cep: 04752-005  - São Paulo - SP"
        'Assinatura4 = "Bloco D / 3°Andar / Estação 394 ou 395"
        'Assinatura5 = "Aos cuidados de: Wellington, Cleverton, Renato ou Márcia."
        
        'Assinatura6 = "Central de Atendimento Santander"
        'Assinatura7 = "4004-2125 (Capitais e Regiões Metropolitanas)"
        'Assinatura8 = "0800 726 2125 (Demais Localidades)"
        'Assinatura9 = "Serviço de Apoio ao Consumidor - SAC: 0800 762 7777 (Atende também Deficientes Auditivos e de Fala)"
        'Assinatura10 = "Ouvidoria: 0800 762 0322 (Atende também Deficientes Auditivos e de Fala)"
        'Assinatura11 = "Acesse: www.santander.com.br"
        
        'With OutMail
        '    .SentOnBehalfOfName = "prpjcadconfirming@santander.com.br"
        '    .To = CStr(Email)
        '    '.To = "marcelohsouza@santander.com.br"
        '    .CC = "prpjcadconfirming@santander.com.br"
        '    .Subject = "CONTRATO MÃE CONFIRMING - " & UCase(NomeFornecedor) & " - " & UCase(NomeAncora)
        '    .Attachments.Add "\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\PDF\" & Caminho & ".PDF"
        '
        '    .HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Arial Size = 3 <BR>" & Linha1 & "<BR><BR>" & linha2 & "<B>" & linha21 & "</B>" & linha211 & "<BR><BR>" & _
        '    "<B>&emsp;&emsp;" & linha3 & "</B>" & linha31 & "<BR><B>&emsp;&emsp;" & linha4 & "</B>" & linha41 & "<BR><B>&emsp;&emsp;" & linha5 & "</B>" & linha51 & "<BR>" & _
        '    "<BR>" & Assinatura01 & "<B><U>" & Assinatura1 & "</U></B><BR><BR><B>" & Assinatura12 & "<BR>" & Assinatura13 & "<BR></B>" & Assinatura2 & "<BR>" & Assinatura3 & "<BR>" & Assinatura4 & "<BR>" & Assinatura5 & "<BR><BR><BR>" & _
        '    "<img src=\\Saont46\apps2\Confirming\PROJETORELATORIOS\Logo.jpg height=50 width=150></FONT><FONT COLOR = BLACK FACE=Arial Size = 2 <BR><BR><B>" & Assinatura6 & _
        '    "</B><BR><BR>" & Assinatura7 & "<BR><BR>" & Assinatura8 & "<BR><BR>" & Assinatura9 & "<BR><BR>" & Assinatura10 & "<BR><BR>" & Assinatura11 & "</BODY></HTML>"
        
        
        '************* novo código Douglas ********************
        
        Linha1 = "Prezado(a) Fornecedor,"
        linha2 = "Você está recebendo o ": linha21 = "Convênio para Antecipação de Recebíveis Confirming®": linha211 = ", abaixo os próximos passos para finalizarmos o processo de cadastro da sua empresa:"
        linha3 = "Retorno do convênio e documentos": ' linha31 = "preencher item 2.1.2 (pessoas autorizadas a fazer operações)"
        linha31 = "Clientes Correntistas Santander com cadastro atualizado:":
        
        linha4 = "•    Em posse do convênio, imprimir, rubricar todas as páginas, preencher os dados das pessoas autorizadas a realizar operações, assinar no campo reservado para sua empresa, levar ao cartório para reconhecimento das firmas e enviar por e-mail para ": linha412 = "prpjcadconfirming@santander.com.br": linha4123 = " com cópia para ": linha41234 = "santanderconfirming@interfile.com.br":
        linha5 = "Clientes Não Correntistas": ' linha51 = " rubricar"
        linha6 = "Além do convênio citado acima, será necessário também o envio dos seguintes documentos digitalizados:":
        linha7 = "•    Cartão de Assinaturas": linha71 = " (Anexo 2) - Deverá ser preenchido e assinado com firma reconhecida para cada um dos representantes legais que assinaram o convênio.":
        linha8 = "•    Contrato Social": linha81 = " devidamente registrado com última alteração ou":  linha82 = "Estatuto Social": linha83 = " e": linha84 = " Ata de Eleição": linha85 = " dos administradores.":
        linha9 = "•    Procuração": linha91 = " devidamente registrada, se for o caso, com poderes vigentes e com as firmas reconhecidas de quem outorgou as procurações.":
        
        Imagem = "\\Saont46\apps2\Confirming\PROJETORELATORIOS\Logo.jpg"
        
        'Assinatura01 = "Após devidamente assinado/preenchido, favor "
        'Assinatura1 = "enviar todas as páginas do contrato para o endereço: VIA: CORREIO(A.R), Sedex ou Portador."
        
        Assinatura12 = "BANCO SANTANDER S/A"
        Assinatura13 = "Depto CONFIRMING"
        
        'Assinatura2 = "Rua Amador Bueno, 474 - Santo Amaro"
        'Assinatura3 = "Cep: 04752-005  - São Paulo - SP"
        'Assinatura4 = "Bloco D / 3°Andar / Estação 394 ou 395"
        'Assinatura5 = "Aos cuidados de: Wellington, Cleverton, Renato ou Márcia."
        
        Assinatura6 = "Central de Atendimento Santander"
        Assinatura7 = "4004-2125 (Capitais e Regiões Metropolitanas)"
        Assinatura8 = "0800 726 2125 (Demais Localidades)"
        Assinatura9 = "Serviço de Apoio ao Consumidor - SAC: 0800 762 7777 (Atende também Deficientes Auditivos e de Fala)"
        Assinatura10 = "Ouvidoria: 0800 762 0322 (Atende também Deficientes Auditivos e de Fala)"
        Assinatura11 = "Acesse: www.santander.com.br"
        
        With OutMail
            .SentOnBehalfOfName = "prpjcadconfirming@santander.com.br"
            .To = CStr(Email)
            '.To = "dwmiranda@santander.com.br; geraldo.ghetti@santander.com.br; joao.pasquini@santander.com.br"
            .cc = "prpjcadconfirming@santander.com.br"
            .Subject = "Confirming Santander - " & UCase(NomeAncora) & " - " & UCase(nomeFornecedor)
            .Attachments.Add "\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\PDF\" & Caminho & ".PDF"
            .Attachments.Add "\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\Anexo2.pdf"
            .HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Arial Size = 3 <BR>" & Linha1 & "<BR><BR>" & linha2 & "<B>" & linha21 & "</B>" & linha211 & _
            "<BR><BR>" & "&emsp;&emsp;&emsp; <B>" & linha3 & "</B>" & _
            "<BR><BR>" & "&emsp;&emsp;&emsp;" & "<FONT COLOR = RED> <B>" & linha31 & "</B></font>" & "<BR>" & _
            "<BR>" & "&emsp;&emsp;&emsp;" & linha4 & "<Font Color = BLUE><B>" & linha412 & "</b></font>" & linha4123 & "<Font Color = BLUE><B>" & linha41234 & "</b></font>" & "<BR>" & _
            "<BR><BR>" & "&emsp;&emsp;&emsp; <Font Color = RED><B>" & linha5 & "</B></Font><BR>" & _
            "<BR>" & "&emsp;&emsp;&emsp;" & linha6 & _
            "<BR><BR>" & "&emsp;&emsp;&emsp;" & "<B>" & linha7 & "</B>" & linha7 & _
            "<BR>" & "&emsp;&emsp;&emsp;" & "<B>" & linha8 & "</B>" & linha81 & "<B>" & linha82 & "</b>" & linha83 & "<B>" & linha84 & "</b>" & linha85 & _
            "<BR>" & "&emsp;&emsp;&emsp;" & "<B>" & linha9 & "</B>" & linha91 & _
            "<BR><BR><BR><BR><img src=\\Saont46\apps2\Confirming\PROJETORELATORIOS\Logo.jpg height=50 width=150></FONT><FONT COLOR = BLACK FACE=Arial Size = 2 <BR><BR><B>" & Assinatura6 & _
            "</B><BR><BR>" & Assinatura7 & "<BR><BR>" & Assinatura8 & "<BR><BR>" & Assinatura9 & "<BR><BR>" & Assinatura10 & "<BR><BR>" & Assinatura11 & "</BODY></HTML>"
        
        '************* novo código Douglas ********************
        
        '.Display
        End With
   OutMail.Send
End Function
Function Converter2PDF(Caminho As String)

    Dim wd As Word.Application, wdocSource As Word.Document, Sel As Word.Selection
    Dim FSO As New FileSystemObject

        Set wd = CreateObject("Word.Application")
        Set wdocSource = wd.Documents.Open("\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\WORD\" & Caminho & ".docx")
        Set Sel = wd.Selection
                
            wdocSource.ExportAsFixedFormat OutputFileName:= _
            "\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\PDF\" & Caminho & ".pdf", ExportFormat:= _
            wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
            wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
            Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
            CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
            BitmapMissingFonts:=True, UseISO19005_1:=False
            
       wdocSource.Close SaveChanges:=False
       wd.Quit

      Set wdocSource = Nothing
      Set wd = Nothing
      
      'Copiar Arquivo PDF Gerado para a pasta da mesa
      FSO.CopyFile "\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\PDF\" & Caminho & ".pdf", "\\Saont46\apps2\Confirming\CONTRATOMAE_ENVIADO\" & Format(UltimoDiaUtil(), "DDMMYYYY") & "\", True
      
End Function
Function Salvar_Arquivo_WORD(nomeFornecedor As String, Convenio, DataInicio, Tipo, CNPJFornecedor, Grupo)
    
    Dim wd As Word.Application, wdocSource As Word.Document, Sel As Word.Selection, Arquivo As String
    Dim CaminhoPadrao As String: CaminhoPadrao = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\MINUTAS\10042017\"
                                          
        If Tipo = "8510" Then: ArquivoBase = "MINUTA_CONFIRMING_PADRAO.docx": Query = "Qry_Minuta_Padrao"
        If Tipo = "8520" Then: ArquivoBase = "MINUTA_CONFIRMING_PRENEGOCIADO.docx": Query = "Qry_Minuta_Pre_Negociado"
        If Tipo = "8530" Then: ArquivoBase = "MINUTA_CONFIRMING_DIRETO.docx": Query = "Qry_Minuta_Direto"
                
        'Criar objeto WORD
        Set wd = CreateObject("Word.Application")
        Set wdocSource = wd.Documents.Open(CaminhoPadrao & ArquivoBase)
        Set Sel = wd.Selection
          wdocSource.MailMerge.MainDocumentType = wdFormLetters
      
              wdocSource.MailMerge.OpenDataSource _
                  Name:="\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\BD_TERMOVIRTUAL.accdb", _
                  LinkToSource:=True, AddToRecentFiles:=False, _
                  Connection:="QUERY qryLabelQuery", _
                  SQLStatement:="SELECT * FROM [" & Query & "]"
          
              With wdocSource.MailMerge
                  .Destination = wdSendToNewDocument
                  .SuppressBlankLines = True
                  With .DataSource
                      .FirstRecord = 1
                      .LastRecord = 1
                  End With
                  .Execute Pause:=False
              End With
                                          
            Arquivo = Replace(Trata_NomeArquivo(nomeFornecedor), "/", "") & "_" & Format(Convenio, "000000000000") & "_" & Format(CNPJFornecedor, "00000000000000") & "_" & Format(DataInicio, "ddmmyyyy")
    
       wd.ActiveDocument.SaveAs ("\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\WORD\" & Arquivo & ".docx")
       wdocSource.Close SaveChanges:=False
       wd.Quit

      Set wdocSource = Nothing
      Set wd = Nothing
      
    'Incluir ultima pagina para convenios agrupados
    If Grupo <> "Não" Then: Call IncluirUltimaPaginaConvenios("\\Saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\MINUTAS\WORD\" & Arquivo & ".docx", Grupo, Tipo)
    
    'Converter WORD para PDF
    Call Converter2PDF(Arquivo)
    
    'Retorna o nome do arquivo para enviar por email
    Salvar_Arquivo_WORD = Arquivo
    
End Function
Function IncluirUltimaPaginaConvenios(Caminho, Grupo, Tipo)
    
    Dim TbConvenios As Recordset
    Dim wd As Word.Application
    Dim wdocSource As Word.Document
        
        Set wd = CreateObject("Word.Application")
        Set wdocSource = wd.Documents.Open(Caminho)
        Set Sel = wd.Selection
        'wd.Visible = True
                    
        If Tipo = "8510" Then: IndexTable = 7
        If Tipo = "8530" Then: IndexTable = 8
                    
            With wdocSource.Tables(IndexTable)
            
              With .Rows.First
                  .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                  .Cells(1).Range.Text = "RAZÃO SOCIAL DO DEVEDOR"
                  .Cells(1).Range.Bold = True
    
                  .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                  .Cells(2).Range.Text = "N° CONVENIO CONFIRMING"
                  .Cells(2).Range.Bold = True
              End With
                            
            Set TbConvenios = DBTVirtual.OpenRecordset("SELECT Format([Tbl_Convenio_Agrupados]![Banco],'0000') & Format([Tbl_Convenio_Agrupados]![Agencia],'0000') & Format([Tbl_Convenio_Agrupados]![Convenio],'000000000000') AS NumConvenio, Tbl_Convenio_Agrupados.Nome_Convenio, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados WHERE (((Tbl_Convenio_Agrupados.Grupo)='" & Grupo & "'));", dbOpenDynaset)
                
                Do While TbConvenios.EOF = False
                    
                    With .Rows.Last
                    
                      .Cells(1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
                      .Cells(1).Range.Text = TbConvenios!nome_convenio
        
                      .Cells(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                      .Cells(2).Range.Text = TbConvenios!NumConvenio
                      
                    End With
                    .Rows.Add
                    TbConvenios.MoveNext
                Loop
              With .Rows.Last
                .Delete
              End With
            End With
            
     wd.ActiveDocument.SaveAs Caminho
     wdocSource.Close SaveChanges:=False
     wd.Quit
     
    Set wd = Nothing
    Set wdocSource = Nothing
    
End Function
Sub GerarWORD_Fornecedores()
'
'    Dim TbDados As Recordset, TbNestle As Recordset: Call AbrirDBTVirtual
'    Dim NOMEARQUIVO As String, TbWhite As Recordset
'    Dim TbBrasilit As Recordset, TbUnilever As Recordset
'
'        'Selecionar os fornecedores agrupados da Nestlé
'        Set TbNestle = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, First(Tbl_Temp.CONVENIO_ANCORA) AS PrimeiroDeCONVENIO_ANCORA, First(Tbl_Temp.CNPJ_ANCORA) AS PrimeiroDeCNPJ_ANCORA, First(Tbl_Temp.NOME_ANCORA) AS PrimeiroDeNOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) AND (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio)" _
'        & " WHERE (((Tbl_Temp.CNPJ_RAIZ) Not In (SELECT Tbl_Enviados.CNPJ_RAIZ FROM Tbl_Convenio_Agrupados INNER JOIN Tbl_Enviados ON (Tbl_Convenio_Agrupados.Agencia = Tbl_Enviados.AGENCIA_ANCORA) AND (Tbl_Convenio_Agrupados.Convenio = Tbl_Enviados.CONVENIO_ANCORA) WHERE (((Tbl_Enviados.DATA_INICI)>#5/10/2017#) AND ((Tbl_Convenio_Agrupados.Grupo)='NESTLE')) GROUP BY Tbl_Enviados.CNPJ_RAIZ;)))" _
'        & " GROUP BY Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo HAVING (((Tbl_Convenio_Agrupados.Grupo)='NESTLE'));", dbOpenDynaset)
'
'            Do While TbNestle.EOF = False
'
'                'Deletar tabela temporaria
'                DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                    'Incluir Informações na tabela temporaria
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC )" _
'                    & " SELECT '" & TbNestle!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbNestle!Agencia_Ancora & "' AS AGENCIA_ANCORA, '" & TbNestle!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbNestle!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbNestle!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbNestle!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & TbNestle!CNPJ & "' AS CNPJ, '" & TbNestle!ENDERECO & "' AS ENDERECO, '" & TbNestle!NUMERO & "' AS NUMERO, '" & TbNestle!BAIRRO & "' AS BAIRRO, '" & TbNestle!CIDADE & "' AS CIDADE, '" & TbNestle!UF & "' AS UF, '" & TbNestle!Email & "' AS EMAIL, '" & TbNestle!DATA_INICI & "' AS DATA_INICI, '" & TbNestle!banco & "' AS BANCO, '" & TbNestle!Agencia & "' AS AGENCIA, '" & TbNestle!Conta & "' AS CONTA, '" & TbNestle!TTR & "' AS TTR, '" & TbNestle!TCO & "' AS TCO, '" & TbNestle!DPC & "' AS DPC;")
'
'                        'Chamar função para gerar arquivo Word
'                        NOMEARQUIVO = Salvar_Arquivo_WORD(TbNestle!NOME_FORNECEDOR, TbNestle!PrimeiroDeCONVENIO_ANCORA, TbNestle!DATA_INICI, TbNestle!TIPO_CONVENIO, TbNestle!Grupo, TbNestle!CNPJ)
'
'                        'Chamar Funcção para enviar o PDF pelo Email
'                        Call EnviarArquivo(NOMEARQUIVO, Trim(TbNestle!Email), TbNestle!NOME_FORNECEDOR, TbNestle!PrimeiroDeNOME_ANCORA)
'
'                    'Incluir dados do arquivo enviado na Tbl_Enviado
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO )" _
'                    & " SELECT '" & TbNestle!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbNestle!Agencia_Ancora & "' AS AGENCIA_ANCORA, '" & TbNestle!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbNestle!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbNestle!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbNestle!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & Format(TbNestle!CNPJ, "00000000000000") & "' AS CNPJ, '" & Left(Format(TbNestle!CNPJ, "00000000000000"), 8) & "' AS CNPJRAIZ, '" & TbNestle!ENDERECO & "' AS ENDERECO, '" & TbNestle!NUMERO & "' AS NUMERO, '" & TbNestle!BAIRRO & "' AS BAIRRO, '" & TbNestle!CIDADE & "' AS CIDADE, '" & TbNestle!UF & "' AS UF, '" & TbNestle!Email & "' AS EMAIL, '" & TbNestle!DATA_INICI & "' AS DATA_INICI, '" & TbNestle!banco & "' AS BANCO, '" & TbNestle!Agencia & "' AS AGENCIA, '" & TbNestle!Conta & "' AS CONTA," _
'                    & " '" & TbNestle!TTR & "' AS TTR, '" & TbNestle!TCO & "' AS TCO, '" & TbNestle!DPC & "' AS DPC, '" & TbNestle!TIPO_CONVENIO & "' AS TIPO_CONVENIO, #" & Format(Date, "mm/dd/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO;")
'
'                    'Atualizar a consulta para
'                    TbNestle.Requery
'
'                    'Verifica se com a atualização a consulta ficou em branco
'                    If TbNestle.EOF = True Then: Exit Do
'
'                TbNestle.MoveNext
'            Loop
'
'    '===============================================================================================================================================================
'
'        Set TbWhite = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, First(Tbl_Temp.CONVENIO_ANCORA) AS PrimeiroDeCONVENIO_ANCORA, First(Tbl_Temp.CNPJ_ANCORA) AS PrimeiroDeCNPJ_ANCORA, First(Tbl_Temp.NOME_ANCORA) AS PrimeiroDeNOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo" _
'        & " FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) AND (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) WHERE (((Tbl_Temp.CNPJ_RAIZ) Not In (SELECT Tbl_Enviados.CNPJ_RAIZ" _
'        & " FROM Tbl_Convenio_Agrupados INNER JOIN Tbl_Enviados ON (Tbl_Convenio_Agrupados.Agencia = Tbl_Enviados.AGENCIA_ANCORA) AND (Tbl_Convenio_Agrupados.Convenio = Tbl_Enviados.CONVENIO_ANCORA) WHERE (((Tbl_Enviados.DATA_INICI)>#5/10/2017#) AND ((Tbl_Convenio_Agrupados.Grupo)='WHITE')) GROUP BY Tbl_Enviados.CNPJ_RAIZ;)))" _
'        & " GROUP BY Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo HAVING (((Tbl_Convenio_Agrupados.Grupo)='WHITE'));", dbOpenDynaset)
'
'            Do While TbWhite.EOF = False
'
'                'Deletar tabela temporaria
'                DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                    'Incluir Informações na tabela temporaria
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC )" _
'                    & " SELECT '" & TbWhite!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbWhite!Agencia_Ancora & "' AS AGENCIA_ANCORA, '" & TbWhite!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbWhite!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbWhite!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbWhite!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & TbWhite!CNPJ & "' AS CNPJ, '" & TbWhite!ENDERECO & "' AS ENDERECO, '" & TbWhite!NUMERO & "' AS NUMERO, '" & TbWhite!BAIRRO & "' AS BAIRRO, '" & TbWhite!CIDADE & "' AS CIDADE, '" & TbWhite!UF & "' AS UF, '" & TbWhite!Email & "' AS EMAIL, '" & TbWhite!DATA_INICI & "' AS DATA_INICI, '" & TbWhite!banco & "' AS BANCO, '" & TbWhite!Agencia & "' AS AGENCIA, '" & TbWhite!Conta & "' AS CONTA, '" & TbWhite!TTR & "' AS TTR, '" & TbWhite!TCO & "' AS TCO, '" & TbWhite!DPC & "' AS DPC;")
'
'                        'Chamar função para gerar arquivo Word
'                        NOMEARQUIVO = Salvar_Arquivo_WORD(TbWhite!NOME_FORNECEDOR, TbWhite!PrimeiroDeCONVENIO_ANCORA, TbWhite!DATA_INICI, TbWhite!Grupo, TbWhite!CNPJ)
'
'                        'Chamar Funcção para enviar o PDF pelo Email
'                        Call EnviarArquivo(NOMEARQUIVO, Trim(TbWhite!Email), TbWhite!NOME_FORNECEDOR, TbWhite!PrimeiroDeNOME_ANCORA)
'
'                    'Incluir dados do arquivo enviado na Tbl_Enviado
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO )" _
'                    & " SELECT '" & TbWhite!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbWhite!Agencia_Ancora & "' AS AGENCIA_ANCORA, '" & TbWhite!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbWhite!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbWhite!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbWhite!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & Format(TbWhite!CNPJ, "00000000000000") & "' AS CNPJ, '" & Left(Format(TbWhite!CNPJ, "00000000000000"), 8) & "' AS CNPJRAIZ, '" & TbWhite!ENDERECO & "' AS ENDERECO, '" & TbWhite!NUMERO & "' AS NUMERO, '" & TbWhite!BAIRRO & "' AS BAIRRO, '" & TbWhite!CIDADE & "' AS CIDADE, '" & TbWhite!UF & "' AS UF, '" & TbWhite!Email & "' AS EMAIL, '" & TbWhite!DATA_INICI & "' AS DATA_INICI, '" & TbWhite!banco & "' AS BANCO, '" & TbWhite!Agencia & "' AS AGENCIA, '" & TbWhite!Conta & "' AS CONTA," _
'                    & " '" & TbWhite!TTR & "' AS TTR, '" & TbWhite!TCO & "' AS TCO, '" & TbWhite!DPC & "' AS DPC, '" & TbWhite!TIPO_CONVENIO & "' AS TIPO_CONVENIO, #" & Format(Date, "mm/dd/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO;")
'
'                    'Atualizar a consulta para
'                    TbWhite.Requery
'
'                    'Verifica se com a atualização a consulta ficou em branco
'                    If TbWhite.EOF = True Then: Exit Do
'
'                TbWhite.MoveNext
'            Loop
'    '===============================================================================================================================================================
'
'        Set TbBrasilit = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.BANCO_ANCORA, First(Tbl_Temp.AGENCIA_ANCORA) AS PrimeiroDeAGENCIA_ANCORA, First(Tbl_Temp.CONVENIO_ANCORA) AS PrimeiroDeCONVENIO_ANCORA, First(Tbl_Temp.CNPJ_ANCORA) AS PrimeiroDeCNPJ_ANCORA, First(Tbl_Temp.NOME_ANCORA) AS PrimeiroDeNOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia)" _
'        & " WHERE (((Tbl_Temp.CNPJ_RAIZ) Not In (SELECT Tbl_Enviados.CNPJ_RAIZ FROM Tbl_Convenio_Agrupados INNER JOIN Tbl_Enviados ON (Tbl_Convenio_Agrupados.Agencia = Tbl_Enviados.AGENCIA_ANCORA) AND (Tbl_Convenio_Agrupados.Convenio = Tbl_Enviados.CONVENIO_ANCORA) WHERE (((Tbl_Enviados.DATA_INICI)>#5/10/2017#) AND ((Tbl_Convenio_Agrupados.Grupo)='BRASILIT')) GROUP BY Tbl_Enviados.CNPJ_RAIZ;)))" _
'        & " GROUP BY Tbl_Temp.BANCO_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo HAVING (((Tbl_Convenio_Agrupados.Grupo)='BRASILIT'));", dbOpenDynaset)
'
'            Do While TbBrasilit.EOF = False
'
'                'Deletar tabela temporaria
'                DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                    'Incluir Informações na tabela temporaria
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC )" _
'                    & " SELECT '" & TbBrasilit!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbBrasilit!PrimeiroDeAGENCIA_ANCORA & "' AS AGENCIA_ANCORA, '" & TbBrasilit!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbBrasilit!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbBrasilit!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbBrasilit!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & TbBrasilit!CNPJ & "' AS CNPJ, '" & TbBrasilit!ENDERECO & "' AS ENDERECO, '" & TbBrasilit!NUMERO & "' AS NUMERO, '" & TbBrasilit!BAIRRO & "' AS BAIRRO, '" & TbBrasilit!CIDADE & "' AS CIDADE, '" & TbBrasilit!UF & "' AS UF, '" & TbBrasilit!Email & "' AS EMAIL, '" & TbBrasilit!DATA_INICI & "' AS DATA_INICI, '" & TbBrasilit!banco & "' AS BANCO, '" & TbBrasilit!Agencia & "' AS AGENCIA, '" & TbBrasilit!Conta & "' AS CONTA, '" & TbBrasilit!TTR & "' AS TTR, '" & TbBrasilit!TCO & "' AS TCO, '" & TbBrasilit!DPC & "' AS DPC;")
'
'                        'Chamar função para gerar arquivo Word
'                        NOMEARQUIVO = Salvar_Arquivo_WORD(TbBrasilit!NOME_FORNECEDOR, TbBrasilit!PrimeiroDeCONVENIO_ANCORA, TbBrasilit!DATA_INICI, TbBrasilit!Grupo, TbBrasilit!CNPJ)
'
'                        'Chamar Funcção para enviar o PDF pelo Email
'                        Call EnviarArquivo(NOMEARQUIVO, Trim(TbBrasilit!Email), TbBrasilit!NOME_FORNECEDOR, TbBrasilit!PrimeiroDeNOME_ANCORA)
'
'                    'Incluir dados do arquivo enviado na Tbl_Enviado
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO )" _
'                    & " SELECT '" & TbBrasilit!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbBrasilit!PrimeiroDeAGENCIA_ANCORA & "' AS AGENCIA_ANCORA, '" & TbBrasilit!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbBrasilit!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbBrasilit!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbBrasilit!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & Format(TbBrasilit!CNPJ, "00000000000000") & "' AS CNPJ, '" & Left(Format(TbBrasilit!CNPJ, "00000000000000"), 8) & "' AS CNPJRAIZ, '" & TbBrasilit!ENDERECO & "' AS ENDERECO, '" & TbBrasilit!NUMERO & "' AS NUMERO, '" & TbBrasilit!BAIRRO & "' AS BAIRRO, '" & TbBrasilit!CIDADE & "' AS CIDADE, '" & TbBrasilit!UF & "' AS UF, '" & TbBrasilit!Email & "' AS EMAIL, '" & TbBrasilit!DATA_INICI & "' AS DATA_INICI, '" & TbBrasilit!banco & "' AS BANCO, '" & TbBrasilit!Agencia & "' AS AGENCIA, '" & TbBrasilit!Conta & "' AS CONTA," _
'                    & " '" & TbBrasilit!TTR & "' AS TTR, '" & TbBrasilit!TCO & "' AS TCO, '" & TbBrasilit!DPC & "' AS DPC, '" & TbBrasilit!TIPO_CONVENIO & "' AS TIPO_CONVENIO, #" & Format(Date, "mm/dd/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO;")
'
'                    'Atualizar a consulta para
'                    TbBrasilit.Requery
'
'                    'Verifica se com a atualização a consulta ficou em branco
'                    If TbBrasilit.EOF = True Then: Exit Do
'
'                TbBrasilit.MoveNext
'            Loop
'
'    '===============================================================================================================================================================
'        Set TbUnilever = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.BANCO_ANCORA, First(Tbl_Temp.AGENCIA_ANCORA) AS PrimeiroDeAGENCIA_ANCORA, First(Tbl_Temp.CONVENIO_ANCORA) AS PrimeiroDeCONVENIO_ANCORA, First(Tbl_Temp.CNPJ_ANCORA) AS PrimeiroDeCNPJ_ANCORA, First(Tbl_Temp.NOME_ANCORA) AS PrimeiroDeNOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) AND (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio)" _
'        & " WHERE (((Tbl_Temp.CNPJ_RAIZ) Not In (SELECT Tbl_Enviados.CNPJ_RAIZ FROM Tbl_Convenio_Agrupados INNER JOIN Tbl_Enviados ON (Tbl_Convenio_Agrupados.Agencia = Tbl_Enviados.AGENCIA_ANCORA) AND (Tbl_Convenio_Agrupados.Convenio = Tbl_Enviados.CONVENIO_ANCORA) WHERE (((Tbl_Enviados.DATA_INICI)>#5/10/2017#) AND ((Tbl_Convenio_Agrupados.Grupo)='UNILEVER')) GROUP BY Tbl_Enviados.CNPJ_RAIZ;)))" _
'        & " GROUP BY Tbl_Temp.BANCO_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo HAVING (((Tbl_Convenio_Agrupados.Grupo)='UNILEVER'));", dbOpenDynaset)
'
'            Do While TbUnilever.EOF = False
'
'                'Deletar tabela temporaria
'                DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                    'Incluir Informações na tabela temporaria
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC )" _
'                    & " SELECT '" & TbUnilever!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbUnilever!PrimeiroDeAGENCIA_ANCORA & "' AS AGENCIA_ANCORA, '" & TbUnilever!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbUnilever!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbUnilever!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbUnilever!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & TbUnilever!CNPJ & "' AS CNPJ, '" & TbUnilever!ENDERECO & "' AS ENDERECO, '" & TbUnilever!NUMERO & "' AS NUMERO, '" & TbUnilever!BAIRRO & "' AS BAIRRO, '" & TbUnilever!CIDADE & "' AS CIDADE, '" & TbUnilever!UF & "' AS UF, '" & TbUnilever!Email & "' AS EMAIL, '" & TbUnilever!DATA_INICI & "' AS DATA_INICI, '" & TbUnilever!banco & "' AS BANCO, '" & TbUnilever!Agencia & "' AS AGENCIA, '" & TbUnilever!Conta & "' AS CONTA, '" & TbUnilever!TTR & "' AS TTR, '" & TbUnilever!TCO & "' AS TCO, '" & TbUnilever!DPC & "' AS DPC;")
'
'                        'Chamar função para gerar arquivo Word
'                        NOMEARQUIVO = Salvar_Arquivo_WORD(TbUnilever!NOME_FORNECEDOR, TbUnilever!PrimeiroDeCONVENIO_ANCORA, TbUnilever!DATA_INICI, TbUnilever!Grupo, TbUnilever!CNPJ)
'
'                        'Chamar Funcção para enviar o PDF pelo Email
'                        Call EnviarArquivo(NOMEARQUIVO, Trim(TbUnilever!Email), TbUnilever!NOME_FORNECEDOR, TbUnilever!PrimeiroDeNOME_ANCORA)
'
'                    'Incluir dados do arquivo enviado na Tbl_Enviado
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO )" _
'                    & " SELECT '" & TbUnilever!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbUnilever!PrimeiroDeAGENCIA_ANCORA & "' AS AGENCIA_ANCORA, '" & TbUnilever!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbUnilever!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbUnilever!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbUnilever!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & Format(TbUnilever!CNPJ, "00000000000000") & "' AS CNPJ, '" & Left(Format(TbUnilever!CNPJ, "00000000000000"), 8) & "' AS CNPJRAIZ, '" & TbUnilever!ENDERECO & "' AS ENDERECO, '" & TbUnilever!NUMERO & "' AS NUMERO, '" & TbUnilever!BAIRRO & "' AS BAIRRO, '" & TbUnilever!CIDADE & "' AS CIDADE, '" & TbUnilever!UF & "' AS UF, '" & TbUnilever!Email & "' AS EMAIL, '" & TbUnilever!DATA_INICI & "' AS DATA_INICI, '" & TbUnilever!banco & "' AS BANCO, '" & TbUnilever!Agencia & "' AS AGENCIA, '" & TbUnilever!Conta & "' AS CONTA," _
'                    & " '" & TbUnilever!TTR & "' AS TTR, '" & TbUnilever!TCO & "' AS TCO, '" & TbUnilever!DPC & "' AS DPC, '" & TbUnilever!TIPO_CONVENIO & "' AS TIPO_CONVENIO, #" & Format(Date, "mm/dd/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO;")
'
'                    'Atualizar a consulta para
'                    TbUnilever.Requery
'
'                    'Verifica se com a atualização a consulta ficou em branco
'                    If TbUnilever.EOF = True Then: Exit Do
'
'                TbUnilever.MoveNext
'            Loop
'
'    '===============================================================================================================================================================
'        'Deletar os demais convenios da Nestlé
'        DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) WHERE (((Tbl_Convenio_Agrupados.Grupo)='NESTLE')) GROUP BY Tbl_Temp.ID)));")
'
'        'Deletar os demais convenios WHITE
'        DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) WHERE (((Tbl_Convenio_Agrupados.Grupo)='WHITE')) GROUP BY Tbl_Temp.ID)));")
'
'        'Deletar os demais convenios Brasilit
'        DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) WHERE (((Tbl_Convenio_Agrupados.Grupo)='BRASILIT')) GROUP BY Tbl_Temp.ID)));")
'
'        'Deletar os demais convenios Unilever
'        DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) WHERE (((Tbl_Convenio_Agrupados.Grupo)='UNILEVER')) GROUP BY Tbl_Temp.ID)));")
'
'        'Deletar os convenios da Raizen
'        DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Raizen ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Raizen.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Raizen.Agencia) GROUP BY Tbl_Temp.ID)));")
'
'    '===============================================================================================================================================================
'
'        'Selecionar todos os fornecedores cadastrados
'        Set TbDados = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.ID, Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.CNPJ_RAIZ, Tbl_Temp.ENDERECO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Format([Tbl_Temp]![AGENCIA_ANCORA],'0000') & Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Temp]![CNPJ_RAIZ],'00000000') AS CHAVE FROM Tbl_Temp" _
'        & " GROUP BY Tbl_Temp.ID, Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.CNPJ_RAIZ, Tbl_Temp.ENDERECO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Format([Tbl_Temp]![AGENCIA_ANCORA],'0000') & Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Temp]![CNPJ_RAIZ],'00000000')" _
'        & " HAVING (((Format([Tbl_Temp]![AGENCIA_ANCORA],'0000') & Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Temp]![CNPJ_RAIZ],'00000000')) Not In (SELECT Format([Tbl_Enviados]![AGENCIA_ANCORA],'0000') & Format([Tbl_Enviados]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Enviados]![CNPJ_RAIZ],'00000000') AS CHAVE FROM Tbl_Enviados WHERE (((Tbl_Enviados.DATA_INICI) > #5/10/2017#)) GROUP BY Format([Tbl_Enviados]![AGENCIA_ANCORA],'0000') & Format([Tbl_Enviados]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Enviados]![CNPJ_RAIZ],'00000000');)));", dbOpenDynaset)
'
'            Do While TbDados.EOF = False
'
'                'Deletar tabela temporaria
'                DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                    'Incluir Informações na tabela temporaria
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, NOME_ANCORA, CNPJ_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC ) SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.CNPJ_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC FROM Tbl_Temp WHERE (((Tbl_Temp.ID)=" & TbDados!ID & "));")
'
'                        'Chamar função para gerar arquivo Word
'                        NOMEARQUIVO = Salvar_Arquivo_WORD(TbDados!NOME_FORNECEDOR, TbDados!Convenio_Ancora, TbDados!DATA_INICI, TbDados!TIPO_CONVENIO, TbDados!CNPJ)
'
'                        'Chamar Funcção para enviar o PDF pelo Email
'                        Call EnviarArquivo(NOMEARQUIVO, Trim(TbDados!Email), TbDados!NOME_FORNECEDOR, TbDados!Nome_Ancora)
'
'                    'Incluir dados do arquivo enviado na Tbl_Enviado
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO ) SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.CNPJ_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Format([Tbl_Temp]![CNPJ],'00000000000000') AS CNPJ, Left(Format([Tbl_Temp]![CNPJ],'00000000000000'),8) AS CNPJRAIZ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, #" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO FROM Tbl_Temp WHERE (((Tbl_Temp.ID)=" & TbDados!ID & "));")
'
'                    'Atualizar a consulta para
'                    TbDados.Requery
'
'                    'Verifica se com a atualização a consulta ficou em branco
'                    If TbDados.EOF = True Then: Exit Do
'
'                TbDados.MoveNext
'            Loop
End Sub
Sub GerarWORD_Fornecedores_V2()
'
'    Dim TbDados As Recordset, TbNestle As Recordset: Call AbrirDBTVirtual
'    Dim NOMEARQUIVO As String, TbWhite As Recordset
'    Dim TbBrasilit As Recordset, TbUnilever As Recordset
'    Dim TbGrupos As Recordset
'
'        Set TbGrupos = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados GROUP BY Tbl_Convenio_Agrupados.Grupo ORDER BY Tbl_Convenio_Agrupados.Grupo;", dbOpenDynaset)
'
'            Do While TbGrupos.EOF = False
'
'                Debug.Print TbGrupos!Grupo
'
'                Set TbDados = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, First(Tbl_Temp.CONVENIO_ANCORA) AS PrimeiroDeCONVENIO_ANCORA, First(Tbl_Temp.CNPJ_ANCORA) AS PrimeiroDeCNPJ_ANCORA, First(Tbl_Temp.NOME_ANCORA) AS PrimeiroDeNOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) AND (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio)" _
'                & " WHERE (((Tbl_Temp.CNPJ_RAIZ) Not In (SELECT Tbl_Enviados.CNPJ_RAIZ FROM Tbl_Convenio_Agrupados INNER JOIN Tbl_Enviados ON (Tbl_Convenio_Agrupados.Agencia = Tbl_Enviados.AGENCIA_ANCORA) AND (Tbl_Convenio_Agrupados.Convenio = Tbl_Enviados.CONVENIO_ANCORA) WHERE (((Tbl_Enviados.DATA_INICI)>#5/10/2017#) AND ((Tbl_Convenio_Agrupados.Grupo)='" & TbGrupos!Grupo & "')) GROUP BY Tbl_Enviados.CNPJ_RAIZ;)))" _
'                & " GROUP BY Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Tbl_Convenio_Agrupados.Grupo HAVING (((Tbl_Convenio_Agrupados.Grupo)='" & TbGrupos!Grupo & "'));", dbOpenDynaset)
'
'                    Do While TbDados.EOF = False
'
'                        'Deletar tabela temporaria
'                        DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                        'Incluir Informações na tabela temporaria
'                        DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC )" _
'                        & " SELECT '" & TbDados!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbDados!Agencia_Ancora & "' AS AGENCIA_ANCORA, '" & TbDados!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbDados!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbDados!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbDados!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & TbDados!CNPJ & "' AS CNPJ, '" & TbDados!ENDERECO & "' AS ENDERECO, '" & TbDados!NUMERO & "' AS NUMERO, '" & TbDados!BAIRRO & "' AS BAIRRO, '" & TbDados!CIDADE & "' AS CIDADE, '" & TbDados!UF & "' AS UF, '" & TbDados!Email & "' AS EMAIL, '" & TbDados!DATA_INICI & "' AS DATA_INICI, '" & TbDados!banco & "' AS BANCO, '" & TbDados!Agencia & "' AS AGENCIA, '" & TbDados!Conta & "' AS CONTA, '" & TbDados!TTR & "' AS TTR, '" & TbDados!TCO & "' AS TCO, '" & TbDados!DPC & "' AS DPC;")
'
'                            'Chamar função para gerar arquivo Word
'                            NOMEARQUIVO = Salvar_Arquivo_WORD(TbDados!NOME_FORNECEDOR, TbDados!PrimeiroDeCONVENIO_ANCORA, TbDados!DATA_INICI, TbDados!TIPO_CONVENIO, TbDados!CNPJ, TbDados!Grupo)
'
'                            'Chamar Funcção para enviar o PDF pelo Email
'                            Call EnviarArquivo(NOMEARQUIVO, Trim(TbDados!Email), Trim(TbDados!NOME_FORNECEDOR), TbDados!PrimeiroDeNOME_ANCORA)
'
'                            'Incluir dados do arquivo enviado na Tbl_Enviado
'                            DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO )" _
'                            & " SELECT '" & TbDados!BANCO_ANCORA & "' AS BANCO_ANCORA, '" & TbDados!Agencia_Ancora & "' AS AGENCIA_ANCORA, '" & TbDados!PrimeiroDeCONVENIO_ANCORA & "' AS CONVENIO_ANCORA, '" & TbDados!PrimeiroDeCNPJ_ANCORA & "' AS CNPJ_ANCORA, '" & TbDados!PrimeiroDeNOME_ANCORA & "' AS NOME_ANCORA, '" & TbDados!NOME_FORNECEDOR & "' AS NOME_FORNECEDOR, '" & Format(TbDados!CNPJ, "00000000000000") & "' AS CNPJ, '" & Left(Format(TbDados!CNPJ, "00000000000000"), 8) & "' AS CNPJRAIZ, '" & TbDados!ENDERECO & "' AS ENDERECO, '" & TbDados!NUMERO & "' AS NUMERO, '" & TbDados!BAIRRO & "' AS BAIRRO, '" & TbDados!CIDADE & "' AS CIDADE, '" & TbDados!UF & "' AS UF, '" & TbDados!Email & "' AS EMAIL, '" & TbDados!DATA_INICI & "' AS DATA_INICI, '" & TbDados!banco & "' AS BANCO, '" & TbDados!Agencia & "' AS AGENCIA, '" & TbDados!Conta & "' AS CONTA," _
'                            & " '" & TbDados!TTR & "' AS TTR, '" & TbDados!TCO & "' AS TCO, '" & TbDados!DPC & "' AS DPC, '" & TbDados!TIPO_CONVENIO & "' AS TIPO_CONVENIO, #" & Format(Date, "mm/dd/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO;")
'
'                            'Atualizar a consulta para
'                            TbDados.Requery
'
'                        'Verifica se com a atualização a consulta ficou em branco
'                        If TbDados.EOF = True Then: Exit Do
'
'                    Loop
'                'Deletar os fornecedores enviados
'                DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Agrupados ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Agrupados.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Agrupados.Agencia) WHERE (((Tbl_Convenio_Agrupados.Grupo)='" & TbGrupos!Grupo & "')) GROUP BY Tbl_Temp.ID)));")
'
'                TbGrupos.MoveNext
'            Loop
'    '===============================================================================================================================================================
'
'        'Deletar os convenios da Raizen
'        DBTVirtual.Execute ("Delete Tbl_Temp.ID FROM Tbl_Temp WHERE (((Tbl_Temp.ID) In (SELECT Tbl_Temp.ID FROM Tbl_Temp INNER JOIN Tbl_Convenio_Raizen ON (Tbl_Temp.CONVENIO_ANCORA = Tbl_Convenio_Raizen.Convenio) AND (Tbl_Temp.AGENCIA_ANCORA = Tbl_Convenio_Raizen.Agencia) GROUP BY Tbl_Temp.ID)));")
'
'    '===============================================================================================================================================================
'
'        'Selecionar todos os fornecedores cadastrados
'        Set TbDados = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.ID, Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.CNPJ_RAIZ, Tbl_Temp.ENDERECO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Format([Tbl_Temp]![AGENCIA_ANCORA],'0000') & Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Temp]![CNPJ_RAIZ],'00000000') AS CHAVE FROM Tbl_Temp" _
'        & " GROUP BY Tbl_Temp.ID, Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.CNPJ_RAIZ, Tbl_Temp.ENDERECO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, Format([Tbl_Temp]![AGENCIA_ANCORA],'0000') & Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Temp]![CNPJ_RAIZ],'00000000')" _
'        & " HAVING (((Format([Tbl_Temp]![AGENCIA_ANCORA],'0000') & Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Temp]![CNPJ_RAIZ],'00000000')) Not In (SELECT Format([Tbl_Enviados]![AGENCIA_ANCORA],'0000') & Format([Tbl_Enviados]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Enviados]![CNPJ_RAIZ],'00000000') AS CHAVE FROM Tbl_Enviados WHERE (((Tbl_Enviados.DATA_INICI) > #5/10/2017#)) GROUP BY Format([Tbl_Enviados]![AGENCIA_ANCORA],'0000') & Format([Tbl_Enviados]![CONVENIO_ANCORA],'000000000000') & Format([Tbl_Enviados]![CNPJ_RAIZ],'00000000');)));", dbOpenDynaset)
'
'            Do While TbDados.EOF = False
'
'                'Deletar tabela temporaria
'                DBTVirtual.Execute ("DELETE Tbl_Temp_WORD.* FROM Tbl_Temp_WORD;")
'
'                    'Incluir Informações na tabela temporaria
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Temp_WORD ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, NOME_ANCORA, CNPJ_ANCORA, NOME_FORNECEDOR, CNPJ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC ) SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.CNPJ_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.CNPJ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC FROM Tbl_Temp WHERE (((Tbl_Temp.ID)=" & TbDados!ID & "));")
'
'                        'Chamar função para gerar arquivo Word
'                        NOMEARQUIVO = Salvar_Arquivo_WORD(TbDados!NOME_FORNECEDOR, TbDados!Convenio_Ancora, TbDados!DATA_INICI, TbDados!TIPO_CONVENIO, TbDados!CNPJ, "Não")
'
'                        'Chamar Funcção para enviar o PDF pelo Email
'                        Call EnviarArquivo(NOMEARQUIVO, Trim(TbDados!Email), Trim(TbDados!NOME_FORNECEDOR), TbDados!Nome_Ancora)
'
'                    'Incluir dados do arquivo enviado na Tbl_Enviado
'                    DBTVirtual.Execute ("INSERT INTO Tbl_Enviados ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ_ANCORA, NOME_ANCORA, NOME_FORNECEDOR, CNPJ, CNPJ_RAIZ, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, DATA_INICI, BANCO, AGENCIA, CONTA, TTR, TCO, DPC, TIPO_CONVENIO, DATA_HORA, USUARIO ) SELECT Tbl_Temp.BANCO_ANCORA, Tbl_Temp.AGENCIA_ANCORA, Tbl_Temp.CONVENIO_ANCORA, Tbl_Temp.CNPJ_ANCORA, Tbl_Temp.NOME_ANCORA, Tbl_Temp.NOME_FORNECEDOR, Format([Tbl_Temp]![CNPJ],'00000000000000') AS CNPJ, Left(Format([Tbl_Temp]![CNPJ],'00000000000000'),8) AS CNPJRAIZ, Tbl_Temp.ENDERECO, Tbl_Temp.NUMERO, Tbl_Temp.BAIRRO, Tbl_Temp.CIDADE, Tbl_Temp.UF, Tbl_Temp.EMAIL, Tbl_Temp.DATA_INICI, Tbl_Temp.BANCO, Tbl_Temp.AGENCIA, Tbl_Temp.CONTA, Tbl_Temp.TTR, Tbl_Temp.TCO, Tbl_Temp.DPC, Tbl_Temp.TIPO_CONVENIO, #" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "HH:MM:SS") & "# AS DATA_HORA, '" & PesqUsername() & "' AS USUARIO FROM Tbl_Temp WHERE (((Tbl_Temp.ID)=" & TbDados!ID & "));")
'
'                    'Atualizar a consulta para
'                    TbDados.Requery
'
'                    'Verifica se com a atualização a consulta ficou em branco
'                    If TbDados.EOF = True Then: Exit Do
'            Loop
End Sub
Function ReplaceString(palavra As String)

    NomeReplace = Replace(palavra, "¸", "Ç")
    NomeReplace = Replace(NomeReplace, "¶", "Ã")
    NomeReplace = Replace(NomeReplace, "ù", "Õ")
    NomeReplace = Replace(NomeReplace, "»", "É")
    NomeReplace = Replace(NomeReplace, "þ", "Ú")
    NomeReplace = Replace(NomeReplace, "ø", "Ó")
    NomeReplace = Replace(NomeReplace, "¼", "Ê")
    NomeReplace = Replace(NomeReplace, "µ", "Á")
    NomeReplace = Replace(NomeReplace, "õ", "Ô")
    NomeReplace = Replace(NomeReplace, "¿", "Í")
    NomeReplace = Replace(NomeReplace, "²", "Â")
    
ReplaceString = NomeReplace

End Function
Public Function Conveter_CSV4XLSX_Convenios()
    
    Dim FSO As New FileSystemObject
    Dim ObjExcel As Object, ObjExcelPlan1 As Object

        Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\Modelo.xlsx"
        Set ObjExcelPlan1 = ObjExcel.Worksheets(1)

            ObjExcelPlan1.Range("A1").Select
            
            With ObjExcelPlan1.QueryTables.Add(Connection:="TEXT;C:\Temp\CONVENIOS.CSV", _
                 Destination:=ObjExcelPlan1.Range("$A$1"))
                 .Name = "CONVENIOS"
                 .FieldNames = True
                 .RowNumbers = False
                 .FillAdjacentFormulas = False
                 .PreserveFormatting = True
                 .RefreshOnFileOpen = False
                 .RefreshStyle = xlInsertDeleteCells
                 .SavePassword = False
                 .SaveData = True
                 .AdjustColumnWidth = True
                 .RefreshPeriod = 0
                 .TextFilePromptOnRefresh = False
                 .TextFilePlatform = 1252
                 .TextFileStartRow = 1
                 '.TextFileParseType = xlDelimited
                 '.TextFileTextQualifier = xlTextQualifierDoubleQuote
                 .TextFileConsecutiveDelimiter = False
                 .TextFileTabDelimiter = True
                 .TextFileSemicolonDelimiter = True
                 .TextFileCommaDelimiter = False
                 .TextFileSpaceDelimiter = False
                 .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
                 , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
                 1, 1, 1, 1, 1, 1, 1)
                 .TextFileTrailingMinusNumbers = True
                 .Refresh BackgroundQuery:=False
             End With
            
            If FSO.FileExists("C:\Temp\CONVENIOS.xlsx") Then: FSO.DeleteFile "C:\Temp\CONVENIOS.xlsx", True
            
        ObjExcelPlan1.SaveAs FileName:="C:\Temp\CONVENIOS.xlsx"
        ObjExcel.activeworkbook.Close SaveChanges:=False
        ObjExcel.Quit

End Function
Function Converter_CSV4XLSX_Fornecedores()
    
    Dim FSO As New FileSystemObject
    Dim ObjExcel As Object, ObjExcelPlan1 As Object

        Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\Modelo.xlsx"
        Set ObjExcelPlan1 = ObjExcel.Worksheets(1)
                  
            ObjExcelPlan1.Range("A1").Select
            With ObjExcelPlan1.QueryTables.Add(Connection:="TEXT;C:\Temp\FORNECEDORES.CSV" _
                , Destination:=ObjExcelPlan1.Range("$A$1"))
            .Name = "FORNECEDORES_210616"
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 1252
            .TextFileStartRow = 1
            .TextFileParseType = 1
            '.TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = True
            .TextFileSemicolonDelimiter = True
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 _
            , 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
            End With
            
            If FSO.FileExists("C:\Temp\FORNECEDORES.xlsx") Then: FSO.DeleteFile "C:\Temp\FORNECEDORES.xlsx", True
            
        ObjExcelPlan1.SaveAs FileName:="C:\Temp\FORNECEDORES.xlsx"
        ObjExcel.activeworkbook.Close SaveChanges:=False
        ObjExcel.Quit

End Function
Function ImportarBaseConvenios()
    
    'Abrir Banco de dados
    Call AbrirDBTVirtual
        
        'Limpar Tabela temporia de convenios
        DBTVirtual.Execute ("DELETE TbL_Temp_Convenios.* FROM TbL_Temp_Convenios;")
            
            'Inserir convenios extraidos do arquivo conveios.csv
            'DBTVirtual.Execute ("INSERT INTO TbL_Temp_Convenios ( DATA_ATUAL, BCO, AGENCIA, NR_CONVENIO, PRODUTO, SUBPRODU, BCO_PREM, AG_PREM, CTA_PREM, ETCA, DT_CAD, DT_ULT_MOV, NOME_CONVENIO, TP_PESS, CPFCNPJ, CONTATO, DDD1, TELEFONE1, RAML1, DDD2, TELEFONE2, RAML2, DFAX, TELEFONE3, CTRL__REME, FIEL_DPOS, RETORNO_LIQUI, WEB, PRE_APROV, IND_REMU_CLIE, EMAIL, MOTIVO_BLOQUEIO, AMBIENTE, TIP_BLOQUEIO, TIPO_COMPROVANTE, ENCARTE, ENDERECO, FORMA_LIQUIDACAO, APUR_PREMIO, FORMA_PGTO_PREMIO, RETORNO_AGENDAM, FRANCESINHA, FLUXO_OPERACIONAL, MODALIDADE_OPERACAO, NOTIFIC, REMU_PRE, STATUS_CONV, NR_SEQ_REMU, NR_SEQ_RETN, SPREAD, PERC. AGEN FORNEC, VLR_TRAVA, VLR_UTILIZ, INDIC. TRAVA VALOR, QTDE_A, QTDE_P, QTDE_S, QTDE_T, VLR_ACU_ATU, VLR_ACU_PRM, VLR_ACU_SEQ, VLR_ACU_TER, VLR_RECE_ATU, VLR_RECE_PRM, VLR_RECE_SEG, VLR_RECE_TER, VLR_BCO_ATU, VLR_BCO_PMR, VLR_BCO_SEG, VLR_BCO_TER, VLR_CLIE_MES_ATU, VLR_CLIE_PMR, VLR_CLIE_SEG, VLR_CLIE_TER, AV_ENV, PERC_REC," _
            & " PERC_ATR, PERC_DES, IND_APRO_OPE, IND_GPO_CONV, BCO_AGR, AGE_AGR, CONV_AGRU, DIAS_REJE, DT_INI_LO, PZ_LO, DT_VENC_LO, PZ_LIM, VL_TRAVA_MASS, AG_BUS, CTA_BU, DIG_BU, USUARIO, DH_ATUALIZ, TIPO_DE_CONTROLE, TIPO_DE_REMESSA, ENCARTE_CT, PERC_AJUS, IND__EQUALIZACAO_CLIENTE, PC_SG_EQUA, IND_NOTIFICACAO_NOTAS_PFORNEC, FDLC, GAROF, PC_FLEX, QT_CPRO_GPO, TP_REJE_CRP, F107 )" _
            & " SELECT Convenios.[DT ATUAL  ], Convenios.[BCO ], Convenios.[AG# ], Convenios.[NR CONVENIO ], Convenios.[PRODUTO ], Convenios.SUBPRODU, Convenios.[BCO PREM], Convenios.[AG# PREM], Convenios.[CTA PREM    ], Convenios.ETCA, Convenios.[DT CAD#   ], Convenios.[DT ULT#MOV], Convenios.[NOME CONVENIO                 ], Convenios.[TP PESS ], Convenios.[CPF/CNPJ       ], Convenios.[CONTATO                       ], Convenios.DDD1, Convenios.TELEFONE1, Convenios.RAML1, Convenios.DDD2, Convenios.TELEFONE2, Convenios.RAML2, Convenios.DFAX, Convenios.TELEFONE3, Convenios.[CTRL# REME#], Convenios.[FIEL DPOS], Convenios.[RETORNO LIQUI# ], Convenios.WEB, Convenios.[PRE APROV#], Convenios.[IND REMU CLIE            ], Convenios.[EMAIL                                                           ], Convenios.[MOTIVO BLOQUEIO     ], Convenios.[AMBIENTE      ], Convenios.[TIP BLOQUEIO]," _
            & " Convenios.[TIPO COMPROVANTE                                              ] , Convenios.[ENCARTE      ], Convenios.[ENDERECO                ], CONVENIOs.[FORMA LIQUIDACAO                  ], Convenios.[APUR PREMIO], Convenios.[FORMA PGTO PREMIO      ], Convenios.[RETORNO AGENDAM# ], Convenios.[FRANCESINHA       ], Convenios.[FLUXO OPERACIONAL         ], Convenios.[MODALIDADE OPERACAO     ], Convenios.NOTIFIC, Convenios.[REMU PRE], Convenios.[STATUS CONV#  ], Convenios.[NR SEQ#REMU# ], Convenios.[NR SEQ#RETN# ], Convenios.[SPREAD       ], Convenios.[PERC. AGEN FORNEC ], Convenios.[VLR TRAVA        ], Convenios.[VLR UTILIZ#      ], Convenios.INDIC. TRAVA VALOR, Convenios.[QTD ACUM MES ATUAL], Convenios.[QTD ACUM PRIMEIRO MES], Convenios.[QTD ACUM SEGUNDO MES], Convenios.[QTD ACUM TERCEIRO MES], Convenios.[VALOR ACUM MES ATUAL]," _
            & " Convenios.[VALOR ACUM PRIMEIRO MES], Convenios.[VALOR ACUM SEGUNDO MES], Convenios.[VALOR ACUM TERCEIRO MES], Convenios.[VALOR RECEITA MES ATUAL], Convenios.[VALOR RECEITA PRIMEIRO MES], " _
            & " Convenios.[VALOR RECEITA SEGUNDO MES], Convenios.[VALOR RECEITA TERCEIRO MES], Convenios.[VALOR BANCO MES ATUAL], Convenios.[VALOR BANCO PRIMEIRO MES ], Convenios.[VALOR BANCO SEGUNDO MES]," _
            & " Convenios.[VALOR BANCO TERCEIRO MES  ] , Convenios.[VALOR CLIENTE MES ATUAL], Convenios.[VALOR CLIENTE PRIMEIRO MES], Convenios.[VALOR CLIENTE SEGUNDO MES], Convenios.[VALOR CLIENTE TERCEIRO MES], Convenios.[NUM ULTIMO AVISO ENVIADO], Convenios.[PERC REC], Convenios.[PERC ATR], Convenios.[PERC DES], Convenios.[IND APRO OPE], Convenios.[IND GPO CONV], Convenios.[BCO AGR], Convenios.[AGE AGR], Convenios.[CONV AGRU   ], Convenios.[DIAS REJE], Convenios.[DT INI LO ], Convenios.[PZ LO ], Convenios.[DT VENC LO], Convenios.[PZ LIM], Convenios.[VL TRAVA MASS    ], Convenios.[AG BUS], Convenios.[CTA BU], Convenios.[DIG BU], Convenios.[USUARIO ], Convenios.[DH ATUALIZ], Convenios.[TIPO DE CONTROLE               ], Convenios.[TIPO DE REMESSA      ], Convenios.[ENCARTE CT], Convenios.[PERC AJUSTE ], Convenios.[IND# EQUALIZACAO CLIENTE   ]," _
            & " Convenios.[PC SG EQUA], Convenios.[IND#NOTIFICACAO NOTAS P/FORNEC  ], Convenios.[FORMA DE CONSOLIDACAO], Convenios.[GAROF       ], Convenios.[PC FLEX], Convenios.[QT CPRO GPO], Convenios.[TP REJE CRP], Convenios.F107 FROM Convenios;")
            
            DBTVirtual.Execute ("INSERT INTO TbL_Temp_Convenios ( DATA_ATUAL, BCO, AGENCIA, NR_CONVENIO, PRODUTO, SUBPRODU, BCO_PREM, AG_PREM, CTA_PREM, ETCA, DT_CAD, DT_ULT_MOV, NOME_CONVENIO, TP_PESS, CPFCNPJ ) SELECT Convenios.[DT ATUAL  ], Convenios.[BCO ], Convenios.[AG# ], Convenios.[NR CONVENIO ], Convenios.[PRODUTO ], Convenios.SUBPRODU, Convenios.[BCO PREM], Convenios.[AG# PREM], Convenios.[CTA PREM    ], Convenios.ETCA, Convenios.[DT CAD#   ], Convenios.[DT ULT#MOV], Convenios.[NOME CONVENIO                 ], Convenios.[TP PESS ], Convenios.[CPF/CNPJ       ] FROM Convenios;")

            'Atualizar formatacao do campo convenio para 12 caracteres
            DBTVirtual.Execute ("UPDATE TbL_Temp_Convenios SET TbL_Temp_Convenios.NR_CONVENIO = Format([TbL_Temp_Convenios]![NR_CONVENIO],'000000000000');")
        
        'Fechar banco de dados
        DBTVirtual.Close

End Function
Function ImpotarBaseFornecedores()
    
    'Abrir Banco de dados
    Call AbrirDBTVirtual
        
        'Limpar Tabela temporia de Fornecedores
        DBTVirtual.Execute ("DELETE Tbl_Temp.* FROM Tbl_Temp;")
            
            'Pesquisar ultimo dia util
            UltDia = Format(UltimoDiaUtil(), "DD/MM/YYYY")
                
                'Inserir convenios extraidos do arquivo Fornecedores.csv
                'Inclusão do CNPJ Raiz
                'DBTVirtual.Execute ("INSERT INTO Tbl_Temp ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ, CNPJ_RAIZ, NOME_FORNECEDOR, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, TTR, TCO, DATA_INICI )" _
                & " SELECT Fornecedores.[BCO ], Fornecedores.AGEN, Fornecedores.[NRO CONVENIO], Format([Fornecedores]![CPF/CNPJ       ],'00000000000000') AS CNPJ, Left(Format([Fornecedores]![CPF/CNPJ       ],'00000000000000'),8) AS CNPJRAIZ, Fornecedores.[RAZAO SOCIAL                            ], Fornecedores.[LOGRADOURO                              ], Fornecedores.[NRO# ], Fornecedores.[COMPLEMENTO         ], Fornecedores.[CIDADE                        ], Fornecedores.UF, Fornecedores.[EMAIL                                                           ], Fornecedores.TTR, Fornecedores.TCO, Fornecedores.[DATA CADAST FORN] FROM Fornecedores WHERE (((Fornecedores.[TIPO CNTR FORN])='V'))" _
                & " GROUP BY Fornecedores.[BCO ], Fornecedores.AGEN, Fornecedores.[NRO CONVENIO], Format([Fornecedores]![CPF/CNPJ       ],'00000000000000'), Left(Format([Fornecedores]![CPF/CNPJ       ],'00000000000000'),8), Fornecedores.[RAZAO SOCIAL                            ], Fornecedores.[LOGRADOURO                              ], Fornecedores.[NRO# ], Fornecedores.[COMPLEMENTO         ], Fornecedores.[CIDADE                        ], Fornecedores.UF, Fornecedores.[EMAIL                                                           ], Fornecedores.TTR, Fornecedores.TCO, Fornecedores.[DATA CADAST FORN] HAVING (((Fornecedores.[DATA CADAST FORN])=#" & Format(UltDia, "mm/dd/yyyy") & "#));")
             
             
                'DBTVirtual.Execute ("INSERT INTO Tbl_Temp ( BANCO_ANCORA, AGENCIA_ANCORA, CONVENIO_ANCORA, CNPJ, CNPJ_RAIZ, NOME_FORNECEDOR, ENDERECO, NUMERO, BAIRRO, CIDADE, UF, EMAIL, TTR, TCO, DATA_INICI )" _
                & " SELECT Fornecedores.[BCO ], Fornecedores.AGEN, Fornecedores.[NRO CONVENIO], Format([Fornecedores]![CPF/CNPJ       ],'00000000000000') AS CNPJ, Left(Format([Fornecedores]![CPF/CNPJ       ],'00000000000000'),8) AS CNPJRAIZ, ReplaceString([Fornecedores]![RAZAO SOCIAL                            ]) AS RAZAOSOCIAL, ReplaceString([Fornecedores]![LOGRADOURO                              ]) AS LOGRADOURO, Fornecedores.[NRO# ], ReplaceString([Fornecedores]![COMPLEMENTO         ]) AS COMPLEMENTO, ReplaceString([Fornecedores]![CIDADE                        ]) AS CIDADE, Fornecedores.UF, Fornecedores.[EMAIL                                                           ], Fornecedores.TTR, Fornecedores.TCO, Fornecedores.[DATA CADAST FORN] FROM Fornecedores WHERE (((Fornecedores.[TIPO CNTR FORN])='V'))" _
                & " GROUP BY Fornecedores.[BCO ], Fornecedores.AGEN, Fornecedores.[NRO CONVENIO], Format([Fornecedores]![CPF/CNPJ       ],'00000000000000'), Left(Format([Fornecedores]![CPF/CNPJ       ],'00000000000000'),8), ReplaceString([Fornecedores]![RAZAO SOCIAL                            ]), ReplaceString([Fornecedores]![LOGRADOURO                              ]), Fornecedores.[NRO# ], ReplaceString([Fornecedores]![COMPLEMENTO         ]), ReplaceString([Fornecedores]![CIDADE                        ]), Fornecedores.UF, Fornecedores.[EMAIL                                                           ], Fornecedores.TTR, Fornecedores.TCO, Fornecedores.[DATA CADAST FORN] HAVING (((Fornecedores.[DATA CADAST FORN])=#" & Format(UltDia, "mm/dd/yyyy") & "#));")
                                  
                Debug.Print "Atualizando tabela de fornecedores " & Time
               
               'Limpar Tabela temporia de Fornecedores
                DBTVirtual.Execute ("DELETE Tbl_Fornecedores.* FROM Tbl_Fornecedores;")
                
                DBTVirtual.Close
                Set DBTVirtual = Nothing
                Call processarImportacaoPlan("C:\temp\FORNECEDORES.xlsx", Format(UltDia, "dd/mm/yyyy"))
                Dim dbCurrent As DAO.Database
                Set dbCurrent = CurrentDb
                dbCurrent.Execute "delete * from Tbl_Temp_Local"
                dbCurrent.Execute "delete * from Tbl_Fornecedores_local"
                dbCurrent.Execute "Qry_InsTempLocal"
                dbCurrent.Execute "Qry_InsFornecedoresLocal"
                dbCurrent.Execute "delete * from Aux_Fornecedores"
                dbCurrent.Execute "delete * from Aux_Fornecedores2"
                dbCurrent.Execute "Qry_InsTempRede"
                dbCurrent.Execute "Qry_InsFornecedoresRede"
                dbCurrent.Execute "delete * from Tbl_Fornecedores_local"
                dbCurrent.Execute "delete * from Tbl_Temp_Local"
                Set dbCurrent = Nothing
                Call AbrirDBTVirtual
        
                'DBTVirtual.Execute ("INSERT INTO Tbl_Fornecedores ( [DATA ATUAL], BCO, AGEN, [NRO CONVENIO], [CPF/CNPJ], [TP PESS], [DT REC FAX], [DT ULT MOV], [RAZAO SOCIAL], LOGRADOURO, NRO, COMPLEMENTO, CIDADE, UF, CT, CEP, DDD, TELEFONE, RMAL, DFAX, [TELEF FAX], CONTATO, DSEC, [TELEF SEC], RMAL1, [SIT FORNECEDOR], [OPER FAX], [MX OPER TERM], [MX DIA TERM], EMAIL, [MOTIVO BLOQUEIO], [TP BLOQ], [LIBE S TERM], TTR, TCO, [VL FATUR], [NEGO TRIM], [TRAV VAL ], [QT AC MES], [QT AC PRMR], [QT AC SEGU], [QT AC TERC], [VL AC MES], [VL AC PRMR], [VL AC SEGU], [VL AC TERC], [VL TRAVA], [VL UTIZ], [VL RECE MES], [VL RECE PRMR], [VL RECE SEGU], [VL RECE TERC], [VL BANC MES], [VL BANC PRMR], [VL BANC SEGU], [VL BANC TERC], [VL CLIE MES], [VL CLIE PRMR], [VL CLIE SEGU], [VL CLIE TERC], [PC RATE], [LIBE CRTO MAE], [DIAS CRTO MAE]," _
                & " [OPER CRTO MAE], [LIBE CRTO ASSN], [DIAS CRTO ASSN], [OPER CRTO ASSN], USUARIO, [ULT ATUALIZACAO], [IND NOTIF NOTAS], [CD ENVIO EMAIL], [CONTRATO MAE], [FX DOCUMENTACAO], [EMAIL SECUN], [ENVIO AUT EMAIL], [TIPO CNTR FORN], [TIPO MEIO ACEITE], [DATA CADAST FORN], [ANTECIP AUTOMAT], [EMAIL FORMALIZ], [CONTA FORNECEDOR 1], [CONTA FORNECEDOR 2], [CONTA FORNECEDOR 3], [CONTA FORNECEDOR 4], [CONTA FORNECEDOR 5], [CONTA FORNECEDOR 6], [CONTA FORNECEDOR 7], [CONTA FORNECEDOR 8], [CONTA FORNECEDOR 9], [CONTA FORNECEDOR 10] )" _
                & " SELECT Fornecedores.[DATA ATUAL], Fornecedores.[BCO ], Fornecedores.AGEN, Format([Fornecedores]![NRO CONVENIO],'000000000000') AS CONVENIO, Fornecedores.[CPF/CNPJ       ], Fornecedores.[TP PESS ], Fornecedores.[DT REC FAX], Fornecedores.[DT ULT MOV], Fornecedores.[RAZAO SOCIAL                            ], Fornecedores.[LOGRADOURO                              ], Fornecedores.[NRO# ], Fornecedores.[COMPLEMENTO         ], Fornecedores.[CIDADE                        ], Fornecedores.UF, Fornecedores.CT, Fornecedores.[CEP      ], Fornecedores.[DDD ], Fornecedores.[TELEFONE  ], Fornecedores.RMAL, Fornecedores.DFAX, Fornecedores.[TELEF FAX ], Fornecedores.[CONTATO                       ], Fornecedores.DSEC, Fornecedores.[TELEF SEC ], Fornecedores.RMAL1, Fornecedores.[SIT FORNECEDOR]," _
                & " Fornecedores.[OPER FAX], Fornecedores.[MX OPER TERM], Fornecedores.[MX DIA TERM], Fornecedores.[EMAIL                                                           ], Fornecedores.[MOTIVO BLOQUEIO     ], Fornecedores.[TP BLOQ              ], Fornecedores.[LIBE S TERM], Fornecedores.TTR, Fornecedores.TCO, Fornecedores.[VL FATUR           ], Fornecedores.[NEGO TRIM          ]," _
                & " Fornecedores.[TRAV VAL ], Fornecedores.[QT AC MES ], Fornecedores.[QT AC PRMR], Fornecedores.[QT AC SEGU], Fornecedores.[QT AC TERC], Fornecedores.[VL AC MES          ], Fornecedores.[VL AC PRMR         ], Fornecedores.[VL AC SEGU         ], Fornecedores.[VL AC TERC         ], Fornecedores.[VL TRAVA           ], Fornecedores.[VL UTIZ            ], Fornecedores.[VL RECE MES        ], Fornecedores.[VL RECE PRMR       ], Fornecedores.[VL RECE SEGU       ], Fornecedores.[VL RECE TERC       ], Fornecedores.[VL BANC MES        ], Fornecedores.[VL BANC PRMR       ], Fornecedores.[VL BANC SEGU       ], Fornecedores.[VL BANC TERC       ], Fornecedores.[VL CLIE MES        ], Fornecedores.[VL CLIE PRMR       ], Fornecedores.[VL CLIE SEGU       ], Fornecedores.[VL CLIE TERC       ]," _
                & " Fornecedores.[PC RATE], Fornecedores.[LIBE CRTO MAE], Fornecedores.[DIAS CRTO MAE], Fornecedores.[OPER CRTO MAE], Fornecedores.[LIBE CRTO ASSN], Fornecedores.[DIAS CRTO ASSN], Fornecedores.[OPER CRTO ASSN], Fornecedores.[USUARIO ],Fornecedores.[ULT ATUALIZACAO           ], Fornecedores.[IND NOTIF NOTAS ], Fornecedores.[CD ENVIO EMAIL                          ], Fornecedores.[CONTRATO MAE        ]," _
                & " Fornecedores.[FX DOCUMENTACAO], Fornecedores.[EMAIL SECUN                                                     ], Fornecedores.[ENVIO AUT EMAIL], Fornecedores.[TIPO CNTR FORN], Fornecedores.[TIPO MEIO ACEITE], Fornecedores.[DATA CADAST FORN], Fornecedores.[ANTECIP AUTOMAT], Fornecedores.[EMAIL FORMALIZ], Fornecedores.[CONTA FORNECEDOR 1       ], Fornecedores.[CONTA FORNECEDOR 2       ], Fornecedores.[CONTA FORNECEDOR 3       ], Fornecedores.[CONTA FORNECEDOR 4       ], Fornecedores.[CONTA FORNECEDOR 5       ], Fornecedores.[CONTA FORNECEDOR 6       ], Fornecedores.[CONTA FORNECEDOR 7       ], Fornecedores.[CONTA FORNECEDOR 8       ], Fornecedores.[CONTA FORNECEDOR 9       ], Fornecedores.[CONTA FORNECEDOR 10      ] FROM Fornecedores" _
                & " GROUP BY Fornecedores.[DATA ATUAL], Fornecedores.[BCO ], Fornecedores.AGEN, Format([Fornecedores]![NRO CONVENIO],'000000000000'), Fornecedores.[CPF/CNPJ       ], Fornecedores.[TP PESS ], Fornecedores.[DT REC FAX], Fornecedores.[DT ULT MOV], Fornecedores.[RAZAO SOCIAL                            ], Fornecedores.[LOGRADOURO                              ], Fornecedores.[NRO# ], Fornecedores.[COMPLEMENTO         ], Fornecedores.[CIDADE                        ], Fornecedores.UF, Fornecedores.CT, Fornecedores.[CEP      ], Fornecedores.[DDD ], Fornecedores.[TELEFONE  ], Fornecedores.RMAL, Fornecedores.DFAX, Fornecedores.[TELEF FAX ], Fornecedores.[CONTATO                       ], Fornecedores.DSEC, Fornecedores.[TELEF SEC ], Fornecedores.RMAL1," _
                & " Fornecedores.[SIT FORNECEDOR], Fornecedores.[OPER FAX], Fornecedores.[MX OPER TERM], Fornecedores.[MX DIA TERM], Fornecedores.[EMAIL                                                           ], Fornecedores.[MOTIVO BLOQUEIO     ], Fornecedores.[TP BLOQ    ], Fornecedores.[LIBE S TERM], Fornecedores.TTR, Fornecedores.TCO, Fornecedores.[VL FATUR           ], Fornecedores.[NEGO TRIM          ], Fornecedores.[TRAV VAL ], Fornecedores.[QT AC MES ], Fornecedores.[QT AC PRMR], Fornecedores.[QT AC SEGU], Fornecedores.[QT AC TERC], Fornecedores.[VL AC MES          ], Fornecedores.[VL AC PRMR         ], Fornecedores.[VL AC SEGU         ], Fornecedores.[VL AC TERC         ], Fornecedores.[VL TRAVA           ], Fornecedores.[VL UTIZ            ], Fornecedores.[VL RECE MES        ], Fornecedores.[VL RECE PRMR       ]," _
                & " Fornecedores.[VL RECE SEGU       ], Fornecedores.[VL RECE TERC       ], Fornecedores.[VL BANC MES        ], Fornecedores.[VL BANC PRMR       ], Fornecedores.[VL BANC SEGU       ], Fornecedores.[VL BANC TERC       ], Fornecedores.[VL CLIE MES        ], Fornecedores.[VL CLIE PRMR       ], Fornecedores.[VL CLIE SEGU       ], Fornecedores.[VL CLIE TERC       ], Fornecedores.[PC RATE], Fornecedores.[LIBE CRTO MAE], Fornecedores.[DIAS CRTO MAE], Fornecedores.[OPER CRTO MAE], Fornecedores.[LIBE CRTO ASSN], Fornecedores.[DIAS CRTO ASSN], Fornecedores.[OPER CRTO ASSN], Fornecedores.[USUARIO ]," _
                & " Fornecedores.[ULT ATUALIZACAO           ], Fornecedores.[IND NOTIF NOTAS ], Fornecedores.[CD ENVIO EMAIL                          ], Fornecedores.[CONTRATO MAE        ], Fornecedores.[FX DOCUMENTACAO], Fornecedores.[EMAIL SECUN                                                     ], Fornecedores.[ENVIO AUT EMAIL], Fornecedores.[TIPO CNTR FORN], Fornecedores.[TIPO MEIO ACEITE], Fornecedores.[DATA CADAST FORN], Fornecedores.[ANTECIP AUTOMAT], Fornecedores.[EMAIL FORMALIZ], Fornecedores.[CONTA FORNECEDOR 1       ], Fornecedores.[CONTA FORNECEDOR 2       ], Fornecedores.[CONTA FORNECEDOR 3       ], Fornecedores.[CONTA FORNECEDOR 4       ], Fornecedores.[CONTA FORNECEDOR 5       ], Fornecedores.[CONTA FORNECEDOR 6       ], Fornecedores.[CONTA FORNECEDOR 7       ]," _
                & " Fornecedores.[CONTA FORNECEDOR 8       ], Fornecedores.[CONTA FORNECEDOR 9       ], Fornecedores.[CONTA FORNECEDOR 10   ] HAVING (((Fornecedores.[DATA CADAST FORN])=#" & Format(UltDia, "mm/dd/yyyy") & "#));")
                                
                Debug.Print "Tabela de fornecedores atualizada " & Time
                
                'Atualizar formatacao do campo convenio para 12 caracteres
                'DBTVirtual.Execute ("UPDATE Tbl_Fornecedores SET Tbl_Fornecedores.[NRO CONVENIO] = Format([Tbl_Fornecedores]![NRO CONVENIO],'000000000000');")
                
                'Atualizar formatacao do campo convenio para 12 caracteres
                DBTVirtual.Execute ("UPDATE Tbl_Temp SET Tbl_Temp.CONVENIO_ANCORA = Format([Tbl_Temp]![CONVENIO_ANCORA],'000000000000');")
                
            'Atualiza nome do Ancora e tipo de convenio de acordo com subproduto
            'Db.Execute ("UPDATE Tbl_Temp INNER JOIN TbL_Temp_Convenios ON (Tbl_Temp.AGENCIA_ANCORA = TbL_Temp_Convenios.AGENCIA) AND (Tbl_Temp.CONVENIO_ANCORA = TbL_Temp_Convenios.NR_CONVENIO) SET Tbl_Temp.NOME_ANCORA = [TbL_Temp_Convenios]![NOME_CONVENIO], Tbl_Temp.TIPO_CONVENIO = [TbL_Temp_Convenios]![SUBPRODU];")
            
            DBTVirtual.Execute ("UPDATE Tbl_Temp INNER JOIN TbL_Temp_Convenios ON (Tbl_Temp.AGENCIA_ANCORA = TbL_Temp_Convenios.AGENCIA) AND (Tbl_Temp.CONVENIO_ANCORA = TbL_Temp_Convenios.NR_CONVENIO) SET Tbl_Temp.NOME_ANCORA = ReplaceString([TbL_Temp_Convenios]![NOME_CONVENIO]), Tbl_Temp.TIPO_CONVENIO = [TbL_Temp_Convenios]![SUBPRODU], Tbl_Temp.CNPJ_ANCORA = [TbL_Temp_Convenios]![CPFCNPJ];")
            
            'Atualiza os dados bancarios do fornecedor
            DBTVirtual.Execute ("UPDATE Tbl_Temp INNER JOIN Tbl_DadosBancarios ON Tbl_Temp.CNPJ = Tbl_DadosBancarios.CNPJ SET Tbl_Temp.BANCO = [Tbl_DadosBancarios]![BANCO], Tbl_Temp.AGENCIA = [Tbl_DadosBancarios]![AGENCIA], Tbl_Temp.CONTA = [Tbl_DadosBancarios]![CONTA];")
            
        'Fechar banco de dados
        DBTVirtual.Close
        Set DBTVirtual = Nothing
End Function
Function Replace_Letras_Fornecedores()

    Dim TbDados As Recordset: Call AbrirDBTVirtual
            
    'TESTE NOVA FUNÇÃO - 29/06/2017 - T683068
            
        DBTVirtual.Execute ("UPDATE Tbl_Temp SET Tbl_Temp.BAIRRO = ReplaceString(Trim([Tbl_Temp]![BAIRRO])), Tbl_Temp.NOME_FORNECEDOR = ReplaceString(Trim([Tbl_Temp]![NOME_FORNECEDOR])), Tbl_Temp.NOME_ANCORA = ReplaceString(Trim([Tbl_Temp]![Nome_Ancora])), Tbl_Temp.ENDERECO = ReplaceString(Trim([Tbl_Temp]![ENDERECO])), Tbl_Temp.CIDADE = ReplaceString(Trim([Tbl_Temp]![CIDADE]));")
              
'        Set TbDados = DBTVirtual.OpenRecordset("SELECT Tbl_Temp.BAIRRO, Tbl_Temp.NOME_FORNECEDOR, Tbl_Temp.NOME_ANCORA, Tbl_Temp.ENDERECO, Tbl_Temp.CIDADE FROM Tbl_Temp;", dbOpenDynaset)
'
'            Do While TbDados.EOF = False
'                    TbDados.Edit
'                        TbDados!BAIRRO = ReplaceString(Trim(TbDados!BAIRRO))
'                        TbDados!Nome_Ancora = ReplaceString(Trim(TbDados!Nome_Ancora))
'                        TbDados!NOME_FORNECEDOR = ReplaceString(Trim(TbDados!NOME_FORNECEDOR))
'                        TbDados!ENDERECO = ReplaceString(Trim(TbDados!ENDERECO))
'                        TbDados!CIDADE = ReplaceString(Trim(TbDados!CIDADE))
'                    TbDados.Update
'                TbDados.MoveNext
'            Loop
'        DBTVirtual.Close

End Function
Sub ImportarBases()

    Dim FSO As New FileSystemObject: UltDia = Format(UltimoDiaUtil(), "DDMMYY")
    Dim ObjExcel As Object, ObjExcelPlan1 As Object, TbCustoDIA As Recordset
    Dim Caminho As String: Caminho = "\\BSBRSP56\confirming relatorio\"
    'Dim Caminho As String: Caminho = "C:\temp\confirming relatorio\"

        If FSO.FolderExists("\\Saont46\apps2\Confirming\CONTRATOMAE_ENVIADO\" & Format(UltimoDiaUtil(), "DDMMYYYY")) = False Then
            FSO.CreateFolder ("\\Saont46\apps2\Confirming\CONTRATOMAE_ENVIADO\" & Format(UltimoDiaUtil(), "DDMMYYYY"))
        End If 'Criar pasta do dia na rede da mesa
                
        ArquivoDia = Caminho & "CONVENIOS_" & UltDia & ".CSV"           'Montar Arquivo Convenios do Dia
            If FSO.FileExists(ArquivoDia) Then                          'valida se o arquivo esta disponivel
                FSO.CopyFile ArquivoDia, "C:\Temp\CONVENIOS.CSV"        'Copiar arquivo do dia para a maquina
                Call Conveter_CSV4XLSX_Convenios                        'Converter aquivo CSV para XLSX
                Call ImportarBaseConvenios                              'Impotar Dados do XLSX Convenios
                Call AtualizaFielDeposPrincial                          'Atualizar Tabela Fiel Depositario da base principal para rodar a voluemtria diaria
            End If
        
        ArquivoDia = Caminho & "FORNECEDORES_" & UltDia & ".CSV"        'Montar Arquivo Convenios do Dia
            If FSO.FileExists(ArquivoDia) Then                          'valida se o arquivo esta disponivel
                FSO.CopyFile ArquivoDia, "C:\Temp\FORNECEDORES.CSV"     'Copiar arquivo do dia para a maquina
                Call Converter_CSV4XLSX_Fornecedores                    'Converter aquivo CSV para XLSX
                Call ImpotarBaseFornecedores                            'Impotar Dados do XLSX Fornecedores
                'Call Replace_Letras_Fornecedores                       'Ajustar as letras com caracteres especiais
                'Call GerarWORD_Fornecedores_V2                          'Função para Gerar WORD, PDF e Enviar Arquivos - Desabilitado 21/03/2018 Emerson
            End If
End Sub


