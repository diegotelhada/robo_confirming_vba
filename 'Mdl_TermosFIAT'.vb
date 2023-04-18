'Mdl_TermosFIAT'

Option Compare Database
'Constantes para caminhos de diretório

'Caminhos Produção
Private Const caminhoPDFRede = "\\Saont46\apps2\Confirming\PROJETORELATORIOS\FIAT_TERMO"

'Caminhos Teste
'Private Const caminhoPDFRede = "C:\Temp"

'Módulo para envio de termos para os fornecedores da FIAT
'Emerson - 03/2018

'Importação Arquivo ARQOPED dia
Public Sub importarArqOpedDia(dia As Variant)

    'Definição de bases, caminhos de arquivo e data
    Dim TbDados As Recordset, TbArq As Recordset
    Dim linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
    Dim Dataarq As Date, DataVenNF As Date, Data_op As Date
    Dim FSO As New FileSystemObject, arq As File, File As String, Contador As String

    DataPesq = dia
    DataPesq2 = Format(DataPesq, "DDMMYY")
    DataPesq1 = Format(DataPesq, "DD/MM/YYYY")

    File = "\\saont46\apps2\Confirming\ArquivosYC\ARQOPED" & DataPesq2 & ".TXT"
    ''Novo Caminho
    File = "\\fscore02\apps2\Confirming\ArquivosYC\ARQOPED" & DataPesq2 & ".TXT"

    'Cópia das informações do arquivo para a tabela temporária
    If (Dir(File) <> "") Then
        
        FSO.CopyFile File, "\\saont46\apps2\Confirming\PROJETORELATORIOS\ARQRESERVA\", True

        Call ReplaceEXE("\\saont46\apps2\Confirming\PROJETORELATORIOS\ARQRESERVA\ARQOPED" & DataPesq2 & ".TXT")

        Call TratarArquivoDeOperacoes("\\saont46\apps2\Confirming\PROJETORELATORIOS\ARQRESERVA\ARQOPED" & DataPesq2 & ".TXT")

        Set arq = FSO.GetFile(File)

        DataCri = arq.DateCreated: DataCri = Format(DataCri, "DD/MM/YYYY")

        'Limpar Tabela temporia de Fornecedores
        AbrirBDTermosFIAT

        BDTermosFIAT.Execute ("DELETE tmpArqopedDia.* FROM tmpArqopedDia;")

        Set TbDados = BDTermosFIAT.OpenRecordset("tmpArqOpedDia", dbOpenDynaset)
        Set TbArq = BDTermosFIAT.OpenRecordset("ARQ", dbOpenDynaset)

        Do While TbArq.EOF = False

            TbDados.AddNew
                TbDados!ARQDATA = Dataarq
                TbDados!SEGMENTO = TbArq!SEGMENTO
                TbDados!Nome_Ancora = SemCaracterEspecial(TbArq!Ancora)
                TbDados!Banco_Ancora = TbArq!banco
                TbDados!Agencia_Ancora = TbArq!Agencia
                TbDados!Convenio_Ancora = TbArq!Convenio
                TbDados!Cnpj_Ancora = Extrai_Zeros(TbArq!Cnpj_Ancora)
                TbDados!Nome_Fornecedor = SemCaracterEspecial(TbArq!Fornecedor)
                TbDados!Cnpj_Fornecedor = Extrai_Zeros(TbArq!Cnpj_Fornecedor)
                TbDados!cod_oper = TbArq!COD_OPERACAO
                TbDados!Modalidade_Oper = TbArq!Mod_Oper
                TbDados!Tipo_Liq = TbArq!Tipo_Liq
                TbDados!Banco_Remet = Extrai_Zeros(TbArq!Banco_Rem)
                TbDados!Agencia_Remet = Extrai_Zeros(TbArq!Ag_Rem)
                TbDados!Conta_Remet = Extrai_Zeros(TbArq!Conta_Rem)
                TbDados!Tipo_Pag = TbArq!Tipo_Pg
                TbDados!Banco_Fav = Extrai_Zeros(TbArq!Bco_Fav)
                TbDados!Agencia_Fav = Extrai_Zeros(TbArq!Ag_Fav)
                TbDados!Conta_Fav = Extrai_Zeros(TbArq!Conta_Fav)

                Data_op = Format(TbArq!Data_op, "DD/MM/YYYY")
                TbDados!Data_op = Data_op

                TbDados!Data_Final = Format(TbArq!Data_Fin, "DD/MM/YYYY")
                TbDados!Prazo_Medio = TbArq!Prazo_Medio

                Juros = TbArq!Juros
                TbDados!Juros = Juros / 10000000

                Custo = TbArq!Custo
                TbDados!Custo = Custo / 10000000

                Spread = TbArq!Spread
                TbDados!Spread = Spread / 10000000

                SpreadAnual = TbArq![Spread Anual]
                TbDados!Spread_Anual = SpreadAnual / 10000000

                ValorOp = TbArq!Valor_Op
                TbDados!Valor_Op = ValorOp / 100

                Valortco = TbArq!Valor_TCO
                TbDados!Valor_TCO = Valortco / 100

                valorttr = TbArq!Valor_TTR
                TbDados!Valot_TTR = valorttr / 100

                TbDados!Compromisso = TbArq!Compromisso

                DataVenNF = Format(TbArq!Data_Venc, "DD/MM/YYYY")
                TbDados!Data_Venc = DataVenNF

                Valornom = TbArq!Valor_Nom
                TbDados!Valor_Nom = Valornom / 100

                Valorabat = TbArq!Valor_aba
                TbDados!Valor_Abat = Valorabat / 100

                Valoracres = TbArq!Valor_Acr
                TbDados!Valor_Acres = Valoracres / 100

                Valorpg = TbArq!Valor_pg
                TbDados!Valor_Pagmto = Valorpg / 100

                ValorJuros = TbArq!Valor_Juros
                TbDados!Valor_Juros = ValorJuros / 100

                valoriof = TbArq!Valor_IOF
                TbDados!Valor_IOF = valoriof / 100

                Valorliq = TbArq!Valor_liq
                TbDados!Valor_Liquido = Valorliq / 100

                ValorCusto = TbArq!Valor_Custo
                TbDados!Valor_Custo = ValorCusto / 100

                SpreadBanco = TbArq!Spread_Bco
                TbDados!Spread_Banco = SpreadBanco / 10000000

                ReceitaBanco = TbArq!Receita_Bco
                TbDados!Receita_Banco = ReceitaBanco / 100


                TbDados!Tp_Apur_prem = TbArq![Tp Apur Prêm]
                TbDados!Tp_Rem_prem = TbArq![Tp Rem Prêm]
                TbDados!Tp_Pgto_prem = TbArq![Tp Pgto Prem]
                TbDados!Dt_pfto_Prem = TbArq![Dt Pfto Prem]
                TbDados!Cod_Bco_Prem = TbArq![Cod Bco Prem]
                TbDados!Cod_Age_Prem = TbArq![Cod Age Prem]
                TbDados!Cod_Conta_prem = TbArq![Cod C/C Prem]

                SpreadClte = TbArq![Rate Spread]
                TbDados!Spread_Clte = SpreadClte / 10000000

                SpreadClte = TbArq![Spread Clte]
                TbDados!Spread_Clte = SpreadClte / 10000000

                ReceitaClte = Extrai_Zeros(TbArq![Receita Clte])
                TbDados!Receita_Clte = ReceitaClte / 100

                TbDados!Prazo_NF = DataVenNF - Data_op

                TbDados!IN_OPER_RETR = TbArq![IN-OPER-RETR]
                TbDados!IN_APRV_OPER = TbArq![IN-APRV-OPER]
                TbDados!QT_DIA_REJE_OPER = TbArq![QT-DIA-REJE-OPER]
                TbDados!CD_USUA_ULTI_ATLZ = TbArq![CD-USUA-ULTI-ATLZ]

                DH_ULTI_ATLZ = Mid(TbArq![DH-ULTI-ATLZ], 9, 2) & "/" & Mid(TbArq![DH-ULTI-ATLZ], 6, 2) & "/" & Left(TbArq![DH-ULTI-ATLZ], 4) & " " & Mid(TbArq![DH-ULTI-ATLZ], 12, 2) & ":" & Mid(TbArq![DH-ULTI-ATLZ], 15, 2) & ":" & Mid(TbArq![DH-ULTI-ATLZ], 18, 2)
                TbDados!DH_ULTI_ATLZ = DH_ULTI_ATLZ

                TbDados!CD_EMPR_SAP = TbArq![CD-EMPR-SAP]
                TbDados!CD_FILI_EMPR_SAP = TbArq![CD-FILI-EMPR-SAP]
                TbDados!NR_DOCT_ORIG = TbArq![NR-DOCT-ORIG]
                TbDados!TP_DOCT_CNTB = TbArq![TP-DOCT-CNTB]
                TbDados!NR_CHAV_SAP = TbArq![NR-CHAV-SAP]
                TbDados!CD_FORN = TbArq![CD-FORN]
                TbDados!DT_EMIS_NOTA = TbArq![DT-EMIS-NOTA]
                TbDados!CD_BARR = TbArq![CD-BARR]
                TbDados!DT_NEGO = TbArq![DT-NEGO]

                DT_CONF = Right(TbArq![DT-CONF], 2) & "/" & Mid(TbArq![DT-CONF], 6, 2) & "/" & Left(TbArq![DT-CONF], 4)
                TbDados!DT_CONF = DT_CONF

                TbDados!VL_NEGO = TbArq![VL-NEGO]
                TbDados!VL_CONF = TbArq![VL-CONF]
                TbDados!Canal = TbArq![Canal]
                TbDados!Usuario = TbArq![Usuario]
            TbDados.Update
            TbArq.MoveNext

        Loop

        TbArq.Close
        TbDados.Close
        Set TbArq = Nothing
        Set TbDados = Nothing

        BDTermosFIAT.Close
        Set BDTermosFIAT = Nothing

    Else
        MsgBox "Arquivo de operações não encontrado, favor verificar"
        Exit Sub
    End If

End Sub

'Importação Arquivo Fornecedores (apenas campos de endereço e email)
Public Sub importarFornecedoresDiario(dia As Variant)

    'Definição de bases, caminhos de arquivo e data
    Dim FSO As New FileSystemObject: dia = Format(dia, "DDMMYY")
    Dim ObjExcel As Object, ObjExcelPlan1 As Object, TbCustoDIA As Recordset
    Dim Caminho As String: Caminho = "\\BSBRSP56\confirming relatorio\"

        ArquivoDia = Caminho & "FORNECEDORES_" & dia & ".CSV"       'Montar nome do arquivo

        If FSO.FileExists(ArquivoDia) Then                          'valida se o arquivo esta disponivel

            FSO.CopyFile ArquivoDia, "C:\Temp\FORNECEDORES.CSV"     'Copiar arquivo do dia para a maquina

           'Limpar Tabela temporia de Fornecedores
            AbrirBDTermosFIAT
            BDTermosFIAT.Execute ("DELETE tmpFornecedoresDia.* FROM tmpFornecedoresDia;")

            Dim ac As Access.Application
            Set ac = CreateObject("Access.Application")

            ac.OpenCurrentDatabase (BDTermosFIAT.Name)
            ac.DoCmd.TransferText acImportDelim, "Fornecedores_FIAT", "tmpFornecedoresDia", "C:\Temp\FORNECEDORES.CSV", True

            ac.CloseCurrentDatabase
            Set ac = Nothing

            BDTermosFIAT.Close
            Set BDTermosFIAT = Nothing
        Else
            MsgBox "Arquivo de fornecedores não encontrado, favor verificar"
            Exit Sub
        End If

End Sub

'Verificação de operações realizadas no dia útil anterior para os convênios cadastrados e complementa com as informações de fornecedores
Public Sub verOperTermosFIAT()

    'Conexão
    AbrirBDTermosFIAT

    'Limpeza tabelas dia anterior
    BDTermosFIAT.Execute ("DELETE tmpDadosTermo.* FROM tmpDadosTermo;")
    BDTermosFIAT.Execute ("DELETE tmpDadosNotas.* FROM tmpDadosNotas;")

    'Inclusão dos dados dos termos, notas e complementos de endereço
    BDTermosFIAT.Execute ("qry_ConsultaOperAntecipada01")
    BDTermosFIAT.Execute ("qry_ConsultaOperAntecipada02")
    BDTermosFIAT.Execute ("qry_ConsultaOperAntecipada03")
    BDTermosFIAT.Execute ("qry_ConsultaOperAntecipada04")
    BDTermosFIAT.Execute ("qry_ConsultaOperAntecipada05")

    'Finalização
    BDTermosFIAT.Close
    Set BDTermosFIAT = Nothing

End Sub

'Criação do termo
Public Sub criarTermoFIAT(ByVal numOper As String)

    'Variaveis para dados do termo
    Dim numTermo As String
    Dim numNotas As String
    Dim dataOper As String
    Dim nomeFornecedor, CNPJFornecedor, EnderecoFornecedor, BairroFornecedor, cidadeFornecedor, UFFornecedor As String
    Dim nomeDevedor, CNPJDevedor, EnderecoDevedor, BairroDevedor, cidadeDevedor, UFDevedor As String
    Dim valorGlobal, taxaDesagio, TCO, TTR, precoCessao As Double
    Dim CC_Credito_Banco, CC_Credito_Ag, CC_Credito_Favorecido As String
    Dim qryDadosTermo As Recordset
    Dim qryDadosConvenio As Recordset

    'Variáveis para dados das notas
    Dim numNota, valorNominal, dtVenc As String
    Dim qryDadosNota As Recordset

    'Captura de dados do termo e convenio
    AbrirBDTermosFIAT
    SQL = "SELECT * FROM tmpDadosTermo WHERE val(Cod_Oper) = val(" & numOper & ")"
    Set qryDadosTermo = BDTermosFIAT.OpenRecordset(SQL)

    SQL = "SELECT * FROM tblConveniosFIAT WHERE (((Convenio_Ancora = " & "'" & qryDadosTermo!Convenio_Ancora & "')" & " AND " & _
    "(Banco_Ancora = " & "'" & Format(qryDadosTermo!Banco_Ancora, "0000") & "')" & " AND " & _
    "(Agencia_Ancora = " & "'" & Format(qryDadosTermo!Agencia_Ancora, "0000") & "')" & _
    "))"
    Set qryDadosConvenio = BDTermosFIAT.OpenRecordset(SQL)

    numTermo = Format(numOper, "000000000000000")

    If Not IsNull(qryDadosTermo!N_COMPROMISSOS) Then numNotas = qryDadosTermo!N_COMPROMISSOS

    If Not IsNull(qryDadosTermo!Data_op) Then dataOper = qryDadosTermo!Data_op

    If Not IsNull(qryDadosTermo!Nome_Fornecedor) Then nomeFornecedor = SemCaracterEspecial(qryDadosTermo!Nome_Fornecedor)
    If Not IsNull(qryDadosTermo!Cnpj_Fornecedor) Then CNPJFornecedor = Format(qryDadosTermo!Cnpj_Fornecedor, "000000000000000")
    If Not IsNull(qryDadosTermo!Endereco_Fornecedor) Then EnderecoFornecedor = SemCaracterEspecial(qryDadosTermo!Endereco_Fornecedor)
    If Not IsNull(qryDadosTermo!Bairro_Fornecedor) Then BairroFornecedor = SemCaracterEspecial(qryDadosTermo!Bairro_Fornecedor)
    If Not IsNull(qryDadosTermo!Cidade_Fornecedor) Then cidadeFornecedor = SemCaracterEspecial(qryDadosTermo!Cidade_Fornecedor)
    If Not IsNull(qryDadosTermo!UF_Fornecedor) Then UFFornecedor = SemCaracterEspecial(qryDadosTermo!UF_Fornecedor)

    If Not IsNull(qryDadosTermo!Nome_Ancora) Then nomeDevedor = SemCaracterEspecial(qryDadosTermo!Nome_Ancora)

    If Not IsNull(qryDadosConvenio!Cnpj_Ancora) Then CNPJDevedor = Format(qryDadosConvenio!Cnpj_Ancora, "000000000000000")
    If Not IsNull(qryDadosConvenio!Endereco_Ancora) Then EnderecoDevedor = SemCaracterEspecial(qryDadosConvenio!Endereco_Ancora)
    If Not IsNull(qryDadosConvenio!Bairro_Ancora) Then BairroDevedor = SemCaracterEspecial(qryDadosConvenio!Bairro_Ancora)
    If Not IsNull(qryDadosConvenio!Cidade_Ancora) Then cidadeDevedor = SemCaracterEspecial(qryDadosConvenio!Cidade_Ancora)
    If Not IsNull(qryDadosConvenio!UF_Ancora) Then UFDevedor = SemCaracterEspecial(qryDadosConvenio!UF_Ancora)

    If Not IsNull(qryDadosTermo!Valor_Global) Then valorGlobal = Format(qryDadosTermo!Valor_Global, "#0.00")
    If Not IsNull(qryDadosTermo!Taxa_Desagio) Then taxaDesagio = Format(qryDadosTermo!Taxa_Desagio, "#0.0000000")
    If Not IsNull(qryDadosTermo!TCO) Then TCO = Format(qryDadosTermo!TCO, "#0.00")
    If Not IsNull(qryDadosTermo!TTR) Then TTR = Format(qryDadosTermo!TTR, "#0.00")
    If Not IsNull(qryDadosTermo!Preco_Cessao) Then precoCessao = Format(qryDadosTermo!Preco_Cessao, "#0.00")

    If Not IsNull(qryDadosTermo!CC_Credito_Agencia) Then CC_Credito_Ag = Format(qryDadosTermo!CC_Credito_Agencia, "00000")
    If Not IsNull(qryDadosTermo!CC_Credito_Favorecido) Then CC_Credito_Favorecido = Format(qryDadosTermo!CC_Credito_Favorecido, "0000000000000")
    If Not IsNull(qryDadosTermo!CC_Credito_Banco) Then CC_Credito_Banco = Format(qryDadosTermo!CC_Credito_Banco, "00000")

    qryDadosTermo.Close
    Set qryDadosTermo = Nothing

    qryDadosConvenio.Close
    Set qryDadosConvenio = Nothing

    BDTermosFIAT.Close
    Set BDTermosFIAT = Nothing

    'Abre Word e preenche dados do termo

    'Criação do Objeto
    Dim wd As Object
    Dim wdocSource As Object

    On Error Resume Next
    Set wd = GetObject(, "Word.Application")
    If wd Is Nothing Then
        Set wd = CreateObject("Word.Application")
    End If
    On Error GoTo 0

    'Criação do Arquivo
    wd.Visible = False
    Set wdocSource = wd.Documents.Open(caminhoPDFRede & "\Modelos\" & "Modelo Termo Cessão.docx")
    wd.ActiveDocument.SaveAs "C:\Temp\Termo Cessão_" & numOper & ".docx"

    'Preenchimento dos dados - Termo
    wd.ActiveDocument.Bookmarks("numTermo1").Range.InsertAfter numTermo
    wd.ActiveDocument.Bookmarks("nomeFornecedor1").Range.InsertAfter nomeFornecedor
    wd.ActiveDocument.Bookmarks("CNPJFornecedor1").Range.InsertAfter CNPJFornecedor
    wd.ActiveDocument.Bookmarks("EnderecoFornecedor1").Range.InsertAfter EnderecoFornecedor
    wd.ActiveDocument.Bookmarks("BairroFornecedor1").Range.InsertAfter BairroFornecedor
    wd.ActiveDocument.Bookmarks("CidadeFornecedor1").Range.InsertAfter cidadeFornecedor
    wd.ActiveDocument.Bookmarks("UFFornecedor1").Range.InsertAfter UFFornecedor
    wd.ActiveDocument.Bookmarks("nomeDevedor1").Range.InsertAfter nomeDevedor
    wd.ActiveDocument.Bookmarks("CNPJDevedor1").Range.InsertAfter CNPJDevedor
    wd.ActiveDocument.Bookmarks("EnderecoDevedor1").Range.InsertAfter EnderecoDevedor
    wd.ActiveDocument.Bookmarks("BairroDevedor1").Range.InsertAfter BairroDevedor
    wd.ActiveDocument.Bookmarks("CidadeDevedor1").Range.InsertAfter cidadeDevedor
    wd.ActiveDocument.Bookmarks("UFDevedor1").Range.InsertAfter UFDevedor
    wd.ActiveDocument.Bookmarks("valorGlobal").Range.InsertAfter Format(CDbl(valorGlobal), "#,###.00")
    wd.ActiveDocument.Bookmarks("valorGlobal_extenso").Range.InsertAfter Extenso_Valor(CDbl(valorGlobal))
    wd.ActiveDocument.Bookmarks("taxaDesagio").Range.InsertAfter taxaDesagio
    wd.ActiveDocument.Bookmarks("taxaDesagio_extenso").Range.InsertAfter Extenso_Taxa(CDbl(taxaDesagio))
    wd.ActiveDocument.Bookmarks("precoCessao").Range.InsertAfter Format(CDbl(precoCessao), "#,###.00")
    wd.ActiveDocument.Bookmarks("precoCessao_extenso").Range.InsertAfter Extenso_Valor(CDbl(precoCessao))
    wd.ActiveDocument.Bookmarks("TCO").Range.InsertAfter TCO
    wd.ActiveDocument.Bookmarks("TTR").Range.InsertAfter TTR
    wd.ActiveDocument.Bookmarks("CC_Credito_Agencia").Range.InsertAfter CC_Credito_Ag
    wd.ActiveDocument.Bookmarks("CC_Credito_Conta").Range.InsertAfter CC_Credito_Favorecido
    wd.ActiveDocument.Bookmarks("CC_Credito_Banco").Range.InsertAfter CC_Credito_Banco
    wd.ActiveDocument.Bookmarks("dtHora_cidade1").Range.InsertAfter StrConv(cidadeFornecedor, vbUpperCase)
    wd.ActiveDocument.Bookmarks("dtHora_data1").Range.InsertAfter _
    StrConv(Format(dataOper, "dd") & " de " & Format(dataOper, "mmmm") & " de " & Format(dataOper, "yyyy"), vbLowerCase)

    'Preenchimento dos dados - Anexo 1
    wd.ActiveDocument.Bookmarks("numTermo2").Range.InsertAfter numTermo
    wd.ActiveDocument.Bookmarks("nomeFornecedor2").Range.InsertAfter nomeFornecedor
    wd.ActiveDocument.Bookmarks("CNPJFornecedor2").Range.InsertAfter CNPJFornecedor
    wd.ActiveDocument.Bookmarks("EnderecoFornecedor2").Range.InsertAfter EnderecoFornecedor
    wd.ActiveDocument.Bookmarks("BairroFornecedor2").Range.InsertAfter BairroFornecedor
    wd.ActiveDocument.Bookmarks("CidadeFornecedor2").Range.InsertAfter cidadeFornecedor
    wd.ActiveDocument.Bookmarks("UFFornecedor2").Range.InsertAfter UFFornecedor
    wd.ActiveDocument.Bookmarks("nomeDevedor2").Range.InsertAfter nomeDevedor
    wd.ActiveDocument.Bookmarks("CNPJDevedor2").Range.InsertAfter CNPJDevedor
    wd.ActiveDocument.Bookmarks("EnderecoDevedor2").Range.InsertAfter EnderecoDevedor
    wd.ActiveDocument.Bookmarks("BairroDevedor2").Range.InsertAfter BairroDevedor
    wd.ActiveDocument.Bookmarks("CidadeDevedor2").Range.InsertAfter cidadeDevedor
    wd.ActiveDocument.Bookmarks("UFDevedor2").Range.InsertAfter UFDevedor
    wd.ActiveDocument.Bookmarks("dtHora_cidade2").Range.InsertAfter StrConv(cidadeFornecedor, vbUpperCase)
    wd.ActiveDocument.Bookmarks("dtHora_data2").Range.InsertAfter _
    StrConv(Format(dataOper, "dd") & " de " & Format(dataOper, "mmmm") & " de " & Format(dataOper, "yyyy"), vbLowerCase)

    'Captura de dados das notas
    AbrirBDTermosFIAT
    SQL = "SELECT * FROM tmpDadosNotas WHERE val(Cod_Oper) = val(" & numOper & ")" & " ORDER BY Codigo"
    Set qryDadosNota = BDTermosFIAT.OpenRecordset(SQL)

    wd.ActiveDocument.Bookmarks("tabelaNotas").Range.Select
    Dim linhas As Integer: linhas = 2
    Dim tabelaNotas As Table

    Set tabelaNotas = wd.ActiveDocument.Tables(9)

    While Not qryDadosNota.EOF

        numNota = qryDadosNota!Compromisso
        tabelaNotas.Cell(linhas, 1).Range.Text = numNota
        valorNominal = Format(CDbl(qryDadosNota!Valor_Nom), "#,###.00")
        tabelaNotas.Cell(linhas, 2).Range.Text = valorNominal
        dtVenc = qryDadosNota!Data_Venc
        tabelaNotas.Cell(linhas, 3).Range.Text = dtVenc

        qryDadosNota.MoveNext

        If Not qryDadosNota.EOF Then
            linhas = linhas + 1
            tabelaNotas.Rows.Add
        End If

    Wend

    qryDadosNota.Close
    Set qryDadosNota = Nothing

    BDTermosFIAT.Close
    Set BDTermosFIAT = Nothing

    'Finalização
    wd.ActiveDocument.Save
    wd.ActiveDocument.ExportAsFixedFormat _
    caminhoPDFRede & "\Termos\" & numOper & "_" & Format(CNPJFornecedor, "00000000000000") & _
    "_" & Format(dataOper, "ddmmyyyy"), wdExportFormatPDF
    wd.ActiveDocument.Close
    wd.Quit
    Set wd = Nothing

    DeleteFile "C:\Temp\Termo Cessão_" & numOper & ".docx"

End Sub

'Envio do termo
Public Sub enviarTermoFIAT(ByVal numOper As String)

    'Varíáveis
    Dim OutApp As Object
    Dim OutMail As Object
    Dim nomeEmpresa, CNPJEmpresa, mailFornecedor, dataOper As String
    Dim qryDadosMail As Recordset

    'Recuperação de dados
    'Captura de dados do termo e convenio
    AbrirBDTermosFIAT
    SQL = "SELECT * FROM tmpDadosTermo WHERE val(Cod_Oper) = val(" & numOper & ")"
    Set qryDadosMail = BDTermosFIAT.OpenRecordset(SQL)

    If Not IsNull(qryDadosMail!Nome_Fornecedor) Then nomeEmpresa = qryDadosMail!Nome_Fornecedor
    If Not IsNull(qryDadosMail!Cnpj_Fornecedor) Then CNPJEmpresa = Format(qryDadosMail!Cnpj_Fornecedor, "000000000000000")
    If Not IsNull(qryDadosMail!Email_Fornecedor) Then mailFornecedor = Format(qryDadosMail!Email_Fornecedor, "000000000000000")
    If Not IsNull(qryDadosMail!Data_op) Then dataOper = qryDadosMail!Data_op

    qryDadosMail.Close
    Set qryDadosMail = Nothing

    BDTermosFIAT.Close
    Set BDTermosFIAT = Nothing

    'Integração com o Outlook
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'De
    OutMail.SentOnBehalfOfName = "prpjcadconfirming@santander.com.br"

    'Para
    If Not IsNull(mailFornecedor) Or mailFornecedor <> "" Then
        OutMail.To = mailFornecedor
    End If

    'CC
    OutMail.cc = "lucrecio.silva@fiat.com.br; pedro.antunes@fiat.com.br; p000renatoliveira@prservicos.com.br"

    'BCC - Cópia para validação
    'OutMail.BCC = "dwmiranda@santander.com.br; ebrant@santander.com.br"

    'Assunto
    OutMail.Subject = "TERMO DE CESSÃO FIAT - " & nomeEmpresa & " - " & CNPJEmpresa

    'Anexos
    Dim myAttachments As Outlook.Attachments
    Set myAttachments = OutMail.Attachments

    Dim file1 As String
    file1 = caminhoPDFRede & "\Termos\" & numOper & "_" & Format(CNPJEmpresa, "00000000000000") & _
    "_" & Format(dataOper, "ddmmyyyy") & ".pdf"
    myAttachments.Add file1, olByValue, 1

    'Montagem da mensagem

    'Inicialização
    OutMail.HTMLBody = ""

    'Conteúdo email
    OutMail.HTMLBody = OutMail.HTMLBody & "<p>Srs, </p>" & "<p>Favor devolver o Termo de Cessão anexo assinado para " & _
    "lucrecio.silva@fiat.com.br" & _
    "<p>e pedro.antunes@fiat.com.br</p>" & _
    "<p>No caso de dúvidas, favor entrar em contato diretamente com a FIAT.</p>" & _
    "<p>Grato</p>"
    
    'Dados para log da mensagem
    logpara = OutMail.To
    logcopia = OutMail.cc

    'Envio da mensagem
    OutMail.Send

    'Registro de envio na tabela
    SQL = "INSERT INTO tblLogEnvio(numOper, dtHoraEnvio, emailPara, emailCopia, arquivoEnviado, user) VALUES (" & _
    "'" & numOper & "', #" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "#, '" & logpara & "', '" & logcopia & _
    "', '" & file1 & "', '" & nomeUser & "')"
    AbrirBDTermosFIAT
    BDTermosFIAT.Execute (SQL)
    BDTermosFIAT.Close
    Set BDTermosFIAT = Nothing

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

'Rotina Completa
Sub envioTermosFIAT()

    'Define o dia que será processado
    dia = Format(UltimoDiaUtil, "dd/mm/yyyy")

    'Importa arquivo de operações antecipadas do dia
    importarArqOpedDia (dia)

    'Importa a base de fornecedores do dia
    importarFornecedoresDiario (dia)

    'Separa as operações para o envio de termos
    verOperTermosFIAT

    'Gera os termos e faz o envio
    Dim qryOper As Recordset
    Dim codOper As String

    AbrirBDTermosFIAT
    SQL = "SELECT * FROM tmpDadosTermo"
    Set qryOper = BDTermosFIAT.OpenRecordset(SQL)

    While Not qryOper.EOF
        codOper = qryOper!cod_oper
        criarTermoFIAT (codOper)
        enviarTermoFIAT (codOper)
        qryOper.MoveNext
    Wend

    qryOper.Close
    Set qryOper = Nothing

    On Error Resume Next
    BDTermosFIAT.Close
    Set BDTermosFIAT = Nothing
    
    Debug.Print "Fim do envio de Termos Fiat " & Now

End Sub
