'MDL_ImportarVInculo'

Option Compare Database
Sub ImportarVinculo()
    
    Dim TbDados As Recordset, TbArq As Recordset
    Dim linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
    Dim Dataarq As Date, DataVenNF As Date, Data_op As Date
    Dim FSO As New FileSystemObject, arq As File, File As String, Contador As String

        'alterar data pra rodar robô
        'DataPesq = "09/09/2020"
        DataPesq = Forms!FrmRelatorios!DataArquivo.Value
        DataPesq2 = Format(DataPesq, "DDMMYY")
        DataPesq1 = Format(DataPesq, "DD/MM/YYYY")
        DataPesq3 = Format(DataPesq, "mm/dd/yyyy")
        
    
    'File = "\\saont46\apps2\Confirming\ArquivosYC\ARQOPED" & DataPesq2 & ".TXT"
    ''Novo Caminho
    File = "\\fscore02\apps2\Confirming\ArquivosYC\ARQOPED" & DataPesq2 & ".TXT"

                    
    If (Dir(File) <> "") Then
        
        'Copiar Arquivo para base local
        Call CopiarBaseToMaquina("DOWNLOAD")
        
        'Abrir bando de dados local
        Call AbrirBDLocal

            'Verificação se a data já foi importada - Emerson 19/03/2018
            Dim testeImport As Recordset
            SQL = "SELECT * FROM tblArqOped WHERE Data_op = #" & DataPesq3 & "#"
            Set testeImport = BDRELocal.OpenRecordset(SQL)
            If testeImport.EOF Then

                FSO.CopyFile File, "\\saont46\apps2\Confirming\PROJETORELATORIOS\ARQRESERVA\", True
    
                  Call ReplaceEXE("\\saont46\apps2\Confirming\PROJETORELATORIOS\ARQRESERVA\ARQOPED" & DataPesq2 & ".TXT")
    
                    Call TratarArquivoDeOperacoes("\\saont46\apps2\Confirming\PROJETORELATORIOS\ARQRESERVA\ARQOPED" & DataPesq2 & ".TXT")
    
                       Set arq = FSO.GetFile(File)
                        
                        DataCri = arq.DateCreated: DataCri = Format(DataCri, "DD/MM/YYYY")
    
                           Set TbDados = BDRELocal.OpenRecordset("TblArqoped", dbOpenDynaset)
                           Set TbArq = BDRELocal.OpenRecordset("ARQ", dbOpenDynaset)
                            
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
                        Set BDRELocal = Nothing
                'break
            End If
            
            'Excluir da base os registros cujo Cod_Oper = 000019021017265 - Solicitação Aloísio/George - Adilson C. S. 03/04/2019
            'Abrir bando de dados local
            Call AbrirBDLocal
            BDRELocal.Execute "Delete * From TblArqoped Where Cod_Oper = '000019021017265'"
            BDRELocal.Close
            Set BDRELocal = Nothing
            
            'MsgBox "Importação Realizada com Sucesso ! "
            
            'Copiar Arquivo para base rede após final da importação - Emerson 01/03/2018
            Call CopiarBaseToMaquina("UPLOAD")
            
            Forms!FrmRelatorios!LbValidaAtivo.Caption = "ENVIADO"
            Call Periodicidade
        
            'MsgBox "Importado."
    End If
Fim:

End Sub
