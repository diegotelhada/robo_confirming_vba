'MDL_VolumetriaDiaria'

Option Compare Database
Global DataPesq1 As String
Global DataPesqAnexo As String
Global Nome As String
'==================================================================================================================================================================
'NOVA VERSÃO DA VOLUMETRIA DIARIA
'MARCELO HENRIQUE DE SOUZA
'29/03/2017 - 10:19:25
'==================================================================================================================================================================
Function AtualizarTabelasVolumetrias(DataPesq)
        
        Dim TbDados As Recordset
        
        'Limpar Tabela de Convenios
        BDVolumetria.Execute ("DELETE Tbl_Convenios.* FROM Tbl_Convenios;")
        
        'Limpar Tabela de Volumetria do Dia - Operações
        BDVolumetria.Execute ("DELETE Tbl_VolumetriaDiaria.* FROM Tbl_VolumetriaDiaria;")
                
        'Limpar Tabela de Volumetria do Dia - Fornecedores
        BDVolumetria.Execute ("DELETE Tbl_VolumetriaFornecedores.* FROM Tbl_VolumetriaFornecedores;")
        
        'Inserir na Tabela Convenios
        BDVolumetria.Execute ("INSERT INTO Tbl_Convenios ( [DT ATUAL  ], BCO, AGENCIA, NR_CONVENIO, [PRODUTO ], SUBPRODU, [BCO PREM], [AG# PREM], [CTA PREM    ], ETCA, [DT CAD#   ], [DT ULT#MOV], [NOME CONVENIO                 ], [TP PESS ], [CPF/CNPJ       ], [CONTATO                       ], DDD1, TELEFONE1, RAML1, DDD2, TELEFONE2, RAML2, DFAX, TELEFONE3, [CTRL# REME#], [FIEL DPOS], [RETORNO LIQUI# ], WEB, [PRE APROV#], [IND REMU CLIE            ], [EMAIL                                                           ], [MOTIVO BLOQUEIO     ], [AMBIENTE      ], [TIP BLOQUEIO], [TIPO COMPROVANTE                                              ], [ENCARTE      ], [ENDERECO                ], [FORMA LIQUIDACAO                  ], [APUR PREMIO], [FORMA PGTO PREMIO      ], [RETORNO AGENDAM# ], [FRANCESINHA       ], [FLUXO OPERACIONAL         ], [MODALIDADE OPERACAO     ], NOTIFIC, [REMU PRE], [STATUS CONV#  ], [NR SEQ#REMU# ], [NR SEQ#RETN# ], [SPREAD       ]," _
        & " [PERC# AGEN FORNEC], [VLR TRAVA        ], [VLR UTILIZ#      ], [INDIC# TRAVA VALOR], [QTD ACUM MES ATUAL], [QTD ACUM PRIMEIRO MES], [QTD ACUM SEGUNDO MES], [QTD ACUM TERCEIRO MES], [VALOR ACUM MES ATUAL], [VALOR ACUM PRIMEIRO MES], [VALOR ACUM SEGUNDO MES], [VALOR ACUM TERCEIRO MES], [VALOR RECEITA MES ATUAL], [VALOR RECEITA PRIMEIRO MES], [VALOR RECEITA SEGUNDO MES]," _
        & " [VALOR RECEITA TERCEIRO MES], [VALOR BANCO MES ATUAL], [VALOR BANCO PRIMEIRO MES ], [VALOR BANCO SEGUNDO MES ], [VALOR BANCO TERCEIRO MES  ], [VALOR CLIENTE MES ATUAL], [VALOR CLIENTE PRIMEIRO MES], [VALOR CLIENTE SEGUNDO MES], [VALOR CLIENTE TERCEIRO MES], [NUM ULTIMO AVISO ENVIADO], [PERC REC], [PERC ATR], [PERC DES], [IND APRO OPE], [IND GPO CONV], [BCO AGR], [AGE AGR], [CONV AGRU   ], [DIAS REJE], [DT INI LO ], [PZ LO ], [DT VENC LO], [PZ LIM], [VL TRAVA MASS    ], [AG BUS], [CTA BU], [DIG BU], [USUARIO ], [DH ATUALIZ], [TIPO DE CONTROLE               ], [TIPO DE REMESSA      ], [ENCARTE CT], [PERC AJUSTE ], [IND# EQUALIZACAO CLIENTE   ], [PC SG EQUA], [IND#NOTIFICACAO NOTAS P/FORNEC  ], [FORMA DE CONSOLIDACAO], [GAROF       ]," _
        & " [PC FLEX], [QT CPRO GPO], [TP REJE CRP], [NRO DEAL GARANTIA], [IND# CONV# SOLD#], [IND# PRAZO DA VALIDADE GARANTIA], [TIPO PRAZO DA GARANTIA], [QTDE DE MESES DA GARANTIA], [DTA LIMITE PRAZO DA GARANTIA], [COD# MODALIDADE PPB], [TAXA OPER ARQ OPER FORM], [VALOR LIQUI ARQ OPER FORM], [VALOR PPB ARQ OPER FORM], [COMP# HIST# ARQ OPER FORM], [TAXA OPER# ARQ LIQUIDACAO], [VALOR LIQUI# ARQ LIQUIDACAO], [VALOR PPB ARQ LIQUIDACAO], [COMP# HIST# ARQ LIQUIDACAO] )" _
        & " SELECT Convenios.[DT ATUAL  ], Convenios.[BCO ], Convenios.[AG# ], Format([Convenios]![NR CONVENIO ],'000000000000') AS CONVENIO, Convenios.[PRODUTO ], Convenios.SUBPRODU, Convenios.[BCO PREM], Convenios.[AG# PREM], Convenios.[CTA PREM    ], Convenios.ETCA, Convenios.[DT CAD#   ], Convenios.[DT ULT#MOV], Convenios.[NOME CONVENIO                 ], Convenios.[TP PESS ], Convenios.[CPF/CNPJ       ], Convenios.[CONTATO                       ], Convenios.DDD1, Convenios.TELEFONE1, Convenios.RAML1, Convenios.DDD2, Convenios.TELEFONE2, Convenios.RAML2, Convenios.DFAX, Convenios.TELEFONE3, Convenios.[CTRL# REME#], Convenios.[FIEL DPOS], Convenios.[RETORNO LIQUI# ], Convenios.WEB, Convenios.[PRE APROV#], Convenios.[IND REMU CLIE            ]," _
        & " Convenios.[EMAIL                                                           ] , Convenios.[MOTIVO BLOQUEIO     ], Convenios.[AMBIENTE      ], Convenios.[TIP BLOQUEIO], Convenios.[TIPO COMPROVANTE                                              ], Convenios.[ENCARTE      ], Convenios.[ENDERECO                ]," _
        & " Convenios.[FORMA LIQUIDACAO                  ], Convenios.[APUR PREMIO], Convenios.[FORMA PGTO PREMIO      ], Convenios.[RETORNO AGENDAM# ], Convenios.[FRANCESINHA       ], Convenios.[FLUXO OPERACIONAL         ], Convenios.[MODALIDADE OPERACAO     ], Convenios.NOTIFIC, Convenios.[REMU PRE], Convenios.[STATUS CONV#  ], Convenios.[NR SEQ#REMU# ], Convenios.[NR SEQ#RETN# ], Convenios.[SPREAD       ], Convenios.[PERC# AGEN FORNEC], Convenios.[VLR TRAVA        ], Convenios.[VLR UTILIZ#      ], Convenios.[INDIC# TRAVA VALOR], Convenios.[QTD ACUM MES ATUAL], Convenios.[QTD ACUM PRIMEIRO MES], Convenios.[QTD ACUM SEGUNDO MES], Convenios.[QTD ACUM TERCEIRO MES], Convenios.[VALOR ACUM MES ATUAL]," _
        & " Convenios.[VALOR ACUM PRIMEIRO MES], Convenios.[VALOR ACUM SEGUNDO MES], Convenios.[VALOR ACUM TERCEIRO MES], Convenios.[VALOR RECEITA MES ATUAL], Convenios.[VALOR RECEITA PRIMEIRO MES], Convenios.[VALOR RECEITA SEGUNDO MES], Convenios.[VALOR RECEITA TERCEIRO MES], Convenios.[VALOR BANCO MES ATUAL]," _
        & " Convenios.[VALOR BANCO PRIMEIRO MES ], Convenios.[VALOR BANCO SEGUNDO MES ], Convenios.[VALOR BANCO TERCEIRO MES  ], Convenios.[VALOR CLIENTE MES ATUAL], Convenios.[VALOR CLIENTE PRIMEIRO MES], Convenios.[VALOR CLIENTE SEGUNDO MES], Convenios.[VALOR CLIENTE TERCEIRO MES], Convenios.[NUM ULTIMO AVISO ENVIADO], Convenios.[PERC REC], Convenios.[PERC ATR], Convenios.[PERC DES], Convenios.[IND APRO OPE], Convenios.[IND GPO CONV], Convenios.[BCO AGR], Convenios.[AGE AGR], Convenios.[CONV AGRU   ], Convenios.[DIAS REJE], Convenios.[DT INI LO ], Convenios.[PZ LO ], Convenios.[DT VENC LO], Convenios.[PZ LIM], Convenios.[VL TRAVA MASS    ], Convenios.[AG BUS], Convenios.[CTA BU], Convenios.[DIG BU], Convenios.[USUARIO ], Convenios.[DH ATUALIZ]," _
        & " Convenios.[TIPO DE CONTROLE               ], Convenios.[TIPO DE REMESSA      ], Convenios.[ENCARTE CT], Convenios.[PERC AJUSTE ], Convenios.[IND# EQUALIZACAO CLIENTE   ], Convenios.[PC SG EQUA], Convenios.[IND#NOTIFICACAO NOTAS P/FORNEC  ], Convenios.[FORMA DE CONSOLIDACAO]," _
        & " Convenios.[GAROF       ] , Convenios.[PC FLEX], Convenios.[QT CPRO GPO], Convenios.[TP REJE CRP], Convenios.[NRO DEAL GARANTIA], Convenios.[IND# CONV# SOLD#], Convenios.[IND# PRAZO DA VALIDADE GARANTIA], Convenios.[TIPO PRAZO DA GARANTIA], Convenios.[QTDE DE MESES DA GARANTIA], Convenios.[DTA LIMITE PRAZO DA GARANTIA], Convenios.[COD# MODALIDADE PPB], Convenios.[TAXA OPER ARQ OPER FORM], Convenios.[VALOR LIQUI ARQ OPER FORM], Convenios.[VALOR PPB ARQ OPER FORM], Convenios.[COMP# HIST# ARQ OPER FORM], Convenios.[TAXA OPER# ARQ LIQUIDACAO], Convenios.[VALOR LIQUI# ARQ LIQUIDACAO], Convenios.[VALOR PPB ARQ LIQUIDACAO], Convenios.[COMP# HIST# ARQ LIQUIDACAO] FROM Convenios;")
        
        'Inserir dados da volumetria do Dia - Operações
        BDVolumetria.Execute ("INSERT INTO Tbl_VolumetriaDiaria ( Area_Geral, SUBPRODU, Nome_Ancora, Agencia_Ancora, Convenio_Ancora, FIEL, PRE, Cod_Oper, Valor_op, Usuario ) SELECT TblSegmentos.Area_Geral, Tbl_Convenios.SUBPRODU, TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, Trim([Tbl_Convenios]![FIEL DPOS]) AS FIEL, Trim([Tbl_Convenios]![PRE APROV#]) AS PRE, TblArqoped.Cod_Oper, TblArqoped.Valor_op, TblArqoped.Usuario" _
        & " FROM (TblArqoped LEFT JOIN Tbl_Convenios ON (TblArqoped.Convenio_Ancora = Tbl_Convenios.NR_CONVENIO) AND (TblArqoped.Agencia_Ancora = Tbl_Convenios.AGENCIA)) LEFT JOIN TblSegmentos ON TblArqoped.Segmento = TblSegmentos.Segmento WHERE (((TblArqoped.Data_op) = #" & Format(DataPesq, "mm/dd/yyyy") & "#)) GROUP BY TblSegmentos.Area_Geral, Tbl_Convenios.SUBPRODU, TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, Trim([Tbl_Convenios]![FIEL DPOS]), Trim([Tbl_Convenios]![PRE APROV#]), TblArqoped.Cod_Oper, TblArqoped.Valor_op, TblArqoped.Usuario;")

        'Inserir dados da Volumetria do dia - Fornecedores
        'BDVolumetria.Execute ("INSERT INTO Tbl_VolumetriaFornecedores ( NomeSegmento, AG, CONVENIO, NOMECONVENIO, SUBPRODU, [CPF/CNPJ], NOMEFORNECEDOR, CONTAFORN, [TIPO CNTR FORN], [SIT FORNECEDOR], [MOTIVO BLOQUEIO] )" _
        & " SELECT TblClientesXSegmento.NomeSegmento, IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',[Tbl_Convenios]![AGE AGR],[Tbl_Convenios]![AGENCIA]) AS AG, IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',Format([Tbl_Convenios]![CONV AGRU],'000000000000'),[Tbl_Convenios]![NR_CONVENIO]) AS CONVENIO, ReplaceString(IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',BuscaNomeAgrupador([Tbl_Convenios]![AGE AGR],Format([Tbl_Convenios]![CONV AGRU],'000000000000')),BuscaNomeAgrupador([Tbl_Convenios]![AGENCIA],[Tbl_Convenios]![NR_CONVENIO]))) AS NOMECONVENIO, Tbl_Convenios.SUBPRODU, Tbl_Fornecedores.[CPF/CNPJ], ReplaceString([Tbl_Fornecedores]![RAZAO SOCIAL]) AS NOMEFORNECEDOR, BuscaNomeBanco(Left([Tbl_Fornecedores]![CONTA FORNECEDOR 1],5)) AS CONTAFORN, Tbl_Fornecedores.[TIPO CNTR FORN], Tbl_Fornecedores.[SIT FORNECEDOR], Tbl_Fornecedores.[MOTIVO BLOQUEIO]" _
        & " FROM (Tbl_Fornecedores INNER JOIN Tbl_Convenios ON (Tbl_Fornecedores.[NRO CONVENIO] = Tbl_Convenios.NR_CONVENIO) AND (Tbl_Fornecedores.AGEN = Tbl_Convenios.AGENCIA)) LEFT JOIN TblClientesXSegmento ON (Tbl_Fornecedores.AGEN = TblClientesXSegmento.Agencia_Ancora) AND (Tbl_Fornecedores.[NRO CONVENIO] = TblClientesXSegmento.Convenio_Ancora) WHERE (((Tbl_Fornecedores.LOGRADOURO)<>''))" _
        & " GROUP BY TblClientesXSegmento.NomeSegmento, IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',[Tbl_Convenios]![AGE AGR],[Tbl_Convenios]![AGENCIA]), IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',Format([Tbl_Convenios]![CONV AGRU],'000000000000'),[Tbl_Convenios]![NR_CONVENIO]), ReplaceString(IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',BuscaNomeAgrupador([Tbl_Convenios]![AGE AGR],Format([Tbl_Convenios]![CONV AGRU],'000000000000')),BuscaNomeAgrupador([Tbl_Convenios]![AGENCIA],[Tbl_Convenios]![NR_CONVENIO]))), Tbl_Convenios.SUBPRODU, Tbl_Fornecedores.[CPF/CNPJ], ReplaceString([Tbl_Fornecedores]![RAZAO SOCIAL]), BuscaNomeBanco(Left([Tbl_Fornecedores]![CONTA FORNECEDOR 1],5)), Tbl_Fornecedores.[TIPO CNTR FORN], Tbl_Fornecedores.[SIT FORNECEDOR], Tbl_Fornecedores.[MOTIVO BLOQUEIO]" _
        & " HAVING (((IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',Format([Tbl_Convenios]![CONV AGRU],'000000000000'),[Tbl_Convenios]![NR_CONVENIO]))<>'008500000035') AND ((ReplaceString(IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',BuscaNomeAgrupador([Tbl_Convenios]![AGE AGR],Format([Tbl_Convenios]![CONV AGRU],'000000000000')),BuscaNomeAgrupador([Tbl_Convenios]![AGENCIA],[Tbl_Convenios]![NR_CONVENIO])))) Not Like '*HUB*') AND ((Tbl_Fornecedores.[MOTIVO BLOQUEIO])<>'CADASTRO VIA ARQUIVO')) ORDER BY ReplaceString(IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',BuscaNomeAgrupador([Tbl_Convenios]![AGE AGR],Format([Tbl_Convenios]![CONV AGRU],'000000000000')),BuscaNomeAgrupador([Tbl_Convenios]![AGENCIA],[Tbl_Convenios]![NR_CONVENIO])));")
        
        'Deletar Contratos da Fiat
        BDVolumetria.Execute ("Delete Tbl_Fornecedores.[NRO CONVENIO] FROM Tbl_Fornecedores WHERE (((Tbl_Fornecedores.[NRO CONVENIO])='008500000035'));")
                
        'Deletar Cadastros via Arquivo
        BDVolumetria.Execute ("Delete Tbl_Fornecedores.[MOTIVO BLOQUEIO] FROM Tbl_Fornecedores WHERE (((Tbl_Fornecedores.[MOTIVO BLOQUEIO])='CADASTRO VIA ARQUIVO'));")
    
        'Deletar Cadastro sem endereço
        BDVolumetria.Execute ("Delete Tbl_Fornecedores.LOGRADOURO FROM Tbl_Fornecedores WHERE (((Tbl_Fornecedores.LOGRADOURO)=''));")
         
        'Atualizar Agencia e Convenio
        BDVolumetria.Execute ("UPDATE Tbl_Fornecedores INNER JOIN Tbl_Convenios ON (Tbl_Fornecedores.[NRO CONVENIO] = Tbl_Convenios.NR_CONVENIO) AND (Tbl_Fornecedores.AGEN = Tbl_Convenios.AGENCIA) SET Tbl_Fornecedores.AGEN = IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',[Tbl_Convenios]![AGE AGR],[Tbl_Convenios]![AGENCIA]), Tbl_Fornecedores.[NRO CONVENIO] = IIf(Left([Tbl_Convenios]![IND GPO CONV],8)='AGRUPADO',Format([Tbl_Convenios]![CONV AGRU],'000000000000'),[Tbl_Convenios]![NR_CONVENIO]);")
      
        'Atualizar nome do Convenio, CNPJ e SubProduto
        BDVolumetria.Execute ("UPDATE Tbl_Fornecedores INNER JOIN Tbl_Convenios ON (Tbl_Fornecedores.[NRO CONVENIO] = Tbl_Convenios.NR_CONVENIO) AND (Tbl_Fornecedores.AGEN = Tbl_Convenios.AGENCIA) SET Tbl_Fornecedores.NOMECONVENIO = [Tbl_Convenios]![NOME CONVENIO                 ], Tbl_Fornecedores.SUBPRODUTO = [Tbl_Convenios]![SUBPRODU], Tbl_Fornecedores.CNPJANCORA = [Tbl_Convenios]![CPF/CNPJ       ];")

        'Atualiza Segmento na tabela fornecedor
        BDVolumetria.Execute ("UPDATE Tbl_Fornecedores LEFT JOIN TblClientesXSegmento ON (Tbl_Fornecedores.AGEN = TblClientesXSegmento.Agencia_Ancora) AND (Tbl_Fornecedores.[NRO CONVENIO] = TblClientesXSegmento.Convenio_Ancora) SET Tbl_Fornecedores.Segmento = [TblClientesXSegmento]![NomeSegmento];")
    
        'Deletar convenios HUB
        BDVolumetria.Execute ("DELETE Tbl_Fornecedores.NOMECONVENIO FROM Tbl_Fornecedores WHERE (((Tbl_Fornecedores.NOMECONVENIO) Like '*HUB*'));")
                       
        'Voluemtria
        BDVolumetria.Execute ("INSERT INTO Tbl_VolumetriaFornecedores ( NomeSegmento, CNPJANCORA, AG, CONVENIO, NOMECONVENIO, SUBPRODU, [CPF/CNPJ], NOMEFORNECEDOR, [TIPO CNTR FORN], [SIT FORNECEDOR], [MOTIVO BLOQUEIO], CONTAFORN ) SELECT Tbl_Fornecedores.SEGMENTO, Tbl_Fornecedores.CNPJANCORA, Tbl_Fornecedores.AGEN, Tbl_Fornecedores.[NRO CONVENIO], Tbl_Fornecedores.NOMECONVENIO, Tbl_Fornecedores.SUBPRODUTO, Tbl_Fornecedores.[CPF/CNPJ], Tbl_Fornecedores.[RAZAO SOCIAL], Tbl_Fornecedores.[TIPO CNTR FORN], Tbl_Fornecedores.[SIT FORNECEDOR], Tbl_Fornecedores.[MOTIVO BLOQUEIO], Left([Tbl_Fornecedores]![CONTA FORNECEDOR 1],5) AS CONTA" _
        & " FROM Tbl_Fornecedores GROUP BY Tbl_Fornecedores.SEGMENTO, Tbl_Fornecedores.CNPJANCORA, Tbl_Fornecedores.AGEN, Tbl_Fornecedores.[NRO CONVENIO], Tbl_Fornecedores.NOMECONVENIO, Tbl_Fornecedores.SUBPRODUTO, Tbl_Fornecedores.[CPF/CNPJ], Tbl_Fornecedores.[RAZAO SOCIAL], Tbl_Fornecedores.[TIPO CNTR FORN], Tbl_Fornecedores.[SIT FORNECEDOR], Tbl_Fornecedores.[MOTIVO BLOQUEIO], Left([Tbl_Fornecedores]![CONTA FORNECEDOR 1],5) ORDER BY Tbl_Fornecedores.NOMECONVENIO;")
        
        'Tirar Caraceres Especiais
        BDVolumetria.Execute ("UPDATE Tbl_VolumetriaFornecedores SET Tbl_VolumetriaFornecedores.NOMECONVENIO = ReplaceString([Tbl_VolumetriaFornecedores]![NOMECONVENIO]), Tbl_VolumetriaFornecedores.NOMEFORNECEDOR = ReplaceString([Tbl_VolumetriaFornecedores]![NOMEFORNECEDOR]), Tbl_VolumetriaFornecedores.[MOTIVO BLOQUEIO] = ReplaceString([Tbl_VolumetriaFornecedores]![MOTIVO BLOQUEIO]);")
    
        'Atualizar Segmentos em Branco
        BDVolumetria.Execute ("UPDATE Tbl_VolumetriaFornecedores LEFT JOIN DBM ON Tbl_VolumetriaFornecedores.CNPJANCORA = DBM.CNPJ SET Tbl_VolumetriaFornecedores.NomeSegmento = IIf([DBM]![Segmento]='GB&M','GBM',[DBM]![Segmento]) WHERE (((Tbl_VolumetriaFornecedores.NomeSegmento) Is Null));")
        
        'Atualizar nome do banco
        BDVolumetria.Execute ("UPDATE Tbl_VolumetriaFornecedores INNER JOIN Tbl_CodigoBanco ON Tbl_VolumetriaFornecedores.CONTAFORN = Tbl_CodigoBanco.Codigo SET Tbl_VolumetriaFornecedores.CONTAFORN = [Tbl_CodigoBanco]![Nome];")
        
End Function
Sub VolumetriaDiariaV2()
        
    Dim ObjExcel, ObjPlan1Excel As Object, ObjPlan0Excel As Object
    Dim ObjPlan2Excel As Object, ObjPlan3Excel As Object, ObjPlan4Excel As Object
    Dim TbAnalitico As Recordset, TbDados As Recordset
            
        'Pesquisar Ultimo dia Uitil
        DataPesq = UltimoDiaUtil()
        
        'Abrir Banco de dados
        Call AbrirBDVolumetria
        Call AbrirBDTermos
        
        'Atualizar Tabelas volumetria diaria
        Call AtualizarTabelasVolumetrias(DataPesq)

        'CAMINHO DA MASCARA PADRÃO
        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
        
        'CRIA EXCEL E CHAMA A MASCARA / SHEET
        Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "Volumetria_DiariaV3.xlsx", ReadOnly:=True
            'ObjExcel.Visible = True
        Set ObjPlan0Excel = ObjExcel.Worksheets("Resumo")
        Set ObjPlan1Excel = ObjExcel.Worksheets("Analitico_Operacoes")
        Set ObjPlan2Excel = ObjExcel.Worksheets("Analitico_Fornecedores")
        Set ObjPlan3Excel = ObjExcel.Worksheets("Analitico_Emails")
        Set ObjPlan4Excel = ObjExcel.Worksheets("Analitico_Termos")
            ObjPlan1Excel.Select
                
        '========   VOLUMETRIA DE OPERAÇÕES ==========='
            Set TbAnalitico = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.SUBPRODU, Tbl_VolumetriaDiaria.Nome_Ancora, Tbl_VolumetriaDiaria.Agencia_Ancora, Tbl_VolumetriaDiaria.Convenio_Ancora, Tbl_VolumetriaDiaria.FIEL, Tbl_VolumetriaDiaria.PRE, Tbl_VolumetriaDiaria.Cod_Oper, Tbl_VolumetriaDiaria.Valor_op, IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='T','MESA',IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='P','PORTAL',IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='A','ARQUIVO',IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='B','ARQUIVO',[Tbl_VolumetriaDiaria]![Usuario])))) AS [USER]" _
            & " FROM Tbl_VolumetriaDiaria GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.SUBPRODU, Tbl_VolumetriaDiaria.Nome_Ancora, Tbl_VolumetriaDiaria.Agencia_Ancora, Tbl_VolumetriaDiaria.Convenio_Ancora, Tbl_VolumetriaDiaria.FIEL, Tbl_VolumetriaDiaria.PRE, Tbl_VolumetriaDiaria.Cod_Oper, Tbl_VolumetriaDiaria.Valor_op, IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='T','MESA',IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='P','PORTAL',IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='A','ARQUIVO',IIf(Left([Tbl_VolumetriaDiaria]![Usuario],1)='B','ARQUIVO',[Tbl_VolumetriaDiaria]![Usuario]))));", dbOpenDynaset)
                
                If TbAnalitico.EOF = False Then
                    
                    'Pegar ultima linha da consulta
                    TbAnalitico.MoveLast: UltimaLinha = TbAnalitico.RecordCount + 1: TbAnalitico.MoveFirst: linha = 2
                    
                    'JOGA OS DADOS DA CONSULTA NO ARQUIVO EXCEL
                    ObjPlan1Excel.Range("A2").CopyFromRecordset TbAnalitico
                    
                    'Fortmatando Analitico
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":J" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan1Excel.Range("I" & linha & ":I" & UltimaLinha).Style = "Currency"
                    ObjPlan1Excel.Columns("A:j").Select
                    ObjPlan1Excel.Columns.AutoFit
                        
                    'Selecionar sheet Resumo
                    ObjPlan0Excel.Select
                            
                 '==== 8510 =====
                     'Pré Aprovado por Segmento
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8510)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE HAVING (((Tbl_VolumetriaDiaria.PRE)='S'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D6") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E6") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G6") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H6") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J6") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K6") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
                         
                     'Portal Cash
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op, IIf([Tbl_VolumetriaDiaria]![Usuario]='PORTAL','PORTAL','MESA') AS Usuario FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8510)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, IIf([Tbl_VolumetriaDiaria]![Usuario]='PORTAL','PORTAL','MESA') HAVING (((Tbl_VolumetriaDiaria.PRE)='N') AND ((IIf([Tbl_VolumetriaDiaria]![Usuario]='PORTAL','PORTAL','MESA'))='PORTAL'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D7") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E7") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G7") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H7") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J7") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K7") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
            
                     'Back Office
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op, IIf([Tbl_VolumetriaDiaria]![Usuario]='PORTAL','PORTAL','MESA') AS Usuario FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8510)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, IIf([Tbl_VolumetriaDiaria]![Usuario]='PORTAL','PORTAL','MESA') HAVING (((Tbl_VolumetriaDiaria.PRE)='N') AND ((IIf([Tbl_VolumetriaDiaria]![Usuario]='PORTAL','PORTAL','MESA'))='MESA'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D8") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E8") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G8") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H8") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J8") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K8") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
                         
                 '==== 8520 =====
                     
                     'Pré Aprovado
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8520)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE HAVING (((Tbl_VolumetriaDiaria.PRE)='S'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D11") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E11") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G11") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H11") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J11") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K11") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
                 
                     'Via Arquivo
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op  FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8520)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE HAVING (((Tbl_VolumetriaDiaria.PRE)='N'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D12") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E12") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G12") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H12") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J12") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K12") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
                 
                 '==== 8530 =====
                 
                     'Pré Aprovado
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8530)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE HAVING (((Tbl_VolumetriaDiaria.PRE)='S'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D15") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E15") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G15") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H15") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J15") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K15") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
                 
                     'Via Arquivo
                     Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE, Count(Tbl_VolumetriaDiaria.Cod_Oper) AS ContarDeCod_Oper, Sum(Tbl_VolumetriaDiaria.Valor_op) AS SomaDeValor_op FROM Tbl_VolumetriaDiaria WHERE (((Tbl_VolumetriaDiaria.SUBPRODU) = 8530)) GROUP BY Tbl_VolumetriaDiaria.Area_Geral, Tbl_VolumetriaDiaria.PRE HAVING (((Tbl_VolumetriaDiaria.PRE)='N'));", dbOpenDynaset)
                         Do While TbDados.EOF = False
                                 If TbDados!Area_Geral = "GBM" Then: ObjPlan0Excel.Range("D16") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("E16") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "CORPORATE" Then: ObjPlan0Excel.Range("G16") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("H16") = TbDados!SomaDeValor_op
                                 If TbDados!Area_Geral = "VAREJO" Then: ObjPlan0Excel.Range("J16") = TbDados!ContarDeCod_Oper: ObjPlan0Excel.Range("K16") = TbDados!SomaDeValor_op
                             TbDados.MoveNext
                         Loop
                         
                    'Formatar Resumo
                    ObjPlan0Excel.Columns("B:N").Select
                    ObjPlan0Excel.Columns.AutoFit
                    ObjPlan0Excel.Range("C19").Select
                End If
        '========   VOLUMETRIA DE FORNECEDORES ==========='
            ObjPlan2Excel.Select
            
            'Excluir convenio ADF FIAT 3377 008500000035
            'Excluir convenios HUB da Nestlé
            'Excluir Fornecedores com endereço em branco
            'Excluir Fornecedores cadastrado via Arquivo
  
            Set TbAnalitico = BDVolumetria.OpenRecordset("SELECT Tbl_VolumetriaFornecedores.NomeSegmento, Tbl_VolumetriaFornecedores.AG, Tbl_VolumetriaFornecedores.CONVENIO, Tbl_VolumetriaFornecedores.NOMECONVENIO, Tbl_VolumetriaFornecedores.SUBPRODU, Tbl_VolumetriaFornecedores.[CPF/CNPJ], Tbl_VolumetriaFornecedores.NOMEFORNECEDOR, Tbl_VolumetriaFornecedores.CONTAFORN, Tbl_VolumetriaFornecedores.[TIPO CNTR FORN], Tbl_VolumetriaFornecedores.[SIT FORNECEDOR], Tbl_VolumetriaFornecedores.[MOTIVO BLOQUEIO]" _
            & " FROM Tbl_VolumetriaFornecedores GROUP BY Tbl_VolumetriaFornecedores.NomeSegmento, Tbl_VolumetriaFornecedores.AG, Tbl_VolumetriaFornecedores.CONVENIO, Tbl_VolumetriaFornecedores.NOMECONVENIO, Tbl_VolumetriaFornecedores.SUBPRODU, Tbl_VolumetriaFornecedores.[CPF/CNPJ], Tbl_VolumetriaFornecedores.NOMEFORNECEDOR, Tbl_VolumetriaFornecedores.CONTAFORN, Tbl_VolumetriaFornecedores.[TIPO CNTR FORN], Tbl_VolumetriaFornecedores.[SIT FORNECEDOR], Tbl_VolumetriaFornecedores.[MOTIVO BLOQUEIO] ORDER BY Tbl_VolumetriaFornecedores.NOMECONVENIO;", dbOpenDynaset)
                                
                If TbAnalitico.EOF = False Then
                
                    'Pegar ultima linha da consulta
                    TbAnalitico.MoveLast: UltimaLinha = TbAnalitico.RecordCount + 1: TbAnalitico.MoveFirst: linha = 2
                    
                    'JOGA OS DADOS DA CONSULTA NO ARQUIVO EXCEL
                    ObjPlan2Excel.Range("A2").CopyFromRecordset TbAnalitico
                    
                    'Fortmatando Analitico
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan2Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan2Excel.Columns("A:K").Select
                    ObjPlan2Excel.Columns.AutoFit
                        
                    'Selecionar sheet Resumo
                    ObjPlan0Excel.Select
                    
                 '==== GCB =====
                Set TbDados = BDVolumetria.OpenRecordset("SELECT QryTbl_VolumetriaFornecedores.NomeSegmento, Sum(IIf([QryTbl_VolumetriaFornecedores]![CONTAFORN] Like '*Santander*',IIf([QryTbl_VolumetriaFornecedores]![TIPO CNTR FORN]='F',1,0))) AS CORRENTISTAFISICO, Sum(IIf([QryTbl_VolumetriaFornecedores]![CONTAFORN] Like '*Santander*',IIf([QryTbl_VolumetriaFornecedores]![TIPO CNTR FORN]='V',1,0))) AS CORRENTISTAVIRTUAL, Sum(IIf([QryTbl_VolumetriaFornecedores]![CONTAFORN] Not Like '*Santander*' Or IsNull([QryTbl_VolumetriaFornecedores]![CONTAFORN]),IIf([QryTbl_VolumetriaFornecedores]![TIPO CNTR FORN]='F',1,0),0)) AS NAOCORRENTISTAFISICO, Sum(IIf([QryTbl_VolumetriaFornecedores]![CONTAFORN] Not Like '*Santander*' Or IsNull([QryTbl_VolumetriaFornecedores]![CONTAFORN]),IIf([QryTbl_VolumetriaFornecedores]![TIPO CNTR FORN]='V',1,0),0)) AS NAOCORRENTISTAVIRTUAL FROM QryTbl_VolumetriaFornecedores GROUP BY QryTbl_VolumetriaFornecedores.NomeSegmento;", dbOpenDynaset)
                    Do While TbDados.EOF = False
                            If TbDados!NomeSegmento = "GBM" Then
                                
                                ObjPlan0Excel.Range("D26") = TbDados!CORRENTISTAFISICO
                                ObjPlan0Excel.Range("E26") = TbDados!CORRENTISTAVIRTUAL
                                
                                ObjPlan0Excel.Range("D27") = TbDados!NAOCORRENTISTAFISICO
                                ObjPlan0Excel.Range("E27") = TbDados!NAOCORRENTISTAVIRTUAL

                            ElseIf TbDados!NomeSegmento = "CORPORATE" Then
                                
                                ObjPlan0Excel.Range("G26") = TbDados!CORRENTISTAFISICO
                                ObjPlan0Excel.Range("H26") = TbDados!CORRENTISTAVIRTUAL
                                
                                ObjPlan0Excel.Range("G27") = TbDados!NAOCORRENTISTAFISICO
                                ObjPlan0Excel.Range("H27") = TbDados!NAOCORRENTISTAVIRTUAL
                            
                            
                            ElseIf TbDados!NomeSegmento = "VAREJO" Then
                            
                                ObjPlan0Excel.Range("J26") = TbDados!CORRENTISTAFISICO
                                ObjPlan0Excel.Range("K26") = TbDados!CORRENTISTAVIRTUAL
                                
                                ObjPlan0Excel.Range("J27") = TbDados!NAOCORRENTISTAFISICO
                                ObjPlan0Excel.Range("K27") = TbDados!NAOCORRENTISTAVIRTUAL
                            
                            End If
                            
                        TbDados.MoveNext
                    Loop
                End If
                
        '========   VOLUMETRIA DE TERMOS  ==========='
            ObjPlan3Excel.Select
            Set TbAnalitico = BDTermos.OpenRecordset("SELECT Format(Left([Tbl_Email]![DataHoraRec],2) & '/' & Mid([Tbl_Email]![DataHoraRec],3,2) & '/' & Mid([Tbl_Email]![DataHoraRec],5,4),'dd/mm/yyyy') AS Data, Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2) AS Hora, Tbl_Email.Remetente, Tbl_Email.Assunto FROM Tbl_Email WHERE (((Format(Left([Tbl_Email]![DataHoraRec], 2) & '/' & Mid([Tbl_Email]![DataHoraRec], 3, 2) & '/' & Mid([Tbl_Email]![DataHoraRec], 5, 4), 'dd/mm/yyyy')) ='" & DataPesq & "')) ORDER BY Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2);", dbOpenDynaset)
                                
                If TbAnalitico.EOF = False Then
                
                    'Pegar ultima linha da consulta
                    TbAnalitico.MoveLast: UltimaLinha = TbAnalitico.RecordCount + 1: TbAnalitico.MoveFirst: linha = 2
                    
                    'JOGA OS DADOS DA CONSULTA NO ARQUIVO EXCEL
                    ObjPlan3Excel.Range("A2").CopyFromRecordset TbAnalitico
                    
                    'Fortmatando Analitico
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan3Excel.Range("A" & linha & ":D" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan3Excel.Columns("A:D").Select
                    ObjPlan3Excel.Columns.AutoFit
                    
                End If
        
            ObjPlan4Excel.Select
            Set TbAnalitico = BDTermos.OpenRecordset("SELECT Format([Tbl_Log]![DataHora_Execucao],'dd/mm/yyyy') AS Data, Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss') AS HORA, Tbl_Log.Remetente, Tbl_Log.Assunto, Tbl_Log.Acao, Tbl_Log.DataHora_Recepcao, Tbl_Log.Usuario FROM Tbl_Log WHERE (((Format([Tbl_Log]![DataHora_Execucao], 'dd/mm/yyyy')) ='" & DataPesq & "')) ORDER BY Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss');", dbOpenDynaset)
                                
                If TbAnalitico.EOF = False Then
                
                    'Pegar ultima linha da consulta
                    TbAnalitico.MoveLast: UltimaLinha = TbAnalitico.RecordCount + 1: TbAnalitico.MoveFirst: linha = 2
                    
                    'JOGA OS DADOS DA CONSULTA NO ARQUIVO EXCEL
                    ObjPlan4Excel.Range("A2").CopyFromRecordset TbAnalitico
                    
                    'Fortmatando Analitico
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan4Excel.Range("A" & linha & ":G" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan4Excel.Columns("A:G").Select
                    ObjPlan4Excel.Columns.AutoFit
                    
                End If
                
            'Selecionar sheet Resumo
            ObjPlan0Excel.Select
                
                'Emails recebidos
                Set TbDados = BDTermos.OpenRecordset("SELECT Format(Left([Tbl_Email]![DataHoraRec],2) & '/' & Mid([Tbl_Email]![DataHoraRec],3,2) & '/' & Mid([Tbl_Email]![DataHoraRec],5,4),'dd/mm/yyyy') AS Data, Sum(IIf(Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2)<='12:00:00',1,0)) AS ATE12, Sum(IIf(Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2)>'12:00:00' And Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2)<='15:00:00',1,0)) AS 12a15, Sum(IIf(Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2)>'15:00:00' And" _
                & " Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2)<='18:00:00',1,0)) AS 15a18, Sum(IIf(Mid([Tbl_Email]![DataHoraRec],10,2) & ':' & Mid([Tbl_Email]![DataHoraRec],12,2) & ':' & Right([Tbl_Email]![DataHoraRec],2)>'18:00:00',1,0)) AS APOS18 FROM Tbl_Email GROUP BY Format(Left([Tbl_Email]![DataHoraRec],2) & '/' & Mid([Tbl_Email]![DataHoraRec],3,2) & '/' & Mid([Tbl_Email]![DataHoraRec],5,4),'dd/mm/yyyy') HAVING (((Format(Left([Tbl_Email]![DataHoraRec],2) & '/' & Mid([Tbl_Email]![DataHoraRec],3,2) & '/' & Mid([Tbl_Email]![DataHoraRec],5,4),'dd/mm/yyyy'))='" & DataPesq & "'));", dbOpenDynaset)
                    If TbDados.EOF = False Then
                        ObjPlan0Excel.Range("D35") = TbDados!ATE12
                        ObjPlan0Excel.Range("D36") = TbDados![12a15]
                        ObjPlan0Excel.Range("D37") = TbDados![15a18]
                        ObjPlan0Excel.Range("D38") = TbDados!APOS18
                    End If
                    
                'Termos Baixados/Pendentes
                Set TbDados = BDTermos.OpenRecordset("SELECT Tbl_Log.Acao, Sum(IIf(Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss')<='12:00:00',1,0)) AS ATE12, Sum(IIf(Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss')>'12:00:00' And Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss')<='15:00:00',1,0)) AS 12a15, Sum(IIf(Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss')>'15:00:00' And Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss')<='18:00:00',1,0)) AS 15a18, Sum(IIf(Format([Tbl_Log]![DataHora_Execucao],'hh:nn:ss')>'18:00:00',1,0)) AS Apos18 FROM Tbl_Log WHERE (((Format([Tbl_Log]![DataHora_Execucao], 'dd/mm/yyyy')) ='" & DataPesq & "')) GROUP BY Tbl_Log.Acao;", dbOpenDynaset)
                    Do While TbDados.EOF = False
                            If UCase(TbDados!Acao) = "BAIXA DE TERMO" Then
                                ObjPlan0Excel.Range("J35") = TbDados!ATE12
                                ObjPlan0Excel.Range("J36") = TbDados![12a15]
                                ObjPlan0Excel.Range("J37") = TbDados![15a18]
                                ObjPlan0Excel.Range("J38") = TbDados!APOS18
                            ElseIf UCase(TbDados!Acao) = "PENDÊNCIA" Then
                                ObjPlan0Excel.Range("G35") = TbDados!ATE12
                                ObjPlan0Excel.Range("G36") = TbDados![12a15]
                                ObjPlan0Excel.Range("G37") = TbDados![15a18]
                                ObjPlan0Excel.Range("G38") = TbDados!APOS18
                            End If
                        TbDados.MoveNext
                    Loop
                
                'Termos Enviados para o GDS
                Set TbDados = BDTermos.OpenRecordset("SELECT Format([Tbl_GDS]![DataHora_Execucao],'dd/mm/yyyy') AS Data, Sum(IIf(Format([Tbl_GDS]![DataHora_Execucao],'hh:nn:ss')<='12:00:00',1,0)) AS Ate12, Sum(IIf(Format([Tbl_GDS]![DataHora_Execucao],'hh:nn:ss')>'12:00:00' And Format([Tbl_GDS]![DataHora_Execucao],'hh:nn:ss')<='15:00:00',1,0)) AS 12a15, Sum(IIf(Format([Tbl_GDS]![DataHora_Execucao],'hh:nn:ss')>'15:00:00' And Format([Tbl_GDS]![DataHora_Execucao],'hh:nn:ss')<='18:00:00',1,0)) AS 15a18, Sum(IIf(Format([Tbl_GDS]![DataHora_Execucao],'hh:nn:ss')>'18:00:00',1,0)) AS apos18 FROM Tbl_GDS GROUP BY Format([Tbl_GDS]![DataHora_Execucao],'dd/mm/yyyy') HAVING (((Format([Tbl_GDS]![DataHora_Execucao],'dd/mm/yyyy'))='" & DataPesq & "'));", dbOpenDynaset)

                    If TbDados.EOF = False Then
                        ObjPlan0Excel.Range("G35") = ObjPlan0Excel.Range("G35") + TbDados!ATE12
                        ObjPlan0Excel.Range("G36") = ObjPlan0Excel.Range("G36") + TbDados![12a15]
                        ObjPlan0Excel.Range("G37") = ObjPlan0Excel.Range("G37") + TbDados![15a18]
                        ObjPlan0Excel.Range("G38") = ObjPlan0Excel.Range("G38") + TbDados!APOS18
                    End If

            Nome = "Volumetria_DiariaV3_" & Format(DataPesq, "DDMMYY")
            
            'SOBREPOR O ARQUIVO SE EXISTIR
            sFname = "\\saont46\apps2\\Confirming\PROJETORELATORIOS\VOLUMETRIA DIARIA\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
            
            'SALVA ARQUIVO
            ObjPlan0Excel.SaveAs FileName:="\\saont46\apps2\\Confirming\PROJETORELATORIOS\VOLUMETRIA DIARIA\" & Nome & ".xlsx"
            ObjExcel.activeworkbook.Close SaveChanges:=False
            ObjExcel.Quit

        'Função para enviar o email - Nao enviar mais o e-mail - Solicitaão Aloísio 27/09/2019
        'Call EnviarEmailVolumetria(Nome, DataPesq)

End Sub
Function EnviarEmailVolumetria(Nome, DataPesq)

    'ANEXO A SER ENVIADO
    File = "\\saont46\apps2\\Confirming\PROJETORELATORIOS\VOLUMETRIA DIARIA\" & Nome & ".xlsx"
           
    'EMAIL DESTINO / COPIA / LOGO
     EmailDestino = "apdmoraes@santander.com.br;hscruz@santander.com.br"
     EmailCopia = "jorge.junior@santander.com.br;"
     Assinatura = "\\saont46\apps2\\Confirming\Produto\Documentação\Documentação\BCDADOS_FORNECEDORES\Logo.Jpg"
        
        Set sbObj = New Scripting.FileSystemObject
               
        'ABRE A TAREFA DO OUTLOOK
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
    
            'ASSUNTO
            oitem.Subject = ("Volumetria Por Segmento - Confirming " & DataPesq & "")
            
            'ENDEREÇO DE ONDE VAI SAIR O EMAIL
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            
            oitem.To = EmailDestino
            oitem.cc = EmailCopia

        'CORPO DO EMAIL
        Corpo1 = "Prezados,"
        Corpo2 = "Segue anexo a "
        Relatorio = "Volumetria por Segmento do Último dia Útil: " & Format(DataPesq, "dd/mm/yyyy")
        Corpo3 = "Atenciosamente."
        Assinatura1 = "Confirming®"
        Assinatura2 = "Processamento de Ativos e Garantias"
        Assinatura3 = "Rua Amador Bueno, 474"
        Assinatura4 = "Ativos Atacado Processamento"
        Assinatura5 = "CEP: 04752-005  São Paulo-SP"
        Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
        Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
        Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
                       
            'FORMATAÇÃO DO CORPO
            oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Calibri <BR>" & Corpo1 & "<BR/>" & _
        "<BR>" & Corpo2 & "<B>" & Relatorio & "</B>" & Corpo21 & "<BR>" & "<BR>" & Corpo3 & "<BR><BR>" & _
        "<img src=" & Assinatura & " height=50 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 3" & "<BR>" & _
        "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 2 <BR>" & Assinatura3 & _
        "<BR/>" & Assinatura4 & "<BR/>" & Assinatura5 & "<BR/></FONT><FONT COLOR = BLACK FACE = Calibri Size = 1 <BR><I>" & Assinatura6 & _
        "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"
            
   'INSERE O ANEXO NO EMAIL
    oitem.Attachments.Add File
    'oitem.Display 'MOSTRA NA TELA ANTES DE ENVIAR
    oitem.Send

    Set olapp = Nothing
    Set oitem = Nothing
Fim:

End Function

