'Mdl_PPB_RDC'

Function ImportarCustoRDCExcel(Caminho As String)

    Dim ObjExcelCusto As Object, TbCustoRDC As Recordset
        
        LinhaCusto = 2

        Set ObjExcelCusto = CreateObject("EXCEL.application")
            ObjExcelCusto.Workbooks.Open FileName:=Caminho
        Set ObjExcelCusto = ObjExcelCusto.Worksheets("Plan1")

            Do While True

                Set TbCustoRDC = BDPPB.OpenRecordset("Tbl_Custo", dbOpenDynaset)

                    TbCustoRDC.AddNew
                        TbCustoRDC!Agencia_Ancora = "2271"
                        TbCustoRDC!Convenio_Ancora = "008500001038"
                        TbCustoRDC!Data = Trim(ObjExcelCusto.Cells(LinhaCusto, 1))
                        TbCustoRDC!Prazo = Trim(ObjExcelCusto.Cells(LinhaCusto, 2))
                        TbCustoRDC!Custo = Trim(ObjExcelCusto.Cells(LinhaCusto, 3))
                        TbCustoRDC!Prazo_2 = Trim(ObjExcelCusto.Cells(LinhaCusto, 4))
                        TbCustoRDC!Vencimento = Trim(ObjExcelCusto.Cells(LinhaCusto, 5))
                    TbCustoRDC.Update

                  LinhaCusto = LinhaCusto + 1

              If Trim(ObjExcelCusto.Cells(LinhaCusto, 1)) = "" Then: Exit Do

            Loop

    ObjExcelCusto.Application.Quit

End Function
Sub AdicionarCustosRDC()

    Dim TbDados As Recordset, CustoRDC As QueryDef, sFname As String
    Dim Percentual As Double, unidade As Double, Etapa As String
    
        Call AbrirBDPPB

          DataInic = Format(DataInicio, "mm/dd/YYYY")
          DataFin = Format(DataFinal, "mm/dd/YYYY")
            
           Call AtualizarStatus("1/5", 1, 0)
            
            Set CustoRDC = BDPPB.CreateQueryDef("CustoRDC", "SELECT Tbl_Custo.Agencia_Ancora, Tbl_Custo.Convenio_Ancora, Tbl_Custo.Data FROM Tbl_Custo GROUP BY Tbl_Custo.Agencia_Ancora, Tbl_Custo.Convenio_Ancora, Tbl_Custo.Data HAVING (((Tbl_Custo.Agencia_Ancora)='2271') AND ((Tbl_Custo.Convenio_Ancora)='008500001038') AND ((Tbl_Custo.Data)>=#" & DataInic & "# And (Tbl_Custo.Data)<=#" & DataFin & "#));")

                BDPPB.Containers.Refresh
            
            Set TbDados = BDPPB.OpenRecordset("SELECT CustoRDC.Data, TblCalendario.Data_dia, TblCalendario.Tipo FROM CustoRDC RIGHT JOIN TblCalendario ON CustoRDC.Data = TblCalendario.Data_dia WHERE (((CustoRDC.Data) Is Null) AND ((TblCalendario.Data_dia)>=#" & DataInic & "# And (TblCalendario.Data_dia)<Date()) AND ((TblCalendario.Tipo)='Util'));", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                    TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                    Contador = 0: unidade = Round((8200 / Qntd), 2): Etapa = "1/5"
                        Do While TbDados.EOF = False
                        Debug.Print TbDados!Data_dia
                            sFname = "\\saont46\apps2\Confirming\PlanCarrefour\Planilhas Custo Carrefour\Planilha Carrefour " & Format(TbDados!Data_dia, "ddmmyyyy") & ".xlsx"
                                If (Dir(sFname) <> "") Then
                                    Call ImportarCustoRDCExcel(sFname)
                                End If
                            TbDados.MoveNext
                           Contador = Contador + 1
                          Percentual = Round(((Contador / Qntd) * 100), 0)
                         Call AtualizarStatus(Etapa, Percentual, unidade)
                        Loop
                    BDPPB.Execute ("Drop Table CustoRDC")
                  BDPPB.Close
                End If
End Sub
Sub Operaçõe_SDF()

    Dim Caminho As String, Fornecedor As String, VALOR_NOTA As String, Nota As String, txvalorop As String, txjuros As String
    Dim txbanco As Double, txconvenio As Double, TxAgencia As Double, txconta As Double, TxCNPJ As Double, txtaxa As Double
    Dim Pos As Double, PosQtde As Double, Ret As Double
    Dim VencimentoNota As Date, VencimentoNota2 As Date
    Dim txdata_op As Date, txdata_contra As Date
    Dim Db As Database, TbDados As Recordset

        Call AbrirBDPPB
            
          DiarioPesq = UltimoDiaUtil(): DataPesq = Format(DiarioPesq, "ddmmyy")

                sFname = "\\saont46\apps2\Confirming\ArquivoLido\ARQRET_CA5D01_DX" & DataPesq & ".TXT"
                    If (Dir(sFname) = "") Then
                        sFname = "\\saont46\apps2\Confirming\ArquivoLido\ARQRET_CA5D02_DX" & DataPesq & ".TXT"
                           If (Dir(sFname) = "") Then: GoTo Fim
                    End If

                        Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPD_RDC_SDF.Data_Op, Tbl_PPD_RDC_SDF.Banco, Tbl_PPD_RDC_SDF.Agencia_Ancora, Tbl_PPD_RDC_SDF.Convenio_Ancora, Tbl_PPD_RDC_SDF.Conta_Ancora, Tbl_PPD_RDC_SDF.Data_Contratacao, Tbl_PPD_RDC_SDF.Nome_Fornecedor, Tbl_PPD_RDC_SDF.CNPJ_Fornecedor, Tbl_PPD_RDC_SDF.Valor_Op, Tbl_PPD_RDC_SDF.Valor_Juros, Tbl_PPD_RDC_SDF.Valor_Liquido, Tbl_PPD_RDC_SDF.Taxa_Mes, Tbl_PPD_RDC_SDF.Numero_Nota, Tbl_PPD_RDC_SDF.Vencimento_Nota, Tbl_PPD_RDC_SDF.Vencimento_Nota2, Tbl_PPD_RDC_SDF.Valor_Nota, Tbl_PPD_RDC_SDF.CHAVE FROM Tbl_PPD_RDC_SDF" _
                        & " GROUP BY Tbl_PPD_RDC_SDF.Data_Op, Tbl_PPD_RDC_SDF.Banco, Tbl_PPD_RDC_SDF.Agencia_Ancora, Tbl_PPD_RDC_SDF.Convenio_Ancora, Tbl_PPD_RDC_SDF.Conta_Ancora, Tbl_PPD_RDC_SDF.Data_Contratacao, Tbl_PPD_RDC_SDF.Nome_Fornecedor, Tbl_PPD_RDC_SDF.CNPJ_Fornecedor, Tbl_PPD_RDC_SDF.Valor_Op, Tbl_PPD_RDC_SDF.Valor_Juros, Tbl_PPD_RDC_SDF.Valor_Liquido, Tbl_PPD_RDC_SDF.Taxa_Mes, Tbl_PPD_RDC_SDF.Numero_Nota, Tbl_PPD_RDC_SDF.Vencimento_Nota, Tbl_PPD_RDC_SDF.Vencimento_Nota2, Tbl_PPD_RDC_SDF.Valor_Nota, Tbl_PPD_RDC_SDF.CHAVE HAVING (((Tbl_PPD_RDC_SDF.Data_Op)=#" & Format(DiarioPesq, "MM/DD/YYYY") & "#));", dbOpenDynaset)
            
                            If TbDados.EOF = True Then
                                
                                Set TbDados = BDPPB.OpenRecordset("Tbl_PPD_RDC_SDF", dbOpenDynaset)
                                
                                        Qntd = 1
                                    
                                            Open sFname For Input As #1
                                            
                                                Line Input #1, FileBuffer
                                            
                                                    X = Mid(FileBuffer, 82, 3)
                                            
                                                    If Mid(FileBuffer, 82, 3) = "RDC" Then
                                                    
                                                 Do While Not EOF(1)
                                                    
                                                        Line Input #1, FileBuffer
                                                        
                                                            X = (Mid(FileBuffer, 151, 8))
                                                                
                                                          If (Mid(FileBuffer, 151, 8)) <> "00000000" Then: Exit Do
                                                                
                                                                If Mid(FileBuffer, 290, 1) <> "0" Then
                                                                  Do While True
                                                                     Line Input #1, FileBuffer
                                                                    If (Mid(FileBuffer, 151, 8)) = "00000000" Then: Exit Do
                                                                  Loop
                                                                  GoTo proximo
                                                                 End If
                                                          
                                                                txbanco = Extrai_Zeros(Mid(FileBuffer, 1, 3))
                                                                txconvenio = Extrai_Zeros(Mid(FileBuffer, 11, 4))
                                                                TxAgencia = Extrai_Zeros(Mid(FileBuffer, 32, 4))
                                                                txconta = Extrai_Zeros(Mid(FileBuffer, 36, 14))
                                                                TxCNPJ = Extrai_Zeros(Mid(FileBuffer, 51, 14))
                                                                txfornecedor = Mid(FileBuffer, 65, 30)
                                                            'Pesq Dia Operação
                                                                txdia = Mid(FileBuffer, 95, 2)
                                                                txmes = Mid(FileBuffer, 97, 2)
                                                                txano = Mid(FileBuffer, 99, 4)
                                                                txdata_op = txdia & "/" & txmes & "/" & txano
                                                            'Pesq Dia da Contratação
                                                                txdiacontra = Mid(FileBuffer, 103, 2)
                                                                txmescontra = Mid(FileBuffer, 105, 2)
                                                                txanocontra = Mid(FileBuffer, 107, 4)
                                                                txdata_contra = txdiacontra & "/" & txmescontra & "/" & txanocontra
                                                            'Pesq Valor da Operação
                                                                txvalorop = Extrai_Zeros(Mid(FileBuffer, 111, 17))
                                                                txvalorop = Extrai_Zeros(txvalorop)
                                                                valoropdecimal = Right(txvalorop, 2)
                                                                Qntd = Len(txvalorop)
                                                                Valoropinteiro = Mid(txvalorop, 1, Qntd - 2)
                                                                txvalorop = Valoropinteiro & "," & valoropdecimal
                                                            'Pesq Valor Juros
                                                                txjuros = Extrai_Zeros(Mid(FileBuffer, 128, 17))
                                                                txjuros = Extrai_Zeros(txjuros)
                                                                txjurosdecimal = Right(txjuros, 2)
                                                                Qntd = Len(txjuros)
                                                                txjurosinteiro = Mid(txjuros, 1, Qntd - 2)
                                                                txjuros = txjurosinteiro & "," & txjurosdecimal
                                                            'Pesq Taxa
                                                                txtaxa = Extrai_Zeros(Mid(FileBuffer, 170, 15))
                                                                Qntd = Len(txtaxa)
                                                                txtaxaprincipal = Left(txtaxa, 1)
                                                                txtaxadecimal = Mid(txtaxa, 2, Qntd)
                                                                txtaxa = txtaxaprincipal & "," & txtaxadecimal
                                                                
                                                                    Line Input #1, FileBuffer
                                                                
                                                                            Do While Not EOF(1)
                                            
                                                                                If (Mid(FileBuffer, 151, 8)) = "00000000" Then: Exit Do
                                                                                 'Pesq Nota
                                                                                    Nota = Extrai_Zeros((Mid(FileBuffer, 31, 13)))
                                                                                    Nota = Extrai_Zeros(Nota)
                                                                                 'Pesq Vencimento
                                                                                    Vencimento = Extrai_Zeros((Mid(FileBuffer, 44, 15)))
                                                                                    Vencimento = Format(Vencimento, "00000000")
                                                                                    Vencdia = Left(Vencimento, 2)
                                                                                    Vencmes = Mid(Vencimento, 3, 2)
                                                                                    Vencano = Right(Vencimento, 4)
                                                                                    VencimentoNota = Vencdia & "/" & Vencmes & "/" & Vencano
                                                                                  'Pesq Vencimento 2
                                                                                    VENCIMENTO2 = Extrai_Zeros((Mid(FileBuffer, 59, 8)))
                                                                                    VENCIMENTO2 = Format(VENCIMENTO2, "00000000")
                                                                                    Vencdia = Left(VENCIMENTO2, 2)
                                                                                    Vencmes = Mid(VENCIMENTO2, 3, 2)
                                                                                    Vencano = Right(VENCIMENTO2, 4)
                                                                                    VencimentoNota2 = Vencdia & "/" & Vencmes & "/" & Vencano
                                                                                  'Pesq Valor Nota
                                                                                    VALOR_NOTA = Extrai_Zeros((Mid(FileBuffer, 67, 15)))
                                                                                    VALOR_NOTA = Extrai_Zeros(VALOR_NOTA)
                                                                                    valornotadecimal = Right(VALOR_NOTA, 2)
                                                                                    Qntd = Len(VALOR_NOTA)
                                                                                    Valornotainteiro = Mid(VALOR_NOTA, 1, Qntd - 2)
                                                                                    VALOR_NOTA = Valornotainteiro & "," & valornotadecimal
                                                                                    
                                                                                        TbDados.AddNew
                                                                                    
                                                                                            TbDados!banco = txbanco
                                                                                            TbDados!Agencia_Ancora = "2271"
                                                                                            TbDados!Convenio_Ancora = "008500001038"
                                                                                            TbDados!Conta_Ancora = txconta
                                                                                            TbDados!Data_Contratacao = Format(txdata_contra, "dd/mm/yyyy")
                                                                                            TbDados!Data_op = Format(txdata_op, "dd/mm/yyyy")
                                                                                            TbDados!Nome_Fornecedor = txfornecedor
                                                                                            TbDados!Cnpj_Fornecedor = TxCNPJ
                                                                                            TbDados!Valor_Op = txvalorop
                                                                                            TbDados!Valor_Juros = txjuros
                                                                                            TbDados!Valor_Liquido = txvalorop - txjuros
                                                                                            TbDados!Taxa_Mes = txtaxa
                                                                                            TbDados!Numero_nota = Nota
                                                                                            TbDados!Vencimento_Nota = VencimentoNota
                                                                                            TbDados!Vencimento_Nota2 = VencimentoNota2
                                                                                            TbDados!VALOR_NOTA = VALOR_NOTA
                                                                                            TbDados!CHAVE = txdata_op & txconvenio & TxCNPJ & Nota & VencimentoNota
                                            
                                                                                        TbDados.Update
                                            
                                                                                 Line Input #1, FileBuffer
                                               
                                                                             Loop
proximo:
                                            
                                                Loop
                                                
                                            Close #1
                                    
                                    TbDados.Close
                                
                                End If
                                
                            End If
Fim:
        
End Sub
Sub GerarPPBRDC()

    Dim Y As Integer, SiglaPesq As String
    Dim ValorJuros As Double, DiferimentoDiario As Double, ValorBanco As Double, ValorPPb As Double, TaxaMinina As Double
    Dim Db2 As Database, TbDados2 As Recordset, TbDados As Recordset, TbIncluir As Recordset
    Dim Percentual As Double, unidade As Double, Etapa As String

        Call AtualizarStatus("3/5", 1, 0)
        
            Call AbrirBDPPB
        
              DataInic = Format(DataInicio, "mm/dd/YYYY")
              DataFin = Format(DataFinal, "mm/dd/YYYY")
            
            'Inserir Dados na Tabela
            'BDPPB.Execute ("INSERT INTO Tbl_PPB ( Agencia_Ancora, Convenio_Ancora, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, CHAVE ) SELECT TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Juros, TblArqoped.Prazo_NF, TblArqoped.Valor_Juros, ([TblArqoped]![Data_op] & [TblArqoped]![Convenio_Ancora] & [TblArqoped]![Cod_Oper] & [TblArqoped]![Compromisso]) AS CHAVE" _
            & " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=2271 Or (TblArqoped.Agencia_Ancora)=2148) AND ((TblArqoped.Convenio_Ancora)='008500000011' Or (TblArqoped.Convenio_Ancora)='008500000422' Or (TblArqoped.Convenio_Ancora)='008500000449' Or (TblArqoped.Convenio_Ancora)='008500001021' Or (TblArqoped.Convenio_Ancora)='008500001038' Or (TblArqoped.Convenio_Ancora)='008500001011') AND ((TblArqoped.Data_op)>=#" & DataInic & "# And (TblArqoped.Data_op)<=#" & DataFin & "#));")
            
            'Inserir Dados na Tabela (Nome e CNPJ do Ancora)
            BDPPB.Execute ("INSERT INTO Tbl_PPB ( Nome_Ancora, Cnpj_Ancora, Agencia_Ancora, Convenio_Ancora, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, CHAVE ) SELECT TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Juros, TblArqoped.Prazo_NF, TblArqoped.Valor_Juros, ([TblArqoped]![Data_op] & [TblArqoped]![Convenio_Ancora] & [TblArqoped]![Cod_Oper] & [TblArqoped]![Compromisso]) AS CHAVE" _
            & " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=2271 Or (TblArqoped.Agencia_Ancora)=2148) AND ((TblArqoped.Convenio_Ancora)='008500000011' Or (TblArqoped.Convenio_Ancora)='008500000422' Or (TblArqoped.Convenio_Ancora)='008500000449' Or (TblArqoped.Convenio_Ancora)='008500001021' Or (TblArqoped.Convenio_Ancora)='008500001038' Or (TblArqoped.Convenio_Ancora)='008500001011') AND ((TblArqoped.Data_op)>=#" & DataInic & "# And (TblArqoped.Data_op)<=#" & DataFin & "#));")

            'Inserir Notas do SDF
            BDPPB.Execute ("INSERT INTO Tbl_PPB ( Agencia_Ancora, Convenio_Ancora, Prazo, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, Taxa, Numero_Nota, Data_Venc, Valor_Nota, CHAVE ) SELECT Tbl_PPD_RDC_SDF.Agencia_Ancora, Tbl_PPD_RDC_SDF.Convenio_Ancora, ([Tbl_PPD_RDC_SDF]![Vencimento_Nota]-[Tbl_PPD_RDC_SDF]![Data_op]) AS PRAZO, Tbl_PPD_RDC_SDF.Data_Op, Tbl_PPD_RDC_SDF.Nome_Fornecedor, Tbl_PPD_RDC_SDF.CNPJ_Fornecedor, Tbl_PPD_RDC_SDF.Taxa_Mes, Tbl_PPD_RDC_SDF.Numero_Nota, Tbl_PPD_RDC_SDF.Vencimento_Nota, Tbl_PPD_RDC_SDF.Valor_Nota, Tbl_PPD_RDC_SDF.CHAVE FROM Tbl_PPD_RDC_SDF WHERE (((Tbl_PPD_RDC_SDF.Data_Op)>=#" & DataInic & "# And (Tbl_PPD_RDC_SDF.Data_Op)<=#" & DataFin & "#));")

            'Inserir Dados e Custo TBL FINAL
            'BDPPB.Execute ("INSERT INTO Tbl_PPB_FINAL ( Agencia_Ancora, Convenio_Ancora, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, Custo, Valor_Banco, PPB_Bruto, CHAVE ) SELECT Tbl_PPB.Agencia_Ancora, Tbl_PPB.Convenio_Ancora, Tbl_PPB.Data_Op, Tbl_PPB.Nome_Fornecedor, Tbl_PPB.CNPJ_Fornecedor, Tbl_PPB.COD_OPERACAO, Tbl_PPB.Numero_Nota, Tbl_PPB.Data_Venc, Tbl_PPB.Valor_Nota, Tbl_PPB.Taxa, Tbl_PPB.Prazo, Tbl_PPB.Valor_Juros, Tbl_Custo.Custo, Tbl_PPB.Valor_Banco, Tbl_PPB.PPB_Bruto, Tbl_PPB.CHAVE FROM Tbl_PPB INNER JOIN Tbl_Custo ON (Tbl_PPB.Data_Op = Tbl_Custo.Data) AND (Tbl_PPB.Prazo = Tbl_Custo.Prazo)" _
             & " WHERE (((Tbl_Custo.Agencia_Ancora)='2271') AND ((Tbl_Custo.Convenio_Ancora)='008500001038') AND ((Tbl_PPB.Agencia_Ancora)=2271 Or (Tbl_PPB.Agencia_Ancora)=2148) AND ((Tbl_PPB.Convenio_Ancora)='008500000011' Or (Tbl_PPB.Convenio_Ancora)='008500000422' Or (Tbl_PPB.Convenio_Ancora)='008500000449' Or (Tbl_PPB.Convenio_Ancora)='008500001021' Or (Tbl_PPB.Convenio_Ancora)='008500001038' Or (Tbl_PPB.Convenio_Ancora)='008500001011') AND ((Tbl_PPB.Data_Op)>=#" & DataInic & "# And (Tbl_PPB.Data_Op)<=#" & DataFin & "#));")

            'Inserir Dados e Custo Tbl Final (Nome e CNPJ Ancora)
            BDPPB.Execute ("INSERT INTO Tbl_PPB_FINAL ( Nome_Ancora, CNPJ_Ancora, Agencia_Ancora, Convenio_Ancora, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, Custo, Valor_Banco, PPB_Bruto, CHAVE ) SELECT Tbl_PPB.Nome_Ancora, Tbl_PPB.CNPJ_Ancora, Tbl_PPB.Agencia_Ancora, Tbl_PPB.Convenio_Ancora, Tbl_PPB.Data_Op, Tbl_PPB.Nome_Fornecedor, Tbl_PPB.CNPJ_Fornecedor, Tbl_PPB.COD_OPERACAO, Tbl_PPB.Numero_Nota, Tbl_PPB.Data_Venc, Tbl_PPB.Valor_Nota, Tbl_PPB.Taxa, Tbl_PPB.Prazo, Tbl_PPB.Valor_Juros, Tbl_Custo.Custo, Tbl_PPB.Valor_Banco, Tbl_PPB.PPB_Bruto, Tbl_PPB.CHAVE" _
            & " FROM Tbl_PPB INNER JOIN Tbl_Custo ON (Tbl_PPB.Prazo = Tbl_Custo.Prazo) AND (Tbl_PPB.Data_Op = Tbl_Custo.Data) WHERE (((Tbl_PPB.Agencia_Ancora)=2271 Or (Tbl_PPB.Agencia_Ancora)=2148) AND ((Tbl_PPB.Convenio_Ancora)='008500000011' Or (Tbl_PPB.Convenio_Ancora)='008500000422' Or (Tbl_PPB.Convenio_Ancora)='008500000449' Or (Tbl_PPB.Convenio_Ancora)='008500001021' Or (Tbl_PPB.Convenio_Ancora)='008500001038' Or (Tbl_PPB.Convenio_Ancora)='008500001011') AND ((Tbl_PPB.Data_Op)>=#" & DataInic & "# And (Tbl_PPB.Data_Op)<=#" & DataFin & "#) AND ((Tbl_Custo.Agencia_Ancora)='2271') AND ((Tbl_Custo.Convenio_Ancora)='008500001038'));")

            'Calcular PPB
            Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Código, Tbl_PPB_FINAL.Agencia_Ancora, Tbl_PPB_FINAL.Convenio_Ancora, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.COD_OPERACAO, Tbl_PPB_FINAL.Numero_Nota, Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Prazo, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Custo, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto, Tbl_PPB_FINAL.CHAVE " _
            & " FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora)='2271' Or (Tbl_PPB_FINAL.Agencia_Ancora)='2148') AND ((Tbl_PPB_FINAL.Convenio_Ancora)='008500000011' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500000422' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500000449' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001021' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001038' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001011') AND ((Tbl_PPB_FINAL.Data_Op)>=#" & DataInic & "# And (Tbl_PPB_FINAL.Data_Op)<=#" & DataFin & "#));", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                     TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                     Contador = 0: unidade = Round((8200 / Qntd), 2): Etapa = "3/5"
                        Do While TbDados.EOF = False
                                TaxaMinina = TbDados!Custo / 1
            
                                    ValorJuros = CalcJuros(TbDados!VALOR_NOTA, TbDados!taxa, TbDados!Prazo)
        
                                    ValorBanco = CalcBanco(TbDados!VALOR_NOTA, TaxaMinina, TbDados!Prazo)
                                            
                                    ValorPPb = Round(CalcPPB(ValorJuros, ValorBanco), 2)
         
                                Set TbIncluir = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Código, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Código)=" & TbDados!Código & "));", dbOpenDynaset)
                                    TbIncluir.Edit
                                        TbIncluir!Valor_Juros = ValorJuros
                                        TbIncluir!Valor_Banco = ValorBanco
                                        TbIncluir!PPB_Bruto = ValorPPb
                                    TbIncluir.Update
                                TbIncluir.Close
                            TbDados.MoveNext
                           Contador = Contador + 1
                          Percentual = Round(((Contador / Qntd) * 100), 0)
                         Call AtualizarStatus(Etapa, Percentual, unidade)
                        Loop
                End If
End Sub
Sub DiferimentoTblRDC()

    Dim TbDados1, TbCalendario, TbDados, TbVenc, TbDados2, TbData, TbDataOp, TbPrazo, TbPPB As DAO.Recordset
    Dim DataPesq As Date, Y As Integer, SiglaPesq As String
    Dim ValorJuros As Double, DiferimentoDiario  As Double, ValorDif(500) As Double
    Dim Caminho, MesData, LinhaDiferimento, UltimoVencimento, MesVenc, MesPesq, Limite, AnoPesq, Datames As String
    Dim cont, Qntd As Integer, Percentual As Double, unidade As Double, Etapa As String
    Dim DataDaPesquisa, Resultado(24), DataVencimento, DataOperação As Date
    Dim Data(500) As Date, DataInicioPlan As Date, DiasCalc(500) As Integer

       Call AtualizarStatus("4/5", 1, 0)
        
          Call AbrirBDPPB
            
            Datames = Right(Format(DataFinal, "dd/mm/yyyy"), 7): MesPesq = Mid(Format(DataFinal, "dd/mm/yyyy"), 4, 2): AnoPesq = Right(Format(DataFinal, "dd/mm/yyyy"), 4): Qntd = 1

                Do While True
                    MesPesq = Format(MesPesq, "00")
                    DataDaPesquisa = MesPesq & "/" & AnoPesq
                        Set TbCalendario = CurrentDb.OpenRecordset("SELECT TblCalendario.Data_dia, TblCalendario.Semana, TblCalendario.Feriado, TblCalendario.Tipo FROM TblCalendario WHERE (((TblCalendario.Data_dia) Like '*" & DataDaPesquisa & "*')) ORDER BY TblCalendario.Data_dia DESC;", dbOpenDynaset)
                            Resultado(Qntd) = TbCalendario!Data_dia
                                MesPesq = MesPesq + 1
                                If MesPesq = "13" Then
                                    MesPesq = "01"
                                    AnoPesq = AnoPesq + 1
                                End If
                            Qntd = Qntd + 1
                        If Qntd = 15 Then: Exit Do
                Loop

        Set TbDiferimento = BDPPB.OpenRecordset("Tbl_Diferimento", dbOpenDynaset)
    
            DataFin = Format(DataFinal, "mm/dd/yyyy"): DataInic = Format(DataInicio, "mm/dd/yyyy"): DataInicioPlan = DataInicio
    
        Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Código, Tbl_PPB_FINAL.Agencia_Ancora, Tbl_PPB_FINAL.Convenio_Ancora, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.COD_OPERACAO, Tbl_PPB_FINAL.Numero_Nota, Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Prazo, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Custo, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto, Tbl_PPB_FINAL.CHAVE " _
        & " FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora)='2271' Or (Tbl_PPB_FINAL.Agencia_Ancora)='2148') AND ((Tbl_PPB_FINAL.Convenio_Ancora)='008500000011' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500000422' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500000449' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001021' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001038' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001011') AND ((Tbl_PPB_FINAL.Data_Op)>=#" & DataInic & "# And (Tbl_PPB_FINAL.Data_Op)<=#" & DataFin & "#))ORDER BY Tbl_PPB_FINAL.Data_Op;", dbOpenDynaset)
            
            If TbDados.EOF = False Then
            
                TbDados.MoveLast: Quantidade = TbDados.RecordCount: TbDados.MoveFirst
                Contador = 0: unidade = Round((8200 / Quantidade), 2): Etapa = "4/5"

                    Do While TbDados.EOF = False
                
                       Set TbDados1 = BDPPB.OpenRecordset("SELECT Tbl_Diferimento.Código, Tbl_Diferimento.Agencia_Ancora, Tbl_Diferimento.Convenio_Ancora, Tbl_Diferimento.Nome_Ancora, Tbl_Diferimento.CHAVE, Tbl_Diferimento.Data_Dif, Tbl_Diferimento.Valor_Dif FROM Tbl_Diferimento GROUP BY Tbl_Diferimento.Código, Tbl_Diferimento.Agencia_Ancora, Tbl_Diferimento.Convenio_Ancora, Tbl_Diferimento.Nome_Ancora, Tbl_Diferimento.CHAVE, Tbl_Diferimento.Data_Dif, Tbl_Diferimento.Valor_Dif HAVING (((Tbl_Diferimento.CHAVE)='" & TbDados!CHAVE & "'));", dbOpenDynaset)
                       
                            If TbDados1.EOF = False Then: BDPPB.Execute ("Delete Tbl_Diferimento.Código, Tbl_Diferimento.Agencia_Ancora, Tbl_Diferimento.Convenio_Ancora, Tbl_Diferimento.Nome_Ancora, Tbl_Diferimento.CHAVE, Tbl_Diferimento.Data_Dif, Tbl_Diferimento.Valor_Dif FROM Tbl_Diferimento WHERE (((Tbl_Diferimento.CHAVE)='" & TbDados!CHAVE & "'));")
                
                        DiferimentoDiario = (TbDados!PPB_Bruto / TbDados!Prazo): DataVencimento = TbDados!Data_Venc: DataOperação = TbDados!Data_op
                        
                            If DataVencimento < DataInicioPlan Then
                                Data(1) = DataInicioPlan
                            Else
                                Data(1) = DataOperação
                            End If
                            
                                If DataVencimento > Resultado(1) Then
                                    Data(2) = Resultado(1)
                                Else
                                    Data(2) = DataVencimento
                                End If
                            
                                DiasCalc(3) = Data(2) - Data(1)
                        
                            If (DiferimentoDiario * DiasCalc(3)) = 0 Then
                                ValorDif(1) = "0"
                            Else
                                ValorDif(1) = CalcDiferimento(DiferimentoDiario, DiasCalc(3))
                            End If
                            
                            If ValorDif(1) > 0 Then
                    
                                TbDiferimento.AddNew
                                    TbDiferimento!Agencia_Ancora = "2271"
                                    TbDiferimento!Convenio_Ancora = "008500001038"
                                    TbDiferimento!Nome_Ancora = "RDC"
                                    TbDiferimento!CHAVE = TbDados!CHAVE
                                    TbDiferimento!Data_Dif = Resultado(1)
                                    TbDiferimento!Valor_Dif = ValorDif(1)
                                TbDiferimento.Update
                                
                            End If
                
                        QntdResul = 1: DataResul = 4: ValorResul = 2
                
                                Do While True
                    
                                    If DataVencimento < Resultado(QntdResul) Then
                                        Data(DataResul) = DataVencimento
                                    Else
                                        Data(DataResul) = Resultado(QntdResul)
                                    End If
                    
                                        If DataVencimento > Resultado(QntdResul + 1) Then
                                            Data(DataResul + 1) = Resultado(QntdResul + 1)
                                        Else
                                            Data(DataResul + 1) = DataVencimento
                                        End If
                    
                                    DiasCalc(DataResul + 2) = Data(DataResul + 1) - Data(DataResul)
                    
                                    If (DiferimentoDiario * DiasCalc(DataResul + 2)) = 0 Then
                                        ValorDif(ValorResul) = "0"
                                    Else
                                        ValorDif(ValorResul) = CalcDiferimento(DiferimentoDiario, DiasCalc(DataResul + 2))
                                    End If
                    
                                    If ValorDif(ValorResul) > 0 Then
                
                                        TbDiferimento.AddNew
                                            TbDiferimento!Agencia_Ancora = "2271"
                                            TbDiferimento!Convenio_Ancora = "008500001038"
                                            TbDiferimento!Nome_Ancora = "RDC"
                                            TbDiferimento!CHAVE = TbDados!CHAVE
                                            TbDiferimento!Data_Dif = Resultado(QntdResul + 1)
                                            TbDiferimento!Valor_Dif = ValorDif(ValorResul)
                                        TbDiferimento.Update
                                    End If
                
                                    QntdResul = QntdResul + 1: DataResul = DataResul + 1: ValorResul = ValorResul + 1
                
                                    If ValorResul = 14 Then: Exit Do
                                Loop
                            TbDados.MoveNext
                           Contador = Contador + 1
                          Percentual = Round(((Contador / Quantidade) * 100), 0)
                         Call AtualizarStatus(Etapa, Percentual, unidade)
                        Loop
            End If
End Sub
Sub DiferimentoExcel()

    Dim ObjExcelCusto, ObjExcelppb, ObjExcel, ObjPlan1Excel, ObjPlan2Excel As Object
    Dim TbDados1, TbCalendario, TbDados, TbVenc, TbDados2, TbData, TbDataOp, TbPrazo, TbPPB As DAO.Recordset
    Dim Percentual As Double, unidade As Double, Etapa As String
    Dim DtInicio As Date, DtFim As Date, DataPesq As Date, DataInicioPlan As Date
    Dim ValorJuros, DiferimentoDiario As Double
    Dim Caminho, MesData, LinhaDiferimento, UltimoVencimento, MesVenc, MesPesq, Limite, AnoPesq, Datames As String
    Dim cont, Qntd As Integer, linha As Double, Nome As String
    Dim DataDaPesquisa, Resultado(24), DataVencimento, DataOperação, Databasemais As Date
    
        Call AbrirBDPPB
    
            Call AtualizarStatus("5/5", 1, 0)

                DataFin = Format(DataFinal, "mm/dd/yyyy")
                DataInic = Format(DataInicio, "mm/dd/yyyy")
                           
                    'Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.Numero_Nota, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Prazo, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Custo, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora) = '2271' Or (Tbl_PPB_FINAL.Agencia_Ancora) = '2148') And ((Tbl_PPB_FINAL.Convenio_Ancora) = '008500000011' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500000422' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500000449' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500001021' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500001038' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500001011') And ((Tbl_PPB_FINAL.Data_op) >= #" & DataInic & "# And (Tbl_PPB_FINAL.Data_op) <= #" & DataFin & "#)) ORDER BY Tbl_PPB_FINAL.Data_Op;", dbOpenDynaset)
                    
                    Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Nome_Ancora, Tbl_PPB_FINAL.CNPJ_Ancora, Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.Numero_Nota, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Prazo, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Custo, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Data_Op)>=#" & DataInic & "# And (Tbl_PPB_FINAL.Data_Op)<=#" & DataFin & "#) AND ((Tbl_PPB_FINAL.Agencia_Ancora)='2271' Or (Tbl_PPB_FINAL.Agencia_Ancora)='2148') AND ((Tbl_PPB_FINAL.Convenio_Ancora)='008500000011' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500000422' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500000449' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001021' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001038' Or (Tbl_PPB_FINAL.Convenio_Ancora)='008500001011')) ORDER BY Tbl_PPB_FINAL.Data_Op;", dbOpenDynaset)

                        If TbDados.EOF = False Then
                    
                            Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
        
                            Set ObjExcel = CreateObject("EXCEL.application")
                                ObjExcel.Workbooks.Open FileName:=Caminho & "PPB_RDCV2.xlsx", ReadOnly:=True
                            Set ObjPlan1Excel = ObjExcel.Worksheets("Operações")
                                ObjPlan1Excel.Select
        
                                linha = 7
        
                            MesData = UCase(MonthName(Month(DataFinal)))
        
                                ObjPlan1Excel.Range("B2") = "Cliente:  RDC FOCCAR FACTORING FOMENTO"
                                ObjPlan1Excel.Range("B4") = MesData
                                ObjPlan1Excel.Range("A7").CopyFromRecordset TbDados
           
                                    TbDados.MoveLast
                                        TbLinha = TbDados.RecordCount
                                    UltimaLinha = TbLinha + linha
           
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(7).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(8).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(9).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(10).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(11).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(1).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(2).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(3).LineStyle = 2
                                ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(4).LineStyle = 2
                                ObjPlan1Excel.Range("C" & linha & ":C" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                                ObjPlan1Excel.Range("G" & linha & ":G" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                                ObjPlan1Excel.Range("E" & linha & ":F" & UltimaLinha).NumberFormat = "00000"
                                ObjPlan1Excel.Range("F" & linha & ":L" & UltimaLinha).NumberFormat = "00000"
                                ObjPlan1Excel.Range("H" & linha & ":H" & UltimaLinha).Style = "Currency"
                                ObjPlan1Excel.Range("K" & linha & ":K" & UltimaLinha).Style = "Currency"
                                ObjPlan1Excel.Range("N" & linha & ":N" & UltimaLinha).Style = "Currency"
                                ObjPlan1Excel.Rows(UltimaLinha & ":1000000").Delete Shift:=xlUp
                            TbDados.Close
                
                        Databasemais = DataFinal: Datames = Right(Format(Databasemais, "dd/mm/yyyy"), 7): MesPesq = Mid(Format(Databasemais, "dd/mm/yyyy"), 4, 2): AnoPesq = Right(Format(Databasemais, "dd/mm/yyyy"), 4): Qntd = 1

                        Do While True
                            MesPesq = Format(MesPesq, "00")
                                DataDaPesquisa = MesPesq & "/" & AnoPesq
                                    Set TbCalendario = BDPPB.OpenRecordset("SELECT TblCalendario.Data_dia, TblCalendario.Semana, TblCalendario.Feriado, TblCalendario.Tipo FROM TblCalendario WHERE (((TblCalendario.Data_dia) Like '*" & DataDaPesquisa & "*')) ORDER BY TblCalendario.Data_dia DESC;", dbOpenDynaset)
                                        Resultado(Qntd) = TbCalendario!Data_dia
                                        MesPesq = MesPesq + 1
                                    If MesPesq = "13" Then
                                        MesPesq = "01"
                                        AnoPesq = AnoPesq + 1
                                    End If
                                Qntd = Qntd + 1
                            If Qntd = 15 Then: Exit Do
                        Loop

                Set ObjPlan2Excel = ObjExcel.Worksheets("Alocação PPB")
        
            ObjPlan2Excel.Range("F3") = Format(DataInicio, "MM/DD/YYYY"): ObjPlan2Excel.Range("J3") = Resultado(1): ObjPlan2Excel.Range("N3") = Resultado(2): ObjPlan2Excel.Range("R3") = Resultado(3): ObjPlan2Excel.Range("V3") = Resultado(4): ObjPlan2Excel.Range("Z3") = Resultado(5): ObjPlan2Excel.Range("AD3") = Resultado(6): ObjPlan2Excel.Range("AH3") = Resultado(7): ObjPlan2Excel.Range("AL3") = Resultado(8): ObjPlan2Excel.Range("AP3") = Resultado(9): ObjPlan2Excel.Range("AT3") = Resultado(10): ObjPlan2Excel.Range("AX3") = Resultado(11)
            ObjPlan2Excel.Range("F3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("J3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("N3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("R3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("V3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("Z3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("AD3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("AH3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("AL3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("AP3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("AT3").NumberFormat = "dd/mm/yyyy": ObjPlan2Excel.Range("AX3").NumberFormat = "dd/mm/yyyy"
                
            DataFin = Format(DataFinal, "mm/dd/yyyy"): DataInic = Format(DataInicio, "mm/dd/yyyy"): DataInicioPlan = DataInicio
    
            Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.Numero_Nota, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Prazo, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Custo, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora) = '2271' Or (Tbl_PPB_FINAL.Agencia_Ancora) = '2148') And ((Tbl_PPB_FINAL.Convenio_Ancora) = '008500000011' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500000422' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500000449' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500001021' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500001038' Or (Tbl_PPB_FINAL.Convenio_Ancora) = '008500001011') And ((Tbl_PPB_FINAL.Data_op) >= #" & DataInic & "# And (Tbl_PPB_FINAL.Data_op) <= #" & DataFin & "#)) ORDER BY Tbl_PPB_FINAL.Data_Op;", dbOpenDynaset)
            
            LinhaDiferimento = 4
            
            TbDados.MoveLast: Quantidade = TbDados.RecordCount: TbDados.MoveFirst
            Contador = 0: unidade = Round((8200 / Quantidade), 2): Etapa = "5/5"
            
            Do While TbDados.EOF = False
            
                QntdResul = 1: Coluna = 11
            
                ObjPlan2Excel.Range("A" & LinhaDiferimento) = TbDados!Data_op
                ObjPlan2Excel.Range("B" & LinhaDiferimento) = TbDados!Prazo
                ObjPlan2Excel.Range("C" & LinhaDiferimento) = TbDados!Data_Venc
                ObjPlan2Excel.Range("D" & LinhaDiferimento) = TbDados!PPB_Bruto
            
                    DiferimentoDiario = (TbDados!PPB_Bruto / TbDados!Prazo): DataVencimento = TbDados!Data_Venc: DataOperação = TbDados!Data_op
            
                        ObjPlan2Excel.Range("E" & LinhaDiferimento) = DiferimentoDiario
                    
                             If DataVencimento < DataInicioPlan Then
                                ObjPlan2Excel.Range("G" & LinhaDiferimento) = DataInicioPlan
                            Else
                                ObjPlan2Excel.Range("G" & LinhaDiferimento) = DataOperação
                            End If
                            
                                If DataVencimento > Resultado(1) Then
                                    ObjPlan2Excel.Range("H" & LinhaDiferimento) = Resultado(1)
                                Else
                                    ObjPlan2Excel.Range("H" & LinhaDiferimento) = DataVencimento
                                End If
                            
                                ObjPlan2Excel.Range("I" & LinhaDiferimento) = ObjPlan2Excel.Range("H" & LinhaDiferimento) - ObjPlan2Excel.Range("G" & LinhaDiferimento)
                        
                            If (DiferimentoDiario * ObjPlan2Excel.Range("I" & LinhaDiferimento)) = 0 Then
                                ObjPlan2Excel.Range("J" & LinhaDiferimento) = "0"
                            Else
                                ObjPlan2Excel.Range("J" & LinhaDiferimento) = CalcDiferimento(DiferimentoDiario, ObjPlan2Excel.Range("I" & LinhaDiferimento))
                            End If
                            
                                Do While True
                                        
                                    If DataVencimento < Resultado(QntdResul) Then
                                        ObjPlan2Excel.Cells(LinhaDiferimento, Coluna) = DataVencimento
                                    Else
                                        ObjPlan2Excel.Cells(LinhaDiferimento, Coluna) = Resultado(QntdResul)
                                    End If
                                    
                                        If DataVencimento > Resultado(QntdResul + 1) Then
                                            ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 1) = Resultado(QntdResul + 1)
                                        Else
                                            ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 1) = DataVencimento
                                        End If
                                    
                                    ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 2) = ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 1) - ObjPlan2Excel.Cells(LinhaDiferimento, Coluna)
                                    
                                    If (DiferimentoDiario * ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 2)) = 0 Then
                                       ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 3) = "0"
                                    Else
                                       ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 3) = CalcDiferimento(DiferimentoDiario, ObjPlan2Excel.Cells(LinhaDiferimento, Coluna + 2))
                                    End If
                                
                                        QntdResul = QntdResul + 1: Coluna = Coluna + 4
                                    
                                    If QntdResul = 11 Then: Exit Do
                                Loop
               
                        TbDados.MoveNext
                    LinhaDiferimento = LinhaDiferimento + 1
               
                  Contador = Contador + 1
                 Percentual = Round(((Contador / Quantidade) * 100), 0)
                Call AtualizarStatus(Etapa, Percentual, unidade)
            Loop
        
        TbDados.MoveLast
        TbLinha = TbDados.RecordCount
        UltimaLinha = TbLinha + 4

      ObjPlan2Excel.Rows(UltimaLinha & ":1000000").Delete Shift:=xlUp
    ObjPlan2Excel.Activate
    ObjPlan2Excel.Columns("A:AX").Select
    ObjPlan2Excel.Columns.AutoFit
    
       Nome = "PPB RDC FOCCAR - " & MesData & " - " & Right(Date, 4)
          
       Nome = Trata_NomeArquivo(Nome)
       
        sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            If (Dir(sFname) <> "") Then
                Kill sFname
            End If
        
        ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        ObjExcel.activeworkbook.Close SaveChanges:=False
        ObjExcel.Quit

    End If

End Sub
