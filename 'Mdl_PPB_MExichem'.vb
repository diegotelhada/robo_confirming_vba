'Mdl_PPB_MExichem'

Sub Adc_PPB_Mexichen()

    Dim TbDados As Recordset
    Dim Receita_Cliente As Double, ValorBanco As Double
    Dim Percentual As Double, unidade As Double, Etapa As String
    
        Call AtualizarStatus("1/3", 1, 0)
            
            Call AbrirBDPPB
                       
                DataInic = Format(DataInicio, "mm/dd/YYYY")
                DataFin = Format(DataFinal, "mm/dd/YYYY")
            
                    BDPPB.Execute ("INSERT INTO Tbl_PPB ( Agencia_Ancora, Convenio_Ancora, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, Valor_Banco, CHAVE )" _
                    & " SELECT TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Juros, TblArqoped.Prazo_NF, TblArqoped.Valor_Juros, TblArqoped.Receita_Banco, ([TblArqoped]![Data_op] & [TblArqoped]![Convenio_Ancora] & [TblArqoped]![Cod_Oper] & [TblArqoped]![Compromisso]) AS CHAVE FROM TblArqoped LEFT JOIN Tbl_Exclusao ON TblArqoped.Cnpj_Fornecedor = Tbl_Exclusao.Cnpj_Fornecedor WHERE (((Tbl_Exclusao.Cnpj_Fornecedor) Is Null) AND ((TblArqoped.Agencia_Ancora)=2131) AND ((TblArqoped.Convenio_Ancora)='008500000022') AND ((TblArqoped.Data_op)>=#" & DataInic & "# And (TblArqoped.Data_op)<=#" & DataFin & "#));")
            
                    Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB.Agencia_Ancora, Tbl_PPB.Convenio_Ancora, Tbl_PPB.Data_Op, Tbl_PPB.Valor_Banco, Tbl_PPB.CHAVE FROM Tbl_PPB WHERE (((Tbl_PPB.Agencia_Ancora)=2131) AND ((Tbl_PPB.Convenio_Ancora)='008500000022') AND ((Tbl_PPB.Data_Op)>=#" & DataInic & "# And (Tbl_PPB.Data_Op)<=#" & DataFin & "#));", dbOpenDynaset)
                    
                        If TbDados.EOF = False Then
                    
                          TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                          Contador = 0: unidade = Round((8200 / Qntd), 2): Etapa = "1/3"
            
                                Do While TbDados.EOF = False
                                
                                    ValorBanco = TbDados!Valor_Banco
                                
                                        Receita_Cliente = CalcPPBMEXICHEM(ValorBanco)
                                
                                        CHAVEPESQ = TbDados!CHAVE
                                
                                        Set TbIncluir = BDPPB.OpenRecordset("SELECT Tbl_PPB.Agencia_Ancora, Tbl_PPB.Convenio_Ancora, Tbl_PPB.Data_Op, Tbl_PPB.PPB_Bruto, Tbl_PPB.Valor_Banco, Tbl_PPB.CHAVE FROM Tbl_PPB WHERE (((Tbl_PPB.CHAVE)='" & CHAVEPESQ & "'));", dbOpenDynaset)
                    
                                            TbIncluir.Edit
                                                TbIncluir!PPB_Bruto = Receita_Cliente
                                            TbIncluir.Update
                                        TbIncluir.Close
                                    TbDados.MoveNext
                                   Contador = Contador + 1
                                  Percentual = Round(((Contador / Qntd) * 100), 0)
                                 Call AtualizarStatus(Etapa, Percentual, unidade)
                                Loop
                            BDPPB.Execute ("INSERT INTO Tbl_PPB_FINAL ( Agencia_Ancora, Convenio_Ancora, Data_Op, Nome_Fornecedor, CNPJ_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, Custo, Valor_Banco, PPB_Bruto, CHAVE ) SELECT Tbl_PPB.Agencia_Ancora, Tbl_PPB.Convenio_Ancora, Tbl_PPB.Data_Op, Tbl_PPB.Nome_Fornecedor, Tbl_PPB.CNPJ_Fornecedor, Tbl_PPB.COD_OPERACAO, Tbl_PPB.Numero_Nota, Tbl_PPB.Data_Venc, Tbl_PPB.Valor_Nota, Tbl_PPB.Taxa, Tbl_PPB.Prazo, Tbl_PPB.Valor_Juros, Tbl_PPB.Custo, Tbl_PPB.Valor_Banco, Tbl_PPB.PPB_Bruto, Tbl_PPB.CHAVE FROM Tbl_PPB WHERE (((Tbl_PPB.Agencia_Ancora)=2131) AND ((Tbl_PPB.Convenio_Ancora)='008500000022') AND ((Tbl_PPB.Data_Op)>=#" & DataInic & "# And (Tbl_PPB.Data_Op)<=#" & DataFin & "#));")
                        End If
End Sub
Sub Diferimento_Mexichen()

    Dim TbDados1, TbCalendario, TbDados, TbVenc, TbDados2, TbData, TbDataOp, TbPrazo, TbPPB As DAO.Recordset
    Dim DataPesq As Date, Y As Integer, SiglaPesq As String
    Dim ValorJuros As Double, DiferimentoDiario  As Double, ValorDif(500) As Double
    Dim Caminho, MesData, LinhaDiferimento, UltimoVencimento, MesVenc, MesPesq, Limite, AnoPesq, Datames As String
    Dim cont, Qntd As Integer, Percentual As Double, unidade As Double, Etapa As String
    Dim DataDaPesquisa, Resultado(24), DataVencimento, DataOperação As Date
    Dim Data(500) As Date, DataInicioPlan As Date, DiasCalc(500) As Integer
                
        Call AtualizarStatus("2/3", 1, 0)
        
            Call AbrirBDPPB
                              
                DataInic = Format(DataInicio, "mm/dd/YYYY")
                DataFin = Format(DataFinal, "mm/dd/YYYY")
                                   
                    cont = 1: Qntd = 1: doto = DataInicio: Datames = Right(doto, 7): MesPesq = Format(Mid(doto, 4, 2), "00"): AnoPesq = Right(doto, 4): DataInicioPlan = DataInicio
    
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

    Set TbDiferimento = BDPPB.OpenRecordset("Tbl_Diferimento", dbOpenDynaset)
    
    Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Agencia_Ancora, Tbl_PPB_FINAL.Convenio_Ancora, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.COD_OPERACAO, Tbl_PPB_FINAL.Numero_Nota, Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Prazo, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Custo, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto, Tbl_PPB_FINAL.CHAVE FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora)='2131') AND ((Tbl_PPB_FINAL.Convenio_Ancora)='008500000022') AND ((Tbl_PPB_FINAL.Data_Op)>=#" & DataInic & "# And (Tbl_PPB_FINAL.Data_Op)<=#" & DataFin & "#));", dbOpenDynaset)
        
        If TbDados.EOF = False Then
        
             TbDados.MoveLast: Quantidade = TbDados.RecordCount: TbDados.MoveFirst
             Contador = 0: unidade = Round((8200 / Quantidade), 2): Etapa = "2/3"
                
                Do While TbDados.EOF = False
                
                    Set TbDados1 = BDPPB.OpenRecordset("SELECT Tbl_Diferimento.Agencia_Ancora, Tbl_Diferimento.Convenio_Ancora, Tbl_Diferimento.Nome_Ancora, Tbl_Diferimento.Data_Dif, Tbl_Diferimento.Valor_Dif, Tbl_Diferimento.CHAVE FROM Tbl_Diferimento WHERE (((Tbl_Diferimento.CHAVE)='" & TbDados!CHAVE & "'));", dbOpenDynaset)
                
                        If TbDados1.EOF = False Then: BDPPB.Execute ("Delete Tbl_Diferimento.Agencia_Ancora, Tbl_Diferimento.Convenio_Ancora, Tbl_Diferimento.Nome_Ancora, Tbl_Diferimento.Data_Dif, Tbl_Diferimento.Valor_Dif, Tbl_Diferimento.CHAVE FROM Tbl_Diferimento WHERE (((Tbl_Diferimento.CHAVE)='" & TbDados!CHAVE & "'));")
        
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
                                        TbDiferimento!Agencia_Ancora = "2131"
                                        TbDiferimento!Convenio_Ancora = "008500000022"
                                        TbDiferimento!Nome_Ancora = "MEXICHEM"
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
                                            TbDiferimento!Agencia_Ancora = "2131"
                                            TbDiferimento!Convenio_Ancora = "008500000022"
                                            TbDiferimento!Nome_Ancora = "MEXICHEM"
                                            TbDiferimento!CHAVE = TbDados!CHAVE
                                            TbDiferimento!Data_Dif = Resultado(1)
                                            TbDiferimento!Valor_Dif = ValorDif(1)
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
Sub GerarExcelMexichem()

    Dim ObjExcelCusto, ObjExcelppb, ObjExcel, ObjPlan1Excel, ObjPlan2Excel As Object
    Dim TbDados1, TbCalendario, TbDados, TbVenc, TbDados2, TbData, TbDataOp, TbPrazo, TbPPB As DAO.Recordset
    Dim linha As Double, ValorJuros, DiferimentoDiario As Double
    Dim Nome As String, Percentual As Double, unidade As Double, Etapa As String
    Dim DtInicio As Date, DtFim As Date, DataPesq As Date
    Dim Ret As String, SiglaPesq As String
    Dim Y As Integer, cont, Qntd As Integer
    Dim Caminho, MesData, LinhaDiferimento, UltimoVencimento, MesVenc, MesPesq, Limite, AnoPesq, Datames As String
    Dim DataDaPesquisa, Resultado(24), DataVencimento, DataOperação As Date, DataInicioPlan As Date

        Call AtualizarStatus("3/3", 1, 0)
        
            Call AbrirBDPPB
                
                DataFin = Format(DataFinal, "mm/dd/yyyy")
                DataInic = Format(DataInicio, "mm/dd/yyyy")
        
            Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.COD_OPERACAO, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Data_Op)>=#" & DataInic & "# And (Tbl_PPB_FINAL.Data_Op)<=#" & DataFin & "#) AND ((Tbl_PPB_FINAL.Agencia_Ancora)='2131') AND ((Tbl_PPB_FINAL.Convenio_Ancora)='008500000022')) ORDER BY Tbl_PPB_FINAL.Data_Op;", dbOpenDynaset)
                                             
                If TbDados.EOF = False Then
                             
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
    
                        Set ObjExcel = CreateObject("EXCEL.application")
                            ObjExcel.Workbooks.Open FileName:=Caminho & "PPB_MEXICHEN.xlsx", ReadOnly:=True
                        Set ObjPlan1Excel = ObjExcel.Worksheets("Operações")
                            ObjPlan1Excel.Select
    
                        linha = 7: MesData = UCase(MonthName(Month(DataInicio)))
    
                    ObjPlan1Excel.Range("B4") = "Mês de Referência: " & MesData
                    ObjPlan1Excel.Range("A7").CopyFromRecordset TbDados
    
                        TbDados.MoveLast: TbLinha = TbDados.RecordCount: UltimaLinha = TbLinha + linha
       
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":A" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("F" & linha & ":F" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("I" & linha & ":K" & UltimaLinha).Style = "Currency"
                    ObjPlan1Excel.Rows(UltimaLinha & ":1000000").Delete Shift:=xlUp
                TbDados.Close
    
                cont = 0: Qntd = 1: DataDaPesquisa = DataInicio: Datames = Right(DataDaPesquisa, 7): MesPesq = Mid(DataDaPesquisa, 4, 2): AnoPesq = Right(DataDaPesquisa, 4)
    
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
    
            Set TbDados = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.Data_Op, Tbl_PPB_FINAL.Nome_Fornecedor, Tbl_PPB_FINAL.CNPJ_Fornecedor, Tbl_PPB_FINAL.COD_OPERACAO, Tbl_PPB_FINAL.Valor_Nota, Tbl_PPB_FINAL.Taxa, Tbl_PPB_FINAL.Valor_Juros, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto, Tbl_PPB_FINAL.Data_Venc, Tbl_PPB_FINAL.Prazo FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora) = '2131') And ((Tbl_PPB_FINAL.Convenio_Ancora) = '008500000022') And ((Tbl_PPB_FINAL.Data_op) >= #" & DataInic & "# And (Tbl_PPB_FINAL.Data_op) <= #" & DataFin & "#)) ORDER BY Tbl_PPB_FINAL.Data_Op;", dbOpenDynaset)
    
                LinhaDiferimento = 4: DataInicioPlan = DataInicio
                
                TbDados.MoveLast: Quantidade = TbDados.RecordCount: TbDados.MoveFirst
                Contador = 0: unidade = Round((8200 / Quantidade), 2): Etapa = "3/3"
    
                Do While TbDados.EOF = False
                
                    QntdResul = 1: Coluna = 11
                
                    ObjPlan2Excel.Range("A" & LinhaDiferimento) = TbDados!Data_op
                    ObjPlan2Excel.Range("B" & LinhaDiferimento) = TbDados!Prazo
                    ObjPlan2Excel.Range("C" & LinhaDiferimento) = TbDados!Data_Venc
                    ObjPlan2Excel.Range("D" & LinhaDiferimento) = TbDados!PPB_Bruto
                
                        DiferimentoDiario = (TbDados!PPB_Bruto / TbDados!Prazo): DataVencimento = TbDados!Data_Venc: DataOperação = TbDados!Data_op
                
                     ObjPlan2Excel.Range("E" & LinhaDiferimento) = DiferimentoDiario
            
                     If DataOperação < DataInicioPlan Then
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
              Percentual = Round(((Contador / Qntd) * 100), 0)
             Call AtualizarStatus(Etapa, Percentual, unidade)
            Loop
                
                TbDados.MoveLast: TbLinha = TbDados.RecordCount: UltimaLinha = TbLinha + 4
                
                  ObjPlan2Excel.Rows(UltimaLinha & ":1000000").Delete Shift:=xlUp
                  ObjPlan2Excel.Activate
                  ObjPlan2Excel.Columns("A:AX").Select
                  ObjPlan2Excel.Columns.AutoFit
                                
                   Nome = "PPB MEXICHEM - " & MesData & " - " & Right(Date, 4)
                      
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





