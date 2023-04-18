'mod_GerarPPBYC'

Option Compare Database

Global BDRELocal  As Database
Global BDPPBYC As Database

Public Function AbrirBDLocal()
    'Set BDRELocal = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
    Set BDRELocal = CurrentDb
End Function

Function AbrirBDPPBYC()
    'Set BDPPBYC = OpenDatabase("\\saont46\apps2\Confirming\BD_PPB_YC\BD_PPB_YC.mdb")
    Set BDPPBYC = OpenDatabase("C:\Temp\BD_PPB_YC.mdb")
End Function

Public Function CalcDiferimento(Diferimento As Double, DataDif As Integer) As Double
    CalcDiferimento = (Diferimento * DataDif)
End Function

Sub Adc_PPB_YC()

    Dim Db As Database
    Dim TbDados As Recordset, TbDataIni As Recordset, TbDatafin As Recordset

        Call AbrirBDLocal

            MesPesq = Mid(Format(Date, "dd/mm/yyyy"), 4, 2): AnoPesq = Right(Format(Date, "dd/mm/yyyy"), 4)
    
                DataPesq = MesPesq & "/" & AnoPesq: DataPesq = Format(DataPesq, "mm/yyyy")

                    Set TbDataIni = BDRELocal.OpenRecordset("SELECT TblCalendario.Data_dia FROM TblCalendario WHERE (((TblCalendario.Data_dia) Like '*" & DataPesq & "*')) ORDER BY TblCalendario.Data_dia;", dbOpenDynaset)
                    Set TbDatafin = BDRELocal.OpenRecordset("SELECT TblCalendario.Data_dia FROM TblCalendario WHERE (((TblCalendario.Data_dia) Like '*" & DataPesq & "*')) ORDER BY TblCalendario.Data_dia DESC;", dbOpenDynaset)

                        DatainiPPB = TbDataIni!Data_dia: DataFinPPB = TbDatafin!Data_dia
                        DataInic = Format(DatainiPPB, "mm/dd/YYYY"): DataFin = Format(DataFinPPB, "mm/dd/YYYY")

                            BDRELocal.Execute ("INSERT INTO Tbl_PPB_YC ( Agencia_Ancora, Convenio_Ancora, Nome_Ancora, Cnpj_Ancora, Data_op, Nome_Fornecedor, Cnpj_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, Custo, Valor_Banco, PPB_Bruto, CHAVE )" _
                            & " SELECT TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Juros, TblArqoped.Prazo_NF, TblArqoped.Valor_Juros, TblArqoped.Custo, TblArqoped.Receita_Banco, TblArqoped.Receita_Clte, ([TblArqoped]![Data_op] & [TblArqoped]![Convenio_Ancora] & [TblArqoped]![Cod_Oper] & [TblArqoped]![Compromisso]) AS CHAVE FROM TblArqoped" _
                            & " WHERE (((TblArqoped.Data_op)>=#" & DataInic & "# And (TblArqoped.Data_op)<=#" & DataFin & "#) AND ((TblArqoped.Receita_Clte)>'0')) ORDER BY TblArqoped.Data_op;")
                
                    Set TbDataIni = Nothing
                    Set TbDatafin = Nothing
                
                Call Diferimento_YC_V2(DatainiPPB, DataFinPPB)

End Sub
Function Diferimento_YC_V2(DataInicio, DataFinal)

    Dim TbDados1, TbCalendario, TbDados, TbVenc, TbDados2, TbData, TbDataOp, TbPrazo, TbPPB As DAO.Recordset
    Dim DataPesq As Date, Y As Integer, SiglaPesq As String
    Dim ValorJuros As Double, DiferimentoDiario  As Double, ValorDif(500) As Double
    Dim Caminho, MesData, LinhaDiferimento, UltimoVencimento, MesVenc, MesPesq, Limite, AnoPesq, Datames As String
    Dim cont, Qntd As Integer, Percentual As Double, unidade As Double, Etapa As String
    Dim DataDaPesquisa, Resultado(24), DataVencimento, DataOperação As Date
    Dim Data(500) As Date, DataInicioPlan As Date, DiasCalc(500) As Integer

        Call AbrirBDLocal
                    
        Call AbrirBDPPBYC

            DataInic = Format(DataInicio, "mm/dd/YYYY")
            DataFin = Format(DataFinal, "mm/dd/YYYY")

                DataInicioPlan = DataInicio: cont = 0: Qntd = 1: DataDaPesquisa = DataInicio

                Datames = Right(DataDaPesquisa, 7): MesPesq = Mid(DataDaPesquisa, 4, 2): AnoPesq = Right(DataDaPesquisa, 4)

                    Do While True
                        MesPesq = Format(MesPesq, "00")
                            DataDaPesquisa = MesPesq & "/" & AnoPesq
                                Set TbCalendario = BDRELocal.OpenRecordset("SELECT TblCalendario.Data_dia, TblCalendario.Semana, TblCalendario.Feriado, TblCalendario.Tipo FROM TblCalendario WHERE (((TblCalendario.Data_dia) Like '*" & DataDaPesquisa & "*')) ORDER BY TblCalendario.Data_dia DESC;", dbOpenDynaset)
                                    Resultado(Qntd) = TbCalendario!Data_dia
                                    MesPesq = MesPesq + 1
                                TbCalendario.Close
                                Set TbCalendario = Nothing
                                If MesPesq = "13" Then
                                    MesPesq = "01"
                                    AnoPesq = AnoPesq + 1
                                End If
                            Qntd = Qntd + 1
                        If Qntd = 20 Then: Exit Do
                    Loop
            
            Set TbDiferimento = BDPPBYC.OpenRecordset("Tbl_Diferimento_PPB_YC", dbOpenDynaset)

            Set TbDados = BDPPBYC.OpenRecordset("SELECT Tbl_PPB_YC.Agencia_Ancora, Tbl_PPB_YC.Convenio_Ancora, Tbl_PPB_YC.Nome_Ancora, Tbl_PPB_YC.Data_Op, Tbl_PPB_YC.Nome_Fornecedor, Tbl_PPB_YC.CNPJ_Fornecedor, Tbl_PPB_YC.COD_OPERACAO, Tbl_PPB_YC.Numero_Nota, Tbl_PPB_YC.Data_Venc, Tbl_PPB_YC.Valor_Nota, Tbl_PPB_YC.Taxa, Tbl_PPB_YC.Prazo, Tbl_PPB_YC.Valor_Juros, Tbl_PPB_YC.Custo, Tbl_PPB_YC.Valor_Banco, Tbl_PPB_YC.PPB_Bruto, Tbl_PPB_YC.CHAVE FROM Tbl_PPB_YC WHERE (((Tbl_PPB_YC.Data_op) >= #" & DataInic & "# And (Tbl_PPB_YC.Data_op) <= #" & DataFin & "#)) ORDER BY Tbl_PPB_YC.Data_Op;", dbOpenDynaset)

            If TbDados.EOF = False Then

                 TbDados.MoveLast: Quantidade = TbDados.RecordCount: TbDados.MoveFirst
                 
                    Do While TbDados.EOF = False
                
                        Set TbDados1 = BDPPBYC.OpenRecordset("SELECT Tbl_Diferimento_PPB_YC.Código, Tbl_Diferimento_PPB_YC.CHAVE, Tbl_Diferimento_PPB_YC.Data_Diferimento, Tbl_Diferimento_PPB_YC.Valor_Diferimento FROM Tbl_Diferimento_PPB_YC WHERE (((Tbl_Diferimento_PPB_YC.CHAVE)='" & TbDados!CHAVE & "'));", dbOpenDynaset)
    
                            If TbDados1.EOF = False Then: BDPPBYC.Execute ("DELETE Tbl_Diferimento_PPB_YC.Código, Tbl_Diferimento_PPB_YC.CHAVE, Tbl_Diferimento_PPB_YC.Data_Diferimento, Tbl_Diferimento_PPB_YC.Valor_Diferimento FROM Tbl_Diferimento_PPB_YC WHERE (((Tbl_Diferimento_PPB_YC.CHAVE)='" & TbDados!CHAVE & "'));")

                                DiferimentoDiario = (TbDados!PPB_Bruto / TbDados!Prazo): DataVencimento = TbDados!Data_Venc: DataOperação = TbDados!Data_op
    
                                    If DataOperação < DataInicioPlan Then
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
                                            TbDiferimento!Nome_Ancora = TbDados!Nome_Ancora
                                            TbDiferimento!CHAVE = TbDados!CHAVE
                                            TbDiferimento!Data_Diferimento = Resultado(1)
                                            TbDiferimento!Valor_Diferimento = ValorDif(1)
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
                                    TbDiferimento!Nome_Ancora = TbDados!Nome_Ancora
                                    TbDiferimento!CHAVE = TbDados!CHAVE
                                    TbDiferimento!Data_Diferimento = Resultado(QntdResul + 1)
                                    TbDiferimento!Valor_Diferimento = ValorDif(ValorResul)
                                TbDiferimento.Update
                            End If
        
                            QntdResul = QntdResul + 1: DataResul = DataResul + 1: ValorResul = ValorResul + 1
        
                        If ValorResul = 14 Then: Exit Do
                    Loop
    
                TbDados.MoveNext
               Contador = Contador + 1
            Loop
        End If
    
End Function

