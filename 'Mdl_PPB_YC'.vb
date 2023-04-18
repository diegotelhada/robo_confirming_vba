'Mdl_PPB_YC'


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

                        TbDataIni.Close
                        TbDatafin.Close
                        
                        Set TbDataIni = Nothing
                        Set TbDatafin = Nothing

                            BDRELocal.Execute ("INSERT INTO Tbl_PPB_YC ( Agencia_Ancora, Convenio_Ancora, Nome_Ancora, Cnpj_Ancora, Data_op, Nome_Fornecedor, Cnpj_Fornecedor, COD_OPERACAO, Numero_Nota, Data_Venc, Valor_Nota, Taxa, Prazo, Valor_Juros, Custo, Valor_Banco, PPB_Bruto, CHAVE )" _
                            & " SELECT TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Juros, TblArqoped.Prazo_NF, TblArqoped.Valor_Juros, TblArqoped.Custo, TblArqoped.Receita_Banco, TblArqoped.Receita_Clte, ([TblArqoped]![Data_op] & [TblArqoped]![Convenio_Ancora] & [TblArqoped]![Cod_Oper] & [TblArqoped]![Compromisso]) AS CHAVE FROM TblArqoped" _
                            & " WHERE (((TblArqoped.Data_op)>=#" & DataInic & "# And (TblArqoped.Data_op)<=#" & DataFin & "#) AND ((TblArqoped.Receita_Clte)>'0')) ORDER BY TblArqoped.Data_op;")
                
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
                                'Set TbCalendario = CurrentDb.OpenRecordset("SELECT Data_dia, Semana, Feriado, Tipo FROM [C:\Temp\Relatorios Confirming.mdb].TblCalendario WHERE (((Data_dia) Like '*" & DataDaPesquisa & "*')) ORDER BY Data_dia DESC")
                                    Resultado(Qntd) = TbCalendario!Data_dia
                                    MesPesq = MesPesq + 1
                                If MesPesq = "13" Then
                                    MesPesq = "01"
                                    AnoPesq = AnoPesq + 1
                                End If
                                TbCalendario.Close
                                Set TbCalendario = Nothing
                            Qntd = Qntd + 1
                        If Qntd = 20 Then: Exit Do
                    Loop
            
            Set TbDiferimento = BDPPBYC.OpenRecordset("Tbl_Diferimento_PPB_YC", dbOpenDynaset)

            Set TbDados = BDPPBYC.OpenRecordset("SELECT Tbl_PPB_YC.Agencia_Ancora, Tbl_PPB_YC.Convenio_Ancora, Tbl_PPB_YC.Nome_Ancora, Tbl_PPB_YC.Data_Op, Tbl_PPB_YC.Nome_Fornecedor, Tbl_PPB_YC.CNPJ_Fornecedor, Tbl_PPB_YC.COD_OPERACAO, Tbl_PPB_YC.Numero_Nota, Tbl_PPB_YC.Data_Venc, Tbl_PPB_YC.Valor_Nota, Tbl_PPB_YC.Taxa, Tbl_PPB_YC.Prazo, Tbl_PPB_YC.Valor_Juros, Tbl_PPB_YC.Custo, Tbl_PPB_YC.Valor_Banco, Tbl_PPB_YC.PPB_Bruto, Tbl_PPB_YC.CHAVE FROM Tbl_PPB_YC WHERE (((Tbl_PPB_YC.Data_op) >= #" & DataInic & "# And (Tbl_PPB_YC.Data_op) <= #" & DataFin & "#)) ORDER BY Tbl_PPB_YC.Data_Op;", dbOpenDynaset)

            If TbDados.EOF = False Then

                 TbDados.MoveLast: Quantidade = TbDados.RecordCount: TbDados.MoveFirst
                 
                    Do While TbDados.EOF = False
                
                        Set TbDados1 = BDPPBYC.OpenRecordset("SELECT Tbl_Diferimento_PPB_YC.Código, Tbl_Diferimento_PPB_YC.CHAVE, Tbl_Diferimento_PPB_YC.Data_Diferimento, Tbl_Diferimento_PPB_YC.Valor_Diferimento FROM Tbl_Diferimento_PPB_YC WHERE (((Tbl_Diferimento_PPB_YC.CHAVE)='" & TbDados!CHAVE & "'));", dbOpenDynaset)
                        'Set TbDados1 = CurrentDb.OpenRecordset("SELECT Tbl_Diferimento_PPB_YC.Código, Tbl_Diferimento_PPB_YC.CHAVE, Tbl_Diferimento_PPB_YC.Data_Diferimento, Tbl_Diferimento_PPB_YC.Valor_Diferimento FROM [\\saont46\apps2\Confirming\BD_PPB_YC\BD_PPB_YC.mdb].Tbl_Diferimento_PPB_YC WHERE (((Tbl_Diferimento_PPB_YC.CHAVE)='" & TbDados!CHAVE & "'))")
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
                    
                    If Not TbDados1 Is Nothing Then
                        TbDados1.Close
                        Set TbDados1 = Nothing
                    End If
                    
                TbDados.MoveNext
               Contador = Contador + 1
            Loop
        End If
    
End Function
Sub RelatorioPPBGerencial()

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim TbDados As Recordset, Nome As String
    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
    linha = 4
    
        MensalMes = Mid(Mensal, 4, 2)
        MensalMes = MensalMes - 1
        MensalAno = Right(Mensal, 2)
        If MensalMes = 0 Then
            MensalMes = 12
            MensalAno = MensalAno - 1
        End If
                    
        pesqmensal = MensalMes & "/20" & MensalAno
        pesqmensal = Format(pesqmensal, "MM/YYYY")
    
        Call AbrirBDPPBYC
    
            Set TbDados = BDPPBYC.OpenRecordset("TRANSFORM Sum(Tbl_Diferimento_PPB_YC.Valor_Diferimento) AS SomaDeValor_Diferimento SELECT Tbl_PPB_YC.Cnpj_Ancora FROM Tbl_PPB_YC INNER JOIN Tbl_Diferimento_PPB_YC ON Tbl_PPB_YC.CHAVE = Tbl_Diferimento_PPB_YC.CHAVE" _
            & " WHERE (((Tbl_Diferimento_PPB_YC.Data_Diferimento) Like '*" & pesqmensal & "*')) GROUP BY Tbl_PPB_YC.Cnpj_Ancora ORDER BY Tbl_Diferimento_PPB_YC.Data_Diferimento PIVOT Tbl_Diferimento_PPB_YC.Data_Diferimento;", dbOpenDynaset)
            
                If TbDados.EOF = False Then
                
                    Set ObjExcel = CreateObject("EXCEL.application")
                    ObjExcel.Workbooks.Open FileName:=Caminho & "MascaraPPBGerencial.xlsx", ReadOnly:=True
                    Set ObjPlan1Excel = ObjExcel.Worksheets("PPB")
                    ObjPlan1Excel.Activate
                                                            
                    'ObjExcel.Visible = True
                                                            
                    ObjPlan1Excel.Range("B1") = pesqmensal
                    ObjPlan1Excel.Range("A4").CopyFromRecordset TbDados
                    TbDados.MoveLast: UltimaLinha = TbDados.RecordCount + 4
                    
                    
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(7).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(8).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(9).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(10).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(11).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(1).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(2).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(3).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).Borders(4).LineStyle = 2
                        ObjPlan1Excel.Range("A" & linha & ":A" & UltimaLinha).NumberFormat = "00000000000000"
                        ObjPlan1Excel.Range("B" & linha & ":B" & UltimaLinha).Style = "Currency"
                        ObjPlan1Excel.Range("B" & UltimaLinha) = "=SUM(B4:B" & UltimaLinha - 1 & ")"
                        ObjPlan1Excel.Range("B" & UltimaLinha).Style = "Currency"
                        ObjPlan1Excel.Range("B" & UltimaLinha).Font.Bold = True
                        ObjPlan1Excel.Columns("A:B").Select
                        ObjPlan1Excel.Columns.AutoFit

                    Data = Date
                    Nome = "Relatorio PPB Gerencial - " & Format(Data, "ddmmyy")
                    Nome = Trata_NomeArquivo(Nome)
                   
                        sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                        'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                        
                        If (Dir(sFname) <> "") Then
                            Kill sFname
                        End If
                      
                    ObjPlan1Excel.SaveAs FileName:=sFname
                    ObjExcel.activeworkbook.Close SaveChanges:=False
                    ObjExcel.Quit
                
                    TbDados.Close
                
                'GoTo PuloEmail
                
        '=============ENVIO DO RELATORIO
                    Set ofs = CreateObject("Scripting.FileSystemObject")
                
                        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                        EmailDestino = "emanuela.conceicao@santander.com.br"
                        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                
                        Set sbObj = New Scripting.FileSystemObject
                        Set olapp = CreateObject("Outlook.Application")
                        Set oitem = olapp.CreateItem(0)
                
                            oitem.Subject = ("RELATORIO PPB GERENCIAL MENSAL")
                            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
                            oitem.To = "jorge.junior@santander.com.br;lfrossi@santander.com.br;wellington.da.silva@santander.com.br"
                            oitem.cc = "emanuela.conceicao@santander.com.br"
                
                            Corpo1 = "Prezados,"
                            Corpo2 = "Segue anexo o "
                            Relatorio = "RELATÓRIO DE PPB GERENCIAL"
                            Corpo21 = " referente ao(s) convênio(s) de Confirming® Santander."
                            Corpo3 = "Em caso de dúvida entrar em contato com a "
                            Confirming = "Confirming Desk"
                            Corpo31 = " ou através do(s) tel(s):"
                            Fones = " 0800-725-8090 / (11)3012-6390."
                            Corpo4 = "Atenciosamente,"
                            Assinatura1 = "Confirming®"
                            Assinatura2 = "Global Transaction Banking"
                            Assinatura3 = "Av. Juscelino Kubitschek, 2.235"
                            Assinatura4 = "Meios - Operações e Serviços"
                            Assinatura5 = "CEP: 04543-011  São Paulo-SP"
                            Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
                            Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
                            Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
                                     
                         oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Calibri <BR>" & Corpo1 & "<BR/>" & _
                         "<BR>" & Corpo2 & "<B>" & Relatorio & "</B>" & Corpo21 & "<BR>" & "<BR>" & Corpo3 & "<B>" & Confirming & "</B>" & Corpo31 & "<B>" & Fones & "</B>" & "<BR>" & "<BR>" & Corpo4 & "<BR><BR><BR>" & _
                         " <img src=" & Assinatura & " height=50 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 3" & "<BR>" & _
                         "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 2 <BR>" & Assinatura3 & _
                         "<BR/>" & Assinatura4 & "<BR/>" & Assinatura5 & "<BR/></FONT><FONT COLOR = BLACK FACE = Calibri Size = 1 <BR><I>" & Assinatura6 & _
                         "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"
                         
                         oitem.Attachments.Add File
                         oitem.Attachments.Add "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                        
                         oitem.Send
                            
                    'oitem.DISPLAY True
                    Set olapp = Nothing
                    Set oitem = Nothing
                End If

     
PuloEmail:
                            

End Sub
