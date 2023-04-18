'Mdl_Relatorios'

Sub RelatorioInditex()

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim TbDados As Recordset, Nome As String
    Dim TbEnviar As Recordset
        
        'Abrir base local
        Call AbrirBDLocal

        MensalMes = Mid(Mensal, 4, 2)               '
        MensalMes = MensalMes - 1
        MensalAno = Right(Mensal, 2)
        If MensalMes = 0 Then
            MensalMes = 12
            MensalAno = MensalAno - 1
        End If
                    '
        pesqmensal = MensalMes & "/20" & MensalAno    '
        pesqmensal = Format(pesqmensal, "MM/YYYY")

            Set TbDados = BDRELocal.OpenRecordset("SELECT TblArqoped.Compromisso, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, 'BRL' AS Divisadelpago, TblArqoped.Valor_Nom, TblArqoped.Data_Venc, TblArqoped.Data_op, CInt([TblArqoped]![Prazo_NF]) AS Prazo_NF, TblArqoped.Juros, 0 AS Comision, TblArqoped.Valor_Juros, '0' AS Comisionliquidadaalproveedor, 'BRL' AS Divisadebonificacion, CDbl([TblArqoped]![Receita_Clte]) AS Receita_Clte, 0 AS Bonifcomisiones, 'Brazil' AS Origenproveedor" _
            & " FROM Tbl_GrupoInditex INNER JOIN TblArqoped ON (Tbl_GrupoInditex.Convenio_Ancora = TblArqoped.Convenio_Ancora) AND (Tbl_GrupoInditex.Agencia_Ancora = TblArqoped.Agencia_Ancora) WHERE (((TblArqoped.Data_op) Like '*" & pesqmensal & "*'));", dbOpenDynaset)
    
                If TbDados.EOF = True Then GoTo Fim
 
                Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
            
                    Set ObjExcel = CreateObject("EXCEL.application")
                    ObjExcel.Workbooks.Open FileName:=Caminho & "Inditex.xlsx", ReadOnly:=True
                    Set ObjPlan1Excel = ObjExcel.Worksheets("BonificacionesBRL")
                    linha = 6
                   
                        ObjPlan1Excel.Range("B1") = "Liquidación de Reparto - " & MesesEspanhol(CStr(Format(MensalMes, "00"))) & "/" & Format(Mensal, "YYYY")
                        Titulo = "01/" & Format(Date, "mm/yyyy")
                        ObjPlan1Excel.Range("B2") = Format(Titulo, "mm/dd/yyyy")
                        ObjPlan1Excel.Range("A6").CopyFromRecordset TbDados
                        
                        UltimaLinha = TbDados.RecordCount
                        UltimaLinha = UltimaLinha + 5
                   
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(7).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(8).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(9).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(10).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(11).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(1).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(2).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(3).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Borders(4).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":P" & UltimaLinha).Font.Size = 8
                            ObjPlan1Excel.Range("F" & linha & ":G" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                            ObjPlan1Excel.Range("A" & linha & ":A" & UltimaLinha).NumberFormat = "00000"
                            ObjPlan1Excel.Range("B" & linha & ":B" & UltimaLinha).NumberFormat = "00000000000000"
                            ObjPlan1Excel.Range("E" & linha & ":E" & UltimaLinha).Style = "Currency"
                            ObjPlan1Excel.Range("K" & linha & ":K" & UltimaLinha).Style = "Currency"
                            ObjPlan1Excel.Range("N" & linha & ":N" & UltimaLinha).Style = "Currency"
                            ObjPlan1Excel.Range("I" & linha & ":I" & UltimaLinha).Style = "Percent"
                            ObjPlan1Excel.Range("I" & linha & ":I" & UltimaLinha).NumberFormat = "0.00000%"
                            ObjPlan1Excel.Columns("A:P").Select
                            ObjPlan1Excel.Columns.AutoFit
                        
                        ObjPlan1Excel.Rows("6:" & UltimaLinha).RowHeight = 11.75
                        ObjPlan1Excel.Range("A1").Select
                
                   Data = Date
                   Nome = "Relatorio de Notas Antecipadas - INDITEX - " & Format(Data, "ddmmyy")
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
                
                'GoTo lblPuloEmail
                
        '=============ENVIO DO RELATORIO
        
                Set TbEnviar = BDRELocal.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
                
                    Set ofs = CreateObject("Scripting.FileSystemObject")
                
                        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                
                            EmailDestino = "EnriqueFD@inditex.com;NataliaLB@inditex.com;alejandrogrodri@inditex.com;elienejgr@br.inditex.com;leticiafst@br.inditex.com;milenkaga@itxtrading.com;luciench@itxtrading.com;controlling@itxtrading.com"
                            EmailCopia = "josealjunior@santander.com.br;vitosantos@santander.com.br;javier.hernandezal@servexternos.gruposantander.com"
                            Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                
                    Set sbObj = New Scripting.FileSystemObject
                    Set olapp = CreateObject("Outlook.Application")
                    Set oitem = olapp.CreateItem(0)
                
                        oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS - INDITEX")
                        
                        oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
                        
                        oitem.To = EmailDestino
                        oitem.cc = EmailCopia
                
                            Corpo1 = "Prezado Cliente,"
                            Corpo2 = "Segue anexo o "
                            Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
                        
                        TbEnviar.Close
                        
lblPuloEmail:
                        
                        Set TbDados = BDRELocal.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
                
                            TbDados.AddNew
                                TbDados!Agencia_Ancora = Agencia_Ancora
                                TbDados!Convenio_Ancora = Convenio_Ancora
                                TbDados!Cnpj_Ancora = CNPJcliente
                                TbDados!Nome_Ancora = NomeCliente
                                TbDados!Relatorio_Enviado = "Notas Antecipadas - Inditex"
                                TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                                TbDados!Data_Envio = Date
                                TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                                TbDados!Usuario = NomeUsuario
                                TbDados.Update
                            TbDados.Close
                
Fim:

End Sub
Sub RelatNotasAntComTaxa_SPREAD_CUSTO_PPB() ' Relatorios de Notas Antecipadas (Com Taxa)

'On Error GoTo Listagem

    Dim Db1 As Database, TbUsuarios As Recordset
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double, TbDados As Recordset
    Dim DataPesq As Date, TbData As Recordset

        SiglaPesq = UCase(PesqUsername)

            Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

                Set TbUsuarios = Db1.OpenRecordset("TblUsuarios", dbOpenDynaset)

                    TbUsuarios.FindFirst "Sigla like '*" & SiglaPesq & "'"
                        
                    If TbUsuarios.NoMatch = False Then
                        NomeUsuario = TbUsuarios!Nome
                        EmailUsuario = TbUsuarios!Email
                    End If

        '-----------------------------------------------------------------------------------------------------
        MensalMes = Mid(Mensal, 4, 2)
        MensalMes = MensalMes - 1
        MensalAno = Right(Mensal, 2)
        If MensalMes = 0 Then
            MensalMes = 12
            MensalAno = MensalAno - 1
        End If
                    '
        pesqmensal = MensalMes & "/20" & MensalAno
        pesqmensal = Format(pesqmensal, "MM/YYYY")
            
            Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
                DiarioPesq = Date - 1
                    Do While True
                        TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                            If TbData.NoMatch = False Then
                                If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                                If TbData!Tipo = "UTIL" Then: Exit Do
                            End If
                    Loop
            TbData.Close
        '-----------------------------------------------------------------------------------------------------
        'Alterado em 10/12/2018 -
                
            vString = "SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros, CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]) AS CustoAncora, TblArqoped.Receita_Clte, (CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]))/(CInt([Prazo_NF])*CDbl([Valor_Pagmto])/30) AS TaxaAncora "
            vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
            
            vCriterio = ""
            
            If PeriodicidadeRel = "Diario" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros, CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]) AS CustoAncora, TblArqoped.Receita_Clte, (CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]))/(CInt([Prazo_NF])*CDbl([Valor_Pagmto])/30) AS TaxaAncora" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) Like '*" & DiarioPesq & "*')) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) Like '*" & DiarioPesq & "*')) ORDER BY TblArqoped.Data_op; "
            ElseIf PeriodicidadeRel = "Semanal" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros, CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]) AS CustoAncora, TblArqoped.Receita_Clte, (CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]))/(CInt([Prazo_NF])*CDbl([Valor_Pagmto])/30) AS TaxaAncora" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) = Date()-7)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) >= Date()-7)) ORDER BY TblArqoped.Data_op; "
            ElseIf PeriodicidadeRel = "Mensal" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros, CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]) AS CustoAncora, TblArqoped.Receita_Clte, (CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]))/(CInt([Prazo_NF])*CDbl([Valor_Pagmto])/30) AS TaxaAncora" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) Like '*" & pesqmensal & "*')) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) Like '*" & pesqmensal & "*')) ORDER BY TblArqoped.Data_op; "
            ElseIf PeriodicidadeRel = "Quinzenal" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros, CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]) AS CustoAncora, TblArqoped.Receita_Clte, (CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]))/(CInt([Prazo_NF])*CDbl([Valor_Pagmto])/30) AS TaxaAncora" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) = Date()-15)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) >= Date()-15)) ORDER BY TblArqoped.Data_op; "
            End If

            If vCriterio = "" Then GoTo Fim
            Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)

            If TbDados1.EOF = True Then GoTo Fim
                
            Agencia_Ancora = TbDados1!Agencia_Ancora
            Convenio_Ancora = TbDados1!Convenio_Ancora
            Cnpj_Ancora = TbDados1!Cnpj_Ancora

                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                        Set ObjExcel = CreateObject("EXCEL.application")
                        ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasComTaxaPPB.xlsx", ReadOnly:=True
                        Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                            NomeCliente = TbDados1!Nome_Ancora
                            CNPJcliente = TbDados1!Cnpj_Ancora
                            dataPlan = Format(Date, "MM/DD/YYYY")
                   
                        linha = 9
                   
                        ObjPlan1Excel.Range("C2") = NomeCliente
                        ObjPlan1Excel.Range("C4") = dataPlan
                        ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
                        ObjPlan1Excel.Range("K2") = CNPJcliente
                        ObjPlan1Excel.Range("K2").NumberFormat = "00000"
                        ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
                        
                        UltimaLinha = TbDados1.RecordCount
                        UltimaLinha = UltimaLinha + 8
                   
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(7).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(8).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(9).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(10).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(11).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(1).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(2).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(3).LineStyle = 2
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Borders(4).LineStyle = 2
                            ObjPlan1Excel.Range("C" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
                            ObjPlan1Excel.Range("A" & linha & ":Z" & UltimaLinha).Font.Size = 8
                            ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
                            ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
                            ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                            ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
                            ObjPlan1Excel.Range("T" & linha & ":V" & UltimaLinha).Style = "Currency"
                            ObjPlan1Excel.Range("W" & linha & ":W" & UltimaLinha).NumberFormat = "00000"
                            ObjPlan1Excel.Range("X" & linha & ":Y" & UltimaLinha).Style = "Currency"
                            ObjPlan1Excel.Range("Z" & linha & ":Z" & UltimaLinha).Style = "Percent"
                            ObjPlan1Excel.Range("Z" & linha & ":Z" & UltimaLinha).NumberFormat = "0.00000%"
                            
                            ObjPlan1Excel.Columns("A:Z").Select
                            ObjPlan1Excel.Columns.AutoFit
                            ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
                            ObjPlan1Excel.Range("C2").Select
                
                   Data = Date
                   Nome = "Relatorio de Notas Antecipadas - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
                   Nome = Trata_NomeArquivo(Nome)
                   
                    sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                    'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                    
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                
            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname


        'ACS - Pular envio de e-mail
        'GoTo lblPuloEmail

        'Enviar Arquivo
        '-----------------------------------------------------------------------------------------------------
                Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
                
                    Set ofs = CreateObject("Scripting.FileSystemObject")
                
                        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                
                            EmailDestino = TbDados1!Email_Ancora
                            EmailCopia = TbDados1!Email_Trader
                            Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                            
                    'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
                    Dim nomeRelatorio As String
                    nomeRelatorio = "Notas Antecipadas(Com Taxa/Custo/PPB)"
                    Dim tbMailPersonalizado As Recordset
                    SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
                    "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
                    "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
                    Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
                    
                    If Not tbMailPersonalizado.EOF Then
                        EmailDestino = tbMailPersonalizado!Email_Para
                        EmailCopia = tbMailPersonalizado!Email_Copia
                    End If
                    
                    tbMailPersonalizado.Close
                    Set tbMailPersonalizado = Nothing
                
                    Set sbObj = New Scripting.FileSystemObject
                    Set olapp = CreateObject("Outlook.Application")
                    Set oitem = olapp.CreateItem(0)
                
                        oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente)
                        'oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
                        
                        oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
                        
                        oitem.To = EmailDestino
                        oitem.cc = EmailCopia
                
                            Corpo1 = "Prezado Cliente,"
                            Corpo2 = "Segue anexo o "
                            Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
                        
                        TbDados1.Close
                        
lblPuloEmail:
                        
                        Set TbDados = Db1.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
                
                        ''Inserir aqui todos os registros da AuxConvenio
                        Dim rsDao As DAO.Recordset
                        Set rsDao = Db1.OpenRecordset("Select * From AuxConvenio", 4)
                        
                        If rsDao.EOF = False Then rsDao.MoveFirst
                        Do While rsDao.EOF = False
                        
                            TbDados.AddNew
                                TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                                TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                                TbDados!Cnpj_Ancora = CNPJcliente
                                TbDados!Nome_Ancora = NomeCliente
                                TbDados!Relatorio_Enviado = "Notas Antecipadas(Com Taxa)"
                                TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                                TbDados!Data_Envio = Date
                                TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                                TbDados!Usuario = NomeUsuario
                            TbDados.Update
                        
                            rsDao.MoveNext
                        Loop
                        
                        rsDao.Close
                        Set rsDao = Nothing
            
                        TbDados.Close
                
Fim:

End Sub
Sub RelatNotasAntComTaxa() ' Relatorios de Notas Antecipadas (Com Taxa)

    'On Error GoTo Listagem
    
    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
    
    
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
        'Alterado em 07/12/2018 - ACS
        vString = "Select TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        
        vCriterio = ""
        
            If PeriodicidadeRel = "Diario" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) Like '*" & DiarioPesq & "*'))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) Like '*" & DiarioPesq & "*')) ORDER BY TblArqoped.Data_op;"
            ElseIf PeriodicidadeRel = "Semanal" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) >=Date()-7))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) >=Date()-7)) ORDER BY TblArqoped.Data_op;"
            ElseIf PeriodicidadeRel = "Mensal" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) Like '*" & pesqmensal & "')) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) Like '*" & pesqmensal & "')) ORDER BY TblArqoped.Data_op;"
            ElseIf PeriodicidadeRel = "Quinzenal" Then
                'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros" _
                '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora) =" & Agencia_Ancora & ") And ((TblArqoped.Convenio_Ancora) ='" & Convenio_Ancora & "') And ((TblArqoped.Data_op) >=Date()-15))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                vCriterio = " WHERE (((TblArqoped.Data_op) >=Date()-15))ORDER BY TblArqoped.Data_op;"
            End If
            
            If vCriterio = "" Then GoTo Fim
            Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)
            
            If TbDados1.EOF = True Then GoTo Fim
            
            Agencia_Ancora = TbDados1!Agencia_Ancora
            Convenio_Ancora = TbDados1!Convenio_Ancora
            Cnpj_Ancora = TbDados1!Cnpj_Ancora
            
    '==== GERA O RELATORIO EM EXCEL
    
        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                    
            Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasComTaxa.xlsx", ReadOnly:=True
            Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
            
            NomeCliente = TbDados1!Nome_Ancora
            CNPJcliente = TbDados1!Cnpj_Ancora
            dataPlan = Format(Date, "MM/DD/YYYY")
            
            linha = 9
            
            ObjPlan1Excel.Range("C2") = NomeCliente
            ObjPlan1Excel.Range("C4") = dataPlan
            ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("K2") = CNPJcliente
            ObjPlan1Excel.Range("K2").NumberFormat = "00000"
            ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
            
            UltimaLinha = TbDados1.RecordCount
            UltimaLinha = UltimaLinha + 8
            
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(7).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(8).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(9).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(10).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(11).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(1).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(2).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(3).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(4).LineStyle = 2
            ObjPlan1Excel.Range("C" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Font.Size = 8
            ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("T" & linha & ":V" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("W" & linha & ":W" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Columns("A:W").Select
            ObjPlan1Excel.Columns.AutoFit
            ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
            ObjPlan1Excel.Range("C2").Select
            
            Data = Date
            Nome = "Relatorio de Notas Antecipadas - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
            Nome = Trata_NomeArquivo(Nome)
            
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'Testes
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
            
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
                        
            'ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'ObjExcel.Visible = True
            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname
            


            'ACS - Pular envio de e-mail
            'GoTo lblPuloEmail

    '==== ENVIAR RELATORIO POR EMAIL
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
            Set ofs = CreateObject("Scripting.FileSystemObject")

            File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            
            EmailDestino = TbDados1!Email_Ancora
            EmailCopia = TbDados1!Email_Trader
            Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
            
            'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
            Dim nomeRelatorio As String
            nomeRelatorio = "Notas Antecipadas(Com Taxa)"
            Dim tbMailPersonalizado As Recordset
            SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
            "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
            "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
            Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
            
            If Not tbMailPersonalizado.EOF Then
                EmailDestino = tbMailPersonalizado!Email_Para
                EmailCopia = tbMailPersonalizado!Email_Copia
            End If
            
            tbMailPersonalizado.Close
            Set tbMailPersonalizado = Nothing
            
            Set sbObj = New Scripting.FileSystemObject
            Set olapp = CreateObject("Outlook.Application")
            Set oitem = olapp.CreateItem(0)
            
                    oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente)
                    'oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
                    
                    oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
                    oitem.To = EmailDestino
                    oitem.cc = EmailCopia & ""
            
            Corpo1 = "Prezado Cliente,"
            Corpo2 = "Segue anexo o "
            Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
        
        TbDados1.Close
        
lblPuloEmail:
        
    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas Antecipadas(Com Taxa)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Date
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
            
                rsDao.MoveNext
            Loop
            
            rsDao.Close
            Set rsDao = Nothing
            
            TbDados.Close
            

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatNotasAntSemTaxa() ' Relatorios de Notas Antecipadas (Sem Taxa)

'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
    
    
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
    
        'Alterado em 07/12/2018 - ACS
        vString = " SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        
        vCriterio = ""
    
        If PeriodicidadeRel = "Diario" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) Like '*" & DiarioPesq & "*'));", dbOpenDynaset)
            vCriterio = "  WHERE (((TblArqoped.Data_op) Like '*" & DiarioPesq & "*')); "
        ElseIf PeriodicidadeRel = "Semanal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) >=Date()-7))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = "  WHERE (((TblArqoped.Data_op) >=Date()-7))ORDER BY TblArqoped.Data_op; "
        ElseIf PeriodicidadeRel = "Mensal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) Like '*" & pesqmensal & "'))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = "  WHERE (((TblArqoped.Data_op) Like '*" & pesqmensal & "'))ORDER BY TblArqoped.Data_op; "
        ElseIf PeriodicidadeRel = "Quinzenal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) >=Date()-15))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = "  WHERE (((TblArqoped.Data_op) >=Date()-15))ORDER BY TblArqoped.Data_op; "
        End If
        
        If vCriterio = "" Then GoTo Fim
        Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)

        If TbDados1.EOF = True Then GoTo Fim

        Agencia_Ancora = TbDados1!Agencia_Ancora
        Convenio_Ancora = TbDados1!Convenio_Ancora
        Cnpj_Ancora = TbDados1!Cnpj_Ancora

    '==== GERA O RELATORIO EM EXCEL
        
        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
            
            Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasSemTaxa.xlsx", ReadOnly:=True
            Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
            
            NomeCliente = TbDados1!Nome_Ancora
            CNPJcliente = TbDados1!Cnpj_Ancora
            dataPlan = Format(Date, "MM/DD/YYYY")
            
            linha = 9
            
            ObjPlan1Excel.Range("C2") = NomeCliente
            ObjPlan1Excel.Range("C4") = dataPlan
            ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("K2") = CNPJcliente
            ObjPlan1Excel.Range("K2").NumberFormat = "00000"
            ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
            
            UltimaLinha = TbDados1.RecordCount
            UltimaLinha = UltimaLinha + 8
            
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(7).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(8).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(9).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(10).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(11).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(1).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(2).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(3).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(4).LineStyle = 2
            ObjPlan1Excel.Range("C" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Font.Size = 8
            ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("T" & linha & ":V" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Columns("A:V").Select
            ObjPlan1Excel.Columns.AutoFit
            ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
            ObjPlan1Excel.Range("C2").Select
            
            Data = Date
            Nome = "Relatorio de Notas Antecipadas1 - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
            Nome = Trata_NomeArquivo(Nome)
            
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If

            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname


'ACS - Pular envio de e-mail
'GoTo lblPuloEmail


    '==== ENVIAR RELATORIO POR EMAIL

        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados1!Email_Ancora
        EmailCopia = TbDados1!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Notas Antecipadas(Sem Taxa)"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
                
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
                
            oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente)
            'oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
        
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
        
        TbDados1.Close
                
lblPuloEmail:
                
    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas Antecipadas(Sem Taxa)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            
            rsDao.Close
            Set rsDao = Nothing
            
            TbDados.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatNotasAntSemTaxaJuros() ' Relatorios de Notas Antecipadas (Sem Taxa\Juros\Valor Liq)

'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
        
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
        'Alterado em 07/12/2018 - ACS
        vString = " SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        
        vCriterio = ""
    
        If PeriodicidadeRel = "Diario" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) Like '*" & DiarioPesq & "*'));", dbOpenDynaset)
            vCriterio = "  WHERE (((TblArqoped.Data_op) Like '*" & DiarioPesq & "*')); "
        ElseIf PeriodicidadeRel = "Semanal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) >=Date()-7))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op) >=Date()-7))ORDER BY TblArqoped.Data_op;"
        ElseIf PeriodicidadeRel = "Mensal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) Like '*" & pesqmensal & "'))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op) Like '*" & pesqmensal & "'))ORDER BY TblArqoped.Data_op; "
        ElseIf PeriodicidadeRel = "Quinzenal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) >=Date()-15))ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op) >=Date()-15))ORDER BY TblArqoped.Data_op; "
        End If
        
        If vCriterio = "" Then GoTo Fim
        Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)
        
        If TbDados1.EOF = True Then GoTo Fim
        
        Agencia_Ancora = TbDados1!Agencia_Ancora
        Convenio_Ancora = TbDados1!Convenio_Ancora
        Cnpj_Ancora = TbDados1!Cnpj_Ancora
        
    '==== GERA O RELATORIO EM EXCEL

        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
            
            Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasSemTaxaJuros.xlsx", ReadOnly:=True
            Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
            
            NomeCliente = TbDados1!Nome_Ancora
            CNPJcliente = TbDados1!Cnpj_Ancora
            dataPlan = Format(Date, "MM/DD/YYYY")
            
            linha = 9
            
            ObjPlan1Excel.Range("C2") = NomeCliente
            ObjPlan1Excel.Range("C4") = dataPlan
            ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("K2") = CNPJcliente
            ObjPlan1Excel.Range("K2").NumberFormat = "00000"
            ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
            
            UltimaLinha = TbDados1.RecordCount
            UltimaLinha = UltimaLinha + 8
            
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(7).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(8).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(9).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(10).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(11).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(1).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(2).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(3).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(4).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Font.Size = 8
            ObjPlan1Excel.Range("B" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("T" & linha & ":T" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Columns("A:T").Select
            ObjPlan1Excel.Columns.AutoFit
            ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
            ObjPlan1Excel.Range("C2").Select
            
            Data = Date
            Nome = "Relatorio de Notas Antecipadas2 - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
            Nome = Trata_NomeArquivo(Nome)
            
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If

            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname

'ACS - Pular envio de e-mail
'GoTo lblPuloEmail

    '==== ENVIAR RELATORIO POR EMAIL

        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados1!Email_Ancora
        EmailCopia = TbDados1!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Notas Antecipadas(Sem Taxa/Juros/Val Liq)"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
                
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
    
            oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente)
            'oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
                               
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
        
        'oitem.Display True
        Set olapp = Nothing
        Set oitem = Nothing
        TbDados1.Close
    
lblPuloEmail:
    
    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas Antecipadas(Sem Taxa/Juros/Val Liq)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            
            rsDao.Close
            Set rsDao = Nothing
                
            TbDados.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatNotasAvencSemTaxa() 'Relatorio de Notas a Vencer (Sem Taxa)

'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
        
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
        DataPesq = Format(DiarioPesq, "mm/dd/yyyy")
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
    'Alterado em 10/12/2018 - Para agrupamento de convênios em relatório.
    
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
    
        'Set TbDados = Db.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido" _
        '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_Venc)>#" & DataPesq & "#));", dbOpenDynaset)
        
        vString = "SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        vCriterio = "  WHERE (((TblArqoped.Data_Venc)>#" & DataPesq & "#)); "
        
        Set TbDados = Db.OpenRecordset(vString & vCriterio, dbOpenDynaset)
        
        If TbDados.EOF = True Then GoTo Fim
        
        Agencia_Ancora = TbDados!Agencia_Ancora
        Convenio_Ancora = TbDados!Convenio_Ancora
        Cnpj_Ancora = TbDados!Cnpj_Ancora
        
    '==== GERA O RELATORIO EM EXCEL

        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
        
        Set ObjExcel = CreateObject("EXCEL.application")
        ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAvencerSemTaxa.xlsx", ReadOnly:=True
        Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
        
        NomeCliente = TbDados!Nome_Ancora
        CNPJcliente = TbDados!Cnpj_Ancora
        dataPlan = Format(Date, "MM/DD/YYYY")
        
        linha = 9
        
        ObjPlan1Excel.Range("C2") = NomeCliente
        ObjPlan1Excel.Range("C4") = dataPlan
        ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("K2") = CNPJcliente
        ObjPlan1Excel.Range("K2").NumberFormat = "00000"
        ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
        
        UltimaLinha = TbDados.RecordCount
        UltimaLinha = UltimaLinha + 8
        
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(7).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(8).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(9).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(10).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(11).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(1).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(2).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(3).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Borders(4).LineStyle = 2
        ObjPlan1Excel.Range("C" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("A" & linha & ":V" & UltimaLinha).Font.Size = 8
        ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
        ObjPlan1Excel.Range("T" & linha & ":V" & UltimaLinha).Style = "Currency"
        ObjPlan1Excel.Columns("A:V").Select
        ObjPlan1Excel.Columns.AutoFit
        ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
        ObjPlan1Excel.Range("C2").Select
        
        Data = Date
        Nome = "Relatorio de Notas a Vencer - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
        Nome = Trata_NomeArquivo(Nome)
        
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
            
        'ObjPlan1Excel.SaveAs FileName:=sFname
        'ObjExcel.activeworkbook.Close SaveChanges:=False
        'ObjExcel.Quit
        'TbDados.Close

            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname

''Pular envio de e-mail
'GoTo lblPuloEmail

    '==== ENVIAR RELATORIO POR EMAIL

        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados1!Email_Ancora
        EmailCopia = TbDados1!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Notas a Vencer(Sem Taxa)"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
        
            oitem.Subject = ("RELATORIO DE NOTAS A VENCER -  " & NomeCliente & " - " & CNPJcliente)
            'oitem.Subject = ("RELATORIO DE NOTAS A VENCER -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
        
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS A VENCER"
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
        TbDados1.Close
                
lblPuloEmail:
            
    '==== ABRE O BANCO DE DADOS
    
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas a Vencer(Sem Taxa)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            rsDao.Close
            Set rsDao = Nothing
            
            TbDados.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatNotasAvencComTaxa() 'Relatorios de Notas a Vencer (Com Taxa)

'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
        
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
                
        DataPesq = Format(DiarioPesq, "mm/dd/yyyy")
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
        'Alterado em 10/12/2018 - Para agrupamento de convênios em relatório.
        
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido,TblArqoped.Juros " _
        '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_Venc)>#" & DataPesq & "#));", dbOpenDynaset)
        
        vString = "SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido,TblArqoped.Juros "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        vCriterio = " WHERE (((TblArqoped.Data_Venc)>#" & DataPesq & "#));"
        
        Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)
        
        If TbDados1.EOF = True Then GoTo Fim
        
        Agencia_Ancora = TbDados1!Agencia_Ancora
        Convenio_Ancora = TbDados1!Convenio_Ancora
        Cnpj_Ancora = TbDados1!Cnpj_Ancora
        
    '==== GERA O RELATORIO EM EXCEL
    
        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
            
            Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAvencerComTaxa.xlsx", ReadOnly:=True
            Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
            
            NomeCliente = TbDados1!Nome_Ancora
            CNPJcliente = TbDados1!Cnpj_Ancora
            dataPlan = Format(Date, "MM/DD/YYYY")
            
            linha = 9
            
            ObjPlan1Excel.Range("C2") = NomeCliente
            ObjPlan1Excel.Range("C4") = dataPlan
            ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("K2") = CNPJcliente
            ObjPlan1Excel.Range("K2").NumberFormat = "00000"
            ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
            
            UltimaLinha = TbDados1.RecordCount
            UltimaLinha = UltimaLinha + 8
            
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(7).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(8).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(9).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(10).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(11).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(1).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(2).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(3).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Borders(4).LineStyle = 2
            ObjPlan1Excel.Range("C" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("A" & linha & ":W" & UltimaLinha).Font.Size = 8
            ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("T" & linha & ":V" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("W" & linha & ":W" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Columns("A:W").Select
            ObjPlan1Excel.Columns.AutoFit
            ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
            ObjPlan1Excel.Range("C2").Select
            
            Data = Date
            Nome = "Relatorio de Notas a Vencer2 - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
            Nome = Trata_NomeArquivo(Nome)
            
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
            
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
                           
            
            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname
            

''Pular envio de e-mail
'GoTo lblPuloEmail

    '==== ENVIAR RELATORIO POR EMAIL

        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados1!Email_Ancora
        EmailCopia = TbDados1!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Notas a Vencer(Com Taxa)"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
        
            oitem.Subject = ("RELATORIO DE NOTAS A VENCER -  " & NomeCliente & " - " & CNPJcliente)
            'oitem.Subject = ("RELATORIO DE NOTAS A VENCER -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
         
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS A VENCER"
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
        TbDados1.Close
        
lblPuloEmail:

    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas a Vencer(Com Taxa)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            rsDao.Close
            Set rsDao = Nothing
            
            TbDados.Close
Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatNotasAvencSemTaxaJuros() 'Relatorio de Notas a Vencer (Sem Taxa\Juros\ Valor Liq)

'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
        
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
        DataPesq = Format(DiarioPesq, "mm/dd/yyyy")
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
        'Alterado em 10/12/2018 - Para agrupamento de convênios em relatório.
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

        'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto" _
        '& " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_Venc)>#" & DataPesq & "#));", dbOpenDynaset)
        
        vString = "SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        vCriterio = " WHERE (((TblArqoped.Data_Venc)>#" & DataPesq & "#)); "
        
        Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)
        
        If TbDados1.EOF = True Then GoTo Fim
        
        Agencia_Ancora = TbDados1!Agencia_Ancora
        Convenio_Ancora = TbDados1!Convenio_Ancora
        Cnpj_Ancora = TbDados1!Cnpj_Ancora
        
    '==== GERA O RELATORIO EM EXCEL
    
        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
            
            Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAvencerSemTaxaJuros.xlsx", ReadOnly:=True
            Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
            
            NomeCliente = TbDados1!Nome_Ancora
            CNPJcliente = TbDados1!Cnpj_Ancora
            dataPlan = Format(Date, "MM/DD/YYYY")
            
            linha = 9
            
            ObjPlan1Excel.Range("C2") = NomeCliente
            ObjPlan1Excel.Range("C4") = dataPlan
            ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("K2") = CNPJcliente
            ObjPlan1Excel.Range("K2").NumberFormat = "00000"
            ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
            
            UltimaLinha = TbDados1.RecordCount
            UltimaLinha = UltimaLinha + 8
            
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(7).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(8).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(9).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(10).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(11).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(1).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(2).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(3).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Borders(4).LineStyle = 2
            ObjPlan1Excel.Range("A" & linha & ":T" & UltimaLinha).Font.Size = 8
            ObjPlan1Excel.Range("B" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("F" & linha & ":M" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("Q" & linha & ":S" & UltimaLinha).NumberFormat = "00000"
            ObjPlan1Excel.Range("N" & linha & ":O" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
            ObjPlan1Excel.Range("P" & linha & ":P" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Range("T" & linha & ":T" & UltimaLinha).Style = "Currency"
            ObjPlan1Excel.Columns("A:T").Select
            ObjPlan1Excel.Columns.AutoFit
            ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
            ObjPlan1Excel.Range("C2").Select
            
            Data = Date
            Nome = "Relatorio de Notas a Vencer3 - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
            Nome = Trata_NomeArquivo(Nome)
            
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
            
            'ObjPlan1Excel.SaveAs FileName:=sFname
            'ObjExcel.activeworkbook.Close SaveChanges:=False
            'ObjExcel.Quit
            'TbDados1.Close
            
            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname
            
''Pular envio de e-mail
'GoTo lblPuloEmail
            
    '==== ENVIAR RELATORIO POR EMAIL

        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados2 = Db2.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados2!Email_Ancora
        EmailCopia = TbDados2!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Notas a Vencer(Sem Taxa/Juros/Val Liq)"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
        
            oitem.Subject = ("RELATORIO DE NOTAS A VENCER -  " & NomeCliente & " - " & CNPJcliente)
            'oitem.Subject = ("RELATORIO DE NOTAS A VENCER -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia & ""
        
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS A VENCER"
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
        TbDados2.Close

lblPuloEmail:

    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas a Vencer(Sem Taxa/Juros/Val Liq)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            rsDao.Close
            Set rsDao = Nothing
            
            TbDados.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub FornecedoresCadastrados() 'Relatorio de Fornecedores Cadastrados
'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
        
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
        'Alterado em 10/12/2018 - Para agrupamento de convênios em relatório.
        
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqForn.Bco_Ancora, TblArqForn.Agencia_Ancora, TblArqForn.Convenio_Ancora, TblArqForn.Nome_Ancora, TblArqForn.CNPJ_Fornecedor, TblArqForn.Nome_Fornecedor, TblArqForn.Status_Fornecedor, TblArqForn.End_Fornecedor, TblArqForn.Num_Fornecedor, TblArqForn.Bairro_Fornecedor, TblArqForn.Cidade_Fornecedor, TblArqForn.UF_Fornecedor, TblArqForn.CEP_Fornecedor, TblArqForn.DDD1, TblArqForn.Fone1, TblArqForn.DDDFAX, TblArqForn.FAX, TblArqForn.DDD2, TblArqForn.FONE2, TblArqForn.Banco_Fornecedor, TblArqForn.Agencia_Fornecedor, TblArqForn.Conta_Fornecedor, TblArqForn.Contato_Fornecedor, TblArqForn.Email_Fornecedor, TblArqForn.TipoBlo_Fornecedor" _
        '& " FROM TblArqForn WHERE (((TblArqForn.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblArqForn.Convenio_Ancora)='" & Convenio_Ancora & "'));", dbOpenDynaset)
        
        vString = "SELECT TblArqForn.Bco_Ancora, TblArqForn.Agencia_Ancora, TblArqForn.Convenio_Ancora, TblArqForn.Nome_Ancora, TblArqForn.CNPJ_Fornecedor, TblArqForn.Nome_Fornecedor, TblArqForn.Status_Fornecedor, TblArqForn.End_Fornecedor, TblArqForn.Num_Fornecedor, TblArqForn.Bairro_Fornecedor, TblArqForn.Cidade_Fornecedor, TblArqForn.UF_Fornecedor, TblArqForn.CEP_Fornecedor, TblArqForn.DDD1, TblArqForn.Fone1, TblArqForn.DDDFAX, TblArqForn.FAX, TblArqForn.DDD2, TblArqForn.FONE2, TblArqForn.Banco_Fornecedor, TblArqForn.Agencia_Fornecedor, TblArqForn.Conta_Fornecedor, TblArqForn.Contato_Fornecedor, TblArqForn.Email_Fornecedor, TblArqForn.TipoBlo_Fornecedor "
        vString = vString & " FROM TblArqForn INNER JOIN AuxConvenio ON TblArqForn.Agencia_Ancora = AuxConvenio.Agencia_Ancora And TblArqForn.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        vCriterio = " "
        
        Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)
        
        If TbDados1.EOF = True Then GoTo Fim
    
        Agencia_Ancora = TbDados1!Agencia_Ancora
        Convenio_Ancora = TbDados1!Convenio_Ancora
        
        vString = "Select Cnpj_Ancora FROM TblArqoped Where TblArqoped.Agencia_Ancora = " & Agencia_Ancora & " And TblArqoped.Convenio_Ancora = '" & Convenio_Ancora & "' "
        Set rsDao = Db1.OpenRecordset(vString, dbOpenDynaset)
        
        Cnpj_Ancora = ""
        If rsDao.EOF = False Then
            Cnpj_Ancora = rsDao!Cnpj_Ancora
        End If
        
        rsDao.Close
        Set rsDao = Nothing
        
        
    '==== GERA O RELATORIO EM EXCEL
        
        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
        
        Set ObjExcel = CreateObject("EXCEL.application")
        ObjExcel.Workbooks.Open FileName:=Caminho & "FornecedoresCadastrados.xlsx", ReadOnly:=True
        Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
        
        NomeCliente = TbDados1!Nome_Ancora
        dataPlan = Format(Date, "MM/DD/YYYY")
        linha = 9
        
        ObjPlan1Excel.Range("C2") = NomeCliente
        ObjPlan1Excel.Range("C4") = dataPlan
        ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("K2").NumberFormat = "00000"
        ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
        
        UltimaLinha = TbDados1.RecordCount
        UltimaLinha = UltimaLinha + 8
        
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(7).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(8).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(9).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(10).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(11).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(1).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(2).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(3).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Borders(4).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":Y" & UltimaLinha).Font.Size = 8
        ObjPlan1Excel.Range("A" & linha & ":C" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("E" & linha & ":E" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("H" & linha & ":H" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("M" & linha & ":V" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Columns("A:V").Select
        ObjPlan1Excel.Columns.AutoFit
        ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
        ObjPlan1Excel.Range("C2").Select
        
        Data = Date
        Nome = "Relatorio de Fornecedores Cadastrados - " & NomeCliente & " - " & Format(Data, "ddmmyy")
        Nome = Trata_NomeArquivo(Nome)
        
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
            'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
        
            'ObjPlan1Excel.SaveAs FileName:=sFname
            'ObjExcel.activeworkbook.Close SaveChanges:=False
            'ObjExcel.Quit
            'TbDados1.Close
        
            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname

''Pular envio de e-mail
'GoTo lblPuloEmail

    '==== ENVIAR RELATORIO POR EMAIL

        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados2 = Db2.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados2!Email_Ancora
        EmailCopia = TbDados2!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Fornecedores Cadastrados"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
        
            oitem.Subject = ("RELATORIO DE FORNECEDORES CADASTRADOS -  " & NomeCliente)
            'oitem.Subject = ("RELATORIO DE FORNECEDORES CADASTRADOS -  " & NomeCliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
        
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE FORNECEDORES CADASTRADOS"
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
        
        'oitem.Display True
        Set olapp = Nothing
        Set oitem = Nothing
        TbDados2.Close
    
lblPuloEmail:
    
    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            'Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Fornecedores Cadastrados"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            rsDao.Close
            Set rsDao = Nothing
            
            TbDados.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatNotasAntSimples()

'On Error GoTo Listagem

    Dim Ret As String, DataPesq As Date
    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date, DtFim As Date
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim TbDados As Recordset, TbData As Recordset, TbDados2 As Recordset
        
    '==== PESQUISAR SIGLA DO USUARIO LOGADO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)

    '==== ABRE O BANCO DE DADOS
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                    If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                        TbDados2.Close
                    End If

    '==== ABRE O BANCO DE DADOS
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
            If MensalMes = 0 Then
                MensalMes = 12
                MensalAno = MensalAno - 1
            End If
                        
            pesqmensal = MensalMes & "/20" & MensalAno
            pesqmensal = Format(pesqmensal, "MM/YYYY")

    '==== ABRE A TABELA CALENDARIO
        Set TbData = Db1.OpenRecordset("TblCalendario", dbOpenDynaset)
            DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                        If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                        End If
                Loop
        TbData.Close
        
        
    '==== VALIDA QUAL A PERIODICIDADE DO RELATORIO A SER ENVIADO
        'Alterado em 07/12/2018 - ACS
        vString = " SELECT TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Data_op, TblArqoped.Cod_Oper, TblArqoped.Valor_op "
        vString = vString & " FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        
        vCriterio = ""
    
        If PeriodicidadeRel = "Diario" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Data_op, TblArqoped.Cod_Oper, TblArqoped.Valor_op" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Data_op)Like '*" & DiarioPesq & "*') AND ((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "'));", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op) Like '*" & DiarioPesq & "*')); "
        ElseIf PeriodicidadeRel = "Semanal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Data_op, TblArqoped.Cod_Oper, TblArqoped.Valor_op" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Data_op)>=Date()-7) AND ((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "')) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op)>=Date()-7)) ORDER BY TblArqoped.Data_op; "
        ElseIf PeriodicidadeRel = "Mensal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Data_op, TblArqoped.Cod_Oper, TblArqoped.Valor_op" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Data_op)Like '*" & pesqmensal & "') AND ((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "')) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op)Like '*" & pesqmensal & "')) ORDER BY TblArqoped.Data_op; "
        ElseIf PeriodicidadeRel = "Quinzenal" Then
            'Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Data_op, TblArqoped.Cod_Oper, TblArqoped.Valor_op" _
            '& " FROM TblArqoped WHERE (((TblArqoped.Data_op)>=Date()-15) AND ((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "')) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            vCriterio = " WHERE (((TblArqoped.Data_op)>=Date()-15)) ORDER BY TblArqoped.Data_op; "
        End If
        
        If vCriterio = "" Then GoTo Fim
        Set TbDados1 = Db1.OpenRecordset(vString & vCriterio, dbOpenDynaset)
        
        If TbDados1.EOF = True Then GoTo Fim
    
        Dim rsDao As DAO.Recordset
        
        vString = "Select TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora FROM TblArqoped INNER JOIN AuxConvenio ON TblArqoped.Agencia_Ancora = AuxConvenio.Agencia_AncoraNum And TblArqoped.Convenio_Ancora = AuxConvenio.Convenio_Ancora "
        Set rsDao = Db1.OpenRecordset(vString, dbOpenDynaset)
        
        Agencia_Ancora = rsDao!Agencia_Ancora
        Convenio_Ancora = rsDao!Convenio_Ancora
        Cnpj_Ancora = rsDao!Cnpj_Ancora
        
        rsDao.Close
        Set rsDao = Nothing
    
    '==== GERA O RELATORIO EM EXCEL
    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
    
        Set ObjExcel = CreateObject("EXCEL.application")
        ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasSimples.xlsx", ReadOnly:=True
        Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
        
        NomeCliente = TbDados1!Nome_Ancora
        CNPJcliente = TbDados1!Cnpj_Ancora
        dataPlan = Format(Date, "MM/DD/YYYY")
        
        linha = 9
        
        ObjPlan1Excel.Range("C2") = NomeCliente
        ObjPlan1Excel.Range("C4") = dataPlan
        ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("H2") = CNPJcliente
        ObjPlan1Excel.Range("H2").NumberFormat = "00000"
        ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados1
        
        UltimaLinha = TbDados1.RecordCount
        UltimaLinha = UltimaLinha + 8
        
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
        ObjPlan1Excel.Range("A" & linha & ":B" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("D" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("F" & linha & ":F" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("H" & linha & ":H" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("G" & linha & ":G" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("I" & linha & ":I" & UltimaLinha).Style = "Currency"
        ObjPlan1Excel.Columns("A:I").Select
        ObjPlan1Excel.Columns.AutoFit
        ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
        ObjPlan1Excel.Range("C2").Select
        
        Data = Date
        Nome = "Relatorio de Notas Antecipadas3 - " & PeriodicidadeRel & " - " & NomeCliente & " - " & Format(Data, "ddmmyy")
        Nome = Trata_NomeArquivo(Nome)
        
        sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        'Testes
        'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
        
            If (Dir(sFname) <> "") Then
                Kill sFname
            End If
        
            
        'ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
       ' ObjPlan1Excel.SaveAs FileName:=sFname
        'ObjExcel.activeworkbook.Close SaveChanges:=False
        'ObjExcel.Quit
        'TbDados1.Close
        
            vCaminhoLocal = "c:\temp\ArquivoConfirming\" & Nome & ".xlsx"
            'ObjPlan1Excel.SaveAs FileName:=sFname
            ObjPlan1Excel.SaveAs FileName:=vCaminhoLocal
            ObjExcel.activeworkbook.Close SaveChanges:=False
            
            ObjExcel.Quit
            
            TbDados1.Close
            
            Dim FSO As New FileSystemObject
            FSO.MoveFile vCaminhoLocal, sFname



        'ACS - Pular envio de e-mail
        'GoTo lblPuloEmail
            
    '==== ENVIAR RELATORIO POR EMAIL
    
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados2 = Db2.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados2!Email_Ancora
        EmailCopia = TbDados2!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        'Pesquisa email Lista Personalizada (PPB Makro Reverso) - Emerson 23/04/2018
        Dim nomeRelatorio As String
        nomeRelatorio = "Notas Antecipadas (Simples)"
        Dim tbMailPersonalizado As Recordset
        SQL = "SELECT TblListaPersonalizadaRel.* FROM TblListaPersonalizadaRel WHERE (( (TblListaPersonalizadaRel.Agencia_Ancora='" & Agencia_Ancora & _
        "') AND (TblListaPersonalizadaRel.Convenio_Ancora='" & Convenio_Ancora & "')" & " AND (tblListaPersonalizadaRel.Relatorios='" & nomeRelatorio & _
        "') AND (TblListaPersonalizadaRel.Periodicidade ='" & PeriodicidadeRel & "')" & " AND (tblListaPersonalizadaRel.Ativo=TRUE) ))"
        Set tbMailPersonalizado = Db1.OpenRecordset(SQL)
        
        If Not tbMailPersonalizado.EOF Then
            EmailDestino = tbMailPersonalizado!Email_Para
            EmailCopia = tbMailPersonalizado!Email_Copia
        End If
        
        tbMailPersonalizado.Close
        Set tbMailPersonalizado = Nothing
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
        
            oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente)
            'oitem.Subject = ("RELATORIO DE NOTAS ANTECIPADAS -  " & NomeCliente & " - " & CNPJcliente & " - RETIFICADO")
            
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
        
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
        TbDados2.Close
    
lblPuloEmail:
    
    '==== ABRE O BANCO DE DADOS
        Set Db = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados = Db.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
            
            ''Inserir aqui todos os registros da AuxConvenio
            'Dim rsDao As DAO.Recordset
            Set rsDao = Db.OpenRecordset("Select * From AuxConvenio", 4)
            
            If rsDao.EOF = False Then rsDao.MoveFirst
            Do While rsDao.EOF = False
            
                TbDados.AddNew
                    TbDados!Agencia_Ancora = rsDao!Agencia_AncoraNum        'Agencia_Ancora
                    TbDados!Convenio_Ancora = rsDao!Convenio_Ancora         'Convenio_Ancora
                    TbDados!Cnpj_Ancora = CNPJcliente
                    TbDados!Nome_Ancora = NomeCliente
                    TbDados!Relatorio_Enviado = "Notas Antecipadas (Simples)"
                    TbDados!Periodicidade_Relatorio = PeriodicidadeRel
                    TbDados!Data_Envio = Format(Date, "DD / MM / YYYY")
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
                
                rsDao.MoveNext
            Loop
            
            rsDao.Close
            Set rsDao = Nothing

            TbDados.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub MascaraBRF() 'Mascara BRF

    'On Error GoTo Listagem
    
    Dim Nome As String, Ret As String
    Dim Db As Database, Db1 As Database, Db2 As Database
    Dim DtInicio As Date, DataPesq As Date, DtFim As Date
    Dim ObjExcel As Object, ObjPlan1Excel As Object, ObjPlan2Excel As Object, linha As Double
    Dim TbData As Recordset, TbDados As Recordset, TbDados2 As Recordset, TbDados1 As Recordset, TbDados5 As Recordset

    '==== PESQUISAR SIGLA DE USUARIO
        SiglaUser = String(255, 0)
        Ret = GetUserName(SiglaUser, Len(SiglaUser))
        
        X = 1
        Do While Asc(Mid(SiglaUser, X, 1)) <> 0
            X = X + 1
        Loop
            SiglaUser = Left(SiglaUser, (X - 1))
            SiglaPesq = UCase(SiglaUser)
    
    
    '==== PESQUISAR NOME E EMAIL DO USUARIO
        Set Db2 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
            Set TbDados2 = Db2.OpenRecordset("TblUsuarios", dbOpenDynaset)
                TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    If TbDados2.NoMatch = False Then
                        NomeUsuario = TbDados2!Nome
                        EmailUsuario = TbDados2!Email
                    End If
                TbDados2.Close
    
    '==== PESQUISAR DATA DO ULTIMO MES
        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")

            MensalDia = "01"
            MensalMes = Mid(Mensal, 4, 2)
            MensalMes = MensalMes - 1
            MensalAno = Right(Mensal, 2)
                If MensalMes = "0" Then
                    MensalMes = "12"
                    MensalAno = MensalAno - 1
                End If
            pesqmensal = MensalDia & "/" & MensalMes & "/" & MensalAno
            pesqmensal = Format(pesqmensal, "DD/MM/YYYY")
            
            DataPesq = Format(Date, "dd/mm/yyyy")
        
    '==== CONSULTA OPERAÇÕES REALIZADAS NO PERIODO
        Agencia_Ancora = 2050: Convenio_Ancora = "008500000049"
        
        Set TbDados1 = Db1.OpenRecordset("SELECT TblArqoped.Convenio_Ancora, TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Data_op, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Pagmto, [TblArqoped]![Data_Venc]-Date() AS PRAZO, IIf([TblArqoped]![Data_Venc]-Date()<0,'Liquidado',IIf([TblArqoped]![Data_Venc]-Date()=0,'D0',IIf([TblArqoped]![Data_Venc]-Date()=1,'D1',IIf([TblArqoped]![Data_Venc]-Date()>=2 And [TblArqoped]![Data_Venc]-Date()<=30,'Entre 2 e 30',IIf([TblArqoped]![Data_Venc]-Date()>=31 And [TblArqoped]![Data_Venc]-Date()<=60,'Entre 30 e 60'))))) AS INTERVALO" _
        & " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio_Ancora & "') AND ((TblArqoped.Data_op) >=#" & pesqmensal & "#) AND (([TblArqoped]![Data_Venc]-Date())<=60));", dbOpenDynaset)
        
        If TbDados1.EOF = True Then GoTo Fim

    '==== GERA O RELATORIO EM EXCEL

        Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
        
        Set ObjExcel = CreateObject("EXCEL.application")
        ObjExcel.Workbooks.Open FileName:=Caminho & "BRF_RELATORIOS.xlsx", ReadOnly:=True
        Set ObjPlan1Excel = ObjExcel.Worksheets("Analítico")
               
        NomeCliente = TbDados1!Nome_Ancora
        CNPJcliente = TbDados1!Cnpj_Ancora
        dataPlan = Format(Date, "MM/DD/YYYY")
        
        linha = 2
        
        ObjPlan1Excel.Range("A2").CopyFromRecordset TbDados1
        UltimaLinha = TbDados1.RecordCount
        UltimaLinha = UltimaLinha + 1
        
        ObjPlan1Excel.Select
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(7).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(8).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(9).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(10).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(11).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(1).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(2).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(3).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(4).LineStyle = 2
        ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Font.Size = 8
        ObjPlan1Excel.Range("A" & linha & ":A" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("C" & linha & ":C" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("E" & linha & ":E" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("K" & linha & ":K" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("G" & linha & ":G" & UltimaLinha).NumberFormat = "00000"
        ObjPlan1Excel.Range("H" & linha & ":I" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("F" & linha & ":F" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
        ObjPlan1Excel.Range("J" & linha & ":J" & UltimaLinha).NumberFormat = "General"
        ObjPlan1Excel.Columns("A:K").Select
        ObjPlan1Excel.Columns.AutoFit
        ObjPlan1Excel.Rows("1:" & UltimaLinha).RowHeight = 11.75
        ObjPlan1Excel.Range("C2").Select
        
        Set ObjPlan2Excel = ObjExcel.Worksheets("Máscara")
        
        ObjPlan2Excel.Activate
        ObjPlan2Excel.Range("C2").Select
        ObjPlan2Excel.PivotTables("Tabela dinâmica3").PivotCache.Refresh
        ObjPlan2Excel.PivotTables("Tabela dinâmica2").PivotCache.Refresh
        
        Data = Date
        Nome = "Relatorio BRF - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
        Nome = Trata_NomeArquivo(Nome)
        
            sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                If (Dir(sFname) <> "") Then
                    Kill sFname
                End If
        
        ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        ObjExcel.activeworkbook.Close SaveChanges:=False
        ObjExcel.Quit
        TbDados1.Close

    '==== ENVIAR RELATORIO POR EMAIL

        Set Db1 = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
        
        Set TbDados1 = Db1.OpenRecordset("SELECT TblAncoras.Nome_Ancora, TblAncoras.Email_Trader, TblAncoras.Email_Ancora, TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblAncoras.Cnpj_Ancora FROM TblAncoras WHERE ((TblAncoras.Agencia_Ancora)='" & Agencia_Ancora & "') AND ((TblAncoras.Convenio_Ancora)='" & Convenio_Ancora & "');", dbOpenDynaset)
        
        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        
        EmailDestino = TbDados1!Email_Ancora
        EmailCopia = TbDados1!Email_Trader
        Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
            
            oitem.Subject = ("RELATORIO CONFIRMING - BRF -  " & NomeCliente & " - " & CNPJcliente)
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
            
        Corpo1 = "Prezado Cliente,"
        Corpo2 = "Segue anexo o "
        Relatorio = "RELATÓRIO DE NOTAS ANTECIPADAS"
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
        TbDados1.Close

Listagem:
Select Case Err.Number

Case 3043
Exit Sub

Case Else
End Select

Fim:

End Sub
Sub RelatorioNotificacao()

'On Error GoTo Listagem

    Dim TbUsuario As Recordset, TbDados1 As Recordset, TbAncoras As Recordset
    Dim ObjExcel As Object, ObjPlan1Excel As Object, TbDadosAncora As Recordset
    Dim linha As Integer, Nome As String, Db As Database

      Call AbrirBDLocal

         Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"

            Set TbUsuario = BDRELocal.OpenRecordset("TblUsuarios", dbOpenDynaset)

             SiglaPesq = PesqUsername()

              TbUsuario.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                If TbUsuario.NoMatch = False Then
                    NomeUsuario = TbUsuario!Nome
                    EmailUsuario = TbUsuario!Email
                    TbUsuario.Close
                End If

                 Set TbAncoras = BDRELocal.OpenRecordset("SELECT TblNotificacao.Corpo_Email, TblNotificacao.Agencia_Ancora, TblNotificacao.Convenio_Ancora, TblNotificacao.CNPJ_Ancora, TblNotificacao.Nome_Ancora, TblNotificacao.Grupo_Ancora, TblNotificacao.Email_Ancora, TblNotificacao.Email_Trader, TblNotificacao.Data_Contrato, TblNotificacao.Data_Inclusao, TblNotificacao.Status_Ancora, TblNotificacao.Usuario FROM TblNotificacao WHERE (((TblNotificacao.Status_Ancora)='ATIVO'));", dbOpenDynaset)
                    
                    Do While TbAncoras.EOF = False
                   ' GoTo prox
                   
                            Set TbDados1 = BDRELocal.OpenRecordset("SELECT TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom FROM TblArqoped" _
                            & " WHERE (((TblArqoped.Agencia_Ancora)=" & TbAncoras!Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & TbAncoras!Convenio_Ancora & "')) GROUP BY TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom HAVING (((TblArqoped.Data_op) >= Date() - 7)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)

                                If TbDados1.EOF = False Then
                                
                                    Set TbDadosAncora = BDRELocal.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & TbAncoras!Agencia_Ancora & ") AND ((TblArqoped.Convenio_Ancora)='" & TbAncoras!Convenio_Ancora & "')) GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet;", dbOpenDynaset)
                                    
                                       Set ObjExcel = CreateObject("EXCEL.application")
                                       ObjExcel.Workbooks.Open FileName:=Caminho & "Notificacao.xlsx", ReadOnly:=True
                                       Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                                       
                                        NomeCliente = TbAncoras!Nome_Ancora: CNPJcliente = TbDadosAncora!Cnpj_Ancora
                                        dataPlan = Format(Date, "MM/DD/YYYY"): linha = 11
                                       
                                       ObjPlan1Excel.Range("C2") = NomeCliente
                                       ObjPlan1Excel.Range("G2") = CNPJcliente
                                       ObjPlan1Excel.Range("G2").NumberFormat = "00000"
                                       ObjPlan1Excel.Range("C4") = TbDadosAncora!Banco_Remet
                                       ObjPlan1Excel.Range("E4") = TbDadosAncora!Agencia_Remet
                                       ObjPlan1Excel.Range("E4").NumberFormat = "0000"
                                       ObjPlan1Excel.Range("G4") = TbDadosAncora!Conta_Remet
                                       ObjPlan1Excel.Range("G4").NumberFormat = "0000"
                                       ObjPlan1Excel.Range("D6") = TbAncoras!Convenio_Ancora
                                       ObjPlan1Excel.Range("D6").NumberFormat = "000000000000"
                                       ObjPlan1Excel.Range("G6") = dataPlan
                                       ObjPlan1Excel.Range("B11").CopyFromRecordset TbDados1
                                       
                                       UltimaLinha = TbDados1.RecordCount: UltimaLinha = UltimaLinha + 10
                                       
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(7).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(8).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(9).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(10).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(11).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(1).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(2).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(3).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Borders(4).LineStyle = 2
                                        ObjPlan1Excel.Range("B" & linha & ":G" & UltimaLinha).Font.Size = 8
                                        ObjPlan1Excel.Range("C" & linha & ":C" & UltimaLinha).NumberFormat = "00000"
                                        ObjPlan1Excel.Range("E" & linha & ":E" & UltimaLinha).NumberFormat = "00000"
                                        ObjPlan1Excel.Range("D" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
                                        ObjPlan1Excel.Range("F" & linha & ":F" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                                        ObjPlan1Excel.Range("G" & linha & ":G" & UltimaLinha).Style = "Currency"
                                        ObjPlan1Excel.Columns("B:H").Select
                                        ObjPlan1Excel.Columns.AutoFit
                                        ObjPlan1Excel.Rows("11:" & UltimaLinha).RowHeight = 11.75
                                        ObjPlan1Excel.Range("C2").Select
                                    
                                       Nome = "Notificacao - " & NomeCliente & " - " & CNPJcliente & " - " & Format(dataPlan, "ddmmyy")
                                          
                                       Nome = Trata_NomeArquivo(Nome)
                                       
                                        sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                                        'sFname = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                                         If (Dir(sFname) <> "") Then
                                             Kill sFname
                                         End If
                                          
                                       ObjPlan1Excel.SaveAs FileName:=sFname
                                       ObjExcel.activeworkbook.Close SaveChanges:=False
                                       ObjExcel.Quit
    
                                        'GoTo lblPuloEmail
    
                                        Set ofs = CreateObject("Scripting.FileSystemObject")
    
                                            File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                                            'File = "C:\Users\t705854\Desktop\Confirming\Relatorios\" & Nome & ".xlsx"
                                            
                                                EmailDestino = TbAncoras!Email_Ancora
                                                EmailCopia = TbAncoras!Email_Trader
                                                Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
    
                                                    Set sbObj = New Scripting.FileSystemObject
                                                    Set olapp = CreateObject("Outlook.Application")
                                                    Set oitem = olapp.CreateItem(0)
    
                                                        oitem.Subject = ("RELAÇÃO DOS CRÉDITOS COMERCIAIS CEDIDOS PELOS FORNECEDORES AO BANCO -  " & NomeCliente)
                                                        
                                                        oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
                                                        
                                                        oitem.To = EmailDestino
                                                        oitem.cc = EmailCopia
                                                            
                                                            Assinatura1 = "Confirming®"
                                                            Assinatura2 = "Global Transaction Banking"
                                                            Assinatura3 = "Av. Juscelino Kubitschek, 2.235"
                                                            Assinatura4 = "Meios - Operações e Serviços"
                                                            Assinatura5 = "CEP: 04543-011  São Paulo-SP"
                                                            Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
                                                            Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
                                                            Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
                                                                                 
                                                        oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Tahoma Tahoma Size = 2 <BR>" & TbAncoras!Corpo_Email & "<BR/><BR>" & _
                                                        "<img src=" & Assinatura & " height=50 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Tahoma Size = 2" & "<BR>" & _
                                                        "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Tahoma Size = 1 <BR>" & Assinatura3 & _
                                                        "<BR/>" & Assinatura4 & "<BR/>" & Assinatura5 & "<BR/></FONT><FONT COLOR = BLACK FACE = Tahoma Size = 1 <BR><I>" & Assinatura6 & _
                                                        "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"
                                                                    
                                                    oitem.Attachments.Add File
                                                    oitem.Attachments.Add "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                                                   
                                                    oitem.Send
    
                                            'oitem.DISPLAY True
                                            Set olapp = Nothing
                                            Set oitem = Nothing
                                                            
lblPuloEmail:
                                                    
                                                            
                                        Set TbDados = BDRELocal.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
                                                    
                                            TbDados.AddNew
                                                TbDados!Agencia_Ancora = Agencia_Ancora
                                                TbDados!Convenio_Ancora = Convenio_Ancora
                                                TbDados!Cnpj_Ancora = CNPJcliente
                                                TbDados!Nome_Ancora = NomeCliente
                                                TbDados!Relatorio_Enviado = "Relatorios de Notificaçao NF"
                                                TbDados!Periodicidade_Relatorio = "SEMANAL"
                                                TbDados!Data_Envio = Date
                                                TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                                                TbDados!Usuario = NomeUsuario
                                            TbDados.Update
                                        TbDados.Close
                                    End If
prox:
                        TbAncoras.MoveNext
                    Loop

End Sub
