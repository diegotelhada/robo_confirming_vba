'Form_Frm_ReenvioRelatorios'

Option Compare Database
Function fSetControlOption(ctl As Control)
    Call GetCursor
End Function
Function ValidarCampos()

    If IsNull(Me.TxAgencia) Or Me.TxAgencia = "" Then: MsgBox "Favor inserir a Agencia do Ancora", vbInformation, "Relatorios Confirming": ValidarCampos = "Erro": GoTo Fim
    If IsNull(Me.CbAncora) Or Me.CbAncora = "" Then: MsgBox "Favor selecionar o convenio do Ancora", vbInformation, "Relatorios Confirming": ValidarCampos = "Erro": GoTo Fim
    If IsNull(Me.CbRelatorios) Or Me.CbRelatorios = "" Then: MsgBox "Favor selecinar o relatorio Ancora", vbInformation, "Relatorios Confirming": ValidarCampos = "Erro": GoTo Fim
    If IsNull(Me.TxDtInicio) Or Me.TxDtInicio = "" Then: MsgBox "Favor inserir a data de inicio do relatorio", vbInformation, "Relatorios Confirming": ValidarCampos = "Erro": GoTo Fim
    If IsNull(Me.TxDtFinal) Or Me.TxDtFinal = "" Then: MsgBox "Favor inserir a data final do relatorio", vbInformation, "Relatorios Confirming": ValidarCampos = "Erro": GoTo Fim
    '
Fim:
End Function
Function LimparRelatorio()

    Me.CbRelatorios.Value = ""
    Me.CbRelatorios.RowSourceType = "Value List"
    Me.CbRelatorios.RowSource = ""
    Me.TxDtFinal = Empty
    Me.TxDtInicio = Empty
    Me.LbStatus.Caption = "": DoEvents
End Function
Function LimparAncora()

    Me.CbAncora.Value = ""
    Me.CbAncora.RowSourceType = "Value List"
    Me.CbAncora.RowSource = ""
    Me.CbRelatorios.Value = ""
    Me.CbRelatorios.RowSourceType = "Value List"
    Me.CbRelatorios.RowSource = ""
    Me.TxDtFinal = Empty
    Me.TxDtInicio = Empty
    Me.LbStatus.Caption = "": DoEvents
End Function
Function LimparCampos()

    Me.TxAgencia = Empty
    Me.CbAncora.Value = ""
    Me.CbAncora.RowSourceType = "Value List"
    Me.CbAncora.RowSource = ""
    Me.CbRelatorios.Value = ""
    Me.CbRelatorios.RowSourceType = "Value List"
    Me.CbRelatorios.RowSource = ""
    Me.TxDtFinal = Empty
    Me.TxDtInicio = Empty
    Me.LbStatus.Caption = "": DoEvents
End Function
Function PreencherRelatorio(Agencia As String, Ancora As String)

    Dim TbRelatorios As Recordset
      Call AbrirBDRelatorios: LimparRelatorio
        Set TbRelatorios = BDREL.OpenRecordset("SELECT TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblRelatorios.Relatorios FROM TblAncoras INNER JOIN TblRelatorios ON (TblAncoras.Convenio_Ancora = TblRelatorios.Convenio_Ancora) AND (TblAncoras.Agencia_Ancora = TblRelatorios.Agencia_Ancora) GROUP BY TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora, TblRelatorios.Relatorios HAVING (((TblAncoras.Agencia_Ancora)='" & Agencia & "') AND ((TblAncoras.Convenio_Ancora)='" & Ancora & "'));", dbOpenDynaset)
            If TbRelatorios.EOF = False Then
                Do While TbRelatorios.EOF = False
                        Me.CbRelatorios.AddItem TbRelatorios!Relatorios
                    TbRelatorios.MoveNext
                Loop
            Else
                MsgBox "Não existem relatorios cadastrados para a Agência e Ancora informado!", vbInformation, "Relatorios Confirming": Call LimparCampos
            End If
End Function
Function PreencherAncora(Agencia As String)

    Dim TbAncoras As Recordset
      Call AbrirBDRelatorios: Call LimparAncora
        Set TbAncoras = BDREL.OpenRecordset("SELECT TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora FROM TblAncoras INNER JOIN TblRelatorios ON (TblAncoras.Convenio_Ancora = TblRelatorios.Convenio_Ancora) AND (TblAncoras.Agencia_Ancora = TblRelatorios.Agencia_Ancora) GROUP BY TblAncoras.Agencia_Ancora, TblAncoras.Convenio_Ancora HAVING (((TblAncoras.Agencia_Ancora)='" & Agencia & "'));", dbOpenDynaset)
            If TbAncoras.EOF = False Then
                Do While TbAncoras.EOF = False
                        Me.CbAncora.AddItem TbAncoras!Convenio_Ancora
                    TbAncoras.MoveNext
                Loop
            Else
                MsgBox "Não existem relatorios cadastrados para a agência informada!", vbInformation, "Relatorios Confirming": Call LimparCampos
            End If
End Function
Function PesquisarUsuario()

    Dim TbDados As Recordset
    
        Call AbrirBDRelatorios
            SiglaPesq = PesqUsername()
        
        Set TbDados = BDREL.OpenRecordset("TblUsuarios", dbOpenDynaset)
        
        TbDados.FindFirst "Sigla like '*" & SiglaPesq & "'"
            
            If TbDados.NoMatch = False Then
             
              NomeUsuario = TbDados!Nome
              EmailUsuario = TbDados!Email
              Me.TxUsuario = NomeUsuario
            Else
              Me.TxUsuario = "Usuario não cadastrado"
            End If
        TbDados.Close

End Function
Private Sub CbAncora_AfterUpdate()

    If Not IsNull(Me.TxAgencia) Or Me.TxAgencia <> "" Then
        Call PreencherRelatorio(Me.TxAgencia, Me.CbAncora)
    Else
        Call LimparRelatorio
    End If
    Me.LbStatus.Caption = "": DoEvents
End Sub
Private Sub Form_Open(Cancel As Integer)
    
    Call PesquisarUsuario
    Call LimparCampos
End Sub
Private Sub lb_HOME_Click()

    Dim Form As String: Form = "FrmRelatorios"
    Dim FormClose As String: FormClose = "Frm_ReenvioRelatorios"
        
        DoCmd.OpenForm Form, acNormal
        DoCmd.Close acForm, FormClose

End Sub

Private Sub lb_HOME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption lb_HOME
End Sub

Private Sub TxAgencia_AfterUpdate()
    If Not IsNull(Me.TxAgencia) Or Me.TxAgencia <> "" Then
        Call PreencherAncora(Me.TxAgencia)
    Else
        Call LimparAncora
    End If
    Me.LbStatus.Caption = "": DoEvents
End Sub
Private Sub CmdGerarRelatorio_Click()

    Me.CbRelatorios.SetFocus: Valida = ValidarCampos
        If Valida <> "Erro" Then
          Me.LbStatus.Caption = "Aguarde... Gerando Relatorios...": DoEvents
            If Me.CbRelatorios = "Fornecedores Cadastrados" Then: Call FC(Me.TxAgencia, Me.CbAncora)
            If Me.CbRelatorios = "Notas a Vencer(Com Taxa)" Then: Call NVCT(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            If Me.CbRelatorios = "Notas a Vencer(Sem Taxa)" Then: Call NVST(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            If Me.CbRelatorios = "Notas a Vencer(Sem Taxa/Juros/Val Liq)" Then: Call NVSTJ(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            If Me.CbRelatorios = "Notas Antecipadas(Com Taxa)" Then: Call NACT(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            If Me.CbRelatorios = "Notas Antecipadas(Sem Taxa)" Then: Call NAST(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            If Me.CbRelatorios = "Notas Antecipadas(Sem Taxa/Juros/Val Liq)" Then: Call NASTJ(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            If Me.CbRelatorios = "Notas Antecipadas(Com Taxa/Custo/PPB)" Then: Call NACTCP(Me.TxAgencia, Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
            
            Call LimparCampos
        End If
End Sub
Function NACTCP(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios
        
            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros, CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]) AS CustoAncora, TblArqoped.Receita_Clte, (CDbl([TblArqoped]![Valor_Custo])+CDbl([TblArqoped]![Receita_Banco]))/(CInt([Prazo_NF])*CDbl([Valor_Pagmto])/30) AS TaxaAncora FROM TblArqoped" _
            & " WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_op)>=#" & Format(DataInicio, "mm/dd/yyyy") & "# And (TblArqoped.Data_op)<=#" & Format(DataFinal, "mm/dd/yyyy") & "#)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                        Set ObjExcel = CreateObject("EXCEL.application")
                        ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasComTaxaPPB.xlsx", ReadOnly:=True
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
                        
                         sFname = "C:\Temp\" & Nome & ".xlsx"
                          If (Dir(sFname) <> "") Then
                              Kill sFname
                          End If
                      
                   ObjPlan1Excel.SaveAs FileName:="C:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    TbDados.Close
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If
End Function
Function FC(Agencia As String, Convenio As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbFornecedores As Recordset
    
        Call AbrirBDRelatorios

            Set TbFornecedores = BDREL.OpenRecordset("SELECT TblArqForn.Bco_Ancora, TblArqForn.Agencia_Ancora, TblArqForn.Convenio_Ancora, TblArqForn.Nome_Ancora, TblArqForn.CNPJ_Fornecedor, TblArqForn.Nome_Fornecedor, TblArqForn.Status_Fornecedor, TblArqForn.End_Fornecedor, TblArqForn.Num_Fornecedor, TblArqForn.Bairro_Fornecedor, TblArqForn.Cidade_Fornecedor, TblArqForn.UF_Fornecedor, TblArqForn.CEP_Fornecedor, TblArqForn.DDD1, TblArqForn.Fone1, TblArqForn.DDDFAX, TblArqForn.FAX, TblArqForn.DDD2, TblArqForn.FONE2, TblArqForn.Banco_Fornecedor, TblArqForn.Agencia_Fornecedor, TblArqForn.Conta_Fornecedor, TblArqForn.Contato_Fornecedor, TblArqForn.Email_Fornecedor, TblArqForn.TipoBlo_Fornecedor" _
            & " FROM TblArqForn WHERE (((TblArqForn.Agencia_Ancora)='" & Agencia & "') AND ((TblArqForn.Convenio_Ancora)='" & Convenio & "'));", dbOpenDynaset)

                If TbFornecedores.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                   Set ObjExcel = CreateObject("EXCEL.application")
                   ObjExcel.Workbooks.Open FileName:=Caminho & "FornecedoresCadastrados.xlsx", ReadOnly:=True
                   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbFornecedores!Nome_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                
                   ObjPlan1Excel.Range("C2") = NomeCliente
                   ObjPlan1Excel.Range("C4") = dataPlan
                   ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
                   ObjPlan1Excel.Range("K2").NumberFormat = "00000"
                   ObjPlan1Excel.Range("A9").CopyFromRecordset TbFornecedores
                   
                   UltimaLinha = TbFornecedores.RecordCount
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
                   
                    sFname = "c:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                   
                   ObjPlan1Excel.SaveAs FileName:="c:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If

End Function
Function NVCT(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios

            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido,TblArqoped.Juros " _
            & " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_Venc)>=#" & Format(DataInicio, "mm/dd/yyyy") & "#));", dbOpenDynaset)
            
                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                   Set ObjExcel = CreateObject("EXCEL.application")
                   ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAvencerComTaxa.xlsx", ReadOnly:=True
                   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbDados!Nome_Ancora
                    CNPJcliente = TbDados!Cnpj_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                
                   ObjPlan1Excel.Range("C2") = NomeCliente
                   ObjPlan1Excel.Range("C4") = dataPlan
                   ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
                   ObjPlan1Excel.Range("K2") = CNPJcliente
                   ObjPlan1Excel.Range("K2").NumberFormat = "00000"
                   ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
                   
                   UltimaLinha = TbDados.RecordCount
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
                   
                    sFname = "c:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                   
                   ObjPlan1Excel.SaveAs FileName:="c:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If

End Function
Function NVST(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios

            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido" _
            & " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_Venc)>=#" & Format(DataInicio, "mm/dd/yyyy") & "#));", dbOpenDynaset)
            
                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                   Set ObjExcel = CreateObject("EXCEL.application")
                   ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAvencerSemTaxa.xlsx", ReadOnly:=True
                   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbDados!Nome_Ancora
                    CNPJcliente = TbDados!Cnpj_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                
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
                   
                    sFname = "c:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                   
                   ObjPlan1Excel.SaveAs FileName:="c:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If
End Function
Function NVSTJ(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios

            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto" _
            & " FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_Venc)>=#" & Format(DataInicio, "mm/dd/yyyy") & "#));", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                   Set ObjExcel = CreateObject("EXCEL.application")
                   ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAvencerSemTaxaJuros.xlsx", ReadOnly:=True
                   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbDados!Nome_Ancora
                    CNPJcliente = TbDados!Cnpj_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                
                   ObjPlan1Excel.Range("C2") = NomeCliente
                   ObjPlan1Excel.Range("C4") = dataPlan
                   ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
                   ObjPlan1Excel.Range("K2") = CNPJcliente
                   ObjPlan1Excel.Range("K2").NumberFormat = "00000"
                   ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
                   
                   UltimaLinha = TbDados.RecordCount
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
                   
                    sFname = "c:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                   
                   ObjPlan1Excel.SaveAs FileName:="c:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If

End Function
Function NACT(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios

            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op," _
            & " TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido, TblArqoped.Juros FROM TblArqoped WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_op)>=#" & Format(DataInicio, "mm/dd/yyyy") & "# And (TblArqoped.Data_op)<=#" & Format(DataFinal, "mm/dd/yyyy") & "#)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)

                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                   Set ObjExcel = CreateObject("EXCEL.application")
                   ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasComTaxa.xlsx", ReadOnly:=True
                   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbDados!Nome_Ancora
                    CNPJcliente = TbDados!Cnpj_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                   
                   ObjPlan1Excel.Range("C2") = NomeCliente
                   ObjPlan1Excel.Range("C4") = dataPlan
                   ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
                   ObjPlan1Excel.Range("K2") = CNPJcliente
                   ObjPlan1Excel.Range("K2").NumberFormat = "00000"
                   ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
                   
                   UltimaLinha = TbDados.RecordCount
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
                   
                    sFname = "c:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                      
                   ObjPlan1Excel.SaveAs FileName:="c:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If

End Function
Function NAST(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios

            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto, TblArqoped.Valor_Juros, TblArqoped.Valor_Liquido FROM TblArqoped" _
            & " WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_op)>=#" & Format(DataInicio, "mm/dd/yyyy") & "# And (TblArqoped.Data_op)<=#" & Format(DataFinal, "mm/dd/yyyy") & "#)) ORDER BY TblArqoped.Data_op; ", dbOpenDynaset)
            
                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                    Set ObjExcel = CreateObject("EXCEL.application")
                    ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasSemTaxa.xlsx", ReadOnly:=True
                    Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbDados!Nome_Ancora
                    CNPJcliente = TbDados!Cnpj_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                   
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
                   
                   Nome = "Relatorio de Notas Antecipadas1 - " & PeriodicidadeRel & " - " & NomeCliente & " - " & CNPJcliente & " - " & Format(Data, "ddmmyy")
                   
                   Nome = Trata_NomeArquivo(Nome)
                   
                    sFname = "c:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                   
                   ObjPlan1Excel.SaveAs FileName:="C:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If
End Function
Function NASTJ(Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Dim ObjExcel As Object, ObjPlan1Excel As Object
    Dim Nome As String, TbDados As Recordset
    
        Call AbrirBDRelatorios

            Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Cod_Oper, TblArqoped.Data_op, TblArqoped.Data_Venc, TblArqoped.Valor_op, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Compromisso, TblArqoped.Valor_Pagmto FROM TblArqoped" _
            & " WHERE (((TblArqoped.Agencia_Ancora)=" & Agencia & ") AND ((TblArqoped.Convenio_Ancora)='" & Convenio & "') AND ((TblArqoped.Data_op)>=#" & Format(DataInicio, "mm/dd/yyyy") & "# And (TblArqoped.Data_op)<=#" & Format(DataFinal, "mm/dd/yyyy") & "#)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)
            
                If TbDados.EOF = False Then
                
                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                
                   Set ObjExcel = CreateObject("EXCEL.application")
                   ObjExcel.Workbooks.Open FileName:=Caminho & "NotasAntecipadasSemTaxaJuros.xlsx", ReadOnly:=True
                   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                   
                    NomeCliente = TbDados!Nome_Ancora
                    CNPJcliente = TbDados!Cnpj_Ancora
                    dataPlan = Format(Date, "MM/DD/YYYY"): linha = 9
                
                   ObjPlan1Excel.Range("C2") = NomeCliente
                   ObjPlan1Excel.Range("C4") = dataPlan
                   ObjPlan1Excel.Range("C4").NumberFormat = "dd/mm/yyyy"
                   ObjPlan1Excel.Range("K2") = CNPJcliente
                   ObjPlan1Excel.Range("K2").NumberFormat = "00000"
                   ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
                   
                   UltimaLinha = TbDados.RecordCount
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
                   
                    sFname = "C:\Temp\" & Nome & ".xlsx"
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                
                   ObjPlan1Excel.SaveAs FileName:="C:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                    MsgBox "Relatorio de " & Me.CbRelatorios & " Salvo no diretorio abaixo:" & Chr(13) & "c:\Temp\", vbExclamation, "Relatorios Confirming": Me.LbStatus.Caption = "": DoEvents
                End If


End Function

