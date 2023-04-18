'Form_Frm_PPB'

Option Compare Database

Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function

Function ZerarBarra()

    Me.LbStatus.Width = 0
    Me.LbEtapas.Caption = " "
    Me.LbPercentual.Caption = "0 %"
    Me.LbPercentual.ForeColor = &H0&       'Preto
    'Me.LbPercentual.ForeColor = &HFFFFFF   'Branco
    
    Me.TxDtFinal.Value = Empty: Me.TxDtInicio.Value = Empty
    
End Function

Function ZerarBarraDuranteProcesso()

    Me.LbStatus.Width = 0
    Me.LbPercentual.Caption = "0 %"
    Me.LbPercentual.ForeColor = &H0&       'Preto
    
End Function

Function GerarPPB(Ancora As String, DtInicio As String, DtFinal As String)

    DataInicio = DtInicio
    DataFinal = DtFinal
    
        If UCase(Ancora) = "B2W" Then
            
            Call GerarPPBB2W: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblB2W: Call ZerarBarraDuranteProcesso
            Call GerarExcelB2W: Call ZerarBarra
            
        ElseIf UCase(Ancora) = "DIA BRASIL" Then
                
            Call AdicionarCustoDIA: Call ZerarBarraDuranteProcesso
            Call GerarPPBDIA: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblDIA: Call ZerarBarraDuranteProcesso
            Call GerarExcel: Call ZerarBarra
                
        ElseIf UCase(Ancora) = "LASA" Then
        
            Call GerarPPBLASA: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblLASA: Call ZerarBarraDuranteProcesso
            Call GerarExcelLASA: Call ZerarBarra
        
        ElseIf UCase(Ancora) = "MEXICHEN" Then
        
            Call Adc_PPB_Mexichen: Call ZerarBarraDuranteProcesso
            Call Diferimento_Mexichen: Call ZerarBarraDuranteProcesso
            Call GerarExcelMexichem: Call ZerarBarra
        
        ElseIf UCase(Ancora) = "RAIZEN" Then
        
            Call GerarPPBRAIZEN: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblRAIZEN: Call ZerarBarraDuranteProcesso
            Call GerarExcelRAIZEN: Call ZerarBarra
        
        ElseIf UCase(Ancora) = "RDC" Then
        
            Call AdicionarCustosRDC: Call ZerarBarraDuranteProcesso
            Call Operaçõe_SDF: Call ZerarBarraDuranteProcesso
            Call GerarPPBRDC: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblRDC: Call ZerarBarraDuranteProcesso
            Call DiferimentoExcel: Call ZerarBarra
        
        ElseIf UCase(Ancora) = "SALOBO" Then
        
            Call GerarPPBSALOBO: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblSALOBO: Call ZerarBarraDuranteProcesso
            Call GerarExcelSALOBO: Call ZerarBarra
        
        ElseIf UCase(Ancora) = "VALE" Then

            Call GerarPPBVale: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblVALE: Call ZerarBarraDuranteProcesso
            Call GerarExcelVALE: Call ZerarBarra
            
        ElseIf UCase(Ancora) = "VALE" Then

            Call GerarPPBProssegur: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblProssegur: Call ZerarBarraDuranteProcesso
            Call GerarExcelProssegur: Call ZerarBarra
            
        ElseIf UCase(Ancora) = "CARAJAS" Then

            Call AdicionarCustosCarajas: Call ZerarBarraDuranteProcesso
            Call GerarPPBCarajas: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblCarajas: Call ZerarBarraDuranteProcesso
            Call DiferimentoExcelCarajas: Call ZerarBarra
        
        ElseIf UCase(Ancora) = "YOUCOM" Then

            Call AdicionarCustosYouCom: Call ZerarBarraDuranteProcesso
            Call GerarPPBYouCom: Call ZerarBarraDuranteProcesso
            Call DiferimentoTblYouCom: Call ZerarBarraDuranteProcesso
            Call DiferimentoExcelYouCom: Call ZerarBarra
                    
        End If
        

    MsgBox "PPB " & Ancora & " Gerado com sucesso!", vbInformation, "Relatorios Confirming"

End Function

Function ValidaCampos()

    If IsNull(Me.CbAncora) Or Me.CbAncora.Value = "" Then: MsgBox "Selecione um Âncora pra gerar o PPB", vbInformation, "Relatorios Confirming": ValidaCampos = "Erro": GoTo Fim
    If IsNull(Me.TxDtInicio) Or TxDtInicio.Value = "" Then: MsgBox "Digite a Data de Inicio", vbInformation, "Relatorios Confirming": ValidaCampos = "Erro": GoTo Fim
    If IsNull(Me.TxDtFinal) Or TxDtFinal.Value = "" Then: MsgBox "Digite a Data Final", vbInformation, "Relatorios Confirming": ValidaCampos = "Erro": GoTo Fim
    
        If Me.TxDtInicio > Me.TxDtFinal Then: MsgBox "Data de Inicio maior que a data Final!": GoTo Fim
        If Me.TxDtInicio = Me.TxDtFinal Then: MsgBox "As Datas digitadas são iguais!": GoTo Fim
Fim:
End Function

Function PreencherComboAncora()

    Dim Db As Database, TbAncoras As Recordset
    
        Set Db = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\BDPPB.mdb")
                
            Set TbAncoras = Db.OpenRecordset("SELECT Tbl_ANCORAS_PPB.Ancora FROM Tbl_ANCORAS_PPB WHERE Tbl_ANCORAS_PPB.ID < 12 GROUP BY Tbl_ANCORAS_PPB.Ancora HAVING ((Not (Tbl_ANCORAS_PPB.Ancora) Is Null)) ORDER BY Tbl_ANCORAS_PPB.Ancora;", dbOpenDynaset)
              
              Me.CbAncora.RowSourceType = "Value List"
              
                Do While TbAncoras.EOF = False
                        Me.CbAncora.AddItem TbAncoras!Ancora
                    TbAncoras.MoveNext
                Loop

End Function

Function PesquisarUsuario()

    Dim Db As Database, TbDados As Recordset

     SiglaPesq = PesqUsername()
        
        Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")
        
        Set TbDados = Db.OpenRecordset("TblUsuarios", dbOpenDynaset)
        
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
    Call ZerarBarra
End Sub

Function ExcluirCustoPPB(Ancora As String, Agencia As String, Convenio As String, DataInicio As String, DataFinal As String)

    Call AbrirBDPPB
    
        DataInic = Format(DataInicio, "mm/dd/yyyy"): DataFin = Format(DataFinal, "mm/dd/yyyy")
        
            If UCase(Ancora) = "DIA BRASIL" Or UCase(Ancora) = "RDC" Then: BDPPB.Execute ("Delete Tbl_Custo.Agencia_Ancora, Tbl_Custo.Convenio_Ancora, Tbl_Custo.Data FROM Tbl_Custo WHERE (((Tbl_Custo.Agencia_Ancora)='" & Agencia & "') AND ((Tbl_Custo.Convenio_Ancora)='" & Convenio & "') AND ((Tbl_Custo.Data)>=#" & DataInic & "# And (Tbl_Custo.Data)<=#" & DataFin & "#));")
            
            If UCase(Ancora) = "CARAJAS" Then: BDPPB.Execute ("Delete Tbl_Custo_Carajas.Agencia_Ancora, Tbl_Custo_Carajas.Convenio_Ancora, Tbl_Custo_Carajas.Data FROM Tbl_Custo_Carajas WHERE (((Tbl_Custo_Carajas.Agencia_Ancora)='" & Agencia & "') AND ((Tbl_Custo_Carajas.Convenio_Ancora)='" & Convenio & "') AND ((Tbl_Custo_Carajas.Data)>=#" & DataInic & "# And (Tbl_Custo_Carajas.Data)<=#" & DataFin & "#));")
            
            If UCase(Ancora) = "YOUCOM" Then: BDPPB.Execute ("Delete Tbl_Custo_YouCom.Agencia_Ancora, Tbl_Custo_YouCom.Convenio_Ancora, Tbl_Custo_YouCom.Data FROM Tbl_Custo_YouCom WHERE (((Tbl_Custo_YouCom.Agencia_Ancora)='" & Agencia & "') AND ((Tbl_Custo_YouCom.Convenio_Ancora)='" & Convenio & "') AND ((Tbl_Custo_YouCom.Data)>=#" & DataInic & "# And (Tbl_Custo_YouCom.Data)<=#" & DataFin & "#));")
        
                BDPPB.Execute ("DELETE Tbl_PPB.Agencia_Ancora, Tbl_PPB.Convenio_Ancora, Tbl_PPB.Data_op FROM Tbl_PPB WHERE (((Tbl_PPB.Agencia_Ancora)=" & Agencia & ") AND ((Tbl_PPB.Convenio_Ancora)='" & Convenio & "') AND ((Tbl_PPB.Data_Op)>=#" & DataInic & "# And (Tbl_PPB.Data_Op)<=#" & DataFin & "#));")
                
                BDPPB.Execute ("DELETE Tbl_PPB_FINAL.Agencia_Ancora, Tbl_PPB_FINAL.Convenio_Ancora, Tbl_PPB_FINAL.Data_Op FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.Agencia_Ancora)='" & Agencia & "') AND ((Tbl_PPB_FINAL.Convenio_Ancora)='" & Convenio & "') AND ((Tbl_PPB_FINAL.Data_Op)>=#" & DataInic & "# And (Tbl_PPB_FINAL.Data_Op)<=#" & DataFin & "#));")
        
        MsgBox "Custos Excluidos com Sucesso!", vbInformation, "Relatorios Confirming"
        
End Function

Private Sub CmdExcluirCusto_Click()
        
    Dim TbCustos As Recordset
        
      Call AbrirBDPPB: Me.CmdGerarPPB.SetFocus
    
        Valida = ValidaCampos
            
            If Valida <> "Erro" Then
                
                VbPgt = MsgBox("Tem Certeza que deseja excluir o custo do Âncora " & Me.CbAncora & " ?", vbYesNo, "Relatorios Confirming")
                              
                    If VbPgt = vbYes Then
                    
                        Set TbCustos = BDPPB.OpenRecordset("SELECT Tbl_ANCORAS_PPB.ID, Tbl_ANCORAS_PPB.Agencia, Tbl_ANCORAS_PPB.Convenio, Tbl_ANCORAS_PPB.Ancora, Tbl_ANCORAS_PPB.Pagamento FROM Tbl_ANCORAS_PPB WHERE (((Tbl_ANCORAS_PPB.Ancora)='" & Me.CbAncora & "'));", dbOpenDynaset)
                            
                            If TbCustos.EOF = False Then
                                Call ExcluirCustoPPB(Me.CbAncora, TbCustos!Agencia, TbCustos!Convenio, Me.TxDtInicio, Me.TxDtFinal)
                            Else
                                MsgBox "Âncora selecionado invalalido!", vbInformation, "Relatorios Confirming"
                            End If
                    End If
            End If
End Sub

Private Sub CmdGerarPPB_Click()
    
  Me.CmdExcluirCusto.SetFocus

    Valida = ValidaCampos
        
        If Valida <> "Erro" Then
            Call GerarPPB(Me.CbAncora, Me.TxDtInicio, Me.TxDtFinal)
        End If

End Sub

Private Sub Form_Open(Cancel As Integer)

 '   DoCmd.RunMacro "MCROculta"
    
        Call PesquisarUsuario
        Call PreencherComboAncora
        Call ZerarBarra
    
End Sub

Private Sub lb_HOME_Click()

    Dim Form As String: Form = "FrmRelatorios"
    Dim FormClose As String: FormClose = "Frm_PPB"
        
        DoCmd.OpenForm Form, acNormal
        DoCmd.Close acForm, FormClose

End Sub

Private Sub lb_HOME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption lb_HOME
End Sub

