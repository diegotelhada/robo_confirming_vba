Form_Frm_VencidosYD

Option Compare Database
Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function
Function AtualizarStatus(Status)
    Me.LbStatus.Caption = Status: DoEvents
End Function
Private Sub CmdExcluirCusto_Click()
    
    Call AtualizarStatus("Consultando arquivo YD42 no WSFGER..")
    Call GerarYD42
    Call AtualizarStatus("Importando arquivo YD42...")
    Call ImportarYD42
    Call AtualizarStatus("Importando arquivos de Convenios")
    Call ImportarArquivoConvenios
    Call AtualizarStatus("Exportando arquivo YD42...")
    Call GerarRelatorioYD42
    Call AtualizarStatus("Arquivo salvo com sucesso...")
    
End Sub
Private Sub Form_Open(Cancel As Integer)
    DoCmd.RunMacro "MCROculta"
    Me.TxUsuario.Value = PesqUsername
End Sub
Private Sub lb_HOME_Click()

    Dim Form As String: Form = "FrmRelatorios"
    Dim FormClose As String: FormClose = "Frm_VencidosYD"
        
        DoCmd.OpenForm Form, acNormal
        DoCmd.Close acForm, FormClose
End Sub
Private Sub lb_HOME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption lb_HOME
End Sub

