'Form_FrmReatorios'


Option Compare Database
Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function
Function AbrirForm(Form)
    DoCmd.OpenForm Form, acNormal
    DoCmd.Close acForm, "FrmRelatorios"
End Function
Private Sub Activar_Click()

    If Me.LbValidaAtivo.Caption = "Robo Ativo" Then
    
        Me.LbValidaAtivo.Caption = ""
        Me.LbValidaRelatorios.Caption = ""
        Me.DataArquivo.Value = ""
    
    Else
    
        Me.DataArquivo.Value = Format(Date, "dd/mm/yyyy")
        
        Me.LbValidaAtivo.Caption = "Robo Ativo"
        Me.LbValidaAtivo.FontBold = True
        Me.LbValidaAtivo.ForeColor = &HC000&
                
        Me.LbValidaRelatorios.ForeColor = &H800000
        Me.LbValidaRelatorios.FontBold = True
        Me.LbValidaRelatorios.Caption = "Pesquisando Arquivo de Operações..."
    
    End If

End Sub
Private Sub Form_Open(Cancel As Integer)

    Dim TbDados As Recordset: Call AbrirBDRelatorios
        
        'Macro p/ ocultar menus
        DoCmd.RunMacro "MCROculta"
        
        'Pesquisar silga de usuario
        SiglaPesq = PesqUsername()
            
            'Pesquisar nome do funcionario
            Set TbDados = BDREL.OpenRecordset("TblUsuarios", dbOpenDynaset)

                TbDados.FindFirst "Sigla like '*" & SiglaPesq & "'"
                
                If TbDados.NoMatch = False Then
                    NomeUsuario = TbDados!Nome
                    EmailUsuario = TbDados!Email
                    Me.TxUsuario = NomeUsuario
                Else
                    Me.TxUsuario = "Usuario não cadastrado"
                    MsgBox "Usuario não cadastrado para acessar o programa", vbCritical, "Relatorios Confirming"
                    DoCmd.Quit
                End If
            TbDados.Close
            
        Me.LbValidaAtivo.Caption = ""
        Me.LbValidaRelatorios.Caption = ""
        Me.DataArquivo.Value = ""
End Sub
Private Sub Form_Timer()
    If LbValidaAtivo.Caption = "Robo Ativo" And Format(Now(), "HHMM") > "0100" And Format(Now(), "HHMM") < "0600" Then
        Call ImportarVinculo
    End If
End Sub
Private Sub LbContratoMae_Click()
    
    Dim Form As String: Form = "Frm_TermoVirtual"
    Dim FormClose As String: FormClose = "FrmRelatorios"

        DoCmd.OpenForm Form, acNormal
        DoCmd.Close acForm, FormClose
        
End Sub
Private Sub LbContratoMae_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbContratoMae
End Sub
Private Sub LbEnviados_Click()
    Call AbrirForm("FrmArquivosEnviados")
End Sub
Private Sub LbEnviados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbEnviados
End Sub
Private Sub LbFiel_Click()
    Call AbrirForm("FrmFielDepositario")
End Sub
Private Sub LbFiel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbFiel
End Sub
Private Sub LbImportados_Click()
    Call AbrirForm("FrmArquivosImportados")
End Sub
Private Sub LbImportados_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbImportados
End Sub
Private Sub LbInclusao_Click()
    
    If Me.LbUsuario.Visible = False Then
        Call MOInclusao(True)
    Else
        Call MOInclusao(False)
    End If
    Call MOManutencao(False)
    Call MOPesquisa(False)

End Sub
Private Sub LbInclusao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbInclusao
End Sub
Private Sub LbManutencao_Click()
    Call AbrirForm("FrmPeriodicidade")
End Sub
Private Sub LbManutencao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbManutencao
End Sub
Private Sub LbManutencao1_Click()
    If Me.LbManutencao.Visible = False Then
        Call MOManutencao(True)
    Else
        Call MOManutencao(False)
    End If
    Call MOInclusao(False)
    Call MOPesquisa(False)
End Sub
Private Sub LbManutencao1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbManutencao1
End Sub
Private Sub LbOperacao_Click()
    Call AbrirForm("FrmPesqOperacoes")
End Sub
Private Sub LbOperacao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbOperacao
End Sub
Private Sub LbPesquisa_Click()
    If Me.LbEnviados.Visible = False Then
        Call MOPesquisa(True)
    Else
        Call MOPesquisa(False)
    End If
    Call MOInclusao(False)
    Call MOManutencao(False)
End Sub
Private Sub LbPesquisa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbPesquisa
End Sub
Private Sub LbPPB_Click()
    Call AbrirForm("Frm_PPB")
End Sub
Private Sub LbPPB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbPPB
End Sub
Private Sub lbpwaiver_Click()
    Call AbrirForm("FrmWaiverCorporate")
End Sub
Private Sub lbpwaiver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption lbpwaiver
End Sub
'Private Sub LbRaizen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fSetControlOption LbRaizen
'End Sub
'Private Sub LbRDC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fSetControlOption LbRDC
'End Sub
Private Sub LbReenvio_Click()
    Call AbrirForm("Frm_ReenvioRelatorios")
End Sub
Private Sub LbReenvio_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbReenvio
End Sub
Private Sub LbSair_Click()
    DoCmd.Quit
End Sub
Private Sub LbSair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbSair
End Sub
Private Sub LbUsuario_Click()
    Call AbrirForm("FrmUsuario")
End Sub
Private Sub LbUsuario_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbUsuario
End Sub
'Private Sub LbVale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fSetControlOption LbVale
'End Sub
Private Sub LbVencidosYD_Click()
    Call AbrirForm("Frm_VencidosYD")
End Sub
Private Sub LbVencidosYD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbVencidosYD
End Sub
Private Sub LbVencimento_Click()
    Call AbrirForm("FrmPesqaVencer")
End Sub
Private Sub LbVencimento_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbVencimento
End Sub
Private Sub LbWaiver_Click()
    Call AbrirForm("FrmWaiverCorporate")
End Sub
Private Sub LbWaiver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption LbWaiver
End Sub
Private Sub Rótulo102_Click()
    DoCmd.Quit
End Sub
Function MOInclusao(Comando As Boolean) 'Funcao para Mostrar/Ocultar campos de Inclusão

    Me.LbUsuario.Visible = Comando
    Me.CxUsuario.Visible = Comando
  '---
    Me.LbFiel.Visible = Comando
    Me.CxFiel.Visible = Comando
  '---
    Me.CxWaiver.Visible = Comando
    Me.LbWaiver.Visible = Comando

End Function
Function MOManutencao(Comando As Boolean) 'Funcao para Mostrar/Ocultar campos de Manutençao

    Me.CxManutencao.Visible = Comando
    Me.LbManutencao.Visible = Comando
  '---
    Me.CxContratoMae.Visible = Comando
    Me.LbContratoMae.Visible = Comando

End Function
Function MOPesquisa(Comando As Boolean) 'Funcao para Mostrar/Ocultar campos de Pesquisa

    Me.CxEnviados.Visible = Comando
    Me.LbEnviados.Visible = Comando
  '---
    Me.CxImportados.Visible = Comando
    Me.LbImportados.Visible = Comando
  '---
    Me.CxOperacao.Visible = Comando
    Me.LbOperacao.Visible = Comando
  '---
    Me.CxVencimento.Visible = Comando
    Me.LbVencimento.Visible = Comando
  '---
    Me.cxpWaiver.Visible = Comando
    Me.lbpwaiver.Visible = Comando
  '---
    Me.CxReenvio.Visible = Comando
    Me.LbReenvio.Visible = Comando
  '---
    Me.LbVencidosYD.Visible = Comando
    Me.CxVencidosYD.Visible = Comando
    
End Function

