'Form_Frm_TermoVirtual'

Option Compare Database
Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function
Function AtualizarStatus(Status)
    Me.LbStatus.Caption = Status: DoEvents
End Function
Function PesqNomeGrupo(Nome)

    Dim TbGrupo As Recordset
        Set TbGrupo = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.ID, Tbl_Convenio_Agrupados.Agencia, Tbl_Convenio_Agrupados.Convenio, Tbl_Convenio_Agrupados.Nome_Convenio, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados WHERE (((Tbl_Convenio_Agrupados.Grupo) = '" & Nome & "')) ORDER BY Tbl_Convenio_Agrupados.Nome_Convenio;", dbOpenDynaset)
            If TbGrupo.EOF = False Then
                PesqNomeGrupo = "Encontrado"
            End If

End Function
Function PreencherComboGrupo()
    
    Dim TbGrupo As Recordset: Me.CbGrupo.RowSource = ""
        
        Set TbGrupo = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados GROUP BY Tbl_Convenio_Agrupados.Grupo ORDER BY Tbl_Convenio_Agrupados.Grupo;", dbOpenDynaset)
            
            Do While TbGrupo.EOF = False
                    Me.CbGrupo.AddItem TbGrupo!Grupo
                TbGrupo.MoveNext
            Loop
    
End Function
Function PesquisarConveniosAgrupados(Grupo)
    
    Dim TbGrupo As Recordset
        
        Set TbGrupo = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.ID, Tbl_Convenio_Agrupados.Agencia, Tbl_Convenio_Agrupados.Convenio, Tbl_Convenio_Agrupados.Nome_Convenio, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados WHERE (((Tbl_Convenio_Agrupados.Grupo) = '" & Grupo & "')) ORDER BY Tbl_Convenio_Agrupados.Nome_Convenio;", dbOpenDynaset)
            
            If TbGrupo.EOF = False Then: Call PreencherListBox(TbGrupo)
    
End Function
Function PreencherListBox(TbGrupo As Recordset)

    Me.ListConvenios.RowSource = ""
    Me.ListConvenios.RowSourceType = "Value List"
    Me.ListConvenios.ColumnCount = 4
    Me.ListConvenios.ColumnHeads = True
    Me.ListConvenios.AddItem ("ID") & ";" & ("Agencia") & ";" & ("Convênio") & ";" & ("Nome")
    
        Do While TbGrupo.EOF = False
            Me.ListConvenios.AddItem TbGrupo!ID & ";" & TbGrupo!Agencia & ";" & TbGrupo!Convenio & ";" & TbGrupo!nome_convenio
            TbGrupo.MoveNext
        Loop

End Function
Function ValidaCampos()

'    If IsNull(Me.CbAncora) Or Me.CbAncora.Value = "" Then: MsgBox "Selecione um Âncora pra gerar o PPB", vbInformation, "Relatorios Confirming": ValidaCampos = "Erro": GoTo Fim
'    If IsNull(Me.TxDtInicio) Or TxDtInicio.Value = "" Then: MsgBox "Digite a Data de Inicio", vbInformation, "Relatorios Confirming": ValidaCampos = "Erro": GoTo Fim
'    If IsNull(Me.TxDtFinal) Or TxDtFinal.Value = "" Then: MsgBox "Digite a Data Final", vbInformation, "Relatorios Confirming": ValidaCampos = "Erro": GoTo Fim
    
'        If Me.TxDtInicio > Me.TxDtFinal Then: MsgBox "Data de Inicio maior que a data Final!": GoTo Fim
'        If Me.TxDtInicio = Me.TxDtFinal Then: MsgBox "As Datas digitadas são iguais!": GoTo Fim
Fim:
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
Private Sub CbGrupo_AfterUpdate()
    If Not IsNull(Me.CbGrupo.Value) Or Me.CbGrupo.Value <> "" Then: Call PesquisarConveniosAgrupados(Me.CbGrupo.Value)
End Sub
Private Sub CmdIncluir_Click()
    Dim Form As String: Form = "Frm_SubTermoVirtual"
        
        If IsNull(Me.CbGrupo.Value) Or Me.CbGrupo.Value = "" Then
            MsgBox "Selecione o Grupo que deseja incluir convênios!", vbInformation, "Termo Virtual"
        Else
            DoCmd.OpenForm Form, acNormal
            Forms!Frm_SubTermoVirtual!CbGrupo.Value = Me.CbGrupo.Value
        End If
End Sub
Private Sub CmdNovoGrupo_Click()
    
    Dim Form As String: Form = "Frm_SubTermoVirtual"
    
        NomeGrupo = InputBox("Digite o nome do novo Grupo:", "Termo Virtual")
            Do While NomeGrupo = ""
                NomeGrupo = InputBox("Digite o nome do novo Grupo:", "Termo Virtual")
            Loop
        
        Valida = PesqNomeGrupo(NomeGrupo)
            If Valida = "Encontrado" Then
                MsgBox "Nome do Grupo ja cadastrado!", vbInformation, "Termo Virtual"
            Else
                DoCmd.OpenForm Form, acNormal
                Forms!Frm_SubTermoVirtual!CbGrupo.Value = NomeGrupo
            End If
    
End Sub
Private Sub Form_Open(Cancel As Integer)

 '   DoCmd.RunMacro "MCROculta"
        Call AbrirDBTVirtual
        Call PesquisarUsuario
        Call PreencherComboGrupo
   
End Sub
Private Sub lb_HOME_Click()

    Dim Form As String: Form = "FrmRelatorios"
    Dim FormClose As String: FormClose = "Frm_TermoVirtual"
        
        DoCmd.OpenForm Form, acNormal
        DoCmd.Close acForm, FormClose

End Sub
Private Sub lb_HOME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fSetControlOption lb_HOME
End Sub
Private Sub ListConvenios_DblClick(Cancel As Integer)

    Dim Form As String: Form = "Frm_SubTermoVirtual"
    
        If Me.ListConvenios.ListCount > 1 Then
            DoCmd.OpenForm Form, acNormal
            Forms!Frm_SubTermoVirtual!CbGrupo.Value = Me.CbGrupo.Value
            Forms!Frm_SubTermoVirtual!CbAgencia.Value = Me.ListConvenios.Column(1)
            Forms!Frm_SubTermoVirtual!CbConvenio.Value = Me.ListConvenios.Column(2)
            Forms!Frm_SubTermoVirtual!TxNomeConvenio.Value = Me.ListConvenios.Column(3)
            Forms!Frm_SubTermoVirtual!TxID.Value = Me.ListConvenios.Column(0)
        End If
End Sub
Private Sub ListConvenios_KeyDown(KeyCode As Integer, Shift As Integer)
    
        If KeyCode = 46 Then
            If Me.ListConvenios.ListCount > 1 Then
                VbPgt = MsgBox("Deseja excluir o(s) Convenio(s) selecionado(s) ?", vbYesNo, "Termo Virtual")
                    If VbPgt = vbYes Then
                        Call AtualizarStatus("Deletando Convenio " & Me.ListConvenios.Column(3) & "...")
                            DBTVirtual.Execute ("Delete Tbl_Convenio_Agrupados.ID FROM Tbl_Convenio_Agrupados WHERE (((Tbl_Convenio_Agrupados.ID)=" & Me.ListConvenios.Column(0) & "));")
                        Call PesquisarConveniosAgrupados(Me.CbGrupo.Value)
                    End If
            End If
        End If
    Call AtualizarStatus("")
End Sub
