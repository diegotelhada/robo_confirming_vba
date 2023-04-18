'Form_Frm_SubTermoVirtual'

Option Compare Database
Function FecharForm()
    DoCmd.Close acForm, "Frm_SubTermoVirtual"
End Function
Function ValidaConvenioCadastrado(Agencia, Convenio) As Boolean
    Dim TbValida As Recordset
        Set TbValida = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.Agencia, Tbl_Convenio_Agrupados.Convenio FROM Tbl_Convenio_Agrupados WHERE (((Tbl_Convenio_Agrupados.Agencia)='" & Agencia & "') AND ((Tbl_Convenio_Agrupados.Convenio)='" & Convenio & "'));", dbOpenDynaset)
            ValidaConvenioCadastrado = TbValida.EOF
End Function
Function LimparCampos()
    Me.CbAgencia.Value = Empty
    Me.CbConvenio.Value = Empty
    Me.TxNomeConvenio.Value = Empty
End Function
Function PreencherComboGrupo()
    Dim TbGrupo As Recordset: Forms!Frm_TermoVirtual!CbGrupo.RowSource = ""
        Set TbGrupo = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados GROUP BY Tbl_Convenio_Agrupados.Grupo ORDER BY Tbl_Convenio_Agrupados.Grupo;", dbOpenDynaset)
            Do While TbGrupo.EOF = False
                    Forms!Frm_TermoVirtual!CbGrupo.AddItem TbGrupo!Grupo
                TbGrupo.MoveNext
            Loop
        Forms!Frm_TermoVirtual!CbGrupo.Value = Me.CbGrupo.Value
End Function
Function PreencherNomeConvenio(Agencia, Convenio)
    Dim TbConvenio As Recordset:
        Set TbConvenio = DBTVirtual.OpenRecordset("SELECT TblClientes.Nome_Ancora, TblClientes.Agencia_Ancora, TblClientes.Convenio_Ancora FROM TblClientes WHERE (((TblClientes.Agencia_Ancora)=" & Agencia & ") AND ((TblClientes.Convenio_Ancora)='" & Convenio & "'));", dbOpenDynaset)
            If TbConvenio.EOF = False Then: Me.TxNomeConvenio.Value = UCase(Trim(TbConvenio!Nome_Ancora))
End Function
Function PreencherConvenio(Agencia)
    Dim TbConvenio As Recordset: Me.CbConvenio.RowSource = ""
        Set TbConvenio = DBTVirtual.OpenRecordset("SELECT TblClientes.Convenio_Ancora FROM TblClientes WHERE (((TblClientes.Agencia_Ancora) =" & Agencia & ")) GROUP BY TblClientes.Convenio_Ancora ORDER BY TblClientes.Convenio_Ancora;", dbOpenDynaset)
            Do While TbConvenio.EOF = False
                Me.CbConvenio.AddItem TbConvenio!Convenio_Ancora
                TbConvenio.MoveNext
            Loop
End Function
Function PreencherAgencia()
    Dim TbAgencia As Recordset
        Set TbAgencia = DBTVirtual.OpenRecordset("SELECT TblClientes.Agencia_Ancora FROM TblClientes GROUP BY TblClientes.Agencia_Ancora ORDER BY TblClientes.Agencia_Ancora;", dbOpenDynaset)
            Do While TbAgencia.EOF = False
                Me.CbAgencia.AddItem TbAgencia!Agencia_Ancora
                TbAgencia.MoveNext
            Loop
End Function
Function PesquisarConveniosAgrupados(Grupo)
    Dim TbGrupo As Recordset
        Set TbGrupo = DBTVirtual.OpenRecordset("SELECT Tbl_Convenio_Agrupados.ID, Tbl_Convenio_Agrupados.Agencia, Tbl_Convenio_Agrupados.Convenio, Tbl_Convenio_Agrupados.Nome_Convenio, Tbl_Convenio_Agrupados.Grupo FROM Tbl_Convenio_Agrupados WHERE (((Tbl_Convenio_Agrupados.Grupo) = '" & Grupo & "')) ORDER BY Tbl_Convenio_Agrupados.Nome_Convenio;", dbOpenDynaset)
            If TbGrupo.EOF = False Then: Call PreencherListBox(TbGrupo)
End Function
Function PreencherListBox(TbGrupo As Recordset)

    Forms!Frm_TermoVirtual!ListConvenios.RowSource = ""
    Forms!Frm_TermoVirtual!ListConvenios.RowSourceType = "Value List"
    Forms!Frm_TermoVirtual!ListConvenios.ColumnCount = 4
    Forms!Frm_TermoVirtual!ListConvenios.ColumnHeads = True
    Forms!Frm_TermoVirtual!ListConvenios.AddItem ("ID") & ";" & ("Agencia") & ";" & ("Convênio") & ";" & ("Nome")
    
        Do While TbGrupo.EOF = False
            Forms!Frm_TermoVirtual!ListConvenios.AddItem TbGrupo!ID & ";" & TbGrupo!Agencia & ";" & TbGrupo!Convenio & ";" & TbGrupo!nome_convenio
            TbGrupo.MoveNext
        Loop

End Function
Function ValidarCampos()
    
    If IsNull(Me.CbAgencia.Value) Or Me.CbAgencia.Value = "" Then: MsgBox "Selecione a Agencia do Âncora!", vbInformation, "Termo Virtual": ValidarCampos = "Erro": GoTo Fim
    If IsNull(Me.CbConvenio.Value) Or Me.CbConvenio.Value = "" Then: MsgBox "Selecione o Convênio do Âncora!", vbInformation, "Termo Virtual": ValidarCampos = "Erro": GoTo Fim
    If IsNull(Me.TxNomeConvenio) Or Me.TxNomeConvenio.Value = "" Then: MsgBox "Informe o nome do Âncora!", vbInformation, "Termo Virtual": ValidarCampos = "Erro": GoTo Fim
Fim:
End Function
Function SalvarDados()

    Dim TbDados As Recordset

        If IsNull(Me.TxID) Then
            Valida = ValidaConvenioCadastrado(Me.CbAgencia.Value, Me.CbConvenio.Value)
            If Valida = True Then
                DBTVirtual.Execute ("INSERT INTO Tbl_Convenio_Agrupados ( Agencia, Convenio, Nome_Convenio, DataHora, Usuario, Grupo ) SELECT '" & Me.CbAgencia.Value & "' AS Agencia, '" & Me.CbConvenio.Value & "' AS Convenio, '" & Me.TxNomeConvenio.Value & "' AS Nome_Convenio, #" & Date & " " & Time & "# AS DataHora, '" & Forms!Frm_TermoVirtual!TxUsuario.Value & "' AS Usuario, '" & Me.CbGrupo.Value & "' AS Grupo;")
                MsgBox "Convenio Cadastrado com sucesso!", vbInformation, "Termo Virtual"
                Call PesquisarConveniosAgrupados(Me.CbGrupo.Value)
                Call PreencherComboGrupo
                VbPgt = MsgBox("Deseja Cadastrar mais convenios para este grupo?", vbYesNo, "Termo Virtual")
                If VbPgt = vbYes Then: Call LimparCampos
                If VbPgt = vbNo Then: Call FecharForm
            Else
                MsgBox "Convênio já cadastrado!", vbInformation, "Termo Virtual"
                Call LimparCampos
            End If
        Else
            DBTVirtual.Execute ("UPDATE Tbl_Convenio_Agrupados SET Tbl_Convenio_Agrupados.Agencia = '" & Me.CbAgencia.Value & "', Tbl_Convenio_Agrupados.Convenio = '" & Me.CbConvenio.Value & "', Tbl_Convenio_Agrupados.Nome_Convenio = '" & Me.TxNomeConvenio.Value & "', Tbl_Convenio_Agrupados.DataHora = #" & Date & " " & Time & "#, Tbl_Convenio_Agrupados.Usuario = '" & Forms!Frm_TermoVirtual!TxUsuario.Value & "' WHERE (((Tbl_Convenio_Agrupados.ID)=" & Me.TxID & "));")
            MsgBox "Convenio Alterado com sucesso!", vbInformation, "Termo Virtual"
            Call PesquisarConveniosAgrupados(Me.CbGrupo.Value)
            Call FecharForm
        End If
End Function
Private Sub CbAgencia_AfterUpdate()
    If Not IsNull(Me.CbAgencia.Value) Or Me.CbAgencia.Value <> "" Then: Call PreencherConvenio(Me.CbAgencia.Value)
End Sub
Private Sub CbConvenio_AfterUpdate()
    If Not IsNull(Me.CbAgencia.Value) Or Me.CbAgencia.Value <> "" And Not IsNull(Me.CbConvenio.Value) Or Me.CbConvenio.Value <> "" Then: Call PreencherNomeConvenio(Me.CbAgencia.Value, Me.CbConvenio.Value)
End Sub
Private Sub CmdSalvar_Click()
    Me.CbAgencia.SetFocus
    Valida = ValidarCampos
        If Valida <> "Erro" Then
            Call SalvarDados
        End If
End Sub
Private Sub Form_Open(Cancel As Integer)
    Call AbrirDBTVirtual
    Call PreencherAgencia
End Sub
