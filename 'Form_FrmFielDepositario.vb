'Form_FrmFielDepositario"

ption Compare Database
Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function

Private Sub CbConvenio_AfterUpdate()

'On Error GoTo Listagem

Dim Db As Database
Dim TbDados As Recordset
Dim Agencia As Integer
Dim Convenio As String

If Me.TxAgencia.Value = "" Then
MsgBox "Favor inserir numero da Agencia", vbCritical, "Atenção"
GoTo Fim
End If

If IsNull(Me.CbConvenio) Or Me.CbConvenio = "" Then: GoTo Fim

Agencia = Me.TxAgencia.Value
Convenio = Me.CbConvenio.Value

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("SELECT TblClientes.Nome_Ancora, TblClientes.Agencia_Ancora, TblClientes.Convenio_Ancora, TblClientes.Cnpj_Ancora FROM TblClientes WHERE ((TblClientes.Agencia_Ancora)=" & Agencia & ") AND ((TblClientes.Convenio_Ancora)='" & Convenio & "');", dbOpenDynaset)

    Me.TxNomeAnc = TbDados!Nome_Ancora
    Me.TxCNPJ = TbDados!Cnpj_Ancora

Listagem:
Select Case Err.Number
Case 94
MsgBox "Favor inserir um valor valido", vbCritical, "Atenção"
Me.TxAgencia = Empty
Me.TxCNPJ = Empty
Me.CbConvenio.Value = Empty And Me.CbConvenio.RowSource = ""
GoTo Fim

Case -2147352567
MsgBox "Favor inserir um valor valido", vbCritical, "Atenção"
Me.TxAgencia = Empty
Me.TxCNPJ = Empty
Me.CbConvenio.Value = Empty And Me.CbConvenio.RowSource = ""
GoTo Fim

'Case 3022

'Case Else
 'MsgBox "Relate ao Suporte:  " & Err.Number & " - " & Err.Description
End Select
    
Fim:

End Sub

Private Sub Comando54_Click()

''On Error GoTo Listagem

Dim ObjExcel As Object, TbDados1 As Recordset
Dim Db As Database
Dim TbDados As Recordset
Dim Nome As String
Dim AGENCIAPESQ As Integer

If IsNull(Me.TxAgencia) Or Me.TxAgencia = "" Then
    MsgBox "Favor Inserir Agência do Convenio para pesquisa", , "Mensagem"
    GoTo Fim
ElseIf IsNull(Me.CbConvenio) Or Me.CbConvenio = "" Then
    MsgBox "Favor Inserir número do Convenio para pesquisa", , "Mensagem"
    GoTo Fim
'ElseIf IsNull(Me.DtInicio) Or Me.DtInicio = "" Then
'    MsgBox "Favor Inserir Data de inicio da pesquisa", , "Mensagem"
'    GoTo Fim
'ElseIf IsNull(Me.DtFim) Or Me.DtFim = "" Then
'    MsgBox "Favor Inserir Data final da pesquisa", , "Mensagem"
'    GoTo Fim
End If

'DATAFIM = Format(Me.DtFim, "mm/dd/yyyy")
ConvenioPesq = Me.CbConvenio
AGENCIAPESQ = Me.TxAgencia
'DataInicio = Format(Me.DtInicio, "mm/dd/yyyy")

'Me.LbValida.Caption = "Aguarde...  Pesquisando Operações... "
'Me.LbValida.FontBold = True
'Me.LbValida.ForeColor = &H800000

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Banco_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Modalidade_Oper, TblArqoped.Tipo_Liq, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Tipo_Pag, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Data_op, TblArqoped.Data_Final, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Juros, TblArqoped.Custo, TblArqoped.Spread, TblArqoped.Spread_Anual, TblArqoped.Valor_op, TblArqoped.Valor_TCO, TblArqoped.Valot_TTR, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Valor_Abat, TblArqoped.Valor_Acres, TblArqoped.Valor_Pagmto," _
 & " TblArqoped.Valor_Juros , TblArqoped.Valor_IOF, TblArqoped.Valor_Liquido, TblArqoped.Valor_Custo, TblArqoped.Spread_Banco, TblArqoped.Receita_Banco, TblArqoped.Tp_Apur_prem, TblArqoped.Tp_Rem_prem, TblArqoped.Tp_Pgto_prem, TblArqoped.Dt_pfto_Prem , TblArqoped.Cod_Bco_Prem, TblArqoped.Cod_Age_Prem, TblArqoped.Cod_Conta_prem, TblArqoped.Rate_Spread, TblArqoped.Spread_Clte, TblArqoped.Receita_Clte FROM TblArqoped" _
 & " GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Banco_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Nome_Fornecedor, TblArqoped.Cnpj_Fornecedor, TblArqoped.Cod_Oper, TblArqoped.Modalidade_Oper, TblArqoped.Tipo_Liq, TblArqoped.Banco_Remet, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, TblArqoped.Tipo_Pag, TblArqoped.Banco_Fav, TblArqoped.Agencia_Fav, TblArqoped.Conta_fav, TblArqoped.Data_op, TblArqoped.Data_Final, TblArqoped.Prazo_Medio, TblArqoped.Prazo_NF, TblArqoped.Juros, TblArqoped.Custo, TblArqoped.Spread, TblArqoped.Spread_Anual, TblArqoped.Valor_op, TblArqoped.Valor_TCO, TblArqoped.Valot_TTR, TblArqoped.Compromisso, TblArqoped.Data_Venc, TblArqoped.Valor_Nom, TblArqoped.Valor_Abat, TblArqoped.Valor_Acres, TblArqoped.Valor_Pagmto," _
 & " TblArqoped.Valor_Juros , TblArqoped.Valor_IOF, TblArqoped.Valor_Liquido, TblArqoped.Valor_Custo, TblArqoped.Spread_Banco, TblArqoped.Receita_Banco, TblArqoped.Tp_Apur_prem, TblArqoped.Tp_Rem_prem, TblArqoped.Tp_Pgto_prem, TblArqoped.Dt_pfto_Prem, TblArqoped.Cod_Bco_Prem, TblArqoped.Cod_Age_Prem, TblArqoped.Cod_Conta_prem, TblArqoped.Rate_Spread, TblArqoped.Spread_Clte, TblArqoped.Receita_Clte HAVING (((TblArqoped.Agencia_Ancora) = " & AGENCIAPESQ & ") And ((TblArqoped.Convenio_Ancora) = '" & ConvenioPesq & "') And ((TblArqoped.Data_op) >= #" & DataInicio & "# And (TblArqoped.Data_op) <= #" & DATAFIM & "#)) ORDER BY TblArqoped.Data_op;", dbOpenDynaset)


'If TbDados.EOF = True Then
'Me.LbValida.Caption = "O Cliente não possui Operação na data selecionada!"
'Me.LbValida.FontBold = True
'Me.LbValida.ForeColor = &HFF&
'GoTo Fim
'End If

Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"

   Set ObjExcel = CreateObject("EXCEL.application")
   ObjExcel.Workbooks.Open FileName:=Caminho & "PesqOper.xlsx", ReadOnly:=True
   Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
   
    NomeCliente = TbDados!Nome_Ancora
    dataPlan = Format(Date, "MM/DD/YYYY")
   
   linha = 9
   
   ObjPlan1Excel.Range("F2") = NomeCliente
   ObjPlan1Excel.Range("F4") = dataPlan
   ObjPlan1Excel.Range("F4").NumberFormat = "dd/mm/yyyy"
   ObjPlan1Excel.Range("A9").CopyFromRecordset TbDados
   
   UltimaLinha = TbDados.RecordCount
   UltimaLinha = UltimaLinha + 8
   
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(7).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(8).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(9).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(10).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(11).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(1).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(2).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(3).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Borders(4).LineStyle = 2
  ObjPlan1Excel.Range("A" & linha & ":AX" & UltimaLinha).Font.Size = 8
  ObjPlan1Excel.Range("B" & linha & ":E" & UltimaLinha).NumberFormat = "00000"
  ObjPlan1Excel.Range("G" & linha & ":Q" & UltimaLinha).NumberFormat = "00000"
  ObjPlan1Excel.Range("R" & linha & ":S" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
  ObjPlan1Excel.Range("AD" & linha & ":AD" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
  ObjPlan1Excel.Range("T" & linha & ":Y" & UltimaLinha).NumberFormat = "00000"
  ObjPlan1Excel.Range("AE" & linha & ":AL" & UltimaLinha).Style = "Currency"
  ObjPlan1Excel.Range("AN" & linha & ":AN" & UltimaLinha).Style = "Currency"
  ObjPlan1Excel.Range("AR" & linha & ":AR" & UltimaLinha).NumberFormat = "00000"
  ObjPlan1Excel.Range("AS" & linha & ":AX" & UltimaLinha).NumberFormat = "00000"
  ObjPlan1Excel.Columns("A:AX").Select
  ObjPlan1Excel.Columns.AutoFit
  ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
  ObjPlan1Excel.Range("C2").Select

    Data = Date

   Nome = "Relatorio de Operacoes - " & NomeCliente & " - " & Format(Data, "ddmmyy")
      
   Nome = Trata_NomeArquivo(Nome)
   
         
    sFname = "C:\Temp\" & Nome & ".xlsx"
    If (Dir(sFname) <> "") Then
        Kill sFname
    End If
      
   ObjPlan1Excel.SaveAs FileName:="C:\Temp\" & Nome & ".xlsx"
   ObjExcel.activeworkbook.Close SaveChanges:=False
   ObjExcel.Quit
   
'    Me.LbValida.Caption = "Arquivo Salvo com Sucesso"
'    Me.LbValida.FontBold = True
'    Me.LbValida.ForeColor = &H4000&

    CAMINHOABIR = "excel.exe C:\Temp\" & Nome & ".xlsx"
       Call Shell(CAMINHOABIR, 1)
       
TbDados.Close

Fim:

End Sub

'Private Sub Comando54_Exit(Cancel As Integer)
'fSetControlOption Comando54
'End Sub

'Private Sub DtFim_GotFocus()
'Me.DtFim.BackColor = &H80FFFF
'End Sub

'Private Sub DtFim_LostFocus()
'Me.DtFim.BackColor = &HFFFFFF
'End Sub

'Private Sub DtInicio_GotFocus()
'Me.DtInicio.BackColor = &H80FFFF
'End Sub

'Private Sub DtInicio_LostFocus()
'Me.DtInicio.BackColor = &HFFFFFF
'End Sub

Private Sub Form_Open(Cancel As Integer)
Me.CbConvenio.RowSource = ""
Me.CbConvenio.Value = ""
Me.TxNomeAnc = Empty
Me.TxAgencia = Empty
Me.TxCNPJ = Empty

End Sub

Private Sub Imagem32_Click()

Dim Form As String
Dim FormClose As String

Form = "FrmRelatorios"
FormClose = "FrmPesqOperacoes"

DoCmd.OpenForm Form
DoCmd.Close acForm, FormClose


End Sub

Private Sub Imagem32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem32
End Sub

Private Sub Imagem59_Click()

Dim Db As Database
Dim TbDados As Recordset
Dim Agencia As Integer

If IsNull(Me.TxAgencia) Or Me.TxAgencia = "" Then
MsgBox "Favor Inserir Agencia do Cliente"
GoTo Fim
ElseIf IsNull(Me.CbConvenio) Or Me.CbConvenio = "" Then
MsgBox "Favor Inserir o Convenio do Cliente"
GoTo Fim
End If

Agencia = Me.TxAgencia.Value
ConvenioPesq = Me.CbConvenio.Value

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("SELECT TblFielDepositario.Banco, TblFielDepositario.Agencia, TblFielDepositario.Convenio, TblFielDepositario.[Nome do convênio], TblFielDepositario.[Dt cadastro], TblFielDepositario.Situação, TblFielDepositario.Bloqueio, TblFielDepositario.Ambiente FROM TblFielDepositario WHERE (((TblFielDepositario.Agencia)=" & Agencia & ") AND ((TblFielDepositario.Convenio)='" & ConvenioPesq & "'));", dbOpenDynaset)

If TbDados.EOF = True Then

TbDados.AddNew

    TbDados!banco = "33"
    TbDados!Agencia = Agencia
    TbDados!Convenio = Format(ConvenioPesq, "000000000000")
    TbDados![Nome do convênio] = Me.TxNomeAnc
    TbDados![Dt cadastro] = Format(Date, "dd/mm/yyyy")
    TbDados!Situação = "ATIVO"
    TbDados!Bloqueio = "Sem Bloqueio"
    TbDados!Ambiente = "PRODUÇÃO"
    
    Me.CbConvenio.RowSource = ""
    Me.CbConvenio.Value = ""
    Me.TxNomeAnc = Empty
    Me.TxAgencia = Empty
    Me.TxCNPJ = Empty

TbDados.Update

MsgBox "Cliente Cadastrado Com Sucesso!", vbInformation, "Cadastro de Fiel Depositario"

Else

MsgBox "Cliente ja cadastrado na Tabela de Fiel Depositario", vbInformation, "Cadastro de Fiel Depositario"

End If

Fim:

End Sub

Private Sub Imagem59_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem59
End Sub

Private Sub Imagem62_Click()
Me.CbConvenio.RowSource = ""
Me.CbConvenio.Value = ""
Me.TxNomeAnc = Empty
Me.TxAgencia = Empty
Me.TxCNPJ = Empty
End Sub

Private Sub Imagem62_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem62
End Sub

Private Sub TxAgencia_AfterUpdate()

'On Error GoTo Listagem

Dim Db As Database
Dim TbDados As Recordset
Dim Agencia As Integer

Me.CbConvenio.RowSource = ""
Me.TxNomeAnc = Empty

If IsNull(Me.TxAgencia) Or Me.TxAgencia = "" Then: GoTo Fim

Agencia = Me.TxAgencia

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("SELECT TblClientes.Agencia_Ancora, TblClientes.Convenio_Ancora FROM TblClientes GROUP BY TblClientes.Agencia_Ancora, TblClientes.Convenio_Ancora HAVING ((TblClientes.Agencia_Ancora)=" & Agencia & ");", dbOpenDynaset)

Me.CbConvenio.RowSourceType = "Value List"

Do While TbDados.EOF = False

Me.CbConvenio.AddItem TbDados!Convenio_Ancora
                                                 
TbDados.MoveNext
             
Loop


Listagem:
Select Case Err.Number
Case 13
MsgBox "Favor inserir um valor Valido", vbCritical, "Atenção"
Me.TxAgencia = Empty
GoTo Fim

'Case Else
'MsgBox "Relate ao Suporte:  " & Err.Number & " - " & Err.Description
End Select

Fim:

End Sub
