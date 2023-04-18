'Form_FrmPesqWaiver'

Option Compare Database
Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function

Private Sub CbConvenio_AfterUpdate()

'On Error GoTo Listagem

Dim Db As Database
Dim TbDados As Recordset
Dim Agencia As Integer
Dim Convenio As String

'If Me.TxAgencia.Value = "" Then
'MsgBox "Favor inserir numero da Agencia", vbCritical, "Atenção"
'GoTo Fim
'End If

'If IsNull(Me.CbConvenio) Or Me.CbConvenio = "" Then: GoTo Fim

'Agencia = Me.TxAgencia.Value
'Convenio = Me.CbConvenio.Value

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("SELECT TblClientes.Nome_Ancora, TblClientes.Agencia_Ancora, TblClientes.Convenio_Ancora, TblClientes.Cnpj_Ancora FROM TblClientes WHERE ((TblClientes.Agencia_Ancora)=" & Agencia & ") AND ((TblClientes.Convenio_Ancora)='" & Convenio & "');", dbOpenDynaset)

'    Me.TxNomeAnc = TbDados!Nome_Ancora

Listagem:
Select Case Err.Number
Case 94
MsgBox "Favor inserir um valor valido", vbCritical, "Atenção"
'Me.TxAgencia = Empty
'Me.CbConvenio.Value = Empty And Me.CbConvenio.RowSource = ""
GoTo Fim

Case -2147352567
MsgBox "Favor inserir um valor valido", vbCritical, "Atenção"
'Me.TxAgencia = Empty
'Me.CbConvenio.Value = Empty And Me.CbConvenio.RowSource = ""
GoTo Fim

'Case 3022

'Case Else
 'MsgBox "Relate ao Suporte:  " & Err.Number & " - " & Err.Description
End Select
    
Fim:

End Sub

Private Sub Comando54_Click()

Call GerarRelatoriosWaiver

MsgBox "Relatorio salvo em:" & Chr(13) & "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\"

End Sub

Private Sub Comando54_Exit(Cancel As Integer)
fSetControlOption Comando54
End Sub

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

Private Sub Imagem32_Click()

Dim Form As String
Dim FormClose As String

Form = "FrmRelatorios"
FormClose = "FrmPesqWaiver"

DoCmd.OpenForm Form
DoCmd.Close acForm, FormClose

End Sub

Private Sub Imagem32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem32
End Sub


Private Sub Imagem57_Click()
Call atualizabaseAncora
Call atualizabasefornecedor

MsgBox "Base Atualizada com sucesso!", vbInformation, "Waiver Corporate"

Fim:
End Sub

Private Sub Imagem57_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem57
End Sub
