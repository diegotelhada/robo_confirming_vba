'Form_FrmPeriodicidade'

Option Compare Database
Function fSetControlOption(ctl As Control)
 Call GetCursor
End Function
Private Sub Form_Open(Cancel As Integer)

Dim Db As Database
Dim TbDados As Recordset
Dim Db1 As Database
Dim TbDados1 As Recordset

SiglaUser = String(255, 0)
Ret = GetUserName(SiglaUser, Len(SiglaUser))

'elimina os nulos da variavel Usuário
X = 1
Do While Asc(Mid(SiglaUser, X, 1)) <> 0
    X = X + 1
Loop
    SiglaUser = Left(SiglaUser, (X - 1))
    SiglaPesq = UCase(SiglaUser)

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

Set Db1 = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados1 = Db1.OpenRecordset("SELECT TblPeriodicidade.Código, TblPeriodicidade.Diario, TblPeriodicidade.Semanal, TblPeriodicidade.Semanal_1, TblPeriodicidade.Quinzenal, TblPeriodicidade.Quinzenal_1, TblPeriodicidade.Mensal, TblPeriodicidade.Usuario FROM TblPeriodicidade;", dbOpenDynaset)

Me.CbDiario.Value = TbDados1!Diario
Me.cbSemanal.Value = TbDados1!Semanal
Me.CbSemanal1.Value = TbDados1!Semanal_1
Me.CbQuinzenal.Value = TbDados1!Quinzenal
Me.CbQuinzenal1.Value = TbDados1!Quinzenal_1
Me.CbMensal.Value = TbDados1!Mensal
Me.TxUsuarioAlteracao.Value = TbDados1!Usuario

TbDados1.Close


End Sub

Private Sub Imagem32_Click()

Dim Form As String
Dim FormClose As String

FormClose = "FrmPeriodicidade"
Form = "FrmRelatorios"

DoCmd.OpenForm Form, acNormal
DoCmd.Close acForm, FormClose

End Sub

Private Sub Imagem32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem32
End Sub

Private Sub Imagem33_Click()

Dim Db1 As Database
Dim TbDados1 As Recordset

If IsNull(Me.CbDiario.Value) Or Me.CbDiario.Value = "" Then
MsgBox "Favor inserir Valor no Campo Diario", , "Mensagem"
GoTo Fim
ElseIf IsNull(Me.cbSemanal.Value) Or Me.cbSemanal.Value = "" Then
MsgBox "Favor inserir Valor no Campo Semanal", , "Mensagem"
GoTo Fim
'ElseIf IsNull(Me.CbSemanal1.Value) Or Me.CbSemanal1.Value = "" Then
'MsgBox "Favor inserir Valor no Campo Semanal1", , "Mensagem"
'GoTo Fim
ElseIf IsNull(Me.CbQuinzenal.Value) Or Me.CbQuinzenal.Value = "" Then
MsgBox "Favor inserir Valor no Campo Quinzenal", , "Mensagem"
GoTo Fim
ElseIf IsNull(Me.CbQuinzenal1.Value) Or Me.CbQuinzenal1.Value = "" Then
MsgBox "Favor inserir Valor no Campo Quinzenal1", , "Mensagem"
GoTo Fim
ElseIf IsNull(Me.CbMensal.Value) Or Me.CbMensal.Value = "" Then
MsgBox "Favor inserir Valor no Campo Mensal", , "Mensagem"
GoTo Fim
End If

Set Db1 = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Db1.Execute ("Delete TblPeriodicidade.Código, TblPeriodicidade.Diario, TblPeriodicidade.Semanal, TblPeriodicidade.Semanal_1, TblPeriodicidade.Quinzenal, TblPeriodicidade.Quinzenal_1, TblPeriodicidade.Mensal FROM TblPeriodicidade;")

Set TbDados1 = Db1.OpenRecordset("TblPeriodicidade", dbOpenDynaset)

TbDados1.AddNew

TbDados1!Diario = "Diario"
TbDados1!Semanal = Left(Me.cbSemanal.Value, 3)
'TbDados1!Semanal_1 = Left(Me.CbSemanal1.Value, 3)
TbDados1!Quinzenal = Me.CbQuinzenal.Value
TbDados1!Quinzenal_1 = Me.CbQuinzenal1.Value
TbDados1!Mensal = Me.CbMensal.Value
TbDados1!Usuario = Me.TxUsuario.Value

TbDados1.Update

TbDados1.Close

MsgBox "Alterado Com Sucesso", , "Mensagem"

Fim:

End Sub

Private Sub Imagem33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem33

'Call Imagem33_Click

End Sub
