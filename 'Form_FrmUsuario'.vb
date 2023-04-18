'Form_FrmUsuario'

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

End Sub

Private Sub Imagem32_Click()

Dim Form As String
Dim FormClose As String

FormClose = "FrmUsuario"
Form = "FrmRelatorios"

DoCmd.OpenForm Form, acNormal
DoCmd.Close acForm, FormClose




End Sub

Private Sub Imagem32_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem32
End Sub

Private Sub Imagem33_Click()

Dim Db As Database
Dim TbDados As Recordset
Dim Siglapesquisa As String

If IsNull(Me.TxNome.Value) Or Me.TxNome.Value = "" Then
MsgBox "Favor Inserir o Nome do Usuario"
GoTo Fim
ElseIf IsNull(Me.TxEmail.Value) Or Me.TxEmail.Value = "" Then
MsgBox "Favor Inserir o Email do Usuario"
GoTo Fim
ElseIf IsNull(Me.TxSigla.Value) Or Me.TxSigla.Value = "" Then
MsgBox "Favor Inserir a Sigla de acesso do Usuario"
GoTo Fim
End If

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("TblUsuarios", dbOpenDynaset)

Siglapesquisa = Me.TxSigla.Value


TbDados.FindFirst "Sigla like '*" & Siglapesquisa & "'"
    
If TbDados.NoMatch = False Then
      
Me.TxNome = TbDados!Nome
Me.TxEmail = TbDados!Email
Me.TxSigla = TbDados!Sigla

MsgBox "Usuario ja Cadastrado!", , "Mensagem"

TbDados.Close

GoTo Fim

Else

TbDados.AddNew

TbDados!Sigla = Me.TxSigla.Value
TbDados!Email = Me.TxEmail.Value
TbDados!Nome = Me.TxNome.Value

TbDados.Update

TbDados.Close

End If

MsgBox "Usuario Cadastrado com Sucesso", , "Mensagem"

Fim:

End Sub

Private Sub Imagem33_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem33
End Sub

Private Sub Imagem49_Click()

Me.TxNome = Empty
Me.TxEmail = Empty
Me.TxSigla = Empty

End Sub

Private Sub Imagem49_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem49
End Sub

Private Sub Imagem51_Click()

Dim Db As Database
Dim TbDados As Recordset
Dim Siglapesquisa As String

If IsNull(Me.TxSigla.Value) Or Me.TxSigla.Value = "" Then
MsgBox "Favor Inserir a Sigla de acesso do Usuario"
GoTo Fim
End If

Set Db = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")

Set TbDados = Db.OpenRecordset("TblUsuarios", dbOpenDynaset)

Siglapesquisa = Me.TxSigla.Value


TbDados.FindFirst "Sigla like '*" & Siglapesquisa & "'"
    
If TbDados.NoMatch = False Then

Me.TxNome = TbDados!Nome
Me.TxEmail = TbDados!Email
'Me.TxSigla = TbDados!Sigla
      
MsgBox "Usuario ja Cadastrado!", , "Mensagem"

TbDados.Close

GoTo Fim

Else

MsgBox "Usuario Não Cadastrado", , "Mensagem"

End If

Fim:

End Sub


Private Sub Imagem51_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
fSetControlOption Imagem51
End Sub
