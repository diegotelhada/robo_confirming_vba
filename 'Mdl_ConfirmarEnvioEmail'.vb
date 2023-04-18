'Mdl_ConfirmarEnvioEmail'


Option Compare Database

'Módulo adicionado para confirmação de envio - Emerson 20/06/2018
Public Sub confirmarEnvioEmail(ByVal Data As String)

    'Varíáveis
    Dim OutApp As Object
    Dim OutMail As Object
    Dim qryDadosAntecipadas, qryDadosAVencer As Recordset
    Dim enviadosAntecipadas, enviadosAVencer As Integer
    
    'Quantidade enviada de relatórios
    Call AbrirBDLocal
    
    'Notas Antecipadas
    SQL = "SELECT Count(TblRelatoriosEnviados.Código) AS TotalEnviado, TblRelatoriosEnviados.Data_Envio " & _
    "FROM TblRelatoriosEnviados WHERE (((TblRelatoriosEnviados.Relatorio_Enviado) Like '*Antecipadas*') " & _
    "AND ((TblRelatoriosEnviados.Periodicidade_Relatorio)='Diario') AND ((TblRelatoriosEnviados.Data_Envio)=#" & Format(Data, "mm/dd/yyyy") & "#)) " & _
    "GROUP BY TblRelatoriosEnviados.Data_Envio"
    Set qryDadosAntecipadas = BDRELocal.OpenRecordset(SQL)
    enviadosAntecipadas = qryDadosAntecipadas!TotalEnviado
    qryDadosAntecipadas.Close
    Set qryDadosAntecipadas = Nothing
    
    'Notas a Vencer
    SQL = "SELECT Count(TblRelatoriosEnviados.Código) AS TotalEnviado, TblRelatoriosEnviados.Data_Envio " & _
    "FROM TblRelatoriosEnviados WHERE (((TblRelatoriosEnviados.Relatorio_Enviado) Like '*A Vencer*') " & _
    "AND ((TblRelatoriosEnviados.Periodicidade_Relatorio)='Diario') AND ((TblRelatoriosEnviados.Data_Envio)=#" & Format(Data, "mm/dd/yyyy") & "#)) " & _
    "GROUP BY TblRelatoriosEnviados.Data_Envio"
    Set qryDadosAVencer = BDRELocal.OpenRecordset(SQL)
    enviadosAVencer = qryDadosAVencer!TotalEnviado
    qryDadosAVencer.Close
    Set qryDadosAVencer = Nothing
    
    'Integração com o Outlook
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'De
    OutMail.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
    
    'Para
    OutMail.To = "eaolivei@santander.com.br;apdmoraes@santander.com.br"
    
    'CC
    OutMail.cc = "ldtrevisan@santander.com.br;wellington.da.silva@santander.com.br;jomedeiros@santander.com.br;emanuela.conceicao@santander.com.br; erisousa@santander.com.br;lfrossi@santander.com.br"
    
    'Assunto
    OutMail.Subject = "Confirmação de Envio Relatórios Confirming - " & Data
    
'    'Anexos
'    Dim myAttachments As Outlook.Attachments
'    Set myAttachments = OutMail.Attachments
'
'    Dim file1 As String
'    file1 = ""
'    myAttachments.Add file1, olByValue, 1
    
    'Montagem da mensagem
    
    'Inicialização
    OutMail.HTMLBody = ""
    
    'Conteúdo email
    OutMail.HTMLBody = OutMail.HTMLBody & _
    "<p>Relatórios enviados com sucesso para o dia <b>" & Data & "</b>.</p>" & _
    "<h3>Relatórios Diários</h3>" & _
    "<p>" & _
    "E-mails enviados Notas Antecipadas: " & CStr(enviadosAntecipadas) & ".<br>" & _
    "E-mails enviados Notas A Vencer: " & CStr(enviadosAVencer) & ".<br>" & _
    "</p>"
    
    'Envio da mensagem
    'OutMail.Display
    OutMail.Send
    
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
