'MDL_EnviodeTermos'

Option Compare Database
Function EnviarEmailMobile(Comando, Arquivo)

    EmailDestino = "renato.de.oliveira@santander.com.br;tarjunior@santander.com.br"
    'EmailCopia = "jorge.junior@santander.com.br;lfrossi@santander.com.br"
    Assinatura = "\\saont46\apps2\\Confirming\Produto\Documentação\Documentação\BCDADOS_FORNECEDORES\Logo.Jpg"

        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)

            oitem.Subject = ("RELATORIO DE TERMOS NÃO ENVIADOS e RESULTADO MOBILE -  " & Format(Date, "DD/MM/YYYY"))
            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
            oitem.To = EmailDestino
            oitem.cc = EmailCopia
            Corpo1 = "Prezados,"
            Confirming = ""
            Corpo31 = ""
            Fones = ""
            Corpo4 = ""
            Corpo3 = "Atenciosamente."
            Assinatura1 = "Confirming®"
            Assinatura2 = "Manufatura"
            Assinatura3 = "Rua Amador Bueno, 474"
            Assinatura4 = "Meios - Operações e Serviços"
            Assinatura5 = "CEP: 04752-005  São Paulo-SP"
            Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
            Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
            Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
            oitem.Attachments.Add "\\saont46\apps2\\Confirming\Produto\Documentação\Documentação\BCDADOS_FORNECEDORES\Logo.Jpg"
                
                If Comando = "SUCESSO" Then
                
                    Corpo2 = "Segue anexo o "
                    Relatorio = "RELATÓRIO DE TERMOS NÃO ENVIADOS"
                    Corpo21 = " referente ao(s) convênio(s) de Confirming® Santander."
                    oitem.Attachments.Add Arquivo
                
                
                ElseIf Comando = "ARQUIVO EM BRANCO" Then
                    
                    Corpo2 = "Hoje o relatorio "
                    Relatorio = "YCMOBI02"
                    Corpo21 = " foi gerado em branco, por isso não será enviado o Relátorio de Resultado Mobile."
                                
                ElseIf Comando = "NAO ENCONTRADO" Then
                
                    Corpo2 = "O "
                    Relatorio = "YCMOBI02"
                    Corpo21 = " não foi salvo na rede, por isso hoje não será enviado Relátorio de Resultado Mobile."
                
                End If

            oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Calibri <BR>" & Corpo1 & "<BR/>" & _
            "<BR>" & Corpo2 & "<B>" & Relatorio & "</B>" & Corpo21 & "<BR>" & "<BR>" & Corpo3 & "<B>" & Confirming & "</B>" & Corpo31 & "<B>" & Fones & "</B>" & "<BR>" & "<BR>" & Corpo4 & "<BR><BR><BR>" & _
            " <img src=" & Assinatura & " height=50 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 3" & "<BR>" & _
            "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 2 <BR>" & Assinatura3 & _
            "<BR/>" & Assinatura4 & "<BR/>" & Assinatura5 & "<BR/></FONT><FONT COLOR = BLACK FACE = Calibri Size = 1 <BR><I>" & Assinatura6 & _
            "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"

        oitem.Send

    'oitem.DISPLAY True
    Set olapp = Nothing
    Set oitem = Nothing


End Function
Sub EnvioTermos()

    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim ObjPlan2Excel As Object, Ret As String
    Dim Db1 As Database, TbDados As Recordset
    Dim Db As Database, TbTemp As Recordset, TbData As Recordset
    Dim Nome As String, FSO As New FileSystemObject
    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"

        ARQDATA = UltimoDiaUtil()
        DataPesq = Format(ARQDATA, "DDMMYY")
        File = "\\saont46\apps2\Confirming\ArquivosYC\YCMOBI02." & DataPesq & ".TXT"
        ''Novo Caminho
        File = "\\fscore02\apps2\Confirming\ArquivosYC\YCMOBI02." & DataPesq & ".TXT"
            
            If FSO.FileExists(File) = True Then
                    
                Set Mobile = FSO.GetFile(File)
                            
                 Tamanho = Mobile.Size
                    
                    If Tamanho > 0 Then
                    
                        Set ObjExcel = CreateObject("EXCEL.application")
                        ObjExcel.Workbooks.Open FileName:=Caminho & "TermosEnviados.xlsx", ReadOnly:=True
                        Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                        Set ObjPlan2Excel = ObjExcel.Worksheets("ResultadoMobile")
                    
                        dataPlan = UltimoDiaUtil()
                        linhamobile = 7
                    
                        Open Mobile For Input As #1       'Abr  e Arquivo texto
                            Line Input #1, FileBuffer
                                Do While Not EOF(1)
                                    Line Input #1, FileBuffer
                                        If Trim(Mid(FileBuffer, 113, 50)) <> "MENSAGEM ENVIADA COM SUCESSO." Then
                                            ObjPlan2Excel.Range("A" & linhamobile) = Trim(Mid(FileBuffer, 8, 4))
                                            ObjPlan2Excel.Range("B" & linhamobile) = Trim(Mid(FileBuffer, 13, 12))
                                            ObjPlan2Excel.Range("C" & linhamobile) = Trim(Mid(FileBuffer, 26, 15))
                                            ObjPlan2Excel.Range("D" & linhamobile) = Trim(Mid(FileBuffer, 42, 40))
                                            ObjPlan2Excel.Range("E" & linhamobile) = Trim(Mid(FileBuffer, 113, 50))
                                            linhamobile = linhamobile + 1
                                        End If
                                Loop
                        Close #1
                              
                        ObjPlan2Excel.Select
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(7).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(8).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(9).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(10).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(11).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(1).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(2).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(3).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Borders(4).LineStyle = 2
                        ObjPlan2Excel.Range("A7" & ":E" & linhamobile).Font.Size = 8
                        ObjPlan2Excel.Columns("A:E").Select
                        ObjPlan2Excel.Columns.AutoFit
                        ObjPlan2Excel.Rows("7:" & linhamobile).RowHeight = 11.75
                        ObjPlan2Excel.Range("C2").Select
    
                        Data = Date
                        Nome = "Relatorio de Termos nao Enviados - " & Format(Date, "ddmmyy")
                        Nome = Trata_NomeArquivo(Nome)
                        
                        ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                        ObjExcel.activeworkbook.Close SaveChanges:=False
                        ObjExcel.Quit
            
                        File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                            
                        'Enviar Arquivo
                        Call EnviarEmailMobile("SUCESSO", File)
                        
                    Else
                        'Enviar email arquivo em Branco
                        Call EnviarEmailMobile("ARQUIVO EM BRANCO", "")
                    End If
            Else
                'Enviar email arquivo não encontado
                Call EnviarEmailMobile("NAO ENCONTRADO", "")
            End If
Fim:

End Sub
