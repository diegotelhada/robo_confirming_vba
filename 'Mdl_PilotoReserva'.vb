'Mdl_PilotoReserva'


Option Compare Database
Sub Relatorio_Reserva()

    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Nome As String, DtInicio As Date
    Dim DtFim As Date, Ret As String, DataPesq As Date
    Dim TbDados As Recordset, TbDados2 As Recordset, TbData As Recordset
    
        Call AbrirBDRelatorios
    
            SiglaPesq = PesqUsername
        
                Set TbDados2 = BDREL.OpenRecordset("TblUsuarios", dbOpenDynaset)
                
                    TbDados2.FindFirst "Sigla like '*" & SiglaPesq & "'"
                    
                        If TbDados2.NoMatch = False Then
                            NomeUsuario = TbDados2!Nome
                            EmailUsuario = TbDados2!Email
                          TbDados2.Close
                        End If

                    Data1 = Date + 1: Data2 = Date + 2: Data3 = Date + 3: Data4 = Date + 4: Data5 = Date + 5: Data6 = Date + 6: Data7 = Date + 7
                    DataVenc1 = Format(Data1, "MM/DD/YYYY"): DataVenc2 = Format(Data2, "MM/DD/YYYY"): DataVenc3 = Format(Data3, "MM/DD/YYYY"): DataVenc4 = Format(Data4, "MM/DD/YYYY"): DataVenc5 = Format(Data5, "MM/DD/YYYY"): DataVenc6 = Format(Data6, "MM/DD/YYYY"): DataVenc7 = Format(Data7, "MM/DD/YYYY")

                Set TbDados = BDREL.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc1 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC1, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc2 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC2, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc3 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC3, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc4 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC4, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc5 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC5, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc6 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC6, Sum(IIf([TblArqoped]![Data_Venc]=#" & DataVenc7 & "#,[TblArqoped]![Valor_Pagmto])) AS VENC7" _
                & " FROM TblArqoped WHERE (((TblArqoped.Data_Venc) > Date() And (TblArqoped.Data_Venc) <= Date() + 7))" _
                & " GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Cnpj_Ancora, TblArqoped.Agencia_Remet, TblArqoped.Conta_Remet;", dbOpenDynaset)

                    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"
                    
                       Set ObjExcel = CreateObject("EXCEL.application")
                       ObjExcel.Workbooks.Open FileName:=Caminho & "Piloto_de_Reservas.xlsx", ReadOnly:=True
                       Set ObjPlan1Excel = ObjExcel.Worksheets("Relatório")
                       
                        dataPlan = Format(Date, "MM/DD/YYYY"): linha = 8
                    
                       ObjPlan1Excel.Range("J5") = dataPlan: ObjPlan1Excel.Range("E7") = DataVenc1: ObjPlan1Excel.Range("F7") = DataVenc2: ObjPlan1Excel.Range("G7") = DataVenc3: ObjPlan1Excel.Range("H7") = DataVenc4: ObjPlan1Excel.Range("I7") = DataVenc5: ObjPlan1Excel.Range("J7") = DataVenc6: ObjPlan1Excel.Range("K7") = DataVenc7
                       ObjPlan1Excel.Range("E7").NumberFormat = "dd/mm/yyyy": ObjPlan1Excel.Range("F7").NumberFormat = "dd/mm/yyyy": ObjPlan1Excel.Range("G7").NumberFormat = "dd/mm/yyyy": ObjPlan1Excel.Range("H7").NumberFormat = "dd/mm/yyyy": ObjPlan1Excel.Range("I7").NumberFormat = "dd/mm/yyyy": ObjPlan1Excel.Range("J7").NumberFormat = "dd/mm/yyyy": ObjPlan1Excel.Range("K7").NumberFormat = "dd/mm/yyyy"
                       
                       ObjPlan1Excel.Range("A8").CopyFromRecordset TbDados
                       
                       UltimaLinha = TbDados.RecordCount: UltimaLinha = UltimaLinha + 8: LinhaFormula = UltimaLinha
                       
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(7).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(8).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(9).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(10).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(11).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(1).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(2).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(3).LineStyle = 2
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Borders(4).LineStyle = 2
                      ObjPlan1Excel.Range("B" & linha & ":B" & UltimaLinha).NumberFormat = "00000"
                      ObjPlan1Excel.Range("C" & linha & ":D" & UltimaLinha).NumberFormat = "00000"
                      ObjPlan1Excel.Range("E" & linha & ":K" & UltimaLinha).Style = "Currency"
                      ObjPlan1Excel.Range("A" & linha & ":K" & UltimaLinha).Font.Size = 8
                      ObjPlan1Excel.Range("E" & LinhaFormula) = "=SUM(E8:E" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("E" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("E" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Range("F" & LinhaFormula) = "=SUM(F8:F" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("F" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("F" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Range("G" & LinhaFormula) = "=SUM(G8:G" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("G" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("G" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Range("H" & LinhaFormula) = "=SUM(H8:H" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("H" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("H" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Range("I" & LinhaFormula) = "=SUM(I8:I" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("I" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("I" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Range("J" & LinhaFormula) = "=SUM(J8:J" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("J" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("J" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Range("K" & LinhaFormula) = "=SUM(K8:K" & UltimaLinha - 1 & ")"
                      ObjPlan1Excel.Range("K" & LinhaFormula).Style = "Currency"
                      ObjPlan1Excel.Range("K" & LinhaFormula).Font.Bold = True
                      ObjPlan1Excel.Columns("A:K").Select
                      ObjPlan1Excel.Columns.AutoFit
                      ObjPlan1Excel.Rows("9:" & UltimaLinha).RowHeight = 11.75
                      ObjPlan1Excel.Range("C2").Select
                      
                      Data = Date
                    
                       Nome = "1Relatorio Piloto de Reserva - " & Format(Data, "ddmmyy")
                          
                       ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                       ObjExcel.activeworkbook.Close SaveChanges:=False
                       ObjExcel.Quit
                    
                    TbDados.Close
                    
                    'Enviar Arquivo
                    File = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
                    
                    'EmailDestino = "marcelohsouza@santander.com.br"
                    EmailDestino = "PilotodeReservas@santander.com.br"
                    EmailCopia = "wellington.da.silva@santander.com.br;edcampos@santander.com.br;geraldo.ghetti@santander.com.br;rcrodrigues@santander.com.br;lfrossi@santander.com.br"
                    
                    Assinatura = "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                    
                    Set sbObj = New Scripting.FileSystemObject
                           
                    Set olapp = CreateObject("Outlook.Application")
                    Set oitem = olapp.CreateItem(0)
                    
                            oitem.Subject = ("RELATORIO PILOTO DE RESERVA -  " & Format(Data, "dd/mm/yyyy"))
                            
                            oitem.SentOnBehalfOfName = "Processamento PJ - Relatorios Confirming"
                            
                            oitem.To = EmailDestino
                            oitem.cc = EmailCopia
                    
                    Corpo1 = "Prezados,"
                    Corpo2 = "Segue anexo o "
                    Relatorio = "RELATORIO PILOTO DE RESERVA"
                    Corpo21 = " referente aos vencimentos agendados para liquidar na próxima semana de Confirming® Santander."
                    Corpo3 = "Em caso de dúvida entrar em contato com a "
                    Confirming = "Confirming Desk"
                    Corpo31 = " ou através do(s) tel(s):"
                    Fones = " 0800-725-8090 / (11)3012-6390."
                    Corpo4 = "Atenciosamente,"
                    Assinatura1 = "Confirming®"
                    Assinatura2 = "Global Transaction Banking"
                    Assinatura3 = "Av. Juscelino Kubitschek, 2.235"
                    Assinatura4 = "Meios - Operações e Serviços"
                    Assinatura5 = "CEP: 04543-011  São Paulo-SP"
                    Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
                    Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
                    Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
                    
                                           
                            oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Calibri <BR>" & Corpo1 & "<BR/>" & _
                        "<BR>" & Corpo2 & "<B>" & Relatorio & "</B>" & Corpo21 & "<BR>" & "<BR>" & Corpo3 & "<B>" & Confirming & "</B>" & Corpo31 & "<B>" & Fones & "</B>" & "<BR>" & "<BR>" & Corpo4 & "<BR><BR><BR>" & _
                        " <img src=" & Assinatura & " height=50 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 3" & "<BR>" & _
                        "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 2 <BR>" & Assinatura3 & _
                        "<BR/>" & Assinatura4 & "<BR/>" & Assinatura5 & "<BR/></FONT><FONT COLOR = BLACK FACE = Calibri Size = 1 <BR><I>" & Assinatura6 & _
                        "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"
                            
                            oitem.Attachments.Add File
                            oitem.Attachments.Add "\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\LOGO_GBM.jpg"
                           
                            oitem.Send
                            
                            'oitem.DISPLAY True
                            Set olapp = Nothing
                            Set oitem = Nothing
                                                        
                        Set TbDados = BDREL.OpenRecordset("TblRelatoriosEnviados", dbOpenDynaset)
                    
                TbDados.AddNew
                    TbDados!Relatorio_Enviado = "RELATORIO PILOTO DE RESERVA"
                    TbDados!Periodicidade_Relatorio = "Semanal"
                    TbDados!Data_Envio = Date
                    TbDados!Hora_Envio = Format(Time, "HH:MM:SS")
                    TbDados!Usuario = NomeUsuario
                TbDados.Update
            TbDados.Close



End Sub

