'Mdl_VencidosYD42'

Option Compare Database
Dim Telas1(24) As String
Dim Str_linha As String
Dim Telas2(24) As String
Dim TelaConcat(24) As String
Public faz As Boolean
Dim number_of_connections, connection_choice, input_choice
Dim N As Long
Public autECLConnmgr As Object
Public autECLOIA As Object
Public autECLConnList As Object
Public autECLsession As Object
Public autECLPS As Object
Public Usuario As String
Public Senha As String
Function display_info(index)

   If (index > 0 And index <= number_of_connections) Then
      Call display_ConnList(autECLConnmgr.autECLConnList(index))
   Else
      MsgBox ("Errro de Conexão" & connection_choice)
   End If
      
End Function
Function display_ConnList(dis_object)
      
    ThisSessionName = dis_object.Name
    autECLsession.SetConnectionByName (ThisSessionName)

End Function
Sub GerarYD42()
    
    Dim Conta As String, linhaRel As Double
    Dim linha As Double, nome_convenio As String
    Dim LinhaConsut As Double, LinhaFinal As Double: index = 1
        
        'Cria conexão com o PCOM
        Set autECLConnmgr = CreateObject("PCOMM.autECLConnMgr")
        Set autECLConnList = CreateObject("PCOMM.autECLConnList")
        Set autECLsession = CreateObject("PCOMM.autECLSession")
        Set autECLPS = CreateObject("PCOMM.autECLPS")
        Set autECLOIA = CreateObject("PCOMM.autECLOIA")
        
        autECLConnmgr.autECLConnList.Refresh
        number_of_connections = autECLConnmgr.autECLConnList.Count
        input_choice = "1"
        connection_choice = CInt(input_choice)
        display_info (index)

        'Data de Referencia
        DataRef = Format(Date, "ddmmyyyy")
        Arquivo = "C:\temp\YDW0042S_" & DataRef & ".TXT"

        'Abre o arquivo de texto
        Open Arquivo For Output As #1

            LinhaFinal = autECLsession.autECLPS.GetText(2, 57, 5)
            linha = 6: linhaRel = 0: Str_linha = Empty
            
            Do While True

                Do While True
                    If autECLsession.autECLPS.GetText(3, 35, 18) = "S - 001   E -> 080" Then
                        CopTelas1
                        Exit Do
                    Else
                        autECLsession.autECLPS.SendKeys "[Pf10]"
                        autECLsession.autECLPS.SendKeys "[Pf10]"
                    End If
                Loop
                        
                autECLsession.autECLPS.SendKeys "[Pf10]"
                autECLsession.autECLPS.SendKeys "[Pf10]"
                autECLsession.autECLPS.SendKeys "[Pf11]"

                Do While True
                    If autECLsession.autECLPS.GetText(3, 35, 18) = "S - 081   E -> 160" Then: Exit Do
                Loop
                
                'Função para copiar segunda tela
                Call CopTelas2
                
                'Função para concatenar as duas telas
                Call concatena
                
                'Grava linha no arquivo
                Print #1, Str_linha
                linhaRel = linhaRel + 20
                
                autECLsession.autECLPS.SendKeys "[Pf8]"
                
                'Valida final da pagina
                Do While True
                    LinhaConsut = autECLsession.autECLPS.GetText(3, 58, 10)
                    If autECLsession.autECLPS.GetText(3, 58, 10) = Format(linhaRel, "0000000000") Then: Exit Do
                    cont = cont + 1
                        Do While True
                            For i = 1 To 24
                                If autECLsession.autECLPS.GetText(i, 30, 18) = "FINAL DO RELATORIO" Then: GoTo Fim
                            Next
                            Exit Do
                        Loop
                    linha = 5
                Loop
            Loop
Fim:
        Close #1
End Sub
Sub CopTelas1()

    Dim linhaTela As Integer: index = 1

        Set autECLConnmgr = CreateObject("PCOMM.autECLConnMgr")
        Set autECLConnList = CreateObject("PCOMM.autECLConnList")
        Set autECLsession = CreateObject("PCOMM.autECLSession")
        Set autECLPS = CreateObject("PCOMM.autECLPS")
        Set autECLOIA = CreateObject("PCOMM.autECLOIA")
            
            'Cria Conexao com o PCOM
            autECLConnmgr.autECLConnList.Refresh
            number_of_connections = autECLConnmgr.autECLConnList.Count
            input_choice = "1"
            connection_choice = CInt(input_choice)
            display_info (index)

        For linhaTela = 5 To 24
            Telas1(linhaTela) = Empty
        Next
               
        For linhaTela = 5 To 24
            Telas1(linhaTela) = autECLsession.autECLPS.GetText(linhaTela, 1, 80)
        Next
End Sub
Sub CopTelas2()

    Dim linhaTela As Integer: index = 1

        Set autECLConnmgr = CreateObject("PCOMM.autECLConnMgr")
        Set autECLConnList = CreateObject("PCOMM.autECLConnList")
        Set autECLsession = CreateObject("PCOMM.autECLSession")
        Set autECLPS = CreateObject("PCOMM.autECLPS")
        Set autECLOIA = CreateObject("PCOMM.autECLOIA")
            
            'Cria conexão com o PCOM
            autECLConnmgr.autECLConnList.Refresh
            number_of_connections = autECLConnmgr.autECLConnList.Count
            input_choice = "1"
            connection_choice = CInt(input_choice)
            display_info (index)

        For linhaTela = 5 To 24
            Telas2(linhaTela) = Empty
        Next

        For linhaTela = 5 To 24
            Telas2(linhaTela) = autECLsession.autECLPS.GetText(linhaTela, 1, 80)
        Next
End Sub
Sub concatena()

    Str_linha = Empty

        For linhaTela = 5 To 24
            Str_linha = Str_linha & Chr(13) & Chr(10) & Telas1(linhaTela) & Telas2(linhaTela)
        Next
End Sub
Sub ImportarYD42()
    
    Dim TblYD42 As Recordset
        
        'Abrir banco de dados
        Call AbrirBDVencidos
        
        'Data do dia
        DataRef = Format(Date, "ddmmyyyy")

        'Limpar Tabela
        BDVencidos.Execute ("Delete * from TblYD42;")
        
        'Caminho do Arquivo
        Caminho = "C:\temp\YDW0042S_" & DataRef & ".TXT"
        
        'Abre Tabela
        Set TblYD42 = BDVencidos.OpenRecordset("TblYD42")
        
        'Abre Arquivo texto
        Open Caminho For Input As #1
                
            Line Input #1, FileBuffer
        
                Do While Not EOF(1)
            
                    If Mid(FileBuffer, 9, 10) = "CONFIRMING" Then
                        
                        CodProd = Trim(Mid(FileBuffer, 1, 5))
                        Produto = Trim(Mid(FileBuffer, 9, 10))
                        dataOper = Trim(Mid(FileBuffer, 31, 10))
                        DataVecto = Trim(Mid(FileBuffer, 43, 10))
                        numOper = Trim(Mid(FileBuffer, 55, 15))
                        Titulo = Trim(Mid(FileBuffer, 72, 15))
                        valor = Trim(Mid(FileBuffer, 103, 16))
                        Dias = Trim(Mid(FileBuffer, 125, 6))
                            
                            'Pega na proxima linha o saldo devedor
                            Line Input #1, FileBuffer
                                
                                If InStr(FileBuffer, "VALOR P/ PGTO") <> 0 Then
                                    ValorAtualizado = Trim(Mid(FileBuffer, 108, 20))
                                End If
                                
                        TblYD42.AddNew
                            TblYD42!CodProd = CodProd
                            TblYD42!Produto = Produto
                            TblYD42!dataOper = dataOper
                            TblYD42!DataVecto = DataVecto
                            TblYD42!numOper = numOper
                            TblYD42!Titulo = Titulo
                            TblYD42!valor = valor
                            TblYD42!Dias = Dias
                            TblYD42!ValorAtualizado = ValorAtualizado
                        TblYD42.Update
                        
                    End If
                  Line Input #1, FileBuffer
                Loop
        Close #1
                
        'Atualiza informações complementares do ARQOPED
        BDVencidos.Execute ("UPDATE TblArqoped INNER JOIN TblYD42 ON TblArqoped.Cod_Oper = TblYD42.NumOper SET TblYD42.Agencia = [TblArqoped]![Agencia_Ancora], TblYD42.Convenio = [TblArqoped]![Convenio_Ancora], TblYD42.NomeAncora = [TblArqoped]![Nome_Ancora], TblYD42.CNPJAncora = [TblArqoped]![Cnpj_Ancora], TblYD42.Fornecedor = [TblArqoped]![Nome_Fornecedor], TblYD42.CNPJFornecedor = [TblArqoped]![Cnpj_Fornecedor];")
        
End Sub
Sub GerarRelatorioYD42()
    
    Dim FSO As New FileSystemObject
    Dim TbDados As Recordset: Call AbrirBDVencidos
    Dim ObjExcel As Object, ObjPlan1Excel As Object: linha = 2
    Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"

        Set TbDados = BDVencidos.OpenRecordset("SELECT TblYD42.Agencia, TblYD42.Convenio, TblYD42.NomeAncora, TblYD42.CNPJAncora, TblYD42.Fornecedor, TblYD42.CNPJFornecedor, TblYD42.NumOper, TblYD42.DataOper, TblYD42.DataVecto, TblYD42.Titulo, TblYD42.Valor, TblYD42.ValorAtualizado, TblYD42.Dias, IIf([TblYD42]![Liquidacao]='CONTA CORRENTE CONDICIONAL A SALDO','TEIMOSINHA',IIf([TblYD42]![Liquidacao]='BOLETO BANCARIO','BOLETO','EXCLUIR')) AS Liquidacao" _
        & " FROM TblYD42 GROUP BY TblYD42.Agencia, TblYD42.Convenio, TblYD42.NomeAncora, TblYD42.CNPJAncora, TblYD42.Fornecedor, TblYD42.CNPJFornecedor, TblYD42.NumOper, TblYD42.DataOper, TblYD42.DataVecto, TblYD42.Titulo, TblYD42.Valor, TblYD42.ValorAtualizado, TblYD42.Dias, IIf([TblYD42]![Liquidacao]='CONTA CORRENTE CONDICIONAL A SALDO','TEIMOSINHA',IIf([TblYD42]![Liquidacao]='BOLETO BANCARIO','BOLETO','EXCLUIR')) HAVING (((IIf([TblYD42]![Liquidacao]='CONTA CORRENTE CONDICIONAL A SALDO','TEIMOSINHA',IIf([TblYD42]![Liquidacao]='BOLETO BANCARIO','BOLETO','EXCLUIR')))<>'EXLCUIR'));", dbOpenDynaset)
    
            If TbDados.EOF = False Then
                TbDados.MoveLast: UltimaLinha = (TbDados.RecordCount + (linha - 1)): TbDados.MoveFirst
                
                Set ObjExcel = CreateObject("EXCEL.application")
                ObjExcel.Workbooks.Open FileName:=Caminho & "VencidosYD.xlsx", ReadOnly:=True
                Set ObjPlan1Excel = ObjExcel.Worksheets(1)
                ObjPlan1Excel.Range("A2").CopyFromRecordset TbDados
            
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(7).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(8).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(9).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(10).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(11).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(1).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(2).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(3).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Borders(4).LineStyle = 2
                    ObjPlan1Excel.Range("A" & linha & ":N" & UltimaLinha).Font.Size = 8
                    ObjPlan1Excel.Range("H" & linha & ":I" & UltimaLinha).NumberFormat = "dd/mm/yyyy"
                    ObjPlan1Excel.Range("K" & linha & ":L" & UltimaLinha).Style = "Currency"
                    ObjPlan1Excel.Range("F" & linha & ":F" & UltimaLinha).NumberFormat = "000000000000"
                    ObjPlan1Excel.Range("D" & linha & ":D" & UltimaLinha).NumberFormat = "000000000000"
                    ObjPlan1Excel.Columns("A:N").Select
                    ObjPlan1Excel.Columns.AutoFit
                    ObjPlan1Excel.Rows("2:" & UltimaLinha).RowHeight = 11.75
                    ObjPlan1Excel.Range("A8").Select
                               
                    Nome = "VencidosYD42_" & Format(Date, "ddmmyy")
                    sFname = "C:\Temp\" & Nome & ".xlsx"
                     
                     If (Dir(sFname) <> "") Then
                         Kill sFname
                     End If
                                            
                   ObjPlan1Excel.SaveAs FileName:="C:\Temp\" & Nome & ".xlsx"
                   ObjExcel.activeworkbook.Close SaveChanges:=False
                   ObjExcel.Quit
                   
                   If FSO.FileExists("C:\Temp\" & Nome & ".xlsx") Then: FSO.CopyFile "C:\Temp\" & Nome & ".xlsx", "\\bsbrsp3852\Carga\VencidosConfirming\" & Nome & ".xlsx"
            
               Call Shell("excel.exe C:\Temp\" & Nome & ".xlsx", 1)
            Else
                MsgBox "Não temos operações vencidas com base no relatorio de Hoje!", vbInformation, "Vencidos Confirming"
            End If
End Sub
Sub ImportarArquivoConvenios()

    Dim FSO As New FileSystemObject: UltDia = Format(UltimoDiaUtil(), "DDMMYY")
    Dim ObjExcel As Object, ObjExcelPlan1 As Object:: Call AbrirBDVencidos
    Dim Caminho As String: Caminho = "\\BSBRSP56\confirming relatorio\"

        ArquivoDia = Caminho & "CONVENIOS_" & UltDia & ".CSV"                       'Montar Arquivo Convenios do Dia
            If FSO.FileExists(ArquivoDia) Then                                      'valida se o arquivo esta disponivel
                FSO.CopyFile ArquivoDia, "C:\Temp\CONVENIOS.CSV"                    'Copiar arquivo do dia para a maquina
                Call Conveter_CSV4XLSX_Convenios                                    'Converter aquivo CSV para XLSX
                BDVencidos.Execute ("DELETE Tbl_Convenios.* FROM Tbl_Convenios;")   'Limpar tabela
                BDVencidos.Execute ("AdcTbl_Convenios")                              'Adiciona para tabela local
                BDVencidos.Execute ("UpdateLiquidacao")                             'Atualiza campo liquidacao
            End If

End Sub
