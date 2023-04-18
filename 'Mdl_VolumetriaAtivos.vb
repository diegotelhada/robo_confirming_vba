'Mdl_VolumetriaAtivos"

Option Compare Database
Function FormatarData(Data)
    
    If Not IsNull(Data) Then
        Total = Len(Data)
            For i = 1 To Total
                If Mid(Data, i, 1) = "/" Then
                    Mes = Mid(Data, 1, (i - 1))
                        For j = (i + 1) To Total
                            If Mid(Data, j, 1) = "/" Then
                                dia = Mid(Data, (i + 1), ((j - i) - 1))
                                    For Y = (j + 1) To Total
                                        If Mid(Data, Y, 1) = " " Then
                                            Ano = Mid(Data, (j + 1), ((Y - j) - 1))
                                                For X = (Y + 1) To Total
                                                    If Mid(Data, X, 1) = ":" Then
                                                        Hora = Mid(Data, (Y + 1), ((X - Y) - 1))
                                                            For W = (X + 1) To Total
                                                                If Mid(Data, W, 1) = ":" Then
                                                                    Minuto = Mid(Data, (X + 1), ((W - X) - 1))
                                                                        For l = (W + 1) To Total
                                                                            If Mid(Data, l, 1) = " " Then
                                                                                Segundo = Mid(Data, (W + 1), ((l - W) - 1))
                                                                                Exit For
                                                                            End If
                                                                        Next l
                                                                     Exit For
                                                                End If
                                                            Next W
                                                        Exit For
                                                    End If
                                                Next X
                                            Exit For
                                        End If
                                    Next Y
                                Exit For
                            End If
                        Next j
                    Exit For
                End If
            Next i
        FormatarData = CDate(Format(dia & "/" & Mes & "/" & Ano, "dd/mm/yyyy")) & " " & Format(Hora & ":" & Minuto & ":" & Segundo, "hh:mm:ss")
    Else
        FormatarData = ""
    End If
    
End Function
Sub AtualizaVolumetriaAtivos()
    
    Dim DbAtiv As Database
    Dim FSO As New FileSystemObject
        
        Set DbAtiv = OpenDatabase("\\bsbrsp54\scoativos\Scontrol\BaseDados\BdAtividadesG6.mdb")
        
        Debug.Print "Copiando Arquivo " & Format(Time, "hh:mm:ss"): Tempo = CDate(Format(Time, "hh:mm:ss"))
        
        File = "\\NASBEDADOS2\Relatorio_Batch_G6\REL138_Relatório Detalhe Atividade - Osmar TXT_" & Format(Date, "YYYYMMDD") & ".txt"
                
            If FSO.FileExists(File) = True Then
                
                'Copiar Arquivo Detalhe de Atividade
                FSO.CopyFile File, "C:\Temp\REL138.txt", True
                Debug.Print "Arquivo Copiado " & Format(CDate(Format(Time, "hh:mm:ss")) - CDate(Format(Tempo, "hh:mm:ss")), "hh:mm:ss"): Tempo = CDate(Format(Time, "hh:mm:ss"))
                    
                'Deletar tabela relatorio
                DbAtiv.Execute ("DELETE REL138_Temp.* FROM REL138_Temp;")
                
                Debug.Print "Tabela REL138 Deletada " & Format(CDate(Format(Time, "hh:mm:ss")) - CDate(Format(Tempo, "hh:mm:ss")), "hh:mm:ss"): Tempo = CDate(Format(Time, "hh:mm:ss"))
                
                'Inserir dados na tabela temporaria rel38
                'DbAtiv.Execute ("INSERT INTO REL138_Temp ( Operação, [Deal Nº], Origem, [Status Operação], Evento, PE, CNPJ, Cliente, [Unidade Negócio], [Segmento Primário], [Segmento Secundário], Produto, [Sub-Produto], [Data Entrada G6], [Data de Boletagem], [Data de Início], [Data de Vencimento], Moeda, Valor, [Núm Contrato], [Há Garantias], [Level], [Descrição da Fase], [Data/Hora Prim Exec], [Data/Hora Execução], [Status da Fase], Perfil, Usuário, [Motivo Cancelamento], [Dt Canc], [Login Canc] )" _
                & " SELECT REL138.Operação, REL138.[Deal Nº], REL138.Origem, REL138.[Status Operação], REL138.Evento, REL138.PE, REL138.CNPJ, REL138.Cliente, REL138.[Unidade Negócio], REL138.[Segmento Primário], REL138.[Segmento Secundário], REL138.Produto, REL138.[Sub-Produto], Format([REL138]![Data Entrada G6],'dd/mm/yyyy') & ' ' & Format([REL138]![Data Entrada G6],'hh:nn:ss') AS [Data Entrada G6], Format([REL138]![Data de Boletagem],'dd/mm/yyyy') AS [Data de Boletagem], Format([REL138]![Data de Início],'dd/mm/yyyy') AS [Data de Início], Format([REL138]![Data de Vencimento],'dd/mm/yyyy') AS [Data de Vencimento], REL138.Moeda, REL138.Valor, REL138.[Núm Contrato], REL138.[Há Garantias], REL138.Level, REL138.[Descrição da Fase]," _
                & " Format([REL138]![Data/Hora Prim Exec],'dd/mm/yyyy') & ' ' & Format([REL138]![Data/Hora Prim Exec],'hh:nn:ss') AS [Data/Hora Prim Exec], Format([REL138]![Data/Hora Execução],'dd/mm/yyyy') & ' ' & Format([REL138]![Data/Hora Execução],'hh:nn:ss') AS [Data/Hora Execução], REL138.[Status da Fase], REL138.Perfil , REL138.Usuário, REL138.[Motivo Cancelamento], REL138.[Dt Canc], REL138.[Login Canc] FROM REL138" _
                & " GROUP BY REL138.Operação, REL138.[Deal Nº], REL138.Origem, REL138.[Status Operação], REL138.Evento, REL138.PE, REL138.CNPJ, REL138.Cliente, REL138.[Unidade Negócio], REL138.[Segmento Primário], REL138.[Segmento Secundário], REL138.Produto, REL138.[Sub-Produto], Format([REL138]![Data Entrada G6],'dd/mm/yyyy') & ' ' & Format([REL138]![Data Entrada G6],'hh:nn:ss'), Format([REL138]![Data de Boletagem],'dd/mm/yyyy'), Format([REL138]![Data de Início],'dd/mm/yyyy'), Format([REL138]![Data de Vencimento],'dd/mm/yyyy'), REL138.Moeda, REL138.Valor, REL138.[Núm Contrato], REL138.[Há Garantias], REL138.Level, REL138.[Descrição da Fase], Format([REL138]![Data/Hora Prim Exec],'dd/mm/yyyy') & ' ' & Format([REL138]![Data/Hora Prim Exec],'hh:nn:ss'), Format([REL138]![Data/Hora Execução],'dd/mm/yyyy') & ' ' & Format([REL138]![Data/Hora Execução],'hh:nn:ss'), REL138.[Status da Fase], REL138.Perfil, REL138.Usuário, REL138.[Motivo Cancelamento], REL138.[Dt Canc], REL138.[Login Canc]" _
                & " HAVING (((REL138.Operação) Not Like '*&nbsp;*') AND ((REL138.[Status da Fase]) Not Like 'Não Liberada'));")
                
                DbAtiv.Execute ("INSERT INTO REL138_Temp ( Operação, [Deal Nº], Origem, [Status Operação], Evento, PE, CNPJ, Cliente, [Unidade Negócio], [Segmento Primário], [Segmento Secundário], Produto, [Sub-Produto], [Data Entrada G6], [Data de Boletagem], [Data de Início], [Data de Vencimento], Moeda, Valor, [Núm Contrato], [Há Garantias], [Level], [Descrição da Fase], [Data/Hora Prim Exec], [Data/Hora Execução], [Status da Fase], Perfil, Usuário, [Motivo Cancelamento], [Dt Canc], [Login Canc] )" _
                & " SELECT REL138.Operação, REL138.[Deal Nº], REL138.Origem, REL138.[Status Operação], REL138.Evento, REL138.PE, REL138.CNPJ, REL138.Cliente, REL138.[Unidade Negócio], REL138.[Segmento Primário], REL138.[Segmento Secundário], REL138.Produto, REL138.[Sub-Produto], FormatarData([REL138]![Data Entrada G6]) AS [Data Entrada G6], Format([REL138]![Data de Boletagem],'dd/mm/yyyy') AS [Data de Boletagem], Format([REL138]![Data de Início],'dd/mm/yyyy') AS [Data de Início]," _
                & " Format([REL138]![Data de Vencimento],'dd/mm/yyyy') AS [Data de Vencimento], REL138.Moeda, REL138.Valor, REL138.[Núm Contrato], REL138.[Há Garantias], REL138.Level, REL138.[Descrição da Fase], FormatarData([REL138]![Data/Hora Prim Exec]) AS [Data/Hora Prim Exec], FormatarData([REL138]![Data/Hora Execução]) AS [Data/Hora Execução], REL138.[Status da Fase], REL138.Perfil, REL138.Usuário, REL138.[Motivo Cancelamento], REL138.[Dt Canc], REL138.[Login Canc] FROM REL138 " _
                & " GROUP BY REL138.Operação, REL138.[Deal Nº], REL138.Origem, REL138.[Status Operação], REL138.Evento, REL138.PE, REL138.CNPJ, REL138.Cliente, REL138.[Unidade Negócio], REL138.[Segmento Primário], REL138.[Segmento Secundário], REL138.Produto, REL138.[Sub-Produto], FormatarData([REL138]![Data Entrada G6]), Format([REL138]![Data de Boletagem],'dd/mm/yyyy'), Format([REL138]![Data de Início],'dd/mm/yyyy'), Format([REL138]![Data de Vencimento],'dd/mm/yyyy'), REL138.Moeda, REL138.Valor, REL138.[Núm Contrato], REL138.[Há Garantias], REL138.Level, REL138.[Descrição da Fase], FormatarData([REL138]![Data/Hora Prim Exec]), FormatarData([REL138]![Data/Hora Execução]), REL138.[Status da Fase], REL138.Perfil, REL138.Usuário, REL138.[Motivo Cancelamento], REL138.[Dt Canc], REL138.[Login Canc]" _
                & " HAVING (((REL138.Operação) Not Like '*&nbsp;*') AND ((REL138.[Status da Fase]) Not Like 'Não Liberada'));")

                
                Debug.Print "Tabela REL138 Atualizada " & Format(CDate(Format(Time, "hh:mm:ss")) - CDate(Format(Tempo, "hh:mm:ss")), "hh:mm:ss"): Tempo = CDate(Format(Time, "hh:mm:ss"))
                    
                'Deletar Linhas de comentarios
                 DbAtiv.Execute ("DELETE IIf(IsNumeric([REL138_TEMP]![Operação]),'SIM','NAO') AS Valida FROM REL138_TEMP WHERE (((IIf(IsNumeric([REL138_TEMP]![Operação]),'SIM','NAO'))='NAO'));")
                
                Debug.Print "Tabela REL138 Formatada " & Format(CDate(Format(Time, "hh:mm:ss")) - CDate(Format(Tempo, "hh:mm:ss")), "hh:mm:ss"): Tempo = CDate(Format(Time, "hh:mm:ss"))
                    
                'Inserir Dados na tabela atividade G6
                DbAtiv.Execute ("INSERT INTO TblAtividadeG6 ( Operação, [Deal Nº], Origem, [Status Workflow], Evento, CNPJ, Cliente, [Segmento Primário], [Segmento Secundário], Produto, [Sub-Produto], [Data de Entrada G6], [Data de Boletagem], [Data de Início], [Data de Vencimento], Moeda, Valor, [Num Contrato], [Existe Garantia], [Level], Descrição, [Data/Hora Primeira Execução], [Data/Hora], [Status da atividade], Perfil, Usuário )" _
                & " SELECT REL138_Temp.Operação, REL138_Temp.[Deal Nº], REL138_Temp.Origem, REL138_Temp.[Status Operação], REL138_Temp.Evento, REL138_Temp.CNPJ, REL138_Temp.Cliente, REL138_Temp.[Segmento Primário], REL138_Temp.[Segmento Secundário], REL138_Temp.Produto, REL138_Temp.[Sub-Produto], REL138_Temp.[Data Entrada G6], REL138_Temp.[Data de Boletagem], REL138_Temp.[Data de Início], REL138_Temp.[Data de Vencimento], REL138_Temp.Moeda, REL138_Temp.Valor, REL138_Temp.[Núm Contrato], REL138_Temp.[Há Garantias], REL138_Temp.Level, REL138_Temp.[Descrição da Fase], REL138_Temp.[Data/Hora Prim Exec], REL138_Temp.[Data/Hora Execução], REL138_Temp.[Status da Fase], REL138_Temp.Perfil, REL138_Temp.Usuário FROM REL138_Temp;")
                
                Debug.Print "Tabela Atividade Atualizada " & Format(CDate(Format(Time, "hh:mm:ss")) - CDate(Format(Tempo, "hh:mm:ss")), "hh:mm:ss"): Tempo = CDate(Format(Time, "hh:mm:ss"))
    
            End If


End Sub
Sub GerarRelatorioProdutividadeAtivos()

    Dim TbGeral As Recordset, TbDados As Recordset
    Dim ObjExcel As Object, Total As Integer
    Dim ObjPlan1Excel As Object, ObjPlan2Excel As Object, ObjPlan3Excel As Object
    Dim DbAtiv As Database, FSO As New FileSystemObject
        
        Set DbAtiv = OpenDatabase("\\bsbrsp54\scoativos\Scontrol\BaseDados\BdAtividadesG6.mdb")
        
        Caminho = "\\bsbrsp54\scoativos\scontrol\Mascara\PRODUTIVIDADE\"
        
        DataInicio = Format(UltimoDiaUtil(), "mm/dd/yyyy") & " 00:00:01"
        DataFinal = Format(UltimoDiaUtil(), "mm/dd/yyyy") & " 23:59:59"

        
        Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=Caminho & "VOLUMETRIA_G6_V3.xlsm", ReadOnly:=True
        Set ObjPlan1Excel = ObjExcel.Worksheets("Resumo")
        Set ObjPlan2Excel = ObjExcel.Worksheets("Consolidado")
        Set ObjPlan3Excel = ObjExcel.Worksheets("Analitico")
            ObjPlan1Excel.Select: LinhaInicio = 7
            'ObjExcel.Visible = True
                        
            'Operações Processamento
            Set TbGeral = DbAtiv.OpenRecordset("SELECT IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.Descrição) Like '*LIBERA*') AND ((TblAtividadeG6.Perfil) Like '*processamento*') AND ((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL'));", dbOpenDynaset)
                
                'Nome da primeira area
                ObjPlan1Excel.Range("D7") = "PROCESSAMENTO"
                linha = LinhaInicio + 2
                
                Do While TbGeral.EOF = False
                    
                    Total = Total + TbGeral!Qtde
                    
                    ObjPlan1Excel.Range("C" & linha) = "C"
                    ObjPlan1Excel.Range("D" & linha) = TbGeral!GERAL
                    ObjPlan1Excel.Range("G" & linha) = TbGeral!Qtde
                    
                    ObjPlan1Excel.Range("D" & linha & ":H" & linha).Font.Size = 10
                    ObjPlan1Excel.Range("D" & linha & ":H" & linha).Font.Bold = True
                    
                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).LineStyle = 1
                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).Color = RGB(211, 211, 211)
                    
                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).LineStyle = 1
                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).Color = RGB(211, 211, 211)
                    
                    linha = linha + 1
                        
                        Set TbDados = DbAtiv.OpenRecordset("SELECT IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, TblAtividadeG6.Produto, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.Descrição) Like '*LIBERA*') AND ((TblAtividadeG6.Perfil) Like '*processamento*') AND ((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')), TblAtividadeG6.Produto HAVING (((IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')))='" & TbGeral!GERAL & "'));", dbOpenDynaset)

                            Do While TbDados.EOF = False
                                
                                    ObjPlan1Excel.Range("D" & linha) = TbDados!Produto
                                    ObjPlan1Excel.Range("D" & linha).Font.Color = RGB(105, 105, 105)
                                    
                                    ObjPlan1Excel.Range("G" & linha) = TbDados!Qtde
                                    ObjPlan1Excel.Range("G" & linha).Font.Color = RGB(105, 105, 105)
                                                                        
                                    ObjPlan1Excel.Range("D" & linha & ":X" & linha).Font.Size = 9
                                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).LineStyle = 3
                                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).LineStyle = 3
                                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).Color = RGB(211, 211, 211)
                                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).Color = RGB(211, 211, 211)
                                    
                                    
                                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Font.Color = &H80000011
                                    
                                    ObjPlan1Excel.Rows(linha & ":" & linha).EntireRow.Hidden = True
                                    ObjPlan1Excel.Columns("D:H").Select
                                    ObjPlan1Excel.Columns.AutoFit
                                    
                                    linha = linha + 1
                                                                
                                TbDados.MoveNext
                            Loop
                            
                        ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).LineStyle = 1
                        ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).Color = RGB(211, 211, 211)
                        
                        ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).LineStyle = 1
                        ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).Color = RGB(211, 211, 211)
                            
                    TbGeral.MoveNext
                Loop
                
            ObjPlan1Excel.Columns("C:C").EntireColumn.Hidden = True
            ObjPlan1Excel.Range("G7") = Total
            ObjPlan1Excel.Range("C" & linha) = "0"

            LinhaInicio = linha + 3
            Total = 0
                        
            'Operações Processamento
            Set TbGeral = DbAtiv.OpenRecordset("SELECT IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.Perfil) Like '*formalização*') AND ((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL'));", dbOpenDynaset)
                
                ObjPlan1Excel.Range("D" & LinhaInicio) = "FORMALIZAÇÃO"
                ObjPlan1Excel.Range("D" & LinhaInicio & ":H" & LinhaInicio).Font.Size = 11
                ObjPlan1Excel.Range("D" & LinhaInicio & ":H" & LinhaInicio).Font.Bold = True
                ObjPlan1Excel.Range("D" & LinhaInicio & ":E" & LinhaInicio).Borders(3).LineStyle = 1
                ObjPlan1Excel.Range("D" & LinhaInicio & ":E" & LinhaInicio).Borders(3).Color = RGB(128, 128, 128)
                ObjPlan1Excel.Range("D" & LinhaInicio + 1 & ":E" & LinhaInicio + 1).Borders(3).LineStyle = 1
                ObjPlan1Excel.Range("D" & LinhaInicio + 1 & ":E" & LinhaInicio + 1).Borders(3).Color = RGB(128, 128, 128)
                
                ObjPlan1Excel.Range("G" & LinhaInicio & ":H" & LinhaInicio).Borders(3).LineStyle = 1
                ObjPlan1Excel.Range("G" & LinhaInicio & ":H" & LinhaInicio).Borders(3).Color = RGB(128, 128, 128)
                ObjPlan1Excel.Range("G" & LinhaInicio + 1 & ":H" & LinhaInicio + 1).Borders(3).LineStyle = 1
                ObjPlan1Excel.Range("G" & LinhaInicio + 1 & ":H" & LinhaInicio + 1).Borders(3).Color = RGB(128, 128, 128)
                
                
                linha = LinhaInicio + 2
                
                Do While TbGeral.EOF = False
                    
                    Total = Total + TbGeral!Qtde
                    
                    ObjPlan1Excel.Range("C" & linha) = "C"
                    ObjPlan1Excel.Range("D" & linha) = TbGeral!GERAL
                    ObjPlan1Excel.Range("G" & linha) = TbGeral!Qtde
                    ObjPlan1Excel.Range("D" & linha & ":H" & linha).Font.Size = 10
                    ObjPlan1Excel.Range("D" & linha & ":H" & linha).Font.Bold = True
                    
                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).LineStyle = 1
                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).Color = RGB(211, 211, 211)
                    
                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).LineStyle = 1
                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).Color = RGB(211, 211, 211)
                    
                    linha = linha + 1
                        
                        Set TbDados = DbAtiv.OpenRecordset("SELECT IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, TblAtividadeG6.Produto, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.Perfil) Like '*formalização*') AND ((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')), TblAtividadeG6.Produto HAVING (((IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')))='" & TbGeral!GERAL & "'));", dbOpenDynaset)

                            Do While TbDados.EOF = False
                                
                                    ObjPlan1Excel.Range("D" & linha) = TbDados!Produto
                                    ObjPlan1Excel.Range("D" & linha).Font.Color = RGB(105, 105, 105)
                                    
                                    ObjPlan1Excel.Range("G" & linha) = TbDados!Qtde
                                    ObjPlan1Excel.Range("G" & linha).Font.Color = RGB(105, 105, 105)
                                    
                                    
                                    ObjPlan1Excel.Range("D" & linha & ":X" & linha).Font.Size = 9
                                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).LineStyle = 3
                                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).LineStyle = 3
                                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).Color = RGB(211, 211, 211)
                                    ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).Color = RGB(211, 211, 211)
                                    
                                    ObjPlan1Excel.Range("G" & linha & ":H" & linha).Font.Color = &H80000011
                                    
                                    ObjPlan1Excel.Rows(linha & ":" & linha).EntireRow.Hidden = True
                                    ObjPlan1Excel.Columns("D:H").Select
                                    ObjPlan1Excel.Columns.AutoFit
                                    
                                    linha = linha + 1
                                                                
                                TbDados.MoveNext
                            Loop
                            
                        ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).LineStyle = 1
                        ObjPlan1Excel.Range("D" & linha & ":E" & linha).Borders(3).Color = RGB(211, 211, 211)
                        
                        ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).LineStyle = 1
                        ObjPlan1Excel.Range("G" & linha & ":H" & linha).Borders(3).Color = RGB(211, 211, 211)
                            
                    TbGeral.MoveNext
                Loop
                
            ObjPlan1Excel.Columns("C:C").EntireColumn.Hidden = True
            ObjPlan1Excel.Range("G" & LinhaInicio) = Total
            ObjPlan1Excel.Range("C" & linha) = "0"
            
            'Sheet Consolidado
            ObjPlan2Excel.Select: linhaPlan = 2
            Set TbDados = DbAtiv.OpenRecordset("SELECT TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, TblAtividadeG6.Produto, TblAtividadeG6.Descrição, TblAtividadeG6.Evento, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')), TblAtividadeG6.Produto, TblAtividadeG6.Descrição, TblAtividadeG6.Evento HAVING (((TblAtividadeG6.Perfil) Like '*processamento*') AND ((TblAtividadeG6.Descrição) Like '*libera*'));", dbOpenDynaset)

                If TbDados.EOF = False Then
                    TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                    ObjPlan2Excel.Range("A" & linhaPlan).CopyFromRecordset TbDados
                    linhaPlan = linhaPlan + Qntd
                End If
            
            
            Set TbDados = DbAtiv.OpenRecordset("SELECT TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, TblAtividadeG6.Produto, TblAtividadeG6.Descrição, TblAtividadeG6.Evento, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')), TblAtividadeG6.Produto, TblAtividadeG6.Descrição, TblAtividadeG6.Evento HAVING (((TblAtividadeG6.Perfil) Like '*formalização*'));", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                    TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                    ObjPlan2Excel.Range("A" & linhaPlan).CopyFromRecordset TbDados
                    linhaPlan = linhaPlan + Qntd
                End If
            
            ObjPlan2Excel.Range("A2:F" & linhaPlan).Borders(3).LineStyle = 1
            ObjPlan2Excel.Range("A2:F" & linhaPlan).Borders(3).Color = RGB(211, 211, 211)
            ObjPlan2Excel.Columns("A:F").Select
            ObjPlan2Excel.Columns.AutoFit
            ObjPlan1Excel.Select

            'Sheet Analitico
            ObjPlan3Excel.Select: linhaPlan = 2
            Set TbDados = DbAtiv.OpenRecordset("SELECT TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, TblAtividadeG6.Produto, TblAtividadeG6.[Segmento Primário], TblAtividadeG6.Operação, TblAtividadeG6.Cliente, TblAtividadeG6.Descrição, TblAtividadeG6.Usuário, TblAtividadeG6.Evento, Sum(1) AS Qtde FROM TblAtividadeG6 WHERE (((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado'))" _
            & " GROUP BY TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')), TblAtividadeG6.Produto, TblAtividadeG6.[Segmento Primário], TblAtividadeG6.Operação, TblAtividadeG6.Cliente, TblAtividadeG6.Descrição, TblAtividadeG6.Usuário, TblAtividadeG6.Evento HAVING (((TblAtividadeG6.Perfil) Like '*processamento*') AND ((TblAtividadeG6.Descrição) Like '*libera*'));", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                    TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                    ObjPlan3Excel.Range("A" & linhaPlan).CopyFromRecordset TbDados
                    linhaPlan = linhaPlan + Qntd
                End If
            
            Set TbDados = DbAtiv.OpenRecordset("SELECT TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')) AS GERAL, TblAtividadeG6.Produto, TblAtividadeG6.[Segmento Primário], TblAtividadeG6.Operação, TblAtividadeG6.Cliente, TblAtividadeG6.Descrição, TblAtividadeG6.Usuário, TblAtividadeG6.Evento, Sum(1) AS Qtde FROM TblAtividadeG6" _
            & " WHERE (((TblAtividadeG6.[Data/Hora])>=#" & DataInicio & "# And (TblAtividadeG6.[Data/Hora])<=#" & DataFinal & "#) AND ((TblAtividadeG6.[Status da atividade])='Aprovado')) GROUP BY TblAtividadeG6.Perfil, IIf([TblAtividadeG6]![Produto] Like '*Fian*','FIANÇA',IIf([TblAtividadeG6]![Produto] Like '*LEASING*','LEASING','CARTEIRA GERAL')), TblAtividadeG6.Produto, TblAtividadeG6.[Segmento Primário], TblAtividadeG6.Operação, TblAtividadeG6.Cliente, TblAtividadeG6.Descrição, TblAtividadeG6.Usuário, TblAtividadeG6.Evento HAVING (((TblAtividadeG6.Perfil) Like '*formalização*'));", dbOpenDynaset)
                
                If TbDados.EOF = False Then
                    TbDados.MoveLast: Qntd = TbDados.RecordCount: TbDados.MoveFirst
                    ObjPlan3Excel.Range("A" & linhaPlan).CopyFromRecordset TbDados
                    linhaPlan = linhaPlan + Qntd
                End If
            
            ObjPlan3Excel.Range("A2:J" & linhaPlan).Borders(3).LineStyle = 1
            ObjPlan3Excel.Range("A2:J" & linhaPlan).Borders(3).Color = RGB(211, 211, 211)
            ObjPlan3Excel.Columns("A:J").Select
            ObjPlan3Excel.Columns.AutoFit
            ObjPlan1Excel.Select

        sFname = "\\bsbrsp54\scoativos\scontrol\Relatorios\Volumetria_Ativos_" & Format(UltimoDiaUtil(), "ddmmyy") & ".xlsm"
           If (Dir(sFname) <> "") Then
             Kill sFname
            End If

   ObjPlan1Excel.SaveAs FileName:=sFname
   ObjExcel.activeworkbook.Close SaveChanges:=False
   ObjExcel.Quit

        Set ofs = CreateObject("Scripting.FileSystemObject")
        
        Assinatura = "\\bsbrsp54\scoativos\scontrol\Mascara\LOGO.jpg"
        
        Set sbObj = New Scripting.FileSystemObject
        Set olapp = CreateObject("Outlook.Application")
        Set oitem = olapp.CreateItem(0)
        
                oitem.Subject = ("VOLUMETRIA DE ATIVOS -  " & Format(Date, "dd/mm/yyyy"))
        
                oitem.SentOnBehalfOfName = "relatoriosconfirming@santander.com.br"
                
                oitem.To = "apdmoraes@santander.com.br;valdir.rocha.motta@santander.com.br"
                oitem.cc = "hscruz@santander.com.br;lseki@santander.com.br;emanuela.conceicao@santander.com.br;elisabeth.aparecida.moreira@santander.com.br;carlos.falcirolli@santander.com.br;jomedeiros@santander.com.br"
        
                Corpo1 = "Prezados,"
                Corpo2 = "Segue o relatorio de "
                Relatorio = "Volumetria de Ativos"
                Corpo21 = " referente as operações do ultimo dia util."
                Corpo3 = ""
                Confirming = ""
                Corpo31 = ""
                Fones = Linkarq
                Corpo4 = "Atenciosamente."
                Assinatura1 = "Processamento de Ativos e Garantias"
                Assinatura2 = "ativoscontrger@santander.com.br"
                Assinatura3 = "Rua Amador Bueno, 474 Santo Amaro - São Paulo Casa 1 / 3º Andar "
                Assinatura4 = ""
                Assinatura5 = ""
                Assinatura6 = "Favor levar em conta o meio-ambiente antes de imprimir este e-mail."
                Assinatura7 = "Por favor tenga en cuenta el medioambiente antes de imprimir este e-mail."
                Assinatura8 = "Please consider your environmental responsibility before printing this e-mail."
                             
                oitem.HTMLBody = "<HTML><BODY><FONT COLOR = BLACK FACE=Calibri <BR>" & Corpo1 & "<BR/>" & _
            "<BR>" & Corpo2 & "<B>" & Relatorio & "</B>" & Corpo21 & "<BR><BR>" & Corpo4 & "<BR><BR><BR>" & _
            " <img src=" & Assinatura & " height=50 width=150>" & "<BR>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 3" & "<BR>" & _
            "<b>" & Assinatura1 & "<BR/>" & Assinatura2 & "</b><BR/>" & "</FONT><FONT COLOR = BLACK FACE = Calibri Size = 2 <BR>" & Assinatura3 & _
            "<BR/></FONT><FONT COLOR = BLACK FACE = Calibri Size = 1 <BR><I>" & Assinatura6 & _
            "<BR/>" & Assinatura7 & "<BR/>" & Assinatura8 & oitem.HTMLBody & "</BODY></HTML>"
            
                oitem.Attachments.Add "\\bsbrsp54\scoativos\scontrol\Mascara\Logo_Confirming.jpg"
                oitem.Attachments.Add sFname
                
                oitem.Send
                
            'Oitem.Display True
            Set olapp = Nothing
            Set oitem = Nothing

End Sub

