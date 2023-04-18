'Mdl_Waiver_Corporate'

Option Compare Database
Sub Waiver()
    
    Call atualizabaseAncora
    Call atualizabasefornecedor
    Call GerarRelatoriosWaiver

End Sub
Sub GerarRelatoriosWaiver()

    Dim ObjExcel As Object, TbDados1 As Recordset
    Dim ObjPlan1Excel As Object, linha As Double
    Dim Db As Database, Db1 As Database
    Dim TbAncoras As Recordset, TbForn As Recordset, TbAutorizado As Recordset, TbTermos As Recordset, TbPrazoTermos As Recordset
    Dim Nome As String

        Set Db = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\BD_WAIVER_CORPORATE.mdb")
            
            Caminho = "\\saont46\apps2\Confirming\PROJETORELATORIOS\MASCARAS\XLSX\"

                Set ObjExcel = CreateObject("EXCEL.application")
                ObjExcel.Workbooks.Open FileName:=Caminho & "Risco_Waiver.xlsx", ReadOnly:=True
                Set ObjPlan1Excel = ObjExcel.Worksheets("Relatorio")
                    linha = 3
      
                    Set TbAncoras = Db.OpenRecordset("SELECT Tbl_Ancoras_Operantes.Agencia_Ancora, Tbl_Ancoras_Operantes.Convenio_Ancora, Tbl_Ancoras_Operantes.Nome_Ancora, Tbl_Ancoras_Operantes.Cnpj_Ancora, Tbl_Ancoras.Situacao, DBM.[Grupo Final] FROM (Tbl_Ancoras INNER JOIN Tbl_Ancoras_Operantes ON (Tbl_Ancoras.Agencia_Ancora = Tbl_Ancoras_Operantes.Agencia_Ancora) AND (Tbl_Ancoras.Convenio_Ancora = Tbl_Ancoras_Operantes.Convenio_Ancora)) LEFT JOIN DBM ON Tbl_Ancoras_Operantes.Cnpj_Ancora = DBM.CNPJ WHERE (((DBM.Segmento) = 'CORPORATE')) GROUP BY Tbl_Ancoras_Operantes.Agencia_Ancora, Tbl_Ancoras_Operantes.Convenio_Ancora, Tbl_Ancoras_Operantes.Nome_Ancora, Tbl_Ancoras_Operantes.Cnpj_Ancora, Tbl_Ancoras.Situacao, DBM.[Grupo Final] ORDER BY Tbl_Ancoras_Operantes.Nome_Ancora;", dbOpenDynaset)

                        Do While TbAncoras.EOF = False
                            
                            ObjPlan1Excel.Range("B" & linha) = TbAncoras!Cnpj_Ancora
                            ObjPlan1Excel.Range("C" & linha) = TbAncoras!Nome_Ancora
                            ObjPlan1Excel.Range("D" & linha) = TbAncoras![Grupo Final]
                    
                               Set TbForn = Db.OpenRecordset("SELECT TblArqForn.Agencia_Ancora, TblArqForn.Convenio_Ancora, Sum(1) AS Qntd FROM TblArqForn WHERE (((TblArqForn.Status_Fornecedor)='ATIVO') AND ((TblArqForn.TipoBlo_Fornecedor)='SEM BLOQUEIO' Or (TblArqForn.TipoBlo_Fornecedor)='TERMO' Or (TblArqForn.TipoBlo_Fornecedor) Is Null)) GROUP BY TblArqForn.Agencia_Ancora, TblArqForn.Convenio_Ancora HAVING (((TblArqForn.Agencia_Ancora)='" & Format(TbAncoras!Agencia_Ancora, "00") & "') AND ((TblArqForn.Convenio_Ancora)='" & TbAncoras!Convenio_Ancora & "'));", dbOpenDynaset)

                                    If TbForn.EOF = False Then
                                         ObjPlan1Excel.Range("E" & linha) = TbForn!Qntd
                                    Else
                                         ObjPlan1Excel.Range("E" & linha) = "0"
                                    End If
                                    
                                Set TbAutorizado = Db.OpenRecordset("SELECT Tbl_Ancoras_Autorizados.Agencia_Ancora, Tbl_Ancoras_Autorizados.Convenio_Ancora, Tbl_Ancoras_Autorizados.Nome_Ancora, Tbl_Ancoras_Autorizados.CNPJ_Ancora, Tbl_Ancoras_Autorizados.Grupo, Tbl_Ancoras_Autorizados.Dt_Cadastro, Tbl_Ancoras_Autorizados.Dt_Bloqueio, Tbl_Ancoras_Autorizados.Situacao, Tbl_Ancoras_Autorizados.Usuario FROM Tbl_Ancoras_Autorizados WHERE (((Tbl_Ancoras_Autorizados.Agencia_Ancora)=" & TbAncoras!Agencia_Ancora & ") AND ((Tbl_Ancoras_Autorizados.Convenio_Ancora)='" & TbAncoras!Convenio_Ancora & "'));", dbOpenDynaset)
                                    
                                    If TbAutorizado.EOF = False Then
                                        If TbAutorizado!Situacao = "INATIVO" Then
                                            ObjPlan1Excel.Range("F" & linha) = "NÃO"
                                            ObjPlan1Excel.Range("G" & linha) = "SIM"
                                            ObjPlan1Excel.Range("H" & linha) = TbAutorizado!Dt_Bloqueio
                                        Else
                                            ObjPlan1Excel.Range("F" & linha) = "SIM"
                                            ObjPlan1Excel.Range("G" & linha) = "SIM"
                                            ObjPlan1Excel.Range("H" & linha) = TbAutorizado!Dt_Cadastro
                                        End If
                                    Else
                                         ObjPlan1Excel.Range("F" & linha) = "NÃO"
                                         ObjPlan1Excel.Range("G" & linha) = "NÃO"
                                    End If
                    
                                Set TbTermos = Db.OpenRecordset("SELECT Tbl_Termos_Pendentes.Agencia_Ancora, Tbl_Termos_Pendentes.Convenio_Ancora, Sum(1) AS QNTDTERMOS FROM Tbl_Termos_Pendentes GROUP BY Tbl_Termos_Pendentes.Agencia_Ancora, Tbl_Termos_Pendentes.Convenio_Ancora HAVING (((Tbl_Termos_Pendentes.Agencia_Ancora)=" & TbAncoras!Agencia_Ancora & ") AND ((Tbl_Termos_Pendentes.Convenio_Ancora)='" & TbAncoras!Convenio_Ancora & "'));", dbOpenDynaset)
                                
                                Set TbPrazoTermos = Db.OpenRecordset("SELECT Tbl_Termos_Pendentes.Agencia_Ancora, Tbl_Termos_Pendentes.Convenio_Ancora, Tbl_Termos_Pendentes.Dt_Operacao FROM Tbl_Termos_Pendentes GROUP BY Tbl_Termos_Pendentes.Agencia_Ancora, Tbl_Termos_Pendentes.Convenio_Ancora, Tbl_Termos_Pendentes.Dt_Operacao HAVING (((Tbl_Termos_Pendentes.Agencia_Ancora) =" & TbAncoras!Agencia_Ancora & ") And ((Tbl_Termos_Pendentes.Convenio_Ancora) ='" & TbAncoras!Convenio_Ancora & "')) ORDER BY Tbl_Termos_Pendentes.Dt_Operacao;", dbOpenDynaset)

                                    If TbTermos.EOF = False Then
                                         ObjPlan1Excel.Range("I" & linha) = TbTermos!QNTDTERMOS
                                    Else
                                         ObjPlan1Excel.Range("I" & linha) = "0"
                                    End If
                                        
                                    If TbPrazoTermos.EOF = False Then
                                        DiasPendentes = Date - TbPrazoTermos!Dt_Operacao
                                         ObjPlan1Excel.Range("J" & linha) = DiasPendentes
                                    Else
                                         ObjPlan1Excel.Range("J" & linha) = "0"
                                    End If

                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(7).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(8).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(9).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(10).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(11).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(1).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(2).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(3).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(4).LineStyle = 2
                            ObjPlan1Excel.Range("b" & linha & ":K" & linha).Font.Size = 8
                            ObjPlan1Excel.Range("h" & linha & ":h" & linha).NumberFormat = "dd/mm/yyyy"
                            ObjPlan1Excel.Columns("b:K").Select
                            ObjPlan1Excel.Columns.AutoFit
                            ObjPlan1Excel.Rows("3:" & linha).RowHeight = 11.75
                            ObjPlan1Excel.Range("C2").Select
                    
                             linha = linha + 1
                             
                             TbAncoras.MoveNext
                            
                        Loop
                        
       ObjPlan1Excel.Rows(linha & ":300").Delete Shift:=xlUp
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(7).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(8).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(9).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(10).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(11).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(1).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(2).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(3).LineStyle = 2
       ObjPlan1Excel.Range("b" & linha & ":K" & linha).Borders(4).LineStyle = 2

  Data = Date

   Nome = "Relatorio Waiver " & " - " & Format(Data, "ddmmyy")
      
   Nome = Trata_NomeArquivo(Nome)
   
       sFname = "\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
        If (Dir(sFname) <> "") Then
            Kill sFname
        End If
      
   ObjPlan1Excel.SaveAs FileName:="\\saont46\apps2\Confirming\PROJETORELATORIOS\RELATORIOS SALVOS\" & Nome & ".xlsx"
   ObjExcel.activeworkbook.Close SaveChanges:=False
   ObjExcel.Quit
   
End Sub
Sub atualizabaseAncora()

Dim ObjExcelOp As Object
Dim ObjExcelppb As Object
Dim ObjExcel As Object, TbDados1 As Recordset
Dim ObjPlan1Excel As Object, linha As Double
Dim TbDados As Recordset
Dim Db As Database

        File = "C:\Temp\CONVCONF.xlsx"

    Set Db = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\BD_WAIVER_CORPORATE.mdb")

        Db.Execute ("DELETE Tbl_Ancoras.* FROM Tbl_Ancoras;")

    Set ofs = CreateObject("Scripting.FileSystemObject")

        Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=File
                Set ObjExcelOp = ObjExcel.Worksheets(1)
  
        linha = 2

        Set TbDados = Db.OpenRecordset("Tbl_Ancoras", dbOpenDynaset)

            Do While True
            
                If UCase(Trim(ObjExcelOp.Cells(linha, 6))) = "ATIVO" And UCase(Trim(ObjExcelOp.Cells(linha, 7))) = "SEM BLOQUEIO" Then
                    If UCase(Trim(ObjExcelOp.Cells(linha, 8))) <> "TESTE" Then
                
                  Convenio = Trim(ObjExcelOp.Cells(linha, 1))

                    TbDados.AddNew
    
                        TbDados!Agencia_Ancora = Mid(Convenio, 6, 4)
                        TbDados!Convenio_Ancora = Right(Convenio, 12)
                        TbDados!Nome_Ancora = Trim(ObjExcelOp.Cells(linha, 2))
                        TbDados!Dt_Cadastro = Format(Trim(ObjExcelOp.Cells(linha, 4)), "dd/mm/yyyy")
                        TbDados!Dt_UltMovimento = Format(Trim(ObjExcelOp.Cells(linha, 5)), "dd/mm/yyyy")
                        TbDados!Situacao = Trim(ObjExcelOp.Cells(linha, 6))
                        TbDados!Bloqueio = Trim(ObjExcelOp.Cells(linha, 7))
                        TbDados!Ambiente = Trim(ObjExcelOp.Cells(linha, 8))
              
                    TbDados.Update
                    End If
                End If
                    
               linha = linha + 1
                   
              If Trim(ObjExcelOp.Cells(linha, 1)) = "" Then: Exit Do
                
            Loop
            
    ObjExcel.Application.Quit

Db.Execute ("UPDATE Tbl_Ancoras INNER JOIN TblClientes ON (Tbl_Ancoras.Convenio_Ancora = TblClientes.Convenio_Ancora) AND (Tbl_Ancoras.Agencia_Ancora = TblClientes.Agencia_Ancora) SET Tbl_Ancoras.CNPJ_Ancora = [TblClientes]![Cnpj_Ancora];")

End Sub
Sub atualizabasefornecedor()

Dim ObjExcelOp As Object
Dim ObjExcelppb As Object
Dim ObjExcel As Object, TbDados1 As Recordset
Dim ObjPlan1Excel As Object, linha As Double
Dim TbDados As Recordset
Dim Db As Database

        File = "C:\temp\FOLHA-1.xlsx"

    Set Db = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\BD_WAIVER_CORPORATE.mdb")

        Db.Execute ("DELETE Tbl_Termos_Pendentes.* FROM Tbl_Termos_Pendentes;")

    Set ofs = CreateObject("Scripting.FileSystemObject")

        Set ObjExcel = CreateObject("EXCEL.application")
            ObjExcel.Workbooks.Open FileName:=File
                Set ObjExcelOp = ObjExcel.Worksheets(1)
  
        linha = 2

        Set TbDados = Db.OpenRecordset("Tbl_Termos_Pendentes", dbOpenDynaset)

            Do While True
                
                  Convenio = Trim(ObjExcelOp.Cells(linha, 1))

                    TbDados.AddNew
    
                        TbDados!Agencia_Ancora = Mid(Convenio, 6, 4)
                        TbDados!Convenio_Ancora = Right(Convenio, 12)
                        TbDados!Nome_Ancora = Trim(ObjExcelOp.Cells(linha, 2))
                        TbDados!Cnpj_Fornecedor = Trim(ObjExcelOp.Cells(linha, 5))
                        TbDados!Nome_Fornecedor = Trim(ObjExcelOp.Cells(linha, 6))
                        TbDados!COD_OPERACAO = Trim(ObjExcelOp.Cells(linha, 7))
                        TbDados!Dt_Operacao = Format(Trim(ObjExcelOp.Cells(linha, 9)), "dd/mm/yyyy")
                        TbDados!Situação = Trim(ObjExcelOp.Cells(linha, 11))
                        TbDados!Hora = Format(Trim(ObjExcelOp.Cells(linha, 12)), "HH:MM:SS")
                        TbDados!Email_Fornecedor = Trim(ObjExcelOp.Cells(linha, 13))
              
                    TbDados.Update
                    
               linha = linha + 1
                   
              If Trim(ObjExcelOp.Cells(linha, 1)) = "" Then: Exit Do
                
            Loop
            
    ObjExcel.Application.Quit

End Sub
