'Mdl_Funcoes'


Option Compare Database
Private Const C_STR_NOME_INFO As String = "INFO"
Private Const C_STR_NOME_AVISO As String = "AVISO"
Private Const C_STR_NOME_ERRO As String = "ERRO"
Private Const C_STR_NOME_INICIO_PROCESSO As String = "INICIO PROCESSO"
Private Const C_STR_NOME_FIM_PROCESSO As String = "FIM PROCESSO"

'Esta classe depende da classe clsArquivosTXT
Public Enum opcaoLog
    INFO = 1
    AVISO = 2
    ERRO = 3
    INICIO_PROCESSO = 4
    FIM_PROCESSO = 5
End Enum

Public Function CalcJuros(valor As Double, taxa As Double, Prazo As Double) As Double
    CalcJuros = ((valor * taxa * Prazo) / 3000)
    CalcJuros = Round(CalcJuros, 2)
End Function
Public Function CalcBanco(valor As Double, taxa As Double, Prazo As Double) As Double
    CalcBanco = (((valor * taxa * Prazo) / 3000) / 1)
    CalcBanco = Round(CalcBanco, 2)
End Function
Public Function CalcPPB(ValorJuros As Double, ValorBanco As Double) As Double
    CalcPPB = ValorJuros - ValorBanco
    CalcPPB = Round(CalcPPB, 2)
End Function
Public Function CalcPPBMEXICHEM(ValorBanco As Double) As Double
    CalcPPBMEXICHEM = (ValorBanco * 0.49)
    CalcPPBMEXICHEM = Round(CalcPPBMEXICHEM, 2)
End Function
Public Function CalcDiferimento(Diferimento As Double, DataDif As Integer) As Double
    CalcDiferimento = (Diferimento * DataDif)
End Function
Public Function AbrirBDPPB()
    Set BDPPB = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\BDPPB.mdb")
End Function
Public Function AbrirBDRelatorios()
    Set BDREL = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb")
End Function
Public Function AbrirBDLocal()
    Set BDRELocal = OpenDatabase("C:\Temp\Relatorios Confirming.mdb")
End Function
Public Function AbrirBDBoletos()
    Set BDB = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\BoletosLiquidados.mdb")
End Function
Function AbrirBDVolumetria()
    Set BDVolumetria = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Volumetria.accdb")
End Function
Function AbrirDBTVirtual()
    Set DBTVirtual = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\BD_TERMOVIRTUAL.accdb")
End Function
Function AbrirBDPPBYC()
    Set BDPPBYC = OpenDatabase("\\saont46\apps2\Confirming\BD_PPB_YC\BD_PPB_YC.mdb")
End Function
Function AbrirBDTermos()
    Set BDTermos = OpenDatabase("\\bsbrsp54\ConfirmingBack\BD\BD_Confirming.mdb")
End Function
Function AbrirBDVencidos()
    Set BDVencidos = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\VencidosYD.accdb")
End Function
Public Function CopiarBaseToMaquina(Tipo As String)
    
    Dim FSO As New FileSystemObject
        
        If Tipo = "UPLOAD" Then
        
                FSO.CopyFile "C:\Temp\Relatorios Confirming.mdb", "\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb", True
                
        ElseIf Tipo = "DOWNLOAD" Then
            If FSO.FileExists("C:\Temp\Relatorios Confirming.mdb") Then: FSO.DeleteFile "C:\Temp\Relatorios Confirming.mdb", False
            
                FSO.CopyFile "\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\Relatorios Confirming.mdb", "C:\Temp\", True
        End If
End Function
Public Function IncluirValorPPB(CHAVE As String, ValorBanco As Double, ValorPPb As Double)
    
    Call AbrirBDPPB

        Set TbIncluir = BDPPB.OpenRecordset("SELECT Tbl_PPB_FINAL.CHAVE, Tbl_PPB_FINAL.Valor_Banco, Tbl_PPB_FINAL.PPB_Bruto FROM Tbl_PPB_FINAL WHERE (((Tbl_PPB_FINAL.CHAVE)='" & CHAVE & "'));", dbOpenDynaset)
            TbIncluir.Edit
                TbIncluir!Valor_Banco = ValorBanco
                TbIncluir!PPB_Bruto = ValorPPb
            TbIncluir.Update
        TbIncluir.Close

End Function
Public Function PesqUsername()

    SiglaUser = String(255, 0)
    Ret = GetUserName(SiglaUser, Len(SiglaUser))
    
    X = 1
    Do While Asc(Mid(SiglaUser, X, 1)) <> 0
        X = X + 1
    Loop
        SiglaUser = Left(SiglaUser, (X - 1))
        SiglaPesq = UCase(SiglaUser)

    PesqUsername = SiglaPesq

End Function
Public Function AtualizarStatus(Etapa As String, Percentual As Double, unidade As Double)
    
    Forms!Frm_PPB!LbEtapas.Caption = Etapa
    Forms!Frm_PPB!LbStatus.Width = Forms!Frm_PPB!LbStatus.Width + unidade
    Forms!Frm_PPB!LbPercentual.Caption = Percentual & "%"
    If Percentual >= 50 Then
        Forms!Frm_PPB!LbPercentual.ForeColor = &HFFFFFF  '&H0&       'Preto
    Else
        Forms!Frm_PPB!LbPercentual.ForeColor = &H0& '&HFFFFFF   'Branco
    End If
    DoEvents

End Function
Public Function LetraDaColuna(Coluna As Integer) As String

    If Coluna = 1 Then: LetraDaColuna = "A": GoTo Fim
    If Coluna = 2 Then: LetraDaColuna = "B": GoTo Fim
    If Coluna = 3 Then: LetraDaColuna = "C": GoTo Fim
    If Coluna = 4 Then: LetraDaColuna = "D": GoTo Fim
    If Coluna = 5 Then: LetraDaColuna = "E": GoTo Fim
    If Coluna = 6 Then: LetraDaColuna = "F": GoTo Fim
    If Coluna = 7 Then: LetraDaColuna = "G": GoTo Fim
    If Coluna = 8 Then: LetraDaColuna = "H": GoTo Fim
    If Coluna = 9 Then: LetraDaColuna = "i": GoTo Fim
    If Coluna = 10 Then: LetraDaColuna = "J": GoTo Fim
    If Coluna = 11 Then: LetraDaColuna = "K": GoTo Fim
    If Coluna = 12 Then: LetraDaColuna = "l": GoTo Fim
    If Coluna = 13 Then: LetraDaColuna = "M": GoTo Fim
    If Coluna = 14 Then: LetraDaColuna = "N": GoTo Fim
    If Coluna = 15 Then: LetraDaColuna = "O": GoTo Fim
    If Coluna = 16 Then: LetraDaColuna = "P": GoTo Fim
    If Coluna = 17 Then: LetraDaColuna = "Q": GoTo Fim
    If Coluna = 18 Then: LetraDaColuna = "R": GoTo Fim
    If Coluna = 19 Then: LetraDaColuna = "S": GoTo Fim
    If Coluna = 20 Then: LetraDaColuna = "T": GoTo Fim
    If Coluna = 21 Then: LetraDaColuna = "U": GoTo Fim
    If Coluna = 22 Then: LetraDaColuna = "V": GoTo Fim
    If Coluna = 23 Then: LetraDaColuna = "W": GoTo Fim
    If Coluna = 24 Then: LetraDaColuna = "X": GoTo Fim
    If Coluna = 25 Then: LetraDaColuna = "Y": GoTo Fim
    If Coluna = 26 Then: LetraDaColuna = "Z": GoTo Fim
Fim:
End Function
Public Function UltimoDiaUtil()

    Dim TbData As Recordset
    
        Call AbrirBDRelatorios
            Set TbData = BDREL.OpenRecordset("TblCalendario", dbOpenDynaset)
              DiarioPesq = Date - 1
                Do While True
                    TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
                         If TbData.NoMatch = False Then
                            If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1
                            If TbData!Tipo = "UTIL" Then: Exit Do
                         End If
                Loop
            TbData.Close
            'DiarioPesq = DateSerial(2020, 9, 4)
       UltimoDiaUtil = DiarioPesq
End Function
Public Function TratarArquivoDeBoletos(Caminho As String)

    Dim Nota As String, NotaAjustada As String, j As Integer

        ArqDestino = "C:\Temp\BOLETOS.txt"
            Open ArqDestino For Output As #2
        ArqOrigem = Caminho
            Open ArqOrigem For Input As #1
                    Do While Not EOF(1)
                      Line Input #1, FileBuffer
                        If UCase(Trim(Mid(FileBuffer, 22, 30))) = "YPIOCA" Or UCase(Trim(Mid(FileBuffer, 22, 30))) = "SCHENCK PROCESS EQUIP" Or UCase(Trim(Mid(FileBuffer, 22, 30))) = "LIOTECNICA" Then
                            Nota = UCase(Trim(Mid(FileBuffer, 133, 30)))
                            NotaAjustada = Trim(Replace(Nota, ";", " "))
                            QntdChr = Len(NotaAjustada)
                            LimiteChr = 29 - QntdChr
                                For i = 1 To LimiteChr
                                    NotaAjustada = NotaAjustada & " "
                                    QntdChr = Len(NotaAjustada)
                                Next i
                            Print #2, Mid(FileBuffer, 1, 133) & NotaAjustada & Mid(FileBuffer, 163)
                        Else
                            Print #2, FileBuffer
                        End If
                    Loop
                Close #2
            Close #1
End Function
Function TratarArquivoDeLiquidados(Caminho As String)

    Dim Nota As String, NotaAjustada As String, j As Integer
    
        ArqDestino = "C:\Temp\LIQUIDACOES.txt"
            Open ArqDestino For Output As #2
        ArqOrigem = Caminho
            Open ArqOrigem For Input As #1
                    Do While Not EOF(1)
                      Line Input #1, FileBuffer
                        If UCase(Trim(Mid(FileBuffer, 22, 30))) = "YPIOCA" Or UCase(Trim(Mid(FileBuffer, 22, 30))) = "SCHENCK PROCESS EQUIP" Or UCase(Trim(Mid(FileBuffer, 22, 30))) = "LIOTECNICA" Then
                            Nota = UCase(Trim(Mid(FileBuffer, 133, 30)))
                            NotaAjustada = Trim(Replace(Nota, ";", "-"))
                            QntdChr = Len(NotaAjustada)
                            LimiteChr = 29 - QntdChr
                                For i = 1 To LimiteChr
                                    NotaAjustada = NotaAjustada & " "
                                    QntdChr = Len(NotaAjustada)
                                Next i
                            Print #2, Mid(FileBuffer, 1, 133) & NotaAjustada & Mid(FileBuffer, 163)
                        Else
                            Print #2, FileBuffer
                        End If
                    Loop
                Close #2
            Close #1
End Function
Function TratarArquivoDeBaixados(Caminho As String)

    Dim Nota As String, NotaAjustada As String, j As Integer
    
        ArqDestino = "C:\Temp\ARQCOMPBAIXADO.txt"
            Open ArqDestino For Output As #2
        ArqOrigem = Caminho
            Open ArqOrigem For Input As #1
                    Do While Not EOF(1)
                      Line Input #1, FileBuffer
                        If UCase(Trim(Mid(FileBuffer, 71, 40))) <> "NAO HOUVE COMPROMISSOS BAIXADOS NO YD" Then
                            If UCase(Trim(Mid(FileBuffer, 24, 30))) = "YPIOCA" Or UCase(Trim(Mid(FileBuffer, 24, 30))) = "SCHENCK PROCESS EQUIP" Or UCase(Trim(Mid(FileBuffer, 24, 30))) = "LIOTECNICA" Then
                                Nota = UCase(Trim(Mid(FileBuffer, 128, 30)))
                                NotaAjustada = Trim(Replace(Nota, ";", "-"))
                                QntdChr = Len(NotaAjustada)
                                LimiteChr = 29 - QntdChr
                                    For i = 1 To LimiteChr
                                        NotaAjustada = NotaAjustada & " "
                                        QntdChr = Len(NotaAjustada)
                                    Next i
                                Print #2, Mid(FileBuffer, 1, 133) & NotaAjustada & Mid(FileBuffer, 163)
                            Else
                                Print #2, FileBuffer
                            End If
                        End If
                    Loop
                Close #2
            Close #1
End Function
Public Function Extrai_Zeros(Cpf_CNPJ As String)

Dim i        As Integer
Dim num_pos  As Integer

i = 0
For i = 1 To Len(Cpf_CNPJ)
    If InStr(1, Cpf_CNPJ, "0") > 0 Then
        If Left(Cpf_CNPJ, 1) = 0 Then
            num_pos = IIf(i = 1, 1, i - 1)
            Cpf_CNPJ = Right(Cpf_CNPJ, Len(Cpf_CNPJ) - 1)
        End If
    End If
    i = i + 1
Next i

Extrai_Zeros = Cpf_CNPJ

End Function
Function Trata_NomeArquivo(Nome As String) As String

    Dim Data As String
    Dim TotalCampo As Double
    Dim CampoAtu As Double
    Dim ProvAtu As String
    Dim Y As String
    TotalCampo = Len(Nome)
    CampoAtu = 1
    
    Do While Not TotalCampo = CampoAtu
           
       ProvAtu = Trim(Mid(Nome, CampoAtu, 1))
              
       If ProvAtu = "/" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "_" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "Ç" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "C" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "ç" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "C" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "Ã" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "A" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "ã" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "a" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "Õ" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "O" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "õ" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "o" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "É" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "E" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "È" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "E" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "é" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "e" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "è" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "e" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "À" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "A" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "Á" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "A" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "á" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "a" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "à" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "a" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "Ó" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "O" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "Ò" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "O" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "ó" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "o" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "ò" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & "o" & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
        
       ElseIf ProvAtu = "." Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       ElseIf ProvAtu = "" Then
       
       Nome = Trim(Mid(Nome, 1, CampoAtu - 1)) & Trim(Mid(Nome, CampoAtu + 1, TotalCampo))
       
       End If
       
       CampoAtu = CampoAtu + 1
      
    Loop
     
     Data = Nome
     Trata_NomeArquivo = Data

End Function
Public Function SemCaracterEspecial(ByVal palavra As String) As String
'Alterado para tratamento de caracteres especiais ainda não mapeados na tabela de fornecedores - Emerson 17/04/2018
    Dim cont As Integer
    SemCaracterEspecial = palavra
    For cont = 1 To Len(SemCaracterEspecial)
      Select Case Mid(SemCaracterEspecial, cont, 1)
        Case "Ç", "¸"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "C" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
        Case "ç"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "c" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
        Case "á", "à", "ã", "â", "ª"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "a" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
        Case "Á", "À", "Ã", "Â", "¶", "Ñ", "µ"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "A" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
        Case "É", "È", "Ê", "»", "¼"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "E" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
        Case "é", "è", "ê"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "e" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "Í", "Ì", "Î", "¿"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "I" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "í", "ì", "î"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "i" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "Ó", "Ò", "Õ", "Ô", "ý"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "O" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "ó", "ò", "õ", "ô", "º", "°"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "o" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "Ú", "Ù", "Û", "Ü", "þ"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "U" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "ú", "ù", "û", "ü"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "u" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "@", "#", "$", "%", "$", "&", "§"
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "o" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case "", "", ""
          SemCaracterEspecial = Mid(SemCaracterEspecial, 1, cont - 1) + "o" + Mid(SemCaracterEspecial, cont + 1, Len(SemCaracterEspecial))
         Case Else
      End Select
    Next
End Function
Function TratarArquivoDeOperacoes(Caminho As String)

    Dim Nota As String, NotaAjustada As String, j As Integer

        ArqDestino = "C:\Temp\ARQ.txt"
            Open ArqDestino For Output As #2
        ArqOrigem = Caminho
            Open ArqOrigem For Input As #1
                    Do While Not EOF(1)
                      Line Input #1, FileBuffer
                        If UCase(Trim(Mid(FileBuffer, 6, 30))) = "YPIOCA" Or UCase(Trim(Mid(FileBuffer, 6, 30))) = "SCHENCK PROCESS EQUIP" Or UCase(Trim(Mid(FileBuffer, 6, 30))) = "LIOTECNICA" Then
                            Nota = UCase(Trim(Mid(FileBuffer, 343, 30)))
                            NotaAjustada = Trim(Replace(Nota, ";", " "))
                            QntdChr = Len(NotaAjustada)
                            LimiteChr = 29 - QntdChr
                                For i = 1 To LimiteChr
                                    NotaAjustada = NotaAjustada & " "
                                    QntdChr = Len(NotaAjustada)
                                Next i
                            Print #2, Mid(FileBuffer, 1, 343) & NotaAjustada & Mid(FileBuffer, 373)
                        Else
                            Print #2, FileBuffer
                        End If
                    Loop
                Close #2
            Close #1
            
End Function
Function ReplaceEXE(Caminho As String)

    sFname1 = "C:\Temp\ArquivoFinalizado.TXT"
        If (Dir(sFname1) <> "") Then
            Kill sFname1
        End If

    sFname = "C:\Temp\arquivo.TXT"
        If (Dir(sFname) <> "") Then
            Kill sFname
        End If

        Open sFname For Output As #1
            Print #1, Caminho
        Close #1

      stAppName = "\\Saont46\apps2\Confirming\PROJETORELATORIOS\BD\REPLACE\REPLACE_ARQ.EXE"
         Call Shell(stAppName, 1)

    Do While True
        If (Dir(sFname1) <> "") Then
            Exit Do
        End If
    Loop
    
End Function
Function MesesEspanhol(Mes As String)
    
    If Mes = "01" Then: MesesEspanhol = "ENERO": GoTo Fim
    If Mes = "02" Then: MesesEspanhol = "FEBRERO": GoTo Fim
    If Mes = "03" Then: MesesEspanhol = "MARZO": GoTo Fim
    If Mes = "04" Then: MesesEspanhol = "ABRIL": GoTo Fim
    If Mes = "05" Then: MesesEspanhol = "MAYO": GoTo Fim
    If Mes = "06" Then: MesesEspanhol = "JUNIO": GoTo Fim
    If Mes = "07" Then: MesesEspanhol = "JULIO": GoTo Fim
    If Mes = "08" Then: MesesEspanhol = "AGOSTO": GoTo Fim
    If Mes = "09" Then: MesesEspanhol = "SETIEMBRE": GoTo Fim
    If Mes = "10" Then: MesesEspanhol = "OCTUBRE": GoTo Fim
    If Mes = "11" Then: MesesEspanhol = "NOVIEMBRE": GoTo Fim
    If Mes = "12" Then: MesesEspanhol = "DICIEMBRE": GoTo Fim
Fim:
End Function
Function AtualizarTabelaDoDia(Relatorio)
        
    UltimoDia = UltimoDiaUtil()
                   
    'Limpar tabela do Dia
    BDRELocal.Execute ("DELETE Temp_TblOperDia.* FROM Temp_TblOperDia;")
        
        If UCase(Relatorio) = "ANTECIPADAS" Then
            
            UltimoDia = Right(UltimoDia, 4) & "-" & Mid(UltimoDia, 4, 2) & "-" & Left(UltimoDia, 2)
            
            'Atualizar tabela com as operações do ultimo dia util
            BDRELocal.Execute ("INSERT INTO Temp_TblOperDia ( Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora, Nome_Ancora, Data_op, CHAVE ) SELECT Extrai_Zeros([ARQ]![Agencia]) AS Agencia, arq.Convenio, arq.CNPJ_Ancora, arq.Ancora, arq.Data_op, Format([ARQ]![Agencia],'0000') & Format([ARQ]![Convenio],'000000000000') AS CHAVE" _
            & " FROM arq GROUP BY Extrai_Zeros([ARQ]![Agencia]), arq.Convenio, arq.CNPJ_Ancora, arq.Ancora, arq.Data_op, Format([ARQ]![Agencia],'0000') & Format([ARQ]![Convenio],'000000000000')" _
            & " HAVING (((arq.Data_op)='" & UltimoDia & "') AND ((Format([ARQ]![Agencia],'0000') & Format([ARQ]![Convenio],'000000000000')) In (SELECT Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblRelatorios WHERE (((TblRelatorios.Relatorios) Like '*Antecipadas*')) GROUP BY Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000');)));")

        ElseIf UCase(Relatorio) = "VENCER" Then
            
            UltimoDia = Format(UltimoDia, "mm/dd/yyyy")
            
            'Atualizar Tabela com as operações a vencer no dia
            BDRELocal.Execute ("INSERT INTO Temp_TblOperDia ( Nome_Ancora, Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora, CHAVE ) SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblArqoped WHERE (((TblArqoped.Data_Venc) > #" & UltimoDia & "#)) GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000')" _
            & " HAVING (((Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000')) In (SELECT Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000') AS CHAVE FROM TblRelatorios WHERE (((TblRelatorios.Relatorios) Like '*Vencer*')) GROUP BY Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000');)));")
        
        ElseIf UCase(Relatorio) = "SEMANAL" Then
        
            'Atualizar Tabela com as operações da ultima semana
            BDRELocal.Execute ("INSERT INTO Temp_TblOperDia ( Nome_Ancora, Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora, CHAVE ) SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblArqoped WHERE (((TblArqoped.Data_op) >= Date()-7)) GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000')" _
            & " HAVING (((Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000')) In (SELECT Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblRelatorios INNER JOIN TblAncoras ON (TblRelatorios.Convenio_Ancora = TblAncoras.Convenio_Ancora) AND (TblRelatorios.Agencia_Ancora = TblAncoras.Agencia_Ancora) WHERE (((TblAncoras.Status_Ancora) = 'ATIVO') And ((TblRelatorios.Relatorios) Like '*Antecipadas*') And ((TblRelatorios.Periodicidade) = 'Semanal')) GROUP BY Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000');)));")

        ElseIf UCase(Relatorio) = "MENSAL" Then
                    
            'Atualizar Tabela com as operações do ultimo mês
            BDRELocal.Execute ("INSERT INTO Temp_TblOperDia ( Nome_Ancora, Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora, CHAVE ) SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblArqoped WHERE (((TblArqoped.Data_op) >= Date()-34)) GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000') HAVING (((Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000')) In (SELECT Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblRelatorios INNER JOIN TblAncoras ON (TblRelatorios.Convenio_Ancora = TblAncoras.Convenio_Ancora) AND (TblRelatorios.Agencia_Ancora = TblAncoras.Agencia_Ancora) WHERE (((TblAncoras.Status_Ancora) = 'ATIVO') And ((TblRelatorios.Relatorios) Like '*Antecipadas*') And ((TblRelatorios.Periodicidade) = 'Mensal')) GROUP BY Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000');)));")
        
        ElseIf UCase(Relatorio) = "QUINZENAL" Then
                    
            'Atualizar Tabela com as operações do ultimo mês
            BDRELocal.Execute ("INSERT INTO Temp_TblOperDia ( Nome_Ancora, Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora, CHAVE ) SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblArqoped WHERE (((TblArqoped.Data_op) >= Date()-15)) GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000') HAVING (((Format([TblArqoped]![Agencia_Ancora],'0000') & Format([TblArqoped]![Convenio_Ancora],'000000000000')) In (SELECT Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000') AS CHAVE" _
            & " FROM TblRelatorios INNER JOIN TblAncoras ON (TblRelatorios.Convenio_Ancora = TblAncoras.Convenio_Ancora) AND (TblRelatorios.Agencia_Ancora = TblAncoras.Agencia_Ancora) WHERE (((TblAncoras.Status_Ancora) = 'ATIVO') And ((TblRelatorios.Relatorios) Like '*Antecipadas*') And ((TblRelatorios.Periodicidade) = 'Mensal')) GROUP BY Format([TblRelatorios]![Agencia_Ancora],'0000') & Format([TblRelatorios]![Convenio_Ancora],'000000000000');)));")
                
        End If

        ''Para adequação de consolidação de informações de diversos convênios em um único relatório - 06/12/2018
        'Atualizar Temp_TblOperDia.ChaveGrupo com aquilo que está na tabela TblConvenioGrupo.ChaveGrupo
        BDRELocal.Execute "Update Temp_TblOperDia Inner Join TblConvenioGrupo " & _
            "On Temp_TblOperDia.Agencia_Ancora = TblConvenioGrupo.Agencia_Ancora " & _
            "And Temp_TblOperDia.Convenio_Ancora = TblConvenioGrupo.Convenio_Ancora " & _
            "Set Temp_TblOperDia.ChaveGrupo = TblConvenioGrupo.ChaveGrupo"
        'Atualizar a coluna ChaveGrupo com o conteúdo da coluna chave
        BDRELocal.Execute "Update (Select * From Temp_TblOperDia) set [ChaveGrupo] = [Chave] Where ChaveGrupo Is Null Or ChaveGrupo = ''"

End Function
Function BuscaNomeAgrupador(Agencia, Convenio)
    
    Dim TbDados As Recordset
        
        Call AbrirBDVolumetria
    
        Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_Convenios.AGENCIA, Tbl_Convenios.NR_CONVENIO, Tbl_Convenios.[NOME CONVENIO                 ] FROM Tbl_Convenios WHERE (((Tbl_Convenios.AGENCIA)=" & Agencia & ") AND ((Tbl_Convenios.NR_CONVENIO)='" & Convenio & "'));", dbOpenDynaset)
            
            If TbDados.EOF = False Then: BuscaNomeAgrupador = TbDados![NOME CONVENIO                 ]

End Function
Function BuscaNomeBanco(Codigo)
    
    Dim TbDados As Recordset
    
        Call AbrirBDVolumetria
    
        Set TbDados = BDVolumetria.OpenRecordset("SELECT Tbl_CodigoBanco.Nome, Tbl_CodigoBanco.Codigo FROM Tbl_CodigoBanco WHERE (((Tbl_CodigoBanco.Codigo)='" & Codigo & "'));", dbOpenDynaset)
            
            If TbDados.EOF = False Then: BuscaNomeBanco = TbDados!Nome

End Function
Public Function processarImportacaoPlan(nomePlan As String, filtro As String) As Boolean
Const C_STR_NOME_METODO As String = "processarImportacaoPlan"
'On Error GoTo SaidaErro
Debug.Print "ini - " & Now()
Dim dbDAO As DAO.Database
Dim rsDao As DAO.Recordset
Dim contLinhas As Double
processarImportacaoPlan = False
Set dbDAO = Nothing
Set dbDAO = CurrentDb()
dbDAO.Execute "delete * from Aux_Fornecedores2"
dbDAO.Execute "delete * from Aux_Fornecedores"
Set rsDao = dbDAO.OpenRecordset("Select * from Aux_Fornecedores2")
Set xlapp = CreateObject("Excel.Application")
With xlapp
    .Workbooks.Open nomePlan
    .Application.DisplayAlerts = False
    .Visible = False
    .Sheets(1).Select
    contLinhas = .Range("A999999").End(-4162).Row
    .Sheets(1).Range("a1:cl1").Select
    .Selection.AutoFilter
    .Sheets(1).Range("$A$1:$cl$" & contLinhas).AutoFilter Field:=78, Criteria1:= _
        "=" & filtro, Operator:=1
    .Sheets(1).Range("A1").Select
    .Sheets(1).Range("$A$1:$cl$1").Select
    .Range(.Selection, .Selection.End(-4121)).Select
    .Selection.Copy
    .Sheets.Add After:=.activesheet
    .Sheets(2).Range("A1").PasteSpecial -4104
    .Application.CutCopyMode = False
    lin = 2
    Do Until .Sheets(2).Cells(lin, 78) = ""
        If .Sheets(2).Cells(lin, 78) = filtro Then
            rsDao.AddNew
            For col = 0 To rsDao.Fields.Count - 1
                rsDao.Fields(col) = .Sheets(2).Cells(lin, col + 1)
            Next
            rsDao.Update
        End If
        lin = lin + 1
    Loop
    .Workbooks.Close
End With
rsDao.Close
Set rsDao = Nothing
Set xlapp = Nothing
Debug.Print "fim - " & Now()
dbDAO.Execute "INSERT INTO Aux_Fornecedores SELECT Aux_Fornecedores2.* FROM Aux_Fornecedores2;"
Set dbDAO = Nothing
processarImportacaoPlan = True
Exit Function
SaidaErro:
    Call salvaLog(opcaoLog.ERRO, caminhoLog, C_STR_NOME_METODO & Chr(165) & Err.Number & Chr(165) & Err.Description)
End Function
Public Sub escreveArqTxt(ByVal caminhoEnomeArq As String, ByVal texto As String)
Close #1
Open caminhoEnomeArq For Append As #1
    Print #1, texto
Close #1
End Sub

Public Function salvaLog(ByVal opcaoLog As Integer, ByVal arquivoLog As String, ByVal msgLog As String)
Const C_STR_NOME_METODO As String = "salvaLog"
Dim concatena As String
On Error GoTo SaidaErro
    concatena = Format(Now(), "dd/mm/yyyy hh:mm") & "¥" & tipoErro(opcaoLog) & "¥" & _
                Environ("username") & "¥" & Environ("computername") & "¥" & CurrentProject.Name & "¥" & msgLog
    Call escreveArqTxt(arquivoLog, concatena)
Exit Function
SaidaErro:
    'Call salvaLog(ERRO, caminhoLog, C_STR_NOME_METODO & Chr(165) & Err.Number & Chr(165) & Err.Description)
End Function
Private Function tipoErro(ByVal nrOpcao As Integer) As String
    Select Case nrOpcao
        Case 1
            tipoErro = C_STR_NOME_INFO
        Case 2
            tipoErro = C_STR_NOME_AVISO
        Case 3
            tipoErro = C_STR_NOME_ERRO
        Case 4
            tipoErro = C_STR_NOME_INICIO_PROCESSO
        Case 5
            tipoErro = C_STR_NOME_FIM_PROCESSO
    End Select
End Function

Function AbrirBDTermosFIAT()
    'Teste
    'Set BDTermosFIAT = OpenDatabase("C:\Users\t706034\Desktop\baseTermosFIAT.accdb")
    
    'Produção
    Set BDTermosFIAT = OpenDatabase("\\Saont46\apps2\Confirming\PROJETORELATORIOS\FIAT_TERMO\BD\baseTermosFIAT.accdb")
End Function


