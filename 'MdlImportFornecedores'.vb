'MdlImportFornecedores'

Sub ImportarArqOped()

Dim Linha As Double, X As Double, Linha1 As String, linha2 As String, linha3 As String
Dim Dataarq As Date
Dim FSO As New FileSystemObject
Dim arq As File
Dim Contador As String
Dim TbDados As Recordset

Set TbDados = CurrentDb.OpenRecordset("TblArqForn")

Contador = 0

For Each File In FSO.GetFolder("C:\Documents and Settings\T683068\Desktop\ARQFO\").Files 'Local dos Arquivos

NOME = File                              'Copia nome/data do arquivo
NameArq = Mid(NOME, 49, 19)              'Reira apenas a Data
Dia = Mid(NameArq, 8, 2)                 'Reira apenas o Dia
Mes = Mid(NameArq, 10, 2)                'Reira apenas o Mes
Ano = Mid(NameArq, 12, 2)                'Reira apenas o Ano
ArqData = Dia & Mes & Ano                'Integra Dia - Mes - Ano
Arqdata1 = Dia & "/" & Mes & "/" & Ano   'Formata Dia - Mes - Ano
Dataarq = Arqdata1                       'Finaliza a Data
 

Set Caminho = FSO.GetFile(File)

    Open Caminho For Input As #1         'Abre Arquivo texto

        'Line Input #1, FileBuffer

        Do While Not EOF(1)
  
        Line Input #1, FileBuffer
        
    TbDados.AddNew
     
        TbDados!ArqData = Dataarq                                 ' Data do Arquivo
     
        TbDados!Segmento = Trim(Mid(FileBuffer, 1, 3))            'Segmento
        
        If Trim(Mid(FileBuffer, 4, 1)) = ";" Then
        TbDados!Bco_Ancora = Trim(Mid(FileBuffer, 5, 4))          'Banco Ancora
        End If
     
        If Trim(Mid(FileBuffer, 9, 1)) = ";" Then                 ' Agencia Ancora
        TbDados!Agencia_Ancora = Trim(Mid(FileBuffer, 10, 4))
        End If
     
        If Trim(Mid(FileBuffer, 14, 1)) = ";" Then                  'Numero do convenio do Ancora
        TbDados!Convenio_Ancora = Trim(Mid(FileBuffer, 15, 12))
        End If
     
        If Trim(Mid(FileBuffer, 27, 1)) = ";" Then 'Nome do Ancora
        TbDados!Nome_Ancora = Trim(Mid(FileBuffer, 28, 30))
        End If
      
        If Trim(Mid(FileBuffer, 58, 1)) = ";" Then ' CNPJ do Fornecedor
        TbDados!Cnpj_Fornecedor = Trim(Mid(FileBuffer, 59, 15))
        End If
      
        If Trim(Mid(FileBuffer, 74, 1)) = ";" Then ' Nome do Fornecedor
        TbDados!Nome_Fornecedor = Trim(Mid(FileBuffer, 75, 40))
        End If
        
        If Trim(Mid(FileBuffer, 131, 1)) = ";" Then ' Status do Fornecedor
        TbDados!Status_Fornecedor = Trim(Mid(FileBuffer, 132, 12))
        End If
        
        If Trim(Mid(FileBuffer, 144, 1)) = ";" Then ' Endereço do Fornecedor
        TbDados!End_Fornecedor = Trim(Mid(FileBuffer, 145, 40))
        End If
        
        
        If Trim(Mid(FileBuffer, 185, 1)) = ";" Then ' Numero do Endereço do Fornecedor
        TbDados!Num_Fornecedor = Trim(Mid(FileBuffer, 186, 5))
        End If
        
        If Trim(Mid(FileBuffer, 191, 1)) = ";" Then ' Complemento do Fornecedor
        TbDados!Bairro_Fornecedor = Trim(Mid(FileBuffer, 192, 20))
        End If
        
        If Trim(Mid(FileBuffer, 212, 1)) = ";" Then ' Cidade do Fornecedor
        TbDados!Cidade_Fornecedor = Trim(Mid(FileBuffer, 213, 30))
        End If
        
        If Trim(Mid(FileBuffer, 243, 1)) = ";" Then ' Estado do Fornecedor
        TbDados!UF_Fornecedor = Trim(Mid(FileBuffer, 244, 2))
        End If
        
        If Trim(Mid(FileBuffer, 246, 1)) = ";" Then ' CEP do Fornecedor
        TbDados!CEP_Fornecedor = Trim(Mid(FileBuffer, 247, 9))
        End If
        
        If Trim(Mid(FileBuffer, 256, 1)) = ";" Then ' DDD1 do Fornecedor
        TbDados!DDD1 = Extrai_Zeros(Trim(Mid(FileBuffer, 257, 4)))
        End If
        
        If Trim(Mid(FileBuffer, 261, 1)) = ";" Then ' Fone1 do Fornecedor
        TbDados!Fone1 = Extrai_Zeros(Trim(Mid(FileBuffer, 262, 10)))
        End If
        
        If Trim(Mid(FileBuffer, 279, 1)) = ";" Then ' DDDFAX do Fornecedor
        TbDados!DDDFAX = Extrai_Zeros(Trim(Mid(FileBuffer, 280, 4)))
        End If
        
        If Trim(Mid(FileBuffer, 284, 1)) = ";" Then ' FAX do Fornecedor
        TbDados!FAX = Extrai_Zeros(Trim(Mid(FileBuffer, 285, 10)))
        End If
        
        If Trim(Mid(FileBuffer, 295, 1)) = ";" Then ' Banco do Fornecedor
        TbDados!Banco_Fornecedor = Extrai_Zeros(Trim(Mid(FileBuffer, 296, 5)))
        End If
        
        If Trim(Mid(FileBuffer, 301, 1)) = ";" Then ' Agencia do Fornecedor
        TbDados!Agencia_Fornecedor = Extrai_Zeros(Trim(Mid(FileBuffer, 302, 5)))
        End If
        
        If Trim(Mid(FileBuffer, 307, 1)) = ";" Then ' Conta do Fornecedor
        TbDados!Conta_Fornecedor = Extrai_Zeros(Trim(Mid(FileBuffer, 308, 13)))
        End If
        
        If Trim(Mid(FileBuffer, 339, 1)) = ";" Then ' Contato do Fornecedor
        TbDados!Contato_Fornecedor = Trim(Mid(FileBuffer, 340, 30))
        End If
        
        If Trim(Mid(FileBuffer, 370, 1)) = ";" Then ' DDD2 do Fornecedor
        TbDados!DDD2 = Extrai_Zeros(Trim(Mid(FileBuffer, 371, 4)))
        End If
        
        If Trim(Mid(FileBuffer, 375, 1)) = ";" Then ' FONE2 do Fornecedor
        TbDados!FONE2 = Extrai_Zeros(Trim(Mid(FileBuffer, 376, 10)))
        End If
        
        If Trim(Mid(FileBuffer, 393, 1)) = ";" Then ' EMAIL do Fornecedor
        TbDados!Email_Fornecedor = Trim(Mid(FileBuffer, 394, 40))
        End If
        
        If Trim(Mid(FileBuffer, 434, 1)) = ";" Then ' Tipo de bloqueio do Fornecedor
        TbDados!TipoBlo_Fornecedor = Trim(Mid(FileBuffer, 435, 12))
        End If
            
     TbDados.Update
     
     Contador = Contador + 1
     
     If Contador = 1167 Then
     MsgBox Contador
     End If
        
     Loop
    
Close #1

Next File

MsgBox "IMPORTADO COM SUCESSO", , "ATENÇÃO"

End Sub



