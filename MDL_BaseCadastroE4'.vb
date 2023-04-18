MDL_BaseCadastroE4'

Sub GerarBaseCadastroE4()
 
Dim Db As Database
Dim TbDados As Recordset
Dim TbDados1 As Recordset
Dim TbData As Recordset
Dim Str_linha As String

Arquivo = "C:\Temp\CadastroE4.txt"
 
Set Db = OpenDatabase("\\saont46\apps2\Confirming\PROJETORELATORIOS\BD\CadastroE4.mdb")
 
Set TbData = CurrentDb.OpenRecordset("TblCalendario", dbOpenDynaset)

DiarioPesq = Date - 1

Do While True

TbData.FindFirst "Data_dia like '" & DiarioPesq & "'"
     If TbData.NoMatch = False Then

        If TbData!Tipo <> "UTIL" Then: DiarioPesq = DiarioPesq - 1

        If TbData!Tipo = "UTIL" Then: Exit Do

     End If
     
Loop
 
 DataPesq = Format(DiarioPesq, "mm/dd/yyyy")
 
    'Operações do GBM que emitem Termo de Cessão
    Set TbDados = Db.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, 'EMPRESAS' AS SEGMENTO, TblSegmentos.Area_Geral, 'TERMO' AS TIPO, TblArqoped.Cod_Oper, TblArqoped.Valor_op, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor FROM TblSegmentos RIGHT JOIN (Tbl_FielCorporate RIGHT JOIN TblArqoped ON (Tbl_FielCorporate.Convenio = TblArqoped.Convenio_Ancora) AND (Tbl_FielCorporate.Agencia = TblArqoped.Agencia_Ancora)) ON TblSegmentos.Segmento = TblArqoped.Segmento" _
    & " GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, 'EMPRESAS', TblSegmentos.Area_Geral, 'TERMO', TblArqoped.Cod_Oper, TblArqoped.Valor_op, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, Tbl_FielCorporate.Agencia HAVING (((TblSegmentos.Area_Geral)='GBM') AND ((TblArqoped.Data_op)=#" & DataPesq & "#) AND ((Tbl_FielCorporate.Agencia) Is Null));", dbOpenDynaset)

 Open Arquivo For Output As #1
 
Do While TbDados.EOF = False
         
              Str_linha = ""
              Str_linha = Str_linha & TbDados!Agencia_Ancora
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & Right(TbDados!Convenio_Ancora, 4)
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Cnpj_Ancora
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!SEGMENTO
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Area_Geral
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Tipo
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!cod_oper
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Valor_Op
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Data_op
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & Format(Time, "HHMM")
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & Format(Date, "DDMMYYYY")
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Nome_Fornecedor
              Str_linha = Str_linha & """"

        Print #1, Str_linha
        
    
      TbDados.MoveNext
      
    Loop
    
    'Operações do Corporate que podem entregar o Termo em D+1
    Set TbDados = Db.OpenRecordset("SELECT TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, 'EMPRESAS' AS SEGMENTO, TblSegmentos.Area_Geral, 'TERMO' AS TIPO, TblArqoped.Cod_Oper, TblArqoped.Valor_op, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor" _
    & " FROM TblSegmentos RIGHT JOIN (Tbl_FielCorporate RIGHT JOIN TblArqoped ON (Tbl_FielCorporate.Convenio = TblArqoped.Convenio_Ancora) AND (Tbl_FielCorporate.Agencia = TblArqoped.Agencia_Ancora)) ON TblSegmentos.Segmento = TblArqoped.Segmento GROUP BY TblArqoped.Nome_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, 'EMPRESAS', TblSegmentos.Area_Geral, 'TERMO', TblArqoped.Cod_Oper, TblArqoped.Valor_op, TblArqoped.Data_op, TblArqoped.Nome_Fornecedor, Tbl_FielCorporate.Agencia HAVING (((TblArqoped.Agencia_Ancora)=2271 Or (TblArqoped.Agencia_Ancora)=2017) AND ((TblArqoped.Convenio_Ancora)='008500000988' Or (TblArqoped.Convenio_Ancora)='008500000036') AND ((TblArqoped.Data_op)=#" & DataPesq & "#) AND ((Tbl_FielCorporate.Agencia) Is Null));", dbOpenDynaset)

Do While TbDados.EOF = False
         
              Str_linha = ""
              Str_linha = Str_linha & TbDados!Agencia_Ancora
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & Right(TbDados!Convenio_Ancora, 4)
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Cnpj_Ancora
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!SEGMENTO
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Area_Geral
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Tipo
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!cod_oper
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Valor_Op
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Data_op
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & Format(Time, "HHMM")
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & Format(Date, "DDMMYYYY")
              Str_linha = Str_linha & """"
              Str_linha = Str_linha & TbDados!Nome_Fornecedor
              Str_linha = Str_linha & """"

        Print #1, Str_linha
        
    
      TbDados.MoveNext
      
    Loop

Close #1

 
 End Sub
 
