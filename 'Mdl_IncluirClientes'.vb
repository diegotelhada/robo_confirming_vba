'Mdl_IncluirClientes'

Option Compare Database
Sub IncluirTblClientes()

    Call AbrirBDLocal
        
        'Atualizar Tabela Clientes
        BDRELocal.Execute ("INSERT INTO TblClientes ( Nome_Ancora, Banco_Ancora, Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora )" _
        & " SELECT TblArqoped.Nome_Ancora, TblArqoped.Banco_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora" _
        & " FROM TblArqoped LEFT JOIN TblClientes ON (TblArqoped.Convenio_Ancora = TblClientes.Convenio_Ancora) AND (TblArqoped.Agencia_Ancora = TblClientes.Agencia_Ancora)" _
        & " GROUP BY TblClientes.Convenio_Ancora, TblArqoped.Nome_Ancora, TblArqoped.Banco_Ancora, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora" _
        & " HAVING (((TblClientes.Convenio_Ancora) Is Null));")
        
        'Deletar tabela Clintes x Segmentos
        BDRELocal.Execute ("DELETE TblClientesXSegmento.* FROM TblClientesXSegmento;")
        
        'Atualizar Tabela Clientes X Segmentos
        BDRELocal.Execute ("INSERT INTO TblClientesXSegmento ( Segmento, Agencia_Ancora, Convenio_Ancora, Cnpj_Ancora ) SELECT TblArqoped.Segmento, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora" _
        & " FROM TblArqoped LEFT JOIN TblClientesXSegmento ON (TblArqoped.Convenio_Ancora = TblClientesXSegmento.Convenio_Ancora) AND (TblArqoped.Agencia_Ancora = TblClientesXSegmento.Agencia_Ancora)" _
        & " GROUP BY TblArqoped.Segmento, TblArqoped.Agencia_Ancora, TblArqoped.Convenio_Ancora, TblArqoped.Cnpj_Ancora, TblClientesXSegmento.Convenio_Ancora HAVING (((TblClientesXSegmento.Convenio_Ancora) Is Null));")
        
        'Atualizar Nome do Segmento
        BDRELocal.Execute ("UPDATE TblClientesXSegmento INNER JOIN TblSegmentos ON TblClientesXSegmento.Segmento = TblSegmentos.Segmento SET TblClientesXSegmento.NomeSegmento = [TblSegmentos]![Area_Geral];")
                    
End Sub
