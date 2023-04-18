'Mdl_PesqUsuario'


Option Compare Database

    Global NomeUsuario As String
    Global EmailUsuario As String
    
    'Capturar login do usuário da rede/equipamento
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Ret As Long
    Public SiglaUser As String

