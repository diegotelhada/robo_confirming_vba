'MDL_UserData'

'Verifica o nome do usuário logado na máquina
Public Function nomeUser() As String
    Set objSysInfo = CreateObject("ADSystemInfo")
    struser = objSysInfo.UserName
    Set objUser = GetObject("LDAP://" & struser)
    strfullname = objUser.Get("displayName")
    nomeUser = strfullname
End Function

'Verifica o login do usuário logado na máquina
Public Function loginUser() As String
    loginUser = Environ("USERNAME")
End Function

'Verifica o ID (Unidade C:) do computador onde esta macro for executada
Public Function idMaqUser() As String
    'Verificação de ID da máquina (drive C:)
    idMaqUser = CreateObject("Scripting.FileSystemObject").GetDrive("C:\").SerialNumber
End Function
