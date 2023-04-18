'MDL_FilesFolders'

'Abre uma caixa para selcionar uma pasta
Public Function GetFolder(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Selecione a pasta"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem & "\"
    Set fldr = Nothing
End Function

'Abre uma caixa para selecionar um arquivo
Public Function searchFile() As String
    Dim f As Object
    Set f = Application.FileDialog(3)
    f.AllowMultiSelect = False
    If f.Show <> -1 Then
        searchFile = ""
        Exit Function
    Else
        searchFile = Trim(f.SelectedItems(1))
    End If
End Function

'Verifica se um arquivo existe
Public Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function

'Deleta um arquivo, se existir
Public Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub
