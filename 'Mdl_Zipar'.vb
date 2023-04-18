'Mdl_Zipar'
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub Zip(ZipFile As String, InputFile As String)
On Error GoTo ErrHandler
    Dim FSO As Object 'Scripting.FileSystemObject
    Dim oApp As Object 'Shell32.Shell
    Dim oFld As Object 'Shell32.Folder
    Dim oShl As Object 'WScript.Shell
    Dim i As Long
    Dim l As Long

    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Not FSO.FileExists(ZipFile) Then
        'Create empty ZIP file
        FSO.CreateTextFile(ZipFile, True).Write _
            "PK" & Chr(5) & Chr(6) & String(18, vbNullChar)
    End If

    Set oApp = CreateObject("Shell.Application")
    Set oFld = oApp.Namespace(CVar(ZipFile))

    Let i = oFld.Items.Count

    oFld.CopyHere (InputFile)

    Set oShl = CreateObject("WScript.Shell")

    'Search for a Compressing dialog
    Do While oShl.AppActivate("Compressing...") = False
        If oFld.Items.Count > i Then
            'There's a file in the zip file now, but
            'compressing may not be done just yet
            Exit Do
        End If
        If l > 30 Then
            '3 seconds has elapsed and no Compressing dialog
            'The zip may have completed too quickly so exiting
            Exit Do
        End If

        DoEvents

        Sleep 100

        Let l = l + 1
    Loop

    ' Wait for compression to complete before exiting
    Do While oShl.AppActivate("Compressing...") = True
        DoEvents

        Sleep 100
    Loop

ExitProc:
    On Error Resume Next
        Set FSO = Nothing
        Set oFld = Nothing
        Set oApp = Nothing
        Set oShl = Nothing
    Exit Sub
ErrHandler:
Exit Sub


    Resume ExitProc

    Resume
End Sub

Public Sub UnZip(ZipFile As String, Optional TargetFolderPath As String = vbNullString, Optional OverwriteFile As Boolean = False)

On Error GoTo ErrHandler
    Dim oApp As Object
    Dim FSO As Object
    Dim fil As Object
    Dim DefPath As String
    Dim strDate As String


    Set FSO = CreateObject("Scripting.FileSystemObject")
   If Len(TargetFolderPath) = 0 Then
        Let DefPath = CurrentProject.Path & "\"
    Else
        If FSO.FolderExists(TargetFolderPath) Then
            Let DefPath = TargetFolderPath & "\"
        Else
            Err.Raise 53, , "Folder not found"
        End If
    End If


    If FSO.FileExists(ZipFile) = False Then
        MsgBox "System could not find " & ZipFile _
            & " upgrade cancelled.", _
            vbInformation, "Error Unziping File"
        Exit Sub
    Else
        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")


        With oApp.Namespace(ZipFile & "\")
            If OverwriteFile Then
                For Each fil In .Items
                    If FSO.FileExists(DefPath & fil.Name) Then
                        Kill DefPath & fil.Name
                    End If
                Next
            End If
            oApp.Namespace(CVar(DefPath)).CopyHere .Items
        End With


        On Error Resume Next
        Kill Environ("Temp") & "\Temporary Directory*"


        'Kill zip file
        Kill ZipFile
    End If


ExitProc:
    On Error Resume Next
    Set oApp = Nothing
    Exit Sub
ErrHandler:
Exit Sub
    Resume ExitProc
    Resume
End Sub


