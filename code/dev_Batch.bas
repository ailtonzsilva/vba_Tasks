Attribute VB_Name = "dev_Batch"
Option Explicit

Sub tools_ClipBoard_Compact()
Dim c As New Collection: Dim s As New clsRepositorio: Dim strTema As String: strTema = InputBox("Qual o nome do arquivo?", "Ocultar arquivo")

    If strTema <> "" Then
        s.strFileName = Trim(strTema)
        s.strPassword = Trim(InputBox("Digite a senha", "Ocultar arquivo - Senha"))
        s.strFilePath = GetFolder()
        s.strPathSourceApp = pathSource & "pstools\7za.exe"
        s.HideFile c
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub

Sub tools_ClipBoard_BatTips()
Dim c As New Collection: Dim s As New clsRepositorio: Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Repositorio - Tips")

    If retVal = vbOK Then
        s.Tips c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub

Sub tools_scriptBackup()
Dim colNews As New Collection: Dim s As New clsRepositorio: s.createBkp colNews, ThisWorkbook.Path & "\pstools\7za.exe"
    
    execucao colNews, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad
    
End Sub

Sub tools_scriptBackupFolders()
Dim colNews As New Collection: Dim s As New clsRepositorio: s.createBkpFolders colNews, ThisWorkbook.Path & "\pstools\7za.exe"
    
    execucao colNews, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=GetFolder()
    
End Sub

Sub tools_scriptBackupFolderControl()
Dim colNews As New Collection: Dim s As New clsRepositorio: s.createFolderControl colNews
    
    execucao colNews, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=GetFolder()
    
End Sub
