Attribute VB_Name = "dev_Git"
Option Explicit

Sub tools_ClipBoard_GitLocalRemote()
Dim c As New Collection
Dim s As New clsGit
Dim strTema As String: strTema = InputBox("Qual o link do repositorio remoto?", "Link repositorio remoto")

    If strTema <> "" Then
        s.strLinkRemote = Trim(strTema)
        s.git_local_remote c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub


Sub tools_ClipBoard_GitTips()
Dim c As New Collection
Dim s As New clsGit
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Git tips")

    If retVal = vbOK Then
        s.Tips c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub
