Attribute VB_Name = "dev_Gitlab"
Option Explicit

Sub tools_ClipBoard_Gitlab_ci_yml()
Dim c As New Collection: Dim s As New clsGitlab: Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Gitlab file yaml")

    If retVal = vbOK Then
        s.strPROJECT_NAME = Trim("PROJECT_NAME")
        s.strAPP_NAME = Trim("APP_NAME")
        s.strOCP_TEMPLATE = Trim("OCP_TEMPLATE")
        
        s.file_gitlab_ci_yml c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub

