Attribute VB_Name = "dev_FilesTemplates"
Option Explicit

Sub tools_ClipBoard_Mardown()
Dim c As New Collection, s As New clsFilesTemplates
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Repositorio - Mardown")

    If retVal = vbOK Then
        s.markDown c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub

Sub tools_ClipBoard_FileSettings()
Dim c As New Collection, s As New clsFilesTemplates
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Repositorio - settingsXml")

    If retVal = vbOK Then
        s.strUserName = "T0204LTN"
        s.strPassword = "H4eJmQrY"
        s.strNonProxyHosts = "localhost|artifactory.visanet.corp|10.82.*|*ccorp.*"
        s.file_settingsXml c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub

Sub tools_ClipBoard_FileJs()
Dim c As New Collection, s As New clsFilesTemplates
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Repositorio - settingsXml")

    If retVal = vbOK Then
        s.file_js c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub
