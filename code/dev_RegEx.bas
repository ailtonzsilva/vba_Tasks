Attribute VB_Name = "dev_RegEx"
Option Explicit

Sub tools_ClipBoard_RemoverLinhasDuplicadas()
'' #Links
'' https://qastack.com.br/programming/3958350/removing-duplicate-rows-in-notepad#:~:text=O%20Notepad%20%2B%2B%20pode%20fazer,linhas%20duplicadas%20ao%20mesmo%20tempo.&text=As%20caixas%20de%20sele%C3%A7%C3%A3o%20e,%3A%20TextFX%20%2D%3E%20TextFX%20Tools%20.&text=e%20substitua%20por%20nada%20.,a%20%C3%BAltima%20ocorr%C3%AAncia%20no%20arquivo.
''Removendo linhas duplicadas no Notepad

Dim c As New Collection: Dim s As New clsRegEx
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "RegEx - RemoverLinhasDuplicadas")

    If retVal = vbOK Then
        s.RemoverLinhasDuplicadas c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub


Sub tools_ClipBoard_VimTips()
'' #Links
'' https://qastack.com.br/programming/3958350/removing-duplicate-rows-in-notepad#:~:text=O%20Notepad%20%2B%2B%20pode%20fazer,linhas%20duplicadas%20ao%20mesmo%20tempo.&text=As%20caixas%20de%20sele%C3%A7%C3%A3o%20e,%3A%20TextFX%20%2D%3E%20TextFX%20Tools%20.&text=e%20substitua%20por%20nada%20.,a%20%C3%BAltima%20ocorr%C3%AAncia%20no%20arquivo.
''Removendo linhas duplicadas no Notepad

Dim c As New Collection: Dim s As New clsVim
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Vim - Tips")

    If retVal = vbOK Then
        s.Tips c: s.strFilePath = GetFolder()
        execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
    End If

End Sub

