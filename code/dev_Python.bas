Attribute VB_Name = "dev_Python"
Option Explicit

Sub teste_parseJson()
Dim c As New Collection: Dim s As New clsPython: s.parseJson c: s.strFilePath = GetFolder()
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Python tips")

    If retVal = vbOK Then execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath

End Sub
