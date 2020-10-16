Attribute VB_Name = "dev_Java"
Option Explicit

Sub teste_buscaPrateleira()

Dim c As New Collection
Dim s As New clsJava: s.buscaPrateleira c: s.strFilePath = GetFolder()
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Java buscaPrateleira")

    If retVal = vbOK Then execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath

End Sub


Sub teste_SalesCancelContantsEnum()

Dim c As New Collection
Dim s As New clsJava: s.SalesCancelContantsEnum c: s.strFilePath = GetFolder()
Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Java SalesCancelContantsEnum")

    If retVal = vbOK Then execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath

End Sub
