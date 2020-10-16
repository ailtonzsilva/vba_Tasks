Attribute VB_Name = "dev_ClipBoard"
Option Explicit

Sub tools_ClipBoard_setJava(): Dim strTemp As String: strTemp = "set java_home=c:\Java\jdk1.8.0_161": ClipBoardThis strTemp: End Sub

Sub tools_ClipBoard_CopyEmailsReport()
Dim colNews As New Collection: Dim strTemp As String: Dim c As Variant: Dim s As New clsRepositorio: s.emailsReport colNews
    
    For Each c In colNews
        strTemp = strTemp + c + ";"
    Next

    ClipBoardThis Left(strTemp, Len(strTemp) - 1)
    
    MsgBox "Copiado!", vbOKOnly + vbInformation, "tools_ClipBoard_CopyEmailsReport"

End Sub

Sub tools_ClipBoard_CopyPathVariables()
Dim colNews As New Collection: Dim strTemp As String: Dim c As Variant: Dim s As New clsRepositorio: s.pathVariables colNews
    
    For Each c In colNews
        strTemp = strTemp + c
    Next

    ClipBoardThis strTemp
    
    MsgBox "Copiado!", vbOKOnly + vbInformation, "tools_ClipBoard_CopyPathVariables"

End Sub
