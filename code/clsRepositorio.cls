VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepositorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strAppPath As String
    strFileName As String
    strFilePath As String
    strPassword As String
    strPathSourceApp As String
End Type
Private This As ClassType

Property Get strAppPath() As String:                   strAppPath = This.strAppPath:          End Property
Property Let strAppPath(pValue As String):             This.strAppPath = pValue:              End Property
Property Get strFileName() As String:                   strFileName = This.strFileName:       End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:            End Property
Property Get strFilePath() As String:                   strFilePath = This.strFilePath:       End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:            End Property

Property Get strPassword() As String:                   strPassword = This.strPassword:       End Property
Property Let strPassword(pValue As String):             This.strPassword = pValue:            End Property

Property Get strPathSourceApp() As String:                   strPathSourceApp = This.strPathSourceApp:       End Property
Property Let strPathSourceApp(pValue As String):             This.strPathSourceApp = pValue:            End Property


Public Function HideFile(ByRef pCol As Collection)

Dim sFileName As String: sFileName = Left(This.strFileName, Len(This.strFileName) - 4)

    pCol.add "REM"
    pCol.add "REM Excluir arquivos existentes"
    pCol.add "REM"
    pCol.add ""
    pCol.add "IF EXIST logo-" & sFileName & ".jpg DEL /F logo-" & sFileName & ".jpg"
    pCol.add "IF EXIST " & sFileName & ".7z DEL /F " & sFileName & ".7z"
    pCol.add ""
    pCol.add "REM"
    pCol.add "REM Compactar dados"
    pCol.add "REM"
    pCol.add ""
    pCol.add "" & strPathSourceApp & " a " & sFileName & ".7z " & This.strFileName & " -p""" & This.strPassword & """"
    pCol.add ""
    pCol.add "REM"
    pCol.add "REM Anexar arquivos"
    pCol.add "REM"
    pCol.add ""
    pCol.add "copy /b logo.jpg + " & sFileName & ".7z logo-" & sFileName & ".jpg"
    pCol.add "explorer ."
    
    This.strFileName = sFileName & ".bat"

End Function

Public Function Tips(ByRef pCol As Collection)

    pCol.add "copy /b mapa.jpg + slc-credit_v01.7z mapa1.jpg | Guardar arquivo dentro de imagem"
    pCol.add "for /r %i in (*) do @echo %i | For example, to find all files in and under the current directory, use:"
    pCol.add "dir /s /b > listarSubPastas.txt | listar SubPastas"
        
    If strFileName = "" Then This.strFileName = "tips.txt"

End Function

Public Function note_EtapasParaHML(ByRef pCol As Collection)

    pCol.add "1. DEV"
    pCol.add "2. Sonar"
    pCol.add "3. GOCD / Delivery"
    pCol.add "4. Sonartype"
    pCol.add "5. CI_CD / SCP_IAC_Deploy *"
    pCol.add "6. GOCD / Deploy"
    pCol.add "* Openshift - Deploy.yaml ( Alterar nome app )"
    
    If strFileName = "" Then This.strFileName = "EtapasParaHML.txt"
    
End Function

Public Function createFolders(ByRef pCol As Collection)
'' #Links
''Creating folders with for loop in windows batch
''https://stackoverflow.com/questions/15786082/creating-folders-with-for-loop-in-windows-batch

    pCol.add "FOR /L %%I IN (1,1,8) DO call md 20200%%I"
    
    If strFileName = "" Then This.strFileName = "createFolders.bat"

End Function


Public Function createFolderControl(ByRef pCol As Collection)

    pCol.add "set Today=%Date: =0%"
    pCol.add "set Year=%Today:~-4%"
    pCol.add "set Month=%Today:~-7,2%"
    pCol.add "set Day=%Today:~-10,2%"
    pCol.add "set hr=%TIME: =0%"
    pCol.add "set hr=%hr:~0,2%"
    pCol.add "set min=%TIME:~3,2%"
    pCol.add ""
    pCol.add "md %Year%%Month%%Day%-%hr%%min%"

     If strFileName = "" Then This.strFileName = "createFolderControl.bat"

End Function

Public Function createBkpFolders(ByRef pCol As Collection, pathSourceApp As String)

    pCol.add "set Today=%Date: =0%"
    pCol.add "set Year=%Today:~-4%"
    pCol.add "set Month=%Today:~-7,2%"
    pCol.add "set Day=%Today:~-10,2%"
    pCol.add "set sFolderBkp=""C:\Users\T0204LTN\OneDrive - Cielo\wspace\log\_bkp\"""
    pCol.add ""
    pCol.add "set hr=%TIME: =0%"
    pCol.add "set hr=%hr:~0,2%"
    pCol.add "set min=%TIME:~3,2%"
    pCol.add ""
    pCol.add "for /d %%X in (*) do " & pathSourceApp & " a %Year%%Month%%Day%-%hr%%min%_""%%X.7z"" ""%%X\"""
    pCol.add "move *.7z %sFolderBkp%"
    pCol.add "move *.Zip %sFolderBkp%"
    pCol.add ""
'    pCol.add "pause"

    If strFileName = "" Then This.strFileName = "createBkpFolders.bat"

End Function

Public Function createBkp(ByRef pCol As Collection, pathSourceApp As String)
    
    pCol.add "@echo off"
    pCol.add "SET Today=%Date: =0%"
    pCol.add "SET Year=%Today:~-4%"
    pCol.add "SET Month=%Today:~-7,2%"
    pCol.add "SET Day=%Today:~-10,2%"
    pCol.add "SET hr=%TIME: =0%"
    pCol.add "SET hr=%hr:~0,2%"
    pCol.add "SET min=%TIME:~3,2%"
    pCol.add "SET src_folder=%cd%"
    pCol.add "SET dst_folder=%cd%""\bkp"""
    
    pCol.add vbNewLine
    pCol.add "REM"
    pCol.add "REM Criar pasta de bkp em destino escolhido"
    pCol.add "REM"
    pCol.add "mkdir ""%dst_folder%"""
    
    pCol.add vbNewLine
    pCol.add "REM Nome deste script"
    pCol.add "REM SET sFolder=%~n0"
    pCol.add "SET sFolder=" & strFileName
    
    pCol.add vbNewLine
    pCol.add "REM pegar os ultimos 6 caracteres do nome"
    pCol.add "REM SET sFolder=%sFolder:~-6%"
    
    pCol.add vbNewLine
    pCol.add "REM arq tmp para uso do script"
    pCol.add "SET sFile=print.txt"
    
    pCol.add vbNewLine
    pCol.add "REM"
    pCol.add "REM Compactar arquivos em pasta selecionada"
    pCol.add "REM"
    pCol.add "CD ""%src_folder%"""
    pCol.add "for /d %%X in (""%sFolder%"") do " & pathSourceApp & " a %Year%%Month%%Day%-%hr%%min%_""%%X.7z"" ""%%X\"""
    
    pCol.add vbNewLine
    pCol.add "REM"
    pCol.add "REM Gerar arquivo lista para copia"
    pCol.add "REM"
    pCol.add "del ""%sFile%"""
    pCol.add "dir *.7z /b > ""%sFile%"""
    
    pCol.add vbNewLine
    pCol.add "TIMEOUT 3"
    
    pCol.add vbNewLine
    pCol.add "REM"
    pCol.add "REM Efetuar copia de arquivos em lista"
    pCol.add "REM"
    pCol.add "for /f ""tokens=*"" %%i in (print.txt) DO ("
    pCol.add "    xcopy /y ""%src_folder%\%%i"" ""%dst_folder%"""
    pCol.add "    del ""%src_folder%\%%i"""
    pCol.add ")"
    
    pCol.add vbNewLine
    pCol.add "del ""%sFile%"""
    
    pCol.add "pause"
    
    
    If strFileName = "" Then This.strFileName = "createBkp.bat"
    
End Function



Public Function Url(ByRef pCol As Collection)

    pCol.add "#include <File.au3>"
    pCol.add "$fileSource = FileOpen(@ScriptDir & ""\"" & StringLeft(@ScriptName,StringLen(@ScriptName)-4) & "".txt"",0)"
    pCol.add "#Region PreSets"
    pCol.add "    HotKeySet(""^{END}"", ""CaptureEND"")"
    pCol.add "    Opt(""GUIOnEventMode"",1)"
    pCol.add "#EndRegion"
    pCol.add "If $fileSource = -1 Then"
    pCol.add "    MsgBox(0, ""Error"", ""Unable to open file."")"
    pCol.add "    Exit"
    pCol.add "EndIf"
    pCol.add "While 1"
    pCol.add "   $line = FileReadLine($fileSource)"
    pCol.add "   If @error = -1 then ExitLoop"
    pCol.add "   ConsoleWrite($line & @LF)"
    pCol.add "   ShellExecute($line)"
    pCol.add "WEnd"
    pCol.add "FileClose($fileSource)"
    pCol.add "Func CaptureEND()"
    pCol.add " Switch @HotKeyPressed "
    pCol.add "       Case ""^{end}"" "
    pCol.add "          Exit"
    pCol.add " EndSwitch"
    pCol.add "EndFunc"

    If strFileName = "" Then This.strFileName = "url.au3"

End Function

Public Function mouse(ByRef pCol As Collection)

    pCol.add "#include <Timers.au3>"
    pCol.add "#Region PreSets"
    pCol.add "   HotKeySet(""^{END}"", ""CaptureEND"")"
    pCol.add "   Opt(""GUIOnEventMode"",1)"
    pCol.add "#EndRegion"
    pCol.add "While true"
    pCol.add "   MouseMove(10,10,1)"
    pCol.add "   MouseMove(20,10,1)"
    pCol.add "   MouseMove(30,20,1)"
    pCol.add "   MouseMove(40,20,1)"
    pCol.add "WEnd"
    pCol.add "Func CaptureEND()"
    pCol.add "   Switch @HotKeyPressed"
    pCol.add "       Case ""^{end}"""
    pCol.add "       Exit"
    pCol.add "   EndSwitch"
    pCol.add "EndFunc"
    
    If strFileName = "" Then This.strFileName = "mouse.au3"
    
End Function

Public Function createVbs_limparBase(ByRef pCol As Collection)

    pCol.add "'=========================================================="
    pCol.add "' PROPOSITO:: Excluir guias ocultas de planilhas"
    pCol.add "' To-Do: "
    pCol.add "' - Filtrar apenas arquivos excel"
    pCol.add "'=========================================================="
    pCol.add ""
    pCol.add "Dim oFolder, oFile, app, wb"
    pCol.add ""
    pCol.add "Dim oFSO : Set oFSO = CreateObject(""Scripting.FileSystemObject"")"
    pCol.add "Dim WshShell : Set WshShell = WScript.CreateObject(""WScript.Shell"")"
    pCol.add "Dim sScriptDir : sScriptDir = oFSO.GetParentFolderName(WScript.ScriptFullName)"
    pCol.add ""
    pCol.add "' Criar repositorio para limpesa"
    pCol.add "Dim path : path = WshShell.CurrentDirectory & ""\limparBase\"""
    pCol.add ""
    pCol.add "If oFSO.FolderExists(path) = False Then oFSO.CreateFolder(path)"
    pCol.add ""
    pCol.add "Set oFolder = oFSO.GetFolder( path )"
    pCol.add ""
    pCol.add "For Each oFile In oFolder.Files"
    pCol.add "                set app = createobject(""Excel.Application"")"
    pCol.add "                set wb = app.workbooks.open(oFile.Path)"
    pCol.add "                For Each ws In app.Worksheets"
    pCol.add "                 If ws.Visible = xlSheetHidden Then"
    pCol.add "                      msgbox ws.name"
    pCol.add "                      ws.Delete"
    pCol.add "                 End If"
    pCol.add "                Next      "
    pCol.add "                wb.Save"
    pCol.add "                app.Quit "
    pCol.add "Next"
    pCol.add ""
    pCol.add "MsgBox(""Concluido"")"
    
    If strFileName = "" Then This.strFileName = "limparBase.vbs"

End Function

'Sub tools_ClipBoard_BatTips()
'
'Dim c As New Collection, s As New clsRepositorio
'Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Repositorio - Tips")
'
'If retVal = vbOK Then
'    s.Tips c: s.strFilePath = ThisWorkbook.Path & "\"
'    execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
'End If
'
'End Sub

