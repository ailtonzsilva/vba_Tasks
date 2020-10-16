Attribute VB_Name = "tools_json"
Option Explicit

Sub tools_jsonContatos()

''' Criar arquivo .json com dados na guia contatos
'Dim jsonItems As New Collection: Dim jsonDictionary As New Dictionary: Dim jsonFileObject As New FileSystemObject: Dim jsonFileExport As TextStream
'
''' Worksheet
'Worksheets("DOC_contatos").Activate
'Dim ws As Worksheet: Set ws = Worksheets("DOC_contatos")
'Dim lRow As Long, x As Long, t As Variant, tmp As String: tmp = ""
'Dim ColumnIndex  As Integer: ColumnIndex = 2
'lRow = ws.Cells(Rows.Count, ColumnIndex).End(xlUp).Offset(1, 0).Row
'
''' Source
'Dim pathSource As String: pathSource = CreateObject("WScript.Shell").SpecialFolders("Desktop")
'Dim pathExit As String: pathExit = pathSource & "\" & ActiveSheet.Name & ".json"
'If (Dir(pathExit) <> "") Then kill pathExit
'
'    Dim i As Integer
'    For i = 2 To lRow - 1
'        jsonDictionary("area") = CStr(Trim(ws.Range("B" & i).Value))
'        jsonDictionary("nome") = CStr(Trim(ws.Range("C" & i).Value))
'        jsonDictionary("email") = CStr(Trim(ws.Range("D" & i).Value))
'        jsonDictionary("obs") = CStr(Trim(ws.Range("E" & i).Value))
'
'        jsonItems.add jsonDictionary
'        Set jsonDictionary = Nothing
'
'    Next i
'
'    Set jsonFileExport = jsonFileObject.CreateTextFile(pathExit, True)
'    jsonFileExport.WriteLine (JsonConverter.ConvertToJson(jsonItems, Whitespace:=3))

End Sub

Sub tools_jsonLinks()

''' Json
'Dim jsonItems As New Collection: Dim jsonDictionary As New Dictionary: Dim jsonFileObject As New FileSystemObject: Dim jsonFileExport As TextStream

''' Worksheet
'Worksheets("DOC_Links").Activate
'Dim ws As Worksheet: Set ws = Worksheets("DOC_Links")
'Dim lRow As Long: lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
'
''' Source
'Dim pathSource As String: pathSource = CreateObject("WScript.Shell").SpecialFolders("Desktop")
'Dim pathExit As String: pathExit = pathSource & "\" & ActiveSheet.Name & ".json"
'If (Dir(pathExit) <> "") Then kill pathExit
'
'    '' Build file text
'    Dim i As Integer: i = 2
'    Dim t As Variant
'
'    For Each t In ws.Range("C2:C" & lRow).SpecialCells(xlCellTypeVisible)
'        jsonDictionary("descricao") = CStr(Trim(ws.Range("E" & i).Value))
'        jsonDictionary("link") = CStr(Trim(ws.Range("C" & i).Value))
'        jsonItems.add jsonDictionary
'        Set jsonDictionary = Nothing
'        i = i + 1
'    Next
'
'    '' Build script
'    Set jsonFileExport = jsonFileObject.CreateTextFile(pathExit, True)
'    jsonFileExport.WriteLine (JsonConverter.ConvertToJson(jsonItems, Whitespace:=3))
'
'    Shell "notepad.exe " & pathExit, vbMaximizedFocus
    
End Sub
