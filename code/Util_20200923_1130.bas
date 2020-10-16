Attribute VB_Name = "Util_20200923_1130"

''[Links para pesquisa nesse modulo]
''#Dica
''#run
''#tools
''#repositorio
''#pathSource
''#tests
''#customUI
''
''#Links
''https://alvinalexander.com/blog/post/linux-unix/how-set-vim-gvim-default-color-scheme/


''=====================================================================================================================
''#Dica
''###[ imageMSO ]
'' https://bert-toolkit.com/imagemso-list.html
''
''###[ Create, list or delete stored user names, passwords or credentials. ]
'' https://ss64.com/nt/cmdkey.html
'' https://docs.microsoft.com/pt-br/windows-server/administration/windows-commands/mstsc
'' https://docs.microsoft.com/pt-br/windows-server/administration/windows-commands/remote-desktop-services-terminal-services-command-reference
''=====================================================================================================================

''#Dica
'' return name the sheet
'ThisWorkbook.Sheets(ActiveSheet.Name).Name
''=============================================

''#Dica M�dulo com fun��es �teis e Facilitadores
'Necess�rio para utiliza��o:
'Refer�ncia para 'Microsoft Forms 2.0 Object Library'
'Refer�ncia para 'Microsoft Windows Common Controls 6.0 (SP6)'

'=====================================================================================================================
' Module    : Util_v3
' Author    : Sidnei Graciolli
' Date      : 12/05/2016
' Purpose   : Cont�m fun��es �teis e facilitadores
'=====================================================================================================================


'Public Function ListView_SetHeader(ByVal oObjListView As ListView, ByVal aTitulos As Variant)
''=====================================================================================================================
'' Procedure    : ListView_SetHeader
'' Author       : s.graziollijunior (MAI/2016)
'' Type         : Public Function
'' Return       : Void
'' Description  :
'' Params
''   - ByVal oObjListView As ListView
''   - ByVal aTitulos As Variant
''=====================================================================================================================
'    Dim j As Integer
'
'    oObjListView.HideColumnHeaders = False
'    oObjListView.View = lvwReport
'    oObjListView.ColumnHeaders.Clear
'
'    For j = 0 To UBound(aTitulos) Step 2
'        oObjListView.ColumnHeaders.Add , , aTitulos(j), aTitulos(j + 1)
'    Next j
'End Function
'Public Function ListView_Populate(oObjListView As ListView, ByVal aDados As Variant)
''=====================================================================================================================
'' Procedure    : ListView_Populate
'' Author       : s.graziollijunior (MAI/2016)
'' Type         : Public Function
'' Return       : Void
'' Description  :
'' Params
''   - oObjListView As ListView
''   - ByVal aDados As Variant
''=====================================================================================================================
'    Dim j As Integer
'    Dim nPOS As Integer
'
'    oObjListView.ListItems.Add , , aDados(0)  ' ID
'    nPOS = oObjListView.ListItems.Count
'
'    For j = 1 To UBound(aDados)
'        oObjListView.ListItems(nPOS).ListSubItems.Add , , aDados(j)
'    Next j
'End Function

' FUNCTION DECLARATIONS ==============================================================================================
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
(ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


'Windows API calls to handle windows
#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long:
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

' ENUMERATORS ========================================================================================================
Public Enum enumState
    sttNone = 0
    sttInsert = 1
    sttOk = 2
    sttDelete = 3
    sttSearch = 4
    sttExport = 5
    sttCancel = 6
    sttClear = 7
    sttUpdate = 8
End Enum
Public Enum enumLanguage
    langUS_English = 0
    langBR_Portuguese = 1
End Enum
Public Enum enTipo
    tpInt = 0
    tpStr = 1
    tpBool = 2
    tpDate = 3
    tpFK = 4
    tpFloat = 5
    tpTime = 6
End Enum
Public Enum enDriverODBC
    drAccess = 0
    drSqlServer = 1
    drExcel_12 = 2
    drExcel_8 = 3
    drSqlite = 4
End Enum


Public Enum enumOperacao
    opNome = 0
    opExecutar = 1
    opNotepad = 2
End Enum

' VARIABLES DECLARATIONS =============================================================================================
'Public clsUser                              As New User
Public bolRecall                            As Boolean
Private Const mcGWL_STYLE = (-16)
Private Const mcWS_SYSMENU = &H80000
Public Const MASTER_USER                    As String = "fp5738" '"xn4718"
Public boolErrorHandler                     As Boolean
Public boolAppendFile                       As Boolean
Public Const VK_CONTROL                     As Integer = &H11
Public Const NOME_PROJETO                   As String = "PROJETO"
Public Const NOME_EMPRESA                   As String = "EMPRESA"
Public Const TITULO_MSG                     As String = NOME_PROJETO & " :: " & NOME_EMPRESA
Public Const FORMAT_CURRENCY                As String = "#,##0.00;(#,##0.00);-"
Public Const NUMBER_FORMAT                  As String = "#,##0.00;(#,##0.00);0.00"
Public Const DATE_FORMAT                    As String = "dd/mm/yy"
Public sql                                  As String
Public i                                    As Integer
Public J                                    As Integer

''#pathSource
'Private Const pathSourceApp As String = ""
'Private Const pathSource As String = "C:\Users\T0204LTN\OneDrive - Cielo\wspace\log\_temas\"
Function pathSource(): pathSource = ThisWorkbook.Path & "\": End Function

' VARIABLES DECLARATIONS =============================================================================================

Public Function MonthName(pMonthNumber As Integer, Optional pLanguage As enumLanguage) As String
'=====================================================================================================================
' Procedure    : MonthName
' Author       : s.graziollijunior (MAI/2012)
' Type         : Public Function
' Return       : String
' Description  : This function take an integer and returns the name of the corresponding month. Optionally, you can pass the language for return.
' Params
'   - pMonthNumber As Integer
'   - Optional pLanguage As enumLanguage
'=====================================================================================================================
    Dim strMonths As String
    
    If pMonthNumber < 1 Or pMonthNumber > 12 Then
        MonthName = "#InvalidParameter"
    Else
        Select Case pLanguage
            Case langUS_English:        strMonths = " January February March April May June July August September October November December"
            Case langBR_Portuguese:     strMonths = " Janeiro Fevereiro Mar�o Abril Maio Junho Julho Agosto Setembro Outubro Novembro Dezembro"
            Case Else:                  strMonths = " January February March April May June July August September October November December"
        End Select
        
        MonthName = Split(strMonths, " ")(pMonthNumber)
    End If
End Function
Public Function MonthNumber(pMonth As String, Optional pLanguage As enumLanguage) As String
'=====================================================================================================================
' Procedure    : MonthNumber
' Author       : s.graziollijunior (MAI/2012)
' Type         : Public Function
' Return       : String
' Description  : This function take the month name and returns a string of the integer value. Optionally, you can pass the language of the entry.
' Params
'   - pMonth As String
'   - Optional pLanguage As enumLanguage
'=====================================================================================================================
    For i = 1 To 12
        If UCase(pMonth) = UCase(MonthName(i, pLanguage)) Then
            MonthNumber = CStr(Format(i, "00"))
            Exit Function
        End If
    Next i
    
    MonthNumber = "#InvalidParameter"
End Function
Public Function MonthComboBox(pComboBox As ComboBox, Optional pSelected As Integer, Optional pLanguage As enumLanguage)
'=====================================================================================================================
' Procedure    : MonthComboBox
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  : This function populate a ComboBox object with months. Optionally, you can pass the language for return, and the month that will stay selected.
' Params
'   - pComboBox As ComboBox
'   - Optional pSelected As Integer
'   - Optional pLanguage As enumLanguage
'=====================================================================================================================
    With pComboBox
        .Clear
        For i = 1 To 12
            .AddItem MonthName(i, pLanguage)
        Next i
        If pSelected <> 0 Then
            .ListIndex = pSelected
        Else
            .Text = "Select..."
        End If
    End With
End Function
Public Function YearComboBox(pComboBox As ComboBox)
'=====================================================================================================================
' Procedure    : YearComboBox
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  : This function populate a ComboBox object with years.
' Params
'   - pComboBox As ComboBox
'=====================================================================================================================
    With pComboBox
        .Clear
        For i = 1 To 5
            .AddItem (Year(Now) - 3) + i
        Next i
        .ListIndex = 2
    End With
End Function
Public Function ClearStatusBar()
'=====================================================================================================================
' Procedure    : ClearStatusBar
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
'=====================================================================================================================
    ThisWorkbook.Application.StatusBar = vbNullString
End Function
Public Function statusFinal(pDate As Date)
'=====================================================================================================================
' Procedure    : statusFinal
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pDate As Date
'=====================================================================================================================
    ThisWorkbook.Application.StatusBar = "Opera��o conclu�da! Tempo de execu��o: " & Format(Now - pDate, "hh:mm:ss")
    'ShowProgress ThisWorkbook.Application.StatusBar
    ThisWorkbook.Application.OnTime Now() + TimeValue("00:00:15"), "ClearStatusBar"
End Function
Public Function Greet() As String
'=====================================================================================================================
' Procedure    : Greet
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
'=====================================================================================================================
    Select Case Hour(Now)
        Case 6 To 11:               Greet = "Bom dia"
        Case 12 To 17:              Greet = "Boa tarde"
        Case Else:                  Greet = "Boa noite"
    End Select
End Function
Public Function Contains(pCollection As Collection, pKey As Variant) As Boolean
'=====================================================================================================================
' Procedure    : Contains
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Boolean
' Description  :
' Params
'   - pCollection As Collection
'   - pKey As Variant
'=====================================================================================================================
    On Error GoTo NoSuchKey
    If VarType(pCollection.Item(pKey)) = vbObject Then
    End If
    Contains = True
    Exit Function
NoSuchKey:
    Contains = False
End Function
Public Function RaiseError(Optional pSource As String, Optional pDescription As String)
'=====================================================================================================================
' Procedure    : RaiseError
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - Optional pSource As String
'   - Optional pDescription As String
'=====================================================================================================================
    ThisWorkbook.Application.StatusBar = ""
    With err
        If pDescription = vbNullString Then
            .Description = "Ocorreu um erro de sistema: " & .Description
        Else
            .Description = pDescription
        End If
        If pSource <> vbNullString Then .Source = pSource
        'Debug.Print CStr(Time()) + "@" + .Source + " >> " + .Description
        .Raise vbObjectError + 513
    End With
End Function
Public Function ShowError()
'=====================================================================================================================
' Procedure    : ShowError
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
'=====================================================================================================================
    Dim FileNumber As Integer
    Dim strLog As String
    Dim strSource As String
    Dim strDescription As String
    
    ThisWorkbook.Application.StatusBar = ""
    
    With err
        strSource = .Source
        strDescription = .Description
        
'        If clsBanco.possuiAcesso Then
'            strLog = Format(Now(), "dd_mm_yyyy HH:mm:ss") & " --> "
'            strLog = strLog & Left("@" & strSource & space(25), 25)
'            strLog = strLog & Left("ERRO: " & strDescription & space(100), 100)
'            strLog = strLog & Left("USER: " & Environ("UserName") & space(30), 30)
'            strLog = strLog & Left("M�DULO: " & NOME_PROJETO & space(30), 30)
'
'            FileNumber = FreeFile                                                ' Get unused file number
'            Open ThisWorkbook.Path & "\LOG_" & UCase(TITULO_MSG) & ".txt" For Append As #FileNumber    ' Connect to the file
'            Print #FileNumber, strLog: Close #FileNumber                         ' Append our string
'
'            If MsgBox("Ocorreu um erro inesperado. Detalhes do erro:" & vbNewLine & vbNewLine & _
'                    vbTab & "-" & strSource & vbNewLine & vbTab & "-" & strDescription & vbNewLine & vbNewLine & _
'                    "Um log do erro foi criado em " & ThisWorkbook.Path & "\LOG_" & UCase(TITULO_MSG) & ".txt" & vbNewLine & vbNewLine & _
'                    "Deseja abr�-lo?", vbYesNo + vbDefaultButton2, TITULO_MSG) = vbYes Then
'
'                Shell "c:\WINDOWS\notepad.exe " & ThisWorkbook.Path & "\LOG_" & UCase(TITULO_MSG) & ".txt"
'            End If
'        Else
            MsgBox "Ocorreu um erro inesperado. Detalhes do erro:" & vbNewLine & vbNewLine & _
                    vbTab & "-" & strSource & vbNewLine & vbTab & "-" & strDescription '& vbNewLine & vbNewLine & _
                    '"Um log do erro n�o pode ser criado.", vbOKOnly, TITULO_MSG
'        End If
    End With
End Function
Public Function PickFolder(pPath As String, pTitle As String, Optional pSubFolder As String) As String
'=====================================================================================================================
' Procedure    : PickFolder
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pPath As String
'   - pTitle As String
'   - Optional pSubFolder As String
'=====================================================================================================================
    With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = pTitle
    .InitialFileName = pPath
    .Show
        If .SelectedItems.Count > 0 Then
            If Trim(pSubFolder) <> "" Then
                If Dir(.SelectedItems(1) & pSubFolder, vbDirectory) = "" Then
                    MkDir Path:=.SelectedItems(1) & pSubFolder
                End If
            End If
            PickFolder = .SelectedItems(1) & pSubFolder
        End If
    End With
End Function
Public Function PickFiles(pTitle As String, Optional pFilter As Variant, Optional pMultiselect As Boolean) As Collection
'=====================================================================================================================
' Procedure    : PickFiles
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Collection
' Description  :
' Params
'   - pTitle As String
'   - Optional pFilter As Variant
'   - Optional pMultiselect As Boolean
'=====================================================================================================================
    Dim iCol As New Collection
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = pTitle
        .InitialFileName = ThisWorkbook.Path
        .AllowMultiSelect = pMultiselect
        .Filters.Clear
        
        If Not IsMissing(pFilter) Then
            For IFilter = LBound(pFilter) To UBound(pFilter) Step 2
                .Filters.add pFilter(IFilter), pFilter(IFilter + 1)
            Next IFilter
        End If
        
        .Show
        
        For Each Sel In .SelectedItems
            iCol.add Sel
        Next
    End With
    
    Set PickFiles = iCol
End Function
Public Function GetFilesInFolder(pFolder As String) As Collection
'=====================================================================================================================
' Procedure    : GetFilesInFolder
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Collection
' Description  :
' Params
'   - pFolder As String
'=====================================================================================================================
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim iCol As New Collection
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(pFolder)
    
    For Each objFile In objFolder.Files
        iCol.add objFile.Path '& "\" & objFile.Name
    Next objFile
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    
    Set GetFilesInFolder = iCol
End Function
Public Function AppendCollectionFilesInFolder(ByRef pCol As Collection, pFolder As String)
'=====================================================================================================================
' Procedure    : AppendCollectionFilesInFolder
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - ByRef pCol As Collection
'   - pFolder As String
'=====================================================================================================================
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(pFolder)
    
    For Each objFile In objFolder.Files
        If Not Contains(pCol, objFile.Name) Then
            pCol.add objFile.Path, objFile.Name
        End If
    Next objFile
    
    Set objFSO = Nothing
    Set objFolder = Nothing

End Function
Public Function DateToSql(pDateValue As Date) As String
'=====================================================================================================================
' Procedure    : DateToSql
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pDateValue As Date
'=====================================================================================================================
    DateToSql = "'" & Year(pDateValue) & "-" & Month(pDateValue) & "-" & Day(pDateValue) & "'"
End Function
Public Function TimeToSql(pValue As Date) As String
'=====================================================================================================================
' Procedure    : TimeToSql
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pValue As Date
'=====================================================================================================================
    TimeToSql = "'" & Hour(pValue) & ":" & Minute(pValue) & ":" & Second(pValue) & "'"
End Function
Public Function StringToSql(pText As String, Optional pMaxLenght As Integer) As String
'=====================================================================================================================
' Procedure    : StringToSql
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pText As String
'   - Optional pMaxLenght As Integer
'=====================================================================================================================
    Dim iStrAux As String
    iStrAux = Replace(pText, Chr(39), vbNullString)         ' Remove aspa simples: '
    iStrAux = Replace(iStrAux, Chr(145), vbNullString)      ' Remove aspa simples: �
    iStrAux = Replace(iStrAux, Chr(146), vbNullString)      ' Remove aspa simples: �
    
    If pMaxLenght > 0 Then iStrAux = Left(iStrAux, pMaxLenght)
    
    If iStrAux = "" Then
        StringToSql = "NULL"
    Else
        StringToSql = "'" & iStrAux & "'"
    End If
End Function
Public Function NumberToSql(pNumber As Variant) As String
'=====================================================================================================================
' Procedure    : NumberToSql
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pNumber As Variant
'=====================================================================================================================
    Dim i As Integer
    Dim iStrSinal As String
    Dim iStrAux As String
    Dim iCaracteres As String
    iCaracteres = "1234567890,.-"
    iStrAux = vbNullString
    iStrSinal = vbNullString
    
    pNumber = CStr(pNumber)
    
    If pNumber = vbNullString Then
        NumberToSql = "NULL"
    Else
        For i = 1 To Len(pNumber)
            If Mid(pNumber, i, 1) = "(" Then iStrSinal = "-"
            If InStr(iCaracteres, Mid(pNumber, i, 1)) <> 0 Then
                iStrAux = iStrAux & Mid(pNumber, i, 1)
            End If
        Next i
        iStrAux = Replace(iStrAux, Chr(46), vbNullString)       ' Remove pontos
        iStrAux = Replace(iStrAux, Chr(44), Chr(46))            ' Substitui v�rgula por ponto
        NumberToSql = iStrSinal & iStrAux
        If Right(NumberToSql, 1) = "." Then NumberToSql = Left(NumberToSql, Len(NumberToSql) - 1)

        If Not IsNumeric(NumberToSql) Then
            Debug.Print Now & " -:- ERRO Function NumberToSql -:- Imposs�vel formatar " & pNumber & ". Resultado n�o � um n�mero v�lido: " & NumberToSql
            NumberToSql = "0"
        End If
    End If
End Function
Public Function ValidateNumberField(ByRef pTextBox As control, ByVal pKeyAscii As MSForms.ReturnInteger)
'=====================================================================================================================
' Procedure    : ValidateNumberField
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - ByRef pTextBox As control
'   - ByVal pKeyAscii As MSForms.ReturnInteger
'=====================================================================================================================
    With pTextBox
        Select Case pKeyAscii
            Case Asc("0") To Asc("99")
            Case Asc(","): pKeyAscii = 0
                If InStr(1, .Text, ",") <= 0 Then
                    If .Text = vbNullString Then
                        .Text = "0,"
                    Else
                        .Text = .Text & ","
                    End If
                End If
            Case Asc("-"): pKeyAscii = 0
                If (Left(.Text, 1)) <> "-" Then .Text = "-" & .Text
            Case Asc("+"): pKeyAscii = 0
                If (Left(.Text, 1)) = "-" Then .Text = Right(.Text, Len(.Text) - 1)
            Case Else: pKeyAscii = 0
        End Select
    End With
End Function
Public Function ValidateIntegerField(ByRef pTextBox As control, ByVal pKeyAscii As MSForms.ReturnInteger)
'=====================================================================================================================
' Procedure    : ValidateIntegerField
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - ByRef pTextBox As control
'   - ByVal pKeyAscii As MSForms.ReturnInteger
'=====================================================================================================================
    With pTextBox
        Select Case pKeyAscii
            Case Asc("0") To Asc("99")
            Case Else: pKeyAscii = 0
        End Select
    End With
End Function
Public Function TextFile_Write(pFilePath As String, pText As String) As Boolean
'=====================================================================================================================
' Procedure    : WriteTextFile
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Boolean
' Description  :
' Params
'   - pFilePath As String
'   - pText As String
'=====================================================================================================================
    Dim intFNumber As Integer
    intFNumber = FreeFile
    
    If boolErrorHandler Then On Error GoTo ErrorHandler
    
    If boolAppendFile Then
        Open pFilePath For Output As #intFNumber
        Print #intFNumber, pText
        Close #intFNumber
    End If
    
    On Error GoTo 0
    TextFile_Write = True
    Exit Function
ErrorHandler:
    MsgBox "N�o foi poss�vel salvar o arquivo." & vbNewLine & "Verifique o caminho informado e as permiss�es de acesso", vbInformation
End Function
Public Function TextFile_Append(pFilePath As String, pText As String) As Boolean
'=====================================================================================================================
' Procedure    : AppendTextFile
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Boolean
' Description  :
' Params
'   - pFilePath As String
'   - pText As String
'=====================================================================================================================
    Dim intFNumber As Integer
    intFNumber = FreeFile
    
    If boolErrorHandler Then On Error GoTo ErrorHandler
    
    Open pFilePath For Append As #intFNumber
    Print #intFNumber, pText
    Close #intFNumber
    
    On Error GoTo 0
    
    TextFile_Append = True
    Exit Function
ErrorHandler:
    MsgBox "N�o foi poss�vel salvar o arquivo." & vbNewLine & "Verifique o caminho informado e as permiss�es de acesso", vbInformation
End Function
Public Function LockScreen(pLock As Boolean)
'=====================================================================================================================
' Procedure    : LockScreen
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pLock As Boolean
'=====================================================================================================================
    With Application
        .DisplayAlerts = Not (boolErrorHandler And pLock)
        .EnableEvents = Not (boolErrorHandler And pLock)
        .ScreenUpdating = Not (boolErrorHandler And pLock)
        .Cursor = IIf((boolErrorHandler And pLock), xlWait, xlDefault)
    End With
End Function
Public Function ShowProgress(pText As String, Optional pLen As Double, Optional ptextAux As String, Optional pChar As String)
'=====================================================================================================================
' Procedure    : ShowProgress
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pText As String
'   - Optional pLen As Double
'   - Optional ptextAux As String
'   - Optional pChar As String
'=====================================================================================================================
    Dim strAux As String
    strAux = pText
    If Trim(ptextAux) <> "" Then
        strAux = strAux & " " & WorksheetFunction.Rept(IIf(pChar = "", " ", pChar), pLen - Len(strAux)) & " " & ptextAux
    End If
    Application.StatusBar = Left("Atualizando... " & strAux, 255)
    Debug.Print strAux
End Function
Public Function LogError()
'=====================================================================================================================
' Procedure    : LogError
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
'=====================================================================================================================
    Debug.Print Format(Now(), "hh:MM:ss") & "@" & err.Source & " - " & err.Description
End Function
Public Function ClipBoardThis(pText As String)
'=====================================================================================================================
' Procedure    : ClipBoardThis
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pText As String
'=====================================================================================================================
    Dim objData As New MSForms.DataObject
    objData.SetText pText
    objData.PutInClipboard
End Function
Public Function ColumnLetter(pNumber As Long) As String
'=====================================================================================================================
' Procedure    : ColumnLetter
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  : Created by brettdj (Melbourne, Australia) Oct/2012 at http://stackoverflow.com/questions/12796973/vba-function-to-convert-column-number-to-letter
'                Updated by Sidnei A. Graciolli Jr. (SP, Brazil) Nov/2014
' Params
'   - pNumber As Long
'=====================================================================================================================
    If pNumber > Columns.Count Then pNumber = pNumber Mod Columns.Count
    ColumnLetter = Split(Cells(1, pNumber).address(True, False), "$")(0)
End Function
Public Function SetErrorHandler(pRotina As String)
'=====================================================================================================================
' Procedure    : SetErrorHandler
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pRotina As String
'=====================================================================================================================
    boolErrorHandler = (GetKeyState(VK_CONTROL) >= 0)
    If Not boolErrorHandler Then boolErrorHandler = MsgBox("Aten��o! O Sistema identificou que a tecla CTRL estava pressionada ao chamar a rotina " & pRotina & "." & vbNewLine & vbNewLine & _
                                                           "Esta op��o desabilita o tratamento de erros para debbugging e aumenta consideravelmente o tempo de processamento." & vbNewLine & vbNewLine & _
                                                           "A op��o de debugging deve ser utilizada apenas para an�lise do c�digo VBA. Deseja continuar? (Escolha N�O para desabilitar o modo de debugging)", vbYesNo + vbDefaultButton2 + vbExclamation, "Debugging mode") = vbNo
End Function
Public Function WorkbookIsOpen(pName As String) As Boolean
'=====================================================================================================================
' Procedure    : WorkbookIsOpen
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Boolean
' Description  :
' Params
'   - pName As String
'=====================================================================================================================
    Dim iWb As Workbook
    For Each iWb In Application.Workbooks
        If iWb.Name = pName Then
            WorkbookIsOpen = True
            Exit For
        End If
    Next
End Function
Public Function IsMasterUser() As Boolean
'=====================================================================================================================
' Procedure    : IsMasterUser
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Boolean
' Description  :
'=====================================================================================================================
    If (MASTER_USER = Environ("UserName")) Then
        'MsgBox "MASTER_USER Verified!", vbOKOnly + vbExclamation, TITULO_MSG
        IsMasterUser = True
    Else
        IsMasterUser = False
    End If
End Function
Public Function GetFileNameByPath(pFileFullPath As String) As String
'=====================================================================================================================
' Procedure    : GetFileNameByPath
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pFileFullPath As String
'=====================================================================================================================
    GetFileNameByPath = StrReverse(Split(StrReverse(pFileFullPath), "\")(0))
End Function
Public Function TextFile_GetRowCount(pFileFullPath As String) As Long
'=====================================================================================================================
' Procedure    : GetTxtFileRowCount
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Long
' Description  :
' Params
'   - pFileFullPath As String
'=====================================================================================================================
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set theFile = FSO.OpenTextFile(pFileFullPath, 8, True)
    TextFile_GetRowCount = theFile.Line
    Set oFSO = Nothing
    theFile.Close
    Set theFile = Nothing
End Function
Public Function ClearCacheMemory()
'=====================================================================================================================
' Procedure    : ClearCacheMemory
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
'=====================================================================================================================
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    
    wsh.Run "cmd.exe /S /C %windir%\system32\rundll32.exe advapi32.dll,ProcessIdleTasks", windowStyle, waitOnReturn
End Function
Public Function EstimatedTime(pStart As Date, pPosition As Long, pTotal As Long) As Date
'=====================================================================================================================
' Procedure    : EstimatedTime
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Date
' Description  :
' Params
'   - pStart As Date
'   - pPosition As Long
'   - pTotal As Long
'=====================================================================================================================
    Dim iTime As Date
    Dim iFracTime As Date
    
    iTime = Now - pStart
    iFracTime = iTime / pPosition
    
    EstimatedTime = Format(iFracTime * (pTotal - pPosition), "dd/mm/yy hh:mm:ss")
    
End Function
Public Function PreventNullString(pText As Variant) As String
'=====================================================================================================================
' Procedure    : PreventNullString
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pText As Variant
'=====================================================================================================================
    If IsNull(pText) Then
        PreventNullString = ""
    Else
        PreventNullString = CStr(pText)
    End If
End Function
Public Sub SelectAllText(ByRef pTextBox As MSForms.TextBox)
'=====================================================================================================================
' Procedure    : SelectAllText
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Sub
' Return       : Void
' Description  :
' Params
'   - ByRef pTextBox As MSForms.TextBox
'=====================================================================================================================
    pTextBox.SelStart = 0
    pTextBox.SelLength = Len(pTextBox.Text)
End Sub
Public Function KillIfExists(pFileFullName As String)
'=====================================================================================================================
' Procedure    : KillIfExists
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pFileFullName As String
'=====================================================================================================================
    If Dir(pFileFullName) <> vbNullString Then kill pFileFullName
End Function
Public Function CreateBatch(pLines As Variant, Optional pAutoRun As Boolean = True, Optional pAutoDestroy As Boolean = True, Optional pEchoOff As Boolean = True)
'=====================================================================================================================
' Procedure    : CreateBatch
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pLines As Variant
'   - Optional pAutoRun As Boolean = True
'   - Optional pAutoDestroy As Boolean = True
'   - Optional pEchoOff As Boolean = True
'=====================================================================================================================
    Dim batchFullName As String
    Dim s As Integer
    
    If ThisWorkbook.Saved Then
        batchFullName = Environ("TEMP") & "\" & IIf(ThisWorkbook.Saved, UCase(ClearString(CStr(Split(ThisWorkbook.Name, ".")(0)))), "TEMP_BATCH_FILE") & ".bat"
    Else
    
    End If
    
    KillIfExists batchFullName
    
    Sleep 1000
    
    Do While Dir(batchFullName) = vbNullString
        If pEchoOff Then TextFile_Append batchFullName, "@ECHO OFF"
        
        For s = LBound(pLines) To UBound(pLines)
            TextFile_Append batchFullName, CStr(pLines(s))
        Next s
        
        If pAutoDestroy Then
            TextFile_Append batchFullName, "TIMEOUT 3"
            TextFile_Append batchFullName, "DEL /Q /F " & batchFullName
        End If
        
        Sleep 500
    Loop
    
    If pAutoRun Then Shell batchFullName, vbMinimizedNoFocus
End Function
Public Function ClearString(pText As String) As String
'=====================================================================================================================
' Procedure    : ClearString
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : String
' Description  :
' Params
'   - pText As String
'=====================================================================================================================
    codiA = "�����������������������������������������������,. /\?!"
    codiB = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN_______"
    Temp = pText
    For J = 1 To Len(Temp)
        p = InStr(codiA, Mid(Temp, J, 1))
        If p > 0 Then Mid(Temp, J, 1) = Mid(codiB, p, 1)
    Next
    ClearString = Temp
End Function
Public Function RemoveCloseButton(pUserForm As Object)
'=====================================================================================================================
' Procedure    : RemoveCloseButton
' Author       : s.graziollijunior (MAI/2016)
' Type         : Public Function
' Return       : Void
' Description  :
' Params
'   - pUserForm As Object
'=====================================================================================================================
    Dim lngStyle As Long
    Dim lngHWnd As Long

    lngHWnd = FindWindow(vbNullString, pUserForm.Caption)
    lngStyle = GetWindowLong(lngHWnd, mcGWL_STYLE)

    If lngStyle And mcWS_SYSMENU > 0 Then
        SetWindowLong lngHWnd, mcGWL_STYLE, (lngStyle And Not mcWS_SYSMENU)
    End If
End Function
Public Function TextFile_ReadToCollection(pFilePath As String) As Collection
'=====================================================================================================================
' Procedure    : ReadTextFile
' Author       : s.graziollijunior (AGO/2016)
' Type         : Public Function
' Return       : Collection
' Description  :
' Params
'   - pFilePath As String
'=====================================================================================================================
    Dim colResult As New Collection
    Dim intFNumber As Integer
    Dim strTextLine As String
    
    intFNumber = FreeFile
    
    Open pFilePath For Input As #intFNumber
    Do Until EOF(intFNumber)
        Line Input #intFNumber, strTextLine
        colResult.add strTextLine
    Loop
    Close #intFNumber
    
    Set TextFile_ReadToCollection = colResult
    
End Function
Public Function GetFolder() As String

Dim fldr As FileDialog
Dim sItem As String

    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Favor selecionar um novo camimho padr�o."
        .AllowMultiSelect = False
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop") ' Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
'        Debug.Print sItem
    End With
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Public Function Etiqueta(sEtiqueta As String) As String
On Error Resume Next

    Etiqueta = Replace(Replace(ThisWorkbook.Names(sEtiqueta), "=", ""), Chr(34), "")

If err.Number <> 0 Then Etiqueta = "#N/A"
On Error GoTo 0
End Function

Public Function CreateDir(strPath As String)
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Function

Private Function ShuffleArray(InArray() As String) As String()
    Dim N As Long, Temp As Variant
    Dim J As Long, Arr() As String

    Randomize

    'make a copy of the array
    ReDim Arr(LBound(InArray) To UBound(InArray))
    For N = LBound(InArray) To UBound(InArray)
        Arr(N) = InArray(N)
    Next N
    'shuffle the copy
    For N = LBound(Arr) To UBound(Arr)
        J = CLng(((UBound(Arr) - N) * Rnd) + N)
        Temp = Arr(N)
        Arr(N) = Arr(J)
        Arr(J) = Temp
    Next N
    ShuffleArray = Arr 'return the shuffled copy
End Function

Public Function UserName() As String ' %HOMEPATH%\Downloads
    UserName = CreateObject("WScript.Network").UserName
End Function

Sub tools_ExportAllCode_EXCEL() '' Extracao de codigos do projeto
'' #Links
'' https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files

    AddRefGuid

    Dim c As VBComponent
    Dim Sfx As String
    Dim sFileName As String: sFileName = "\" & Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))

    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then

            '''' EXCEL
            CreateDir Application.ActiveWorkbook.Path & sFileName & "\code\"
            c.Export fileName:=Application.ActiveWorkbook.Path & sFileName & "\code\" & c.Name & Sfx
            
        End If
    Next c
        MsgBox "Created source files in " & Application.ActiveWorkbook.Path & sFileName
End Sub

Private Sub AddRefGuid()
On Error Resume Next

    'Add VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3

    Application.VBE.VBProjects(1).References.AddFromGuid _
        "{0002E157-0000-0000-C000-000000000046}", 2, 0

End Sub

Private Sub listarGuias()
Dim ws As Worksheet

Dim sTitle As String:       sTitle = "Listar guias"
Dim sMessage As String:     sMessage = "Deseja listar as guias ?"
Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbOKCancel, sTitle)
            
If (resposta = vbOK) Then
    For Each ws In Worksheets
        Debug.Print ws.Name
    Next
End If

End Sub

Public Sub ClearCollection(ByRef container As Collection)
    Dim index As Long
    For index = 1 To container.Count
        container.Remove 1
    Next
End Sub

Public Function saida(strCaminho As String, strConteudo As String)
    Open strCaminho For Append As #1
    Print #1, strConteudo
    Close #1
End Function

Public Function execucao(pCol As Collection, strFileName As String, Optional strFilePath As String, Optional pOperacao As enumOperacao, Optional strApp As String) 'runUrl.au3
Dim c As Variant, tmp As String: tmp = ""
LockScreen True
    
    '' Path
    If ((strFilePath) = "") Then
        strFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    Else
        strFilePath = strFilePath & "\"
    End If
    
'    KillIfExists strFilePath & strFileName
    
    If (Dir(strFilePath & strFileName) <> "") Then kill strFilePath & strFileName
    
    '' Cria��o
    For Each c In pCol
        tmp = tmp + CStr(c) + vbNewLine
    Next c
    TextFile_Append strFilePath & strFileName, tmp
    
    Dim pathApp As String: pathApp = strApp & " " & strFilePath & strFileName
    Select Case pOperacao
        Case opExecutar
            Shell pathApp
        Case opNotepad
            ClipBoardThis tmp
            Shell pathApp, vbMaximizedFocus
        Case Else
    End Select
    
LockScreen False
End Function




''###########################################
''#run ######################################
''###########################################

Sub run_kill()
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
Dim colNews As New Collection: colNews.add ".au3": colNews.add ".txt": colNews.add ".bat": colNews.add ".bas": colNews.add ".md"
Dim c, f As Variant

    For Each f In GetFilesInFolder(pathExit)
        For Each c In colNews
            If (Dir(pathExit) <> "") And InStr(f, c) Then kill f
        Next
    Next f

End Sub

Sub run_postman(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\postman-portable\postman-portable.exe": Shell pathApp: End Sub

Sub run_Soap(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\SoapUIPortable\SoapUIPortable.exe": Shell pathApp: End Sub

Sub run_eclipse(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\eclipse\eclipse.exe": Shell pathApp: End Sub

Sub run_eclipse_202003(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\eclipse-java-2020-03-R-win32-x86_64\eclipse\eclipse.exe": Shell pathApp: End Sub

Sub run_Sqldeveloper(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\sqldeveloper-19.2.1.247.2212-x64\sqldeveloper\sqldeveloper.exe": Shell pathApp: End Sub

Sub run_Python(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\Python-3.8.2 x64\PyScripter-Launcher.exe": Shell pathApp: End Sub

Sub run_Jmeter(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\apache-jmeter-5.3\bin\jmeter.bat": Shell pathApp: End Sub

Sub run_FirefoxPortable(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\FirefoxPortable\FirefoxPortable.exe": Shell pathApp: End Sub

Sub run_gVimPortable(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\gVimPortable\gVimPortable.exe": Shell pathApp: End Sub

Sub run_gVim(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\gVimPortable\App\vim\vim80\vim.exe": Shell pathApp, vbMaximizedFocus: End Sub

Sub run_freemind(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\freemind\FreemindPortable.exe": Shell pathApp: End Sub

Sub run_CustomUIEditor(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\pstools\CustomVBA\Ribbon\CustomUIEditor\CustomUIEditor.exe": Shell pathApp: End Sub

Sub run_VLCPortable(): Dim pathApp As String: pathApp = ThisWorkbook.Path & "\pstools\VLCPortable\VLCPortable.exe": Shell pathApp: End Sub



''###########################################
''#tools ####################################
''###########################################



Sub tools_listarWorksheets()
Dim strTemp As String

    For Each ws In Worksheets
        strTemp = strTemp & ws.Name & vbNewLine
    Next
    
    Debug.Print strTemp
    ClipBoardThis Left(strTemp, Len(strTemp) - 1)
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "List items"

End Sub

Sub tools_runMouse()
Dim colNews As New Collection
Dim s As New clsRepositorio: s.mouse colNews
    execucao colNews, s.strFileName, strApp:=ThisWorkbook.Path & "\pstools\autoit-v3\install\AutoIt3.exe", pOperacao:=opExecutar
End Sub

Sub tools_scriptLimparBase()
Dim colNews As New Collection
Dim s As New clsRepositorio: s.createVbs_limparBase colNews
    execucao colNews, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad
End Sub

Sub tools_createFileTxt()

'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim lRow As Long: lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

'' Additional
Dim colNews As New Collection
Dim s As New clsRepositorio
Dim t As Variant, tmp As String: tmp = ""

    '' Criar arquivo auxiliar
    For Each t In ws.Range("C2:C" & lRow - 1).SpecialCells(xlCellTypeVisible)
        colNews.add Trim(t)
        tmp = tmp & t & vbNewLine
    Next
    execucao colNews, ActiveSheet.Name & ".txt", strApp:="notepad.exe", pOperacao:=opNotepad
    
    ClipBoardThis tmp

End Sub

Sub tools_createNewFile()

'' Worksheet
Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)

'' Source
Dim pathExit As String: pathExit = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\" & ActiveSheet.Name & ".xlsx"
If (Dir(pathExit) <> "") Then kill pathExit

    ws.Copy
    ActiveWorkbook.SaveAs fileName:=pathExit, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

End Sub

Sub tools_createRunUrls()
'' Carregar links filtrados ou n�o de guia atual
'' Worksheet

'Worksheets("DOC_Links").Activate

Dim ws As Worksheet: Set ws = Worksheets(ActiveSheet.Name)
Dim lRow As Long: lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row

'' Additional
Dim t As Variant
Dim colNews As New Collection
Dim s As New clsRepositorio
Dim resposta As Variant: resposta = MsgBox("Deseja abrir todos os links selecionados ? ", vbQuestion + vbOKCancel, "Abrir Links selecionados.")
    
    If resposta = vbOK Then
    
        '' Criar arquivo auxiliar
        For Each t In ws.Range("C2:C" & lRow - 1).SpecialCells(xlCellTypeVisible)
            colNews.add Trim(t)
        Next
        execucao colNews, ActiveSheet.Name & ".txt"
    
        '' Executar script
        ClearCollection colNews
        s.Url colNews
        execucao colNews, ActiveSheet.Name & ".au3", strApp:=ThisWorkbook.Path & "\pstools\autoit-v3\install\AutoIt3.exe", pOperacao:=opExecutar

    End If

End Sub

Sub tools_createRunTema()
On Error Resume Next
'Dim oFSO: Set oFSO = CreateObject("Scripting.FileSystemObject")

'' Source
Dim pathExit As String: pathExit = GetFolder() & "\_temas\" & Year(Now()) & Format(Month(Now()), "00") & "\"  'CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
Dim strFolderName As String: strFolderName = Year(Now()) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & "_"

'' Folder
Dim strExecucao As Variant
Dim strTema As String
Dim strTemaCaminho As String

strExecucao = MsgBox("Deseja criar um novo tema?", vbQuestion + vbOKCancel, "Tema")
If strExecucao = vbOK Then
    '' Create folder
    strTema = InputBox("Qual o Tema?", "Tema")
    If strTema <> "" Then CreateDir pathExit & strFolderName & strTema
    strTemaCaminho = pathExit & strFolderName & strTema
Else
    CreateDir pathExit
    strTemaCaminho = pathExit
End If

'' Open folder
Shell "explorer.exe " & strTemaCaminho, vbMaximizedFocus

ClipBoardThis strTemaCaminho

End Sub

Sub tools_baseOcultar()
Dim colNews As New Collection: colNews.add "HST"

LockScreen True
    
For Each c In colNews
    For Each ws In Worksheets
        If InStr(c, Left(ws.Name, 3)) Then
            ' [0] - Ocultar / [-1] - Mostrar
            ' ws.Visible = IIf(ws.Visible = 0, -1, 0)
            ws.Delete
        End If
    Next
Next
            
LockScreen False
    
End Sub

''###########################################
''#repositorio ##############################
''###########################################


'Sub tools_baseLimpar()
'' validar exclusao
'Dim valResposta As Variant: valResposta = MsgBox("Deseja excluir guias ocultas", vbYesNo + vbQuestion, "Excluir guias")
'Dim i As Integer: i = Sheets(ActiveSheet.Name).index
''Worksheets(Sheets(ActiveSheet.Name).index)
'
'If valResposta = vbYes Then
'    For Each ws In Worksheets
'        If ws.index = i Then
'            If ws.Visible = xlSheetHidden Then ws.Delete
'        End If
'    Next
'
'    MsgBox "Concluido", vbOKOnly + vbInformation
'End If
'
'End Sub



''###########################################
''#tests ##############################
''###########################################



'Sub teste()
'
'Dim shell_obj As Object
'Dim wshSystemEnv As Object
'Set shell_obj = VBA.CreateObject("WScript.Shell")
'
'' This one does not include the path to the Rscript'
'Debug.Print shell_obj.ExpandEnvironmentStrings("%PATH%")
'Set wshSystemEnv = shell_obj.Environment("SYSTEM")
'' This one includes the path to the Rscript'
'Debug.Print wshSystemEnv("PATH")
'
'End Sub



'Sub tools_createScript()
''' to-do: corrigir nome do arquivo criado
''' to-do: corrigir pulo de linha
'
''' Worksheet
'Dim ws As Worksheet: Set ws = Worksheets("dbScripts")
'Dim lRow As Long: lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
'
''' Source
'Dim pathSource As String: pathSource = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
'Dim pathExit As String: pathExit = pathSource
'Dim fileOld As Boolean: fileOld = True
'
''' Additional
'Dim x As Long, t As Variant, tmp As String: tmp = ""
'
'    '' Build file text
'    x = 2
'    For Each t In ws.Range("C2:C" & lRow).SpecialCells(xlCellTypeVisible)
'        If fileOld Then pathExit = pathExit & ws.Range("B" & x) & ".txt"
'        If (Dir(pathExit) <> "" And fileOld) Then kill pathExit
'        If t <> "" Then
'            tmp = tmp & t & vbNewLine
'            If tmp <> "" Then TextFile_Append pathExit, tmp
'            fileOld = False
'        End If
'        x = x + 1
'    Next
'
'    MsgBox "Concluido!", vbOKOnly + vbInformation, "create_Script"
'
'End Sub



'Private Sub teste_Randomize()
'    Dim strProdutos() As String: strProdutos = Split("40|43|41|10|12|11|70|73|71|82|83|164|165|33|133", "|")
'    Dim strBandeira() As String: strBandeira = Split("1|2|7|3|40|60", "|")
'
'    strProdutos = ShuffleArray(strProdutos)
'    strBandeira = ShuffleArray(strBandeira)
'
'    Debug.Print Join(strProdutos, ", ")
'    Debug.Print Join(strBandeira, ", ")
'
'End Sub



'Private Sub teste_ListNames() ' Conceito
'    Worksheets(ActiveSheet.Name).Range("A1").ListNames
'End Sub
