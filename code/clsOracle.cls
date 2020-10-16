VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strAppPath As String
    strFileName As String
    strFilePath As String
    strTableName As String
End Type
Private This As ClassType

Property Get strAppPath() As String:                   strAppPath = This.strAppPath:          End Property
Property Let strAppPath(pValue As String):             This.strAppPath = pValue:              End Property
Property Get strFileName() As String:                   strFileName = This.strFileName:       End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:            End Property
Property Get strFilePath() As String:                   strFilePath = This.strFilePath:       End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:            End Property
Property Get strTableName() As String:                   strTableName = This.strTableName:       End Property
Property Let strTableName(pValue As String):             This.strTableName = pValue:            End Property

Public Function Tips(ByRef pCol As Collection)
    pCol.add "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY HH24:MI:SS'; -- SESSION"
    pCol.add "select * from all_tables; -- Tips"
End Function

Public Function ALL_TABLES(ByRef pCol As Collection)
    pCol.add "select * from all_tables;"
    
    If strFileName = "" Then This.strFileName = "ALL_TABLES.sql"
    
End Function

Public Function CONSTRAINT_TYPE_P(ByRef pCol As Collection)
''composição de PK

    pCol.add "SELECT"
    pCol.add "  B.COLUMN_NAME"
    pCol.add "FROM"
    pCol.add "  ALL_CONSTRAINTS A,"
    pCol.add "  ALL_CONS_COLUMNS B"
    pCol.add "WHERE"
    pCol.add "  A.CONSTRAINT_NAME=B.CONSTRAINT_NAME AND"
    pCol.add "  A.OWNER=B.OWNER AND"
    pCol.add "  A.TABLE_NAME='" & UCase(strTableName) & "' AND"
    pCol.add "  A.CONSTRAINT_TYPE='P';"
    
    If strFileName = "" Then This.strFileName = "CONSTRAINT_TYPE_P.sql"
    
End Function


Public Function CONSTRAINT_TYPE_R(ByRef pCol As Collection)
''composição de PK

    pCol.add "SELECT a.table_name, a.column_name, a.constraint_name, c.owner,"
    pCol.add "-- referenced pk"
    pCol.add "c.r_owner, c_pk.table_name r_table_name, c_pk.constraint_name r_pk"
    pCol.add "FROM all_cons_columns a"
    pCol.add "JOIN all_constraints c ON a.owner = c.owner"
    pCol.add "AND a.constraint_name = c.constraint_name"
    pCol.add "JOIN all_constraints c_pk ON c.r_owner = c_pk.owner"
    pCol.add "AND c.r_constraint_name = c_pk.constraint_name"
    pCol.add "WHERE c.constraint_type = 'R'"
    pCol.add "AND a.table_name = '" & UCase(strTableName) & "';"
    
    If strFileName = "" Then This.strFileName = "CONSTRAINT_TYPE_R.sql"
    
End Function


Public Function ALL_TAB_COLUMNS(ByRef pCol As Collection)

    pCol.add "select table_name,column_name,data_type,data_length from all_tab_columns where TABLE_NAME in ('" & UCase(strTableName) & "') ORDER BY TABLE_NAME,column_name;"

    If strFileName = "" Then This.strFileName = "ALL_TAB_COLUMNS.sql"
    
End Function


''#MODELOS
''
'Sub sqlOracle_Tips()
'Dim c As New Collection
'Dim s As New clsOracle: s.Tips c: s.strFileName = "tempOracle.sql": s.strFilePath = ThisWorkbook.Path & "\"
'Dim retVal As Variant: retVal = MsgBox("Abrir script ?", vbQuestion + vbOKCancel, "Oracle tips - Scripts")
'
'    If retVal = vbOK Then execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
'
'End Sub
'
'Sub sqlOracle_CONSTRAINT_TYPE_P()
'Dim c As New Collection
'Dim s As New clsOracle
'Dim strTema As String: strTema = InputBox("Qual o nome da tabela?", "Oracle tips - Composição de PK")
'
'If strTema <> "" Then
'    s.strTableName = strTema
'    s.strFilePath = ThisWorkbook.Path & "\"
'    s.CONSTRAINT_TYPE_P c
'    execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
'End If
'
'End Sub
'
'Sub sqlOracle_CONSTRAINT_TYPE_R()
'Dim c As New Collection
'Dim s As New clsOracle
'Dim strTema As String: strTema = InputBox("Qual o nome da tabela?", "Oracle tips - Composição de PK")
'
'If strTema <> "" Then
'    s.strTableName = strTema
'    s.strFilePath = ThisWorkbook.Path & "\"
'    s.CONSTRAINT_TYPE_R c
'    execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
'End If
'
'End Sub
'
'
'Sub sqlOracle_ALL_TAB_COLUMNS()
'Dim c As New Collection
'Dim s As New clsOracle
'Dim strTema As String: strTema = InputBox("Qual o nome da tabela?", "Oracle tips - Listar campos da tabela")
'
'If strTema <> "" Then
'    s.strTableName = strTema
'    s.strFilePath = ThisWorkbook.Path & "\"
'    s.ALL_TAB_COLUMNS c
'    execucao c, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath
'End If
'
'End Sub
