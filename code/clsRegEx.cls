VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strAppPath As String
    strFileName As String
    strFilePath As String
    strLinkRemote As String
End Type
Private This As ClassType

Property Get strAppPath() As String:                   strAppPath = This.strAppPath:          End Property
Property Let strAppPath(pValue As String):             This.strAppPath = pValue:              End Property
Property Get strFileName() As String:                   strFileName = This.strFileName:       End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:            End Property
Property Get strFilePath() As String:                   strFilePath = This.strFilePath:       End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:            End Property
Property Get strLinkRemote() As String:                   strLinkRemote = This.strLinkRemote:       End Property
Property Let strLinkRemote(pValue As String):             This.strLinkRemote = pValue:            End Property

Public Function RemoverLinhasDuplicadas(ByRef pCol As Collection)

    pCol.add "^(.*?)$\s+?^(?=.*^\1$)"
    
    If strFileName = "" Then This.strFileName = "RegEx_RemoverLinhasDuplicadas.txt"

End Function

