VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpringBoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    strAppPath As String
    strFileName As String
    strFilePath As String
End Type
Private This As ClassType

Property Get strAppPath() As String:                   strAppPath = This.strAppPath:          End Property
Property Let strAppPath(pValue As String):             This.strAppPath = pValue:              End Property
Property Get strFileName() As String:                   strFileName = This.strFileName:       End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:            End Property
Property Get strFilePath() As String:                   strFilePath = This.strFilePath:       End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:            End Property

Public Function Tips(ByRef pCol As Collection)

    pCol.add "server.port=8081"
    
    If strFileName = "" Then This.strFileName = "springBoot_tips.txt"

End Function

